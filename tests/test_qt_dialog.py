import unittest
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QApplication, QDialog, QFileDialog, QTableView, QTextBrowser

from app.main_window import MainWindow, PreviewTableModel, WorkbookLoadTask
from app.workbook import SheetPreview, WorkbookPreview


class ExcelFileDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QApplication.instance() or QApplication([])

    def test_workbook_picker_uses_qt_dialog_instead_of_native_gtk(self) -> None:
        window = MainWindow()
        native_dialog_values = []

        def reject_dialog(dialog: QFileDialog) -> QDialog.DialogCode:
            native_dialog_values.append(
                dialog.testOption(QFileDialog.Option.DontUseNativeDialog)
            )
            return QDialog.DialogCode.Rejected

        with patch.object(QFileDialog, "exec", reject_dialog):
            window._browse_workbook()
        window.close()

        self.assertEqual(native_dialog_values, [True])

    def test_macro_output_picker_uses_non_native_xlsm_save_dialog(self) -> None:
        window = MainWindow()
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            output = Path(directory) / "resultat.xlsm"

            def accept_dialog(dialog: QFileDialog) -> QDialog.DialogCode:
                self.assertEqual(dialog.acceptMode(), QFileDialog.AcceptMode.AcceptSave)
                self.assertTrue(dialog.testOption(QFileDialog.Option.DontUseNativeDialog))
                self.assertEqual(dialog.defaultSuffix(), "xlsm")
                dialog.selectFile(str(output))
                return QDialog.DialogCode.Accepted

            with patch.object(QFileDialog, "exec", accept_dialog):
                selected = window._choose_macro_output_path(source)

        window.close()
        self.assertEqual(selected, output.resolve())

    def test_workbook_preview_model_is_lightweight_and_read_only(self) -> None:
        model = PreviewTableModel(SheetPreview("Données", (("A", "B"), ("1", "2"))))
        cell = model.index(1, 1)

        self.assertEqual(model.rowCount(), 2)
        self.assertEqual(model.columnCount(), 2)
        self.assertEqual(model.data(cell, Qt.ItemDataRole.DisplayRole), "2")
        self.assertFalse(model.flags(cell) & Qt.ItemFlag.ItemIsEditable)

    def test_git_roadmap_renders_markdown_headings_and_code_blocks(self) -> None:
        window = MainWindow()
        guide = window.pages.widget(2).findChild(QTextBrowser, "roadmap")

        self.assertIsInstance(guide, QTextBrowser)
        self.assertIn("git clone", guide.toPlainText())
        self.assertIn("<h1", guide.document().toHtml())
        self.assertIn("<pre", guide.document().toHtml())
        window.close()

    def test_loaded_workbook_is_rendered_with_table_model(self) -> None:
        window = MainWindow()
        preview = WorkbookPreview(
            Path("classeur.xlsx"), (SheetPreview("Données", (("Nom",), ("Valeur",))),)
        )

        window._show_workbook_preview(preview)

        table = window.sheet_tabs.widget(0)
        self.assertIsInstance(table, QTableView)
        self.assertEqual(table.model().data(table.model().index(1, 0)), "Valeur")
        self.assertEqual(table.palette().color(QPalette.ColorRole.Base).name(), "#ffffff")
        self.assertEqual(table.palette().color(QPalette.ColorRole.AlternateBase).name(), "#f2f6f3")
        self.assertEqual(table.palette().color(QPalette.ColorRole.Text).name(), "#293831")
        self.assertEqual(table.palette().color(QPalette.ColorRole.HighlightedText).name(), "#203a2c")
        window.close()

    def test_workbook_selection_starts_a_worker_instead_of_reading_in_ui_slot(self) -> None:
        window = MainWindow()
        window.pool = Mock()
        selected_path = Path("classeur-lourd.xlsx")

        with patch.object(QFileDialog, "exec", return_value=QDialog.DialogCode.Accepted), patch.object(
            QFileDialog, "selectedFiles", return_value=[str(selected_path)]
        ), patch("app.main_window.inspect_workbook") as inspect:
            window._browse_workbook()

        task = window.pool.start.call_args.args[0]
        self.assertIsInstance(task, WorkbookLoadTask)
        inspect.assert_not_called()
        self.assertFalse(window.workbook_progress.isHidden())
        self.assertFalse(window.browse_workbook_button.isEnabled())
        self.assertIsNone(window.workbook_preview)
        window.workbook_task = None
        window.close()

    def test_workbook_load_error_is_visible_in_status_and_dialog(self) -> None:
        window = MainWindow()
        window.workbook_task = WorkbookLoadTask(Path("classeur.xlsx"))
        window.workbook_progress.show()

        with patch("app.main_window.QMessageBox.critical") as critical:
            window._workbook_failed("fichier invalide")

        self.assertIn("fichier invalide", window.workbook_status.text())
        self.assertEqual(critical.call_args.args[2], "fichier invalide")
        self.assertTrue(window.workbook_progress.isHidden())
        self.assertIsNone(window.workbook_task)
        window.close()

    def test_macro_error_remains_visible_in_status_and_dialog(self) -> None:
        window = MainWindow()
        window.macro_task = object()
        window.workbook_progress.show()

        with patch("app.main_window.QMessageBox.critical") as critical:
            window._macro_failed("Excel a refusé l’exécution")

        self.assertIn("Excel a refusé l’exécution", window.workbook_status.text())
        self.assertEqual(critical.call_args.args[2], "Excel a refusé l’exécution")
        self.assertTrue(window.workbook_progress.isHidden())
        self.assertIsNone(window.macro_task)
        window.close()

    def test_vba_button_explains_windows_requirement_on_linux(self) -> None:
        window = MainWindow()
        window.workbook_preview = WorkbookPreview(Path("classeur.xlsx"), ())
        module = next(item for item in window.items if item.kind == "VBA" and item.entrypoints)
        window.macro_selector.setCurrentIndex(window.macro_selector.findData(module.relative_path))
        window._update_macro_enabled()
        self.assertTrue(window.run_macro_button.isEnabled())

        with patch("app.main_window.os.name", "posix"), patch(
            "app.main_window.QMessageBox.information"
        ) as information:
            window.run_macro_button.click()

        self.assertIn("Microsoft Excel", information.call_args.args[2])
        self.assertIn("Aucun fichier n’a été modifié", information.call_args.args[2])
        window.close()


if __name__ == "__main__":
    unittest.main()