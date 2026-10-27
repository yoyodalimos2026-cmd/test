import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QTextBrowser,
)

from app.main_window import MainWindow, _show_message


class GitAndEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QApplication.instance() or QApplication([])

    def test_git_page_is_available(self) -> None:
        window = MainWindow()
        labels = [button.text() for button in window.nav_buttons]

        self.assertEqual(labels, ["Bibliothèque", "Gestion Git", "Parcours Git"])
        self.assertTrue(hasattr(window, "git_files"))
        self.assertTrue(hasattr(window, "fetch_branches_button"))
        window.close()

    def test_branch_switch_error_guides_user_to_publish_changes(self) -> None:
        window = MainWindow()
        window.git_switch_task = object()

        with patch.object(window, "_refresh_git_status"), patch(
            "app.main_window._show_message"
        ) as warning:
            window._git_branch_switch_failed(
                "Le dossier contient des modifications non publiées."
            )

        self.assertIn("Cochez les fichiers", warning.call_args.args[2])
        self.assertIn("Publier les fichiers cochés", warning.call_args.args[2])
        window.close()

    def test_new_branch_can_be_published_without_selected_files(self) -> None:
        window = MainWindow()
        window.git_remote.setText("https://example.com/owner/repository.git")
        window.git_target.addItem("Nouvelle branche", "__new_branch__")
        window.git_target.setCurrentIndex(window.git_target.count() - 1)
        window.git_branch.setText("feature/empty")
        window.git_commit_message.clear()

        with patch(
            "app.main_window._show_message",
            return_value=QMessageBox.StandardButton.Yes,
        ) as show_message, patch.object(window.pool, "start") as start_task:
            window._push_git_changes()

        task = start_task.call_args.args[0]
        self.assertTrue(task.branch_only)
        self.assertEqual(task.paths, [])
        self.assertEqual(task.target_branch, "feature/empty")
        self.assertEqual(show_message.call_args.args[1], "Confirmer la création de branche")
        self.assertIn("sans créer de commit", show_message.call_args.args[2])
        window.close()

    def test_long_message_dialog_wraps_and_keeps_buttons_accessible(self) -> None:
        window = MainWindow()
        message = "\n".join(["Une erreur détaillée à lire."] * 100)
        inspected: dict[str, object] = {}

        def inspect_dialog(dialog: QDialog) -> int:
            dialog.show()
            self.application.processEvents()
            label = next(
                label for label in dialog.findChildren(QLabel) if label.text() == message
            )
            scroll_area = dialog.findChild(QScrollArea)
            button_box = dialog.findChild(QDialogButtonBox)
            inspected["label"] = label
            inspected["scroll_area"] = scroll_area
            inspected["button_box"] = button_box
            button_box.button(QDialogButtonBox.StandardButton.Ok).click()
            return int(QDialog.DialogCode.Accepted)

        with patch.object(QDialog, "exec", inspect_dialog):
            answer = _show_message(
                window,
                "Erreur détaillée",
                message,
                QMessageBox.Icon.Critical,
            )

        label = inspected["label"]
        scroll_area = inspected["scroll_area"]
        button_box = inspected["button_box"]
        self.assertEqual(answer, QMessageBox.StandardButton.Ok)
        self.assertEqual(label.text(), message)
        self.assertTrue(label.wordWrap())
        self.assertGreater(scroll_area.verticalScrollBar().maximum(), 0)
        self.assertGreaterEqual(
            button_box.button(QDialogButtonBox.StandardButton.Ok).height(),
            36,
        )
        window.close()

    def test_message_dialog_returns_selected_confirmation_button(self) -> None:
        window = MainWindow()

        for selected_button in (
            QMessageBox.StandardButton.Yes,
            QMessageBox.StandardButton.Cancel,
        ):
            def select_button(dialog: QDialog) -> int:
                button_box = dialog.findChild(QDialogButtonBox)
                standard_button = QDialogButtonBox.StandardButton(selected_button.value)
                self.assertIsNotNone(button_box.button(standard_button))
                button_box.button(standard_button).click()
                return int(QDialog.DialogCode.Accepted)

            with patch.object(QDialog, "exec", select_button):
                answer = _show_message(
                    window,
                    "Confirmer",
                    "Continuer ?",
                    QMessageBox.Icon.Question,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                    QMessageBox.StandardButton.Cancel,
                )
            self.assertEqual(answer, selected_button)

        def reject_dialog(dialog: QDialog) -> int:
            dialog.reject()
            return int(QDialog.DialogCode.Rejected)

        with patch.object(QDialog, "exec", reject_dialog):
            answer = _show_message(
                window,
                "Confirmer",
                "Continuer ?",
                QMessageBox.Icon.Question,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel,
            )
        self.assertEqual(answer, QMessageBox.StandardButton.Cancel)
        window.close()

    def test_branch_targets_include_local_and_remote_branches(self) -> None:
        window = MainWindow()

        window._refresh_branch_targets(
            {"main", "feature/local"},
            {"main", "feature/local", "bugfix/remote"},
        )

        targets = [
            (window.git_target.itemText(index), window.git_target.itemData(index))
            for index in range(window.git_target.count())
        ]
        self.assertIn(("Nouvelle branche", "__new_branch__"), targets)
        self.assertIn(("main", "main"), targets)
        self.assertIn(("feature/local", "feature/local"), targets)
        self.assertIn(("bugfix/remote (distante)", "bugfix/remote"), targets)
        window.git_target.setCurrentIndex(window.git_target.findData("bugfix/remote"))
        self.assertFalse(window.git_branch.isEnabled())
        window.git_target.setCurrentIndex(window.git_target.findData("__new_branch__"))
        self.assertTrue(window.git_branch.isEnabled())
        window.close()

    def test_git_roadmap_renders_markdown_headings_and_code_blocks(self) -> None:
        window = MainWindow()
        guide = window.pages.widget(2).findChild(QTextBrowser, "roadmap")

        self.assertIsInstance(guide, QTextBrowser)
        self.assertIn("git clone", guide.toPlainText())
        self.assertIn("<h1", guide.document().toHtml())
        self.assertIn("<pre", guide.document().toHtml())
        window.close()

    def test_source_editor_saves_file_and_preserves_crlf(self) -> None:
        window = MainWindow()
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "source.py"
            source_path.write_bytes(b"first_line = 1\r\n")
            window.current_item = SimpleNamespace(path=source_path)

            def save_editor(dialog: QDialog) -> QDialog.DialogCode:
                editor = dialog.findChild(QPlainTextEdit)
                editor.setPlainText("second_line = 2\n")
                dialog.findChild(QDialogButtonBox).button(
                    QDialogButtonBox.StandardButton.Save
                ).click()
                return QDialog.DialogCode.Accepted

            with patch.object(QDialog, "exec", save_editor), patch.object(
                window, "_refresh_git_status"
            ):
                window._open_source()

            self.assertEqual(source_path.read_bytes(), b"second_line = 2\r\n")
        window.close()


if __name__ == "__main__":
    unittest.main()
