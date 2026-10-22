import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QPlainTextEdit,
    QTextBrowser,
)

from app.main_window import MainWindow


class GitAndEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.application = QApplication.instance() or QApplication([])

    def test_git_page_is_available(self) -> None:
        window = MainWindow()
        labels = [button.text() for button in window.nav_buttons]

        self.assertEqual(labels, ["Bibliothèque", "Gestion Git", "Parcours Git"])
        self.assertTrue(hasattr(window, "git_files"))
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
