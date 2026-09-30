import tempfile
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook

from app.workbook import inspect_workbook, run_vba_macro


class WorkbookTests(unittest.TestCase):
    def test_preview_is_bounded_and_does_not_modify_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recette.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Données"
            for row in range(1, 8):
                sheet.append([f"{row}:{column}" for column in range(1, 7)])
            workbook.save(path)
            original = path.read_bytes()
            status_messages = []

            preview = inspect_workbook(
                path,
                max_rows=3,
                max_columns=2,
                progress_callback=status_messages.append,
            )

            self.assertEqual(preview.sheets[0].name, "Données")
            self.assertEqual(len(preview.sheets[0].rows), 3)
            self.assertEqual(len(preview.sheets[0].rows[0]), 2)
            self.assertEqual(path.read_bytes(), original)
            self.assertTrue(any("Ouverture du classeur" in message for message in status_messages))
            self.assertTrue(any("feuille 1/1" in message for message in status_messages))

    @unittest.skipIf(__import__("os").name == "nt", "Le contrôle concerne les plateformes sans Excel COM.")
    def test_vba_execution_is_refused_without_windows_excel(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "Microsoft Excel sous Windows"):
            run_vba_macro(Path("classeur.xlsx"), Path("Module1.bas"), "Macro", Path("copie.xlsm"))

    def test_macro_is_saved_to_the_user_selected_path_and_preserves_source(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            module_path = root / "Module1.bas"
            module_path.write_text("Sub MacroTest()\nEnd Sub\n", encoding="utf-8")
            source = root / "source.xlsx"
            source.write_bytes(b"source workbook bytes")
            output = root / "destination" / "resultat.xlsm"
            output.parent.mkdir()

            workbook = types.SimpleNamespace(
                VBProject=types.SimpleNamespace(
                    VBComponents=types.SimpleNamespace(
                        Import=lambda path: types.SimpleNamespace(Name="Module1")
                    )
                ),
                SaveAs=Mock(),
                Save=Mock(),
                Close=Mock(),
            )
            excel = types.SimpleNamespace(
                Workbooks=types.SimpleNamespace(Open=Mock(return_value=workbook)),
                Run=Mock(),
                Quit=Mock(),
            )
            client = types.ModuleType("win32com.client")
            client.DispatchEx = Mock(return_value=excel)
            package = types.ModuleType("win32com")
            package.client = client
            pythoncom = types.ModuleType("pythoncom")
            pythoncom.CoInitialize = Mock()
            pythoncom.CoUninitialize = Mock()
            status_messages = []

            with patch("app.workbook.os.name", "nt"), patch.dict(
                sys.modules,
                {"pythoncom": pythoncom, "win32com": package, "win32com.client": client},
            ):
                result = run_vba_macro(
                    source,
                    module_path,
                    "MacroTest",
                    output,
                    progress_callback=status_messages.append,
                )

            self.assertEqual(result, output.resolve())
            self.assertEqual(workbook.SaveAs.call_count, 2)
            working_copy = Path(workbook.SaveAs.call_args_list[0].args[0])
            self.assertEqual(working_copy.parent, source.parent)
            self.assertNotEqual(working_copy, output.resolve())
            self.assertEqual(
                workbook.SaveAs.call_args_list[1].args[0], str(output.resolve())
            )
            self.assertEqual(workbook.SaveAs.call_args_list[1].kwargs, {"FileFormat": 52})
            excel.Workbooks.Open.assert_called_once_with(str(source.resolve()), UpdateLinks=0, ReadOnly=False)
            excel.Run.assert_called_once_with(f"'{working_copy.name}'!Module1.MacroTest")
            pythoncom.CoInitialize.assert_called_once_with()
            pythoncom.CoUninitialize.assert_called_once_with()
            self.assertEqual(source.read_bytes(), b"source workbook bytes")
            self.assertTrue(any("Exécution de la procédure" in message for message in status_messages))
            self.assertTrue(any("Enregistrement du classeur complet" in message for message in status_messages))

    def test_macro_output_must_be_a_different_xlsm_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.xlsx"
            source.touch()
            module_path = root / "Module1.bas"
            module_path.touch()
            with patch("app.workbook.os.name", "nt"):
                with self.assertRaisesRegex(ValueError, "extension .xlsm"):
                    run_vba_macro(source, module_path, "MacroTest", root / "output.xlsx")
                macro_source = root / "original.xlsm"
                macro_source.touch()
                with self.assertRaisesRegex(ValueError, "différent du classeur original"):
                    run_vba_macro(macro_source, module_path, "MacroTest", macro_source)

    def test_existing_macro_output_is_never_overwritten(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.xlsx"
            source.touch()
            module_path = root / "Module1.bas"
            module_path.touch()
            output = root / "resultat.xlsm"
            output.write_text("keep existing output", encoding="utf-8")
            with patch("app.workbook.os.name", "nt"):
                with self.assertRaisesRegex(FileExistsError, "existe déjà"):
                    run_vba_macro(source, module_path, "MacroTest", output)
            self.assertEqual(output.read_text(encoding="utf-8"), "keep existing output")


if __name__ == "__main__":
    unittest.main()