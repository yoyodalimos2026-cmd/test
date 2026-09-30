import tempfile
import unittest
from pathlib import Path

from app.catalog import related_items, scan_catalog


class CatalogTests(unittest.TestCase):
    def test_finds_vba_entrypoints_and_direct_dependencies(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            group = root / "Groupe"
            group.mkdir()
            (group / "Module1.bas").write_text(
                "Public Sub MiseAJour()\n    Call Nettoyer\nEnd Sub\n", encoding="utf-8"
            )
            (group / "Module2.bas").write_text(
                "Private Sub Nettoyer()\nEnd Sub\n", encoding="utf-8"
            )

            items = scan_catalog(root)

            self.assertEqual(items[0].procedures, ("MiseAJour",))
            self.assertEqual(items[0].entrypoints, ("MiseAJour",))
            self.assertEqual(items[0].calls, ("Nettoyer",))
            self.assertEqual([item.title for item in related_items(items[0], items)], ["Module2"])

    def test_private_procedure_is_not_executable_from_excel(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            module = Path(directory) / "Module1.bas"
            module.write_text(
                "Private Sub Interne()\nEnd Sub\nSub Accessible()\nEnd Sub\n", encoding="utf-8"
            )

            item = scan_catalog(Path(directory))[0]

            self.assertEqual(item.procedures, ("Interne", "Accessible"))
            self.assertEqual(item.entrypoints, ("Accessible",))

    def test_vba_description_uses_operations_without_person_names_or_dates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            module = Path(directory) / "Module1.bas"
            module.write_text(
                'Sub Analyse_FMD()\n'
                'Call Conversion_xlsx_29_01_2024\n'
                'Call Recuperation_FMD_29_01_2024\n'
                'Call MAJ_Etat_FMD_BDD\n'
                'Call Traitement_Jean_Dupont_12_03_2024\n'
                'ActiveWorkbook.Save\n'
                "'Modifié par Jean Dupont le 12/03/2024\n"
                'Sheets("Feuille_Jean_2024").Activate\n'
                'End Sub\n',
                encoding="utf-8",
            )

            item = scan_catalog(Path(directory))[0]

            self.assertIn("conversion de documents FMD en Excel", item.description)
            self.assertIn("récupération des données FMD", item.description)
            self.assertIn("mise à jour de l’état dans la base FMD", item.description)
            self.assertIn("feuille(s) Excel", item.description)
            self.assertNotIn("Jean", item.description)
            self.assertNotIn("12/03/2024", item.description)
            self.assertNotIn("Feuille_Jean_2024", item.description)

    def test_vba_description_detects_pdf_to_csv_conversion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            module = Path(directory) / "Module1.bas"
            module.write_text(
                'Sub Convertir_FMD()\n'
                'CreateObject("Wscript.Shell").Run "xpdf -raw fiche.pdf fiche.csv"\n'
                'End Sub\n',
                encoding="utf-8",
            )

            description = scan_catalog(Path(directory))[0].description

            self.assertIn("convertit des documents PDF en fichiers CSV", description)

    def test_vba_description_does_not_copy_comments(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            module = Path(directory) / "Module1.bas"
            module.write_text(
                "Sub Verifier()\n"
                "'===='================'Vérifie si le dossier FMD existe'================\n"
                "'For x = 2 To Nb_Lignes_Data\n"
                "End Sub\n",
                encoding="utf-8",
            )

            description = scan_catalog(Path(directory))[0].description

            self.assertNotIn("Vérifie si le dossier FMD existe", description)
            self.assertNotIn("For x", description)

    def test_discovers_python_functions_and_describes_each_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "outil.py"
            script.write_text(
                '"""Script de Jean Dupont, mis à jour le 12/03/2024."""\n'
                'import pdfplumber\n'
                'def run(path):\n'
                '    return pdfplumber.open(path).pages[0].extract_table()\n',
                encoding="utf-8",
            )

            item = scan_catalog(root)[0]

            self.assertEqual(item.kind, "Python")
            self.assertEqual(item.procedures, ("run",))
            self.assertIn("extrait les tableaux présents dans des fichiers PDF", item.description)
            self.assertNotIn("Jean", item.description)
            self.assertNotIn("12/03/2024", item.description)

    def test_ignores_application_and_test_sources(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "app").mkdir()
            (root / "tests").mkdir()
            (root / "app" / "internal.py").write_text("", encoding="utf-8")
            (root / "tests" / "test_internal.py").write_text("", encoding="utf-8")
            (root / "main.py").write_text("", encoding="utf-8")
            (root / "Module1.bas").write_text("Sub MacroMetier()\nEnd Sub\n", encoding="utf-8")

            items = scan_catalog(root)

            self.assertEqual([item.relative_path for item in items], ["Module1.bas"])


if __name__ == "__main__":
    unittest.main()