from __future__ import annotations

import ast
import re
from dataclasses import dataclass
from pathlib import Path


_VBA_PROCEDURE = re.compile(
    r"^\s*(?:(Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+([\wÀ-ÿ]+)",
    re.IGNORECASE | re.MULTILINE,
)
_VBA_CALL = re.compile(r"^\s*Call\s+([\wÀ-ÿ]+)", re.IGNORECASE | re.MULTILINE)


@dataclass(frozen=True)
class CatalogItem:
    path: Path
    relative_path: str
    kind: str
    folder: str
    title: str
    description: str
    procedures: tuple[str, ...]
    entrypoints: tuple[str, ...]
    calls: tuple[str, ...]


def _read_source(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _describe_vba(procedures: tuple[str, ...], calls: tuple[str, ...], source: str) -> str:
    lower = source.casefold()
    effects: list[str] = []

    if "removeDuplicates".casefold() in lower:
        effects.append("supprime les lignes en doublon dans le classeur")
    if all(token in lower for token in (".pdf", ".csv", "-raw")):
        effects.append("convertit des documents PDF en fichiers CSV au moyen d’un outil externe")
    if "fileexists" in lower or "dir(" in lower:
        effects.append("vérifie l’existence de fichiers ou de dossiers")
    if ("wscript.shell" in lower or "shell(" in lower) and not effects:
        effects.append("lance un programme externe depuis le code VBA")
    if "name " in lower and " as " in lower:
        effects.append("renomme ou déplace des fichiers")
    if "formula" in lower and ("lien_hypertexte" in lower or "hyperlinks.add" in lower):
        effects.append("crée ou met à jour des liens hypertexte dans Excel")
    if "saveas" in lower and ("pdf" in lower or "xlopenxmlworkbook" in lower):
        effects.append("exporte des résultats en PDF ou en fichier Excel")
    if "interior.color" in lower or "interior.colorindex" in lower:
        effects.append("met en couleur les cellules pour signaler l’état du traitement")
    if re.search(r"\.save\b", lower):
        effects.append("enregistre le classeur")
    if "mkDir".casefold() in lower:
        effects.append("crée des dossiers")
    if "workbooks.open" in lower or re.search(r"\bworkbooks\.add\b", lower):
        effects.append("ouvre ou crée des classeurs")

    call_summaries = []
    for name in calls[:4]:
        key = name.casefold()
        if "conversion_xlsx" in key:
            call_summaries.append("conversion de documents FMD en Excel")
        elif "recuperation_fmd" in key or "récupération_fmd" in key:
            call_summaries.append("récupération des données FMD")
        elif "maj_etat_fmd" in key:
            call_summaries.append("mise à jour de l’état dans la base FMD")
        else:
            call_summaries.append("une étape complémentaire")

    if len(calls) >= 2:
        purpose = f"orchestre plusieurs étapes : {' ; '.join(dict.fromkeys(call_summaries))}."
    elif effects:
        purpose = effects[0] + "."
    else:
        purpose = "Automatise un traitement dans le classeur Excel."
    if len(effects) > 1:
        purpose += f" Le code détecté {effects[1]}."
    if calls and len(calls) < 2:
        purpose += f" Le traitement délègue aussi une étape : {call_summaries[0]}."

    sheets = list(dict.fromkeys(re.findall(r"(?:Sheets|Worksheets)\(\s*[\"']([^\"']+)", source, re.IGNORECASE)))
    if sheets:
        sheet_count = len(sheets)
        purpose += f" Le code agit sur {sheet_count} feuille(s) Excel."
    return f"Le module {purpose}"


def _describe_python(procedures: tuple[str, ...], source: str) -> str:
    try:
        tree = ast.parse(source)
    except SyntaxError:
        tree = None

    lower = source.casefold()
    actions: list[str] = []
    if "pdfplumber" in lower and ("extract_table" in lower or "extract_tables" in lower):
        actions.append("extrait les tableaux présents dans des fichiers PDF")
    if "pd.read_excel" in lower:
        actions.append("lit et contrôle des classeurs Excel existants")
    if "pd.excelwriter" in lower and ".to_excel(" in lower:
        actions.append("regroupe les résultats dans un nouveau classeur Excel")
    elif ".to_excel(" in lower:
        actions.append("enregistre les résultats dans des fichiers Excel")
    if "patternfill" in lower or "apply_styles" in lower:
        actions.append("met en couleur les résultats selon leur statut")

    formats = []
    if tree is not None:
        for node in ast.walk(tree):
            if isinstance(node, ast.Constant) and isinstance(node.value, str):
                if node.value in {"FMD_TI", "FMD_TE", "FMD_ONsemi"}:
                    formats.append(node.value.removeprefix("FMD_"))
    if formats:
        actions.append(f"gère les formats {', '.join(dict.fromkeys(formats))}")

    if "filedialog.askdirectory" in lower or "askopenfilenames" in lower:
        actions.append("propose une interface pour choisir des fichiers ou un dossier")
    if not actions:
        actions.append("automatise un traitement de données")
    return f"Le script {' ; '.join(actions)}."


def _describe(path: Path, kind: str, procedures: tuple[str, ...], source: str, calls: tuple[str, ...]) -> str:
    if kind == "Python":
        return _describe_python(procedures, source)
    return _describe_vba(procedures, calls, source)


def scan_catalog(root: Path) -> list[CatalogItem]:
    """Discover exported VBA modules and Python scripts under the source folders."""
    root = root.expanduser().resolve()
    ignored = {".git", ".venv", "venv", "__pycache__", "node_modules", "app", "tests"}
    items: list[CatalogItem] = []
    if not root.is_dir():
        return items

    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in {".bas", ".py"}:
            continue
        relative_path = path.relative_to(root).as_posix()
        if relative_path == "main.py" or any(
            part in ignored or part.startswith(".") for part in path.relative_to(root).parts
        ):
            continue
        source = _read_source(path)
        kind = "VBA" if path.suffix.lower() == ".bas" else "Python"
        if kind == "VBA":
            matches = list(_VBA_PROCEDURE.finditer(source))
            procedures = tuple(dict.fromkeys(match.group(2) for match in matches))
            entrypoints = tuple(
                dict.fromkeys(
                    match.group(2) for match in matches
                    if (match.group(1) or "").casefold() not in {"private", "friend"}
                )
            )
            calls = tuple(dict.fromkeys(_VBA_CALL.findall(source)))
        else:
            try:
                tree = ast.parse(source)
                procedures = tuple(
                    node.name for node in tree.body
                    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))
                )
                entrypoints = ()
            except SyntaxError:
                procedures = ()
                entrypoints = ()
            calls = ()
        items.append(
            CatalogItem(
                path=path,
                relative_path=relative_path,
                kind=kind,
                folder=path.parent.name,
                title=path.stem,
                description=_describe(path, kind, procedures, source, calls),
                procedures=procedures,
                entrypoints=entrypoints,
                calls=calls,
            )
        )
    return items


def related_items(item: CatalogItem, items: list[CatalogItem]) -> list[CatalogItem]:
    """Return same-folder modules and files that define directly called procedures."""
    called = {name.casefold() for name in item.calls}
    related: list[CatalogItem] = []
    for candidate in items:
        if candidate == item:
            continue
        defines_call = any(name.casefold() in called for name in candidate.procedures)
        same_folder = candidate.folder == item.folder
        if same_folder or defines_call:
            related.append(candidate)
    return related