from __future__ import annotations

import json
import os
import sys
from pathlib import Path

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QProcess,
    QRunnable,
    QThreadPool,
    Qt,
    Signal,
    QObject,
)
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTableView,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from app.catalog import CatalogItem, related_items, scan_catalog
from app.workbook import SheetPreview, WorkbookPreview, inspect_workbook, run_vba_macro


ROOT = Path(__file__).resolve().parents[1]
DESCRIPTIONS_FILE = ROOT / "module_descriptions.json"


class TaskSignals(QObject):
    finished = Signal(str)
    failed = Signal(str)
    progress = Signal(str)


class WorkbookLoadSignals(QObject):
    loaded = Signal(object)
    failed = Signal(str)
    progress = Signal(str)


class WorkbookLoadTask(QRunnable):
    def __init__(self, path: Path):
        super().__init__()
        self.path = path
        self.signals = WorkbookLoadSignals()

    def run(self) -> None:
        try:
            preview = inspect_workbook(self.path, progress_callback=self.signals.progress.emit)
            self.signals.loaded.emit(preview)
        except Exception as error:
            self.signals.failed.emit(str(error))


class PreviewTableModel(QAbstractTableModel):
    def __init__(self, sheet: SheetPreview):
        super().__init__()
        self.sheet = sheet
        self.column_count = max((len(row) for row in sheet.rows), default=0)

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.sheet.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else self.column_count

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> str | None:
        if not index.isValid() or role != Qt.ItemDataRole.DisplayRole:
            return None
        row = self.sheet.rows[index.row()]
        return row[index.column()] if index.column() < len(row) else ""

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> str | None:
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return f"Colonne {section + 1}"
        return str(section + 1)


class MacroTask(QRunnable):
    def __init__(self, workbook: Path, module_path: Path, procedure: str, output_path: Path):
        super().__init__()
        self.workbook = workbook
        self.module_path = module_path
        self.procedure = procedure
        self.output_path = output_path
        self.signals = TaskSignals()

    def run(self) -> None:
        try:
            output = run_vba_macro(
                self.workbook,
                self.module_path,
                self.procedure,
                self.output_path,
                progress_callback=self.signals.progress.emit,
            )
            self.signals.finished.emit(str(output))
        except Exception as error:
            self.signals.failed.emit(str(error))


def _label(text: str, style: str = "") -> QLabel:
    widget = QLabel(text)
    if style:
        widget.setObjectName(style)
    return widget


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Bibliothèque d’outils")
        self.resize(1420, 900)
        self.setMinimumSize(1120, 700)
        self.items: list[CatalogItem] = []
        self.current_item: CatalogItem | None = None
        self.workbook_preview: WorkbookPreview | None = None
        self.workbook_task: WorkbookLoadTask | None = None
        self.python_process: QProcess | None = None
        self.macro_task: MacroTask | None = None
        self.descriptions = self._read_descriptions()
        self.pool = QThreadPool.globalInstance()
        self._build_ui()
        self.refresh_catalog()

    def _read_descriptions(self) -> dict[str, str]:
        try:
            value = json.loads(DESCRIPTIONS_FILE.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _build_ui(self) -> None:
        container = QWidget()
        root_layout = QHBoxLayout(container)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(232)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(22, 24, 18, 20)
        sidebar_layout.setSpacing(10)
        sidebar_layout.addWidget(_label("Bibliothèque", "brand"))
        sidebar_layout.addSpacing(32)

        self.nav_buttons: list[QPushButton] = []
        self.pages = QStackedWidget()
        for index, (text, page) in enumerate(
            (("Bibliothèque", self._build_library_page()),
            ("Classeur Excel", self._build_workbook_page()),
             ("Parcours Git", self._build_roadmap_page()))
        ):
            button = QPushButton(text)
            button.setObjectName("nav_button")
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False, page_index=index: self._navigate(page_index))
            sidebar_layout.addWidget(button)
            self.nav_buttons.append(button)
            self.pages.addWidget(page)
        sidebar_layout.addStretch(1)
        root_layout.addWidget(sidebar)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(0)
        topbar = QWidget()
        topbar.setObjectName("topbar")
        topbar_layout = QHBoxLayout(topbar)
        topbar_layout.setContentsMargins(34, 12, 34, 12)
        topbar_layout.addStretch(1)
        self.refresh_button = QPushButton("Actualiser le catalogue")
        self.refresh_button.clicked.connect(self.refresh_catalog)
        topbar_layout.addWidget(self.refresh_button)
        body_layout.addWidget(topbar)
        body_layout.addWidget(self.pages, 1)
        root_layout.addWidget(body, 1)

        self.setCentralWidget(container)
        self.nav_buttons[0].setChecked(True)
        self._set_style()

    def _page(self) -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(34, 28, 34, 24)
        layout.setSpacing(18)
        return page, layout

    def _build_library_page(self) -> QWidget:
        page, layout = self._page()
        header = QHBoxLayout()
        titles = QVBoxLayout()
        titles.addWidget(_label("Bibliothèque", "page_title"))
        titles.addWidget(_label("Explorez les macros et scripts disponibles dans vos dossiers.", "muted"))
        header.addLayout(titles)
        header.addStretch(1)
        self.catalog_count = _label("", "count_badge")
        header.addWidget(self.catalog_count)
        layout.addLayout(header)

        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Rechercher un outil, une procédure ou un dossier")
        self.search.textChanged.connect(self.filter_catalog)
        filters.addWidget(self.search, 1)
        self.kind_filter = QComboBox()
        self.kind_filter.addItems(["Tous les formats", "VBA", "Python"])
        self.kind_filter.currentTextChanged.connect(self.filter_catalog)
        filters.addWidget(self.kind_filter)
        layout.addLayout(filters)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        self.asset_list = QListWidget()
        self.asset_list.setObjectName("asset_list")
        self.asset_list.currentItemChanged.connect(self._select_asset)
        splitter.addWidget(self.asset_list)

        details = QWidget()
        details_layout = QVBoxLayout(details)
        details_layout.setContentsMargins(26, 22, 24, 22)
        details_layout.setSpacing(12)
        self.detail_type = _label("SÉLECTION", "eyebrow")
        self.detail_title = _label("Choisissez un outil", "detail_title")
        self.detail_path = _label("", "muted")
        details_layout.addWidget(self.detail_type)
        details_layout.addWidget(self.detail_title)
        details_layout.addWidget(self.detail_path)
        details_layout.addSpacing(5)
        details_layout.addWidget(_label("UTILITÉ", "section_label"))
        self.description = _label("", "description")
        self.description.setWordWrap(True)
        details_layout.addWidget(self.description)
        self.description_note = _label("", "team_note")
        self.description_note.setWordWrap(True)
        self.description_note.hide()
        details_layout.addWidget(self.description_note)
        self.edit_description = QPushButton("Ajouter une note métier")
        self.edit_description.clicked.connect(self._edit_description)
        details_layout.addWidget(self.edit_description, 0, Qt.AlignmentFlag.AlignLeft)
        details_layout.addSpacing(4)
        details_layout.addWidget(_label("PROCÉDURES DÉTECTÉES", "section_label"))
        self.procedure_list = QListWidget()
        self.procedure_list.setMaximumHeight(142)
        details_layout.addWidget(self.procedure_list)
        details_layout.addWidget(_label("LIENS ET OUTILS DU MÊME DOSSIER", "section_label"))
        self.related_list = QListWidget()
        self.related_list.currentItemChanged.connect(self._open_related)
        details_layout.addWidget(self.related_list, 1)
        actions = QHBoxLayout()
        self.launch_button = QPushButton("Lancer le script")
        self.launch_button.setObjectName("primary_button")
        self.launch_button.clicked.connect(self._launch_python)
        actions.addWidget(self.launch_button)
        self.reveal_button = QPushButton("Ouvrir le fichier source")
        self.reveal_button.clicked.connect(self._open_source)
        actions.addWidget(self.reveal_button)
        details_layout.addLayout(actions)
        self.script_status = _label("Aucun script lancé", "status_line")
        details_layout.addWidget(self.script_status)
        splitter.addWidget(details)
        splitter.setSizes([390, 700])
        layout.addWidget(splitter, 1)
        return page

    def _build_workbook_page(self) -> QWidget:
        page, layout = self._page()
        layout.addWidget(_label("Classeur Excel", "page_title"))
        layout.addWidget(_label("Aperçu en lecture seule. Une macro crée un classeur complet à l’emplacement choisi.", "muted"))
        path_row = QHBoxLayout()
        self.workbook_path = QLineEdit()
        self.workbook_path.setReadOnly(True)
        self.workbook_path.setPlaceholderText("Sélectionnez un classeur .xlsx ou .xlsm")
        path_row.addWidget(self.workbook_path, 1)
        self.browse_workbook_button = QPushButton("Parcourir…")
        self.browse_workbook_button.clicked.connect(self._browse_workbook)
        path_row.addWidget(self.browse_workbook_button)
        layout.addLayout(path_row)

        self.workbook_status = _label("Aucun classeur chargé", "status_line")
        layout.addWidget(self.workbook_status)
        self.workbook_progress = QProgressBar()
        self.workbook_progress.setRange(0, 0)
        self.workbook_progress.setTextVisible(False)
        self.workbook_progress.hide()
        layout.addWidget(self.workbook_progress)
        self.sheet_tabs = QTabWidget()
        self.sheet_tabs.setDocumentMode(True)
        self.sheet_tabs.setObjectName("sheet_tabs")
        layout.addWidget(self.sheet_tabs, 1)

        run_row = QHBoxLayout()
        run_row.addWidget(_label("MODULE À APPLIQUER", "section_label"))
        self.macro_selector = QComboBox()
        self.macro_selector.currentIndexChanged.connect(self._select_macro_from_workbook)
        run_row.addWidget(self.macro_selector, 1)
        self.macro_procedure = QComboBox()
        self.macro_procedure.setMinimumWidth(240)
        run_row.addWidget(self.macro_procedure)
        self.run_macro_button = QPushButton("Appliquer à une copie")
        self.run_macro_button.setObjectName("primary_button")
        self.run_macro_button.clicked.connect(self._run_macro)
        run_row.addWidget(self.run_macro_button)
        layout.addLayout(run_row)
        self.macro_notice = _label(
            "Excel VBA est pris en charge uniquement sous Windows avec Microsoft Excel. "
            "L’accès au projet VBA doit être autorisé dans le Centre de gestion de la confidentialité.",
            "notice",
        )
        self.macro_notice.setWordWrap(True)
        if os.name != "nt":
            self.macro_notice.setText(
                "Sous Linux, les macros VBA ne peuvent pas être exécutées ici. "
                "Choisissez un classeur et une procédure, puis le bouton indiquera le prérequis. "
                "L’exécution nécessite Windows et Microsoft Excel."
            )
        layout.addWidget(self.macro_notice)
        return page

    def _build_roadmap_page(self) -> QWidget:
        page, layout = self._page()
        layout.addWidget(_label("Parcours Git & transmission", "page_title"))
        layout.addWidget(_label("Les étapes pratiques pour contribuer à un dépôt Git déjà créé.", "muted"))
        guide = QTextBrowser()
        guide.setObjectName("roadmap")
        guide.setOpenLinks(False)
        guide.setOpenExternalLinks(False)
        roadmap = ROOT / "ROADMAP.md"
        markdown = roadmap.read_text(encoding="utf-8") if roadmap.exists() else "# Roadmap\n\nLe fichier ROADMAP.md sera créé."
        guide.document().setDefaultStyleSheet("""
            body { color: #293831; font-family: 'Noto Sans'; font-size: 11pt; }
            h1 { color: #245c42; font-size: 20pt; margin: 0 0 14px 0; }
            h2 { color: #327653; font-size: 14pt; margin: 20px 0 8px 0; }
            h3 { color: #40584a; font-size: 12pt; margin: 16px 0 6px 0; }
            p, li { line-height: 1.45; }
            code { color: #245c42; background-color: #edf4ef; font-family: 'DejaVu Sans Mono'; }
            pre { color: #293831; background-color: #eef3f0; padding: 10px; }
            a { color: #28734f; }
        """)
        guide.setMarkdown(markdown)
        layout.addWidget(guide, 1)
        return page

    def _set_style(self) -> None:
        QApplication.instance().setFont(QFont("Noto Sans", 10))
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f5f8f7; color: #293831; }
            #sidebar { background: #e8f0ed; color: #293831; }
            #brand { color: #28734f; font-size: 23px; font-weight: 800; }
            #brand_caption { color: #64766c; font-size: 9px; font-weight: 700; }
            #source_path { color: #64766c; font-size: 9px; }
            #nav_button { color: #3c5045; border: 0; border-radius: 5px; text-align: left; padding: 12px 14px; font-weight: 600; }
            #nav_button:hover { background: #dce9e3; }
            #nav_button:checked { background: #d3e8da; color: #216944; }
            #topbar { background: #ffffff; border-bottom: 1px solid #e4e8df; }
            #topbar_label, #section_label { color: #000000; font-size: 15px; font-weight: 750; }
            #page_title { color: #263a30; font-size: 25px; font-weight: 750; }
            #detail_title { color: #263a30; font-size: 22px; font-weight: 700; }
            #muted { color: #6f7e75; font-size: 11px; }
            #description { color: #35483d; font-size: 12px; line-height: 1.5; }
            #team_note { color: #67562d; background: #fff5dc; border-left: 3px solid #d3a64e; padding: 9px 11px; }
            #count_badge { background: #edf1dd; color: #4b5b31; border-radius: 4px; padding: 8px 12px; font-weight: 700; }
            #asset_list, QListWidget, QTabWidget::pane { background: #ffffff; border: 1px solid #e2e6de; border-radius: 5px; }
            QListWidget::item { padding: 10px 9px; border-bottom: 1px solid #eef0ea; }
            QListWidget::item:selected { background: #e5f1e9; color: #263a30; }
            QLineEdit, QComboBox { background: #ffffff; border: 1px solid #d9ded6; border-radius: 4px; padding: 9px 10px; min-height: 20px; }
            QPushButton { background: #ffffff; border: 1px solid #d8ded5; border-radius: 4px; padding: 9px 13px; font-weight: 600; }
            QPushButton:hover { background: #edf4ef; }
            QPushButton:disabled { color: #899187; background: #edf0ed; }
            #primary_button { background: #327653; color: #ffffff; border: 1px solid #327653; }
            #primary_button:hover { background: #286746; }
            #roadmap { background: #ffffff; color: #293831; border: 1px solid #e2e6de; border-radius: 5px; padding: 16px; }
            #status_line { background: #eaf2ed; color: #365344; padding: 9px 12px; border-radius: 4px; }
            #notice { color: #745d2f; background: #fff5df; border-left: 3px solid #d3a64e; padding: 10px 12px; }
            QTableView {
                background: #ffffff;
                alternate-background-color: #f2f6f3;
                color: #293831;
                selection-background-color: #d5e9dc;
                selection-color: #203a2c;
                border: 1px solid #e2e6de;
                gridline-color: #edf0ea;
            }
            QTableView::item { background: #ffffff; color: #293831; }
            QTableView::item:alternate { background: #f2f6f3; color: #293831; }
            QTableView::item:selected { background: #d5e9dc; color: #203a2c; }
            QHeaderView::section { background: #edf3ef; border: 0; padding: 7px; font-weight: 650; }
            QProgressBar { border: 0; background: #e5e9e1; height: 4px; }
            QProgressBar::chunk { background: #75ad83; }
        """)

    def _navigate(self, index: int) -> None:
        self.pages.setCurrentIndex(index)
        for button_index, button in enumerate(self.nav_buttons):
            button.setChecked(button_index == index)

    def refresh_catalog(self) -> None:
        self.items = scan_catalog(ROOT)
        self.items = [
            CatalogItem(
                item.path, item.relative_path, item.kind, item.folder, item.title,
                item.description, item.procedures,
                item.entrypoints, item.calls
            )
            for item in self.items
        ]
        vba_count = sum(item.kind == "VBA" for item in self.items)
        python_count = sum(item.kind == "Python" for item in self.items)
        self.catalog_count.setText(f"{vba_count} modules VBA   ·   {python_count} scripts Python")
        self._refresh_macro_selector()
        self.filter_catalog()

    def filter_catalog(self, *_args: object) -> None:
        query = self.search.text().strip().casefold()
        selected_kind = self.kind_filter.currentText()
        current_path = self.current_item.relative_path if self.current_item else ""
        self.asset_list.clear()
        visible: list[CatalogItem] = []
        for item in self.items:
            searchable = " ".join((
                item.title, item.folder, item.relative_path, item.description,
                self.descriptions.get(item.relative_path, ""), *item.procedures
            )).casefold()
            if query and query not in searchable:
                continue
            if selected_kind != "Tous les formats" and item.kind != selected_kind:
                continue
            visible.append(item)
            list_item = QListWidgetItem(f"{item.title}   ·   {item.kind}\n{item.folder}")
            list_item.setData(Qt.ItemDataRole.UserRole, item.relative_path)
            self.asset_list.addItem(list_item)
            if item.relative_path == current_path:
                self.asset_list.setCurrentItem(list_item)
        if self.asset_list.count() and self.asset_list.currentRow() < 0:
            self.asset_list.setCurrentRow(0)
        if not visible:
            self.current_item = None
            self.detail_title.setText("Aucun résultat")
            self.description.setText("Essayez un autre terme ou filtre.")

    def _select_asset(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is None:
            return
        relative_path = current.data(Qt.ItemDataRole.UserRole)
        item = next((candidate for candidate in self.items if candidate.relative_path == relative_path), None)
        if item is None:
            return
        self.current_item = item
        self.detail_type.setText(f"{item.kind.upper()}   /   {item.folder.upper()}")
        self.detail_title.setText(item.title)
        self.detail_path.setText(item.relative_path)
        self.description.setText(item.description)
        note = self.descriptions.get(item.relative_path, "").strip()
        has_note = bool(note) and len(set(note.casefold()) - {" ", "\n", "\t"}) > 1
        self.description_note.setText(f"Note de l’équipe : {note}" if has_note else "")
        self.description_note.setVisible(has_note)
        self.procedure_list.clear()
        for procedure in item.procedures:
            self.procedure_list.addItem(QListWidgetItem(procedure))
        self.related_list.clear()
        for related in related_items(item, self.items):
            label = f"{related.title}   ·   {related.folder}"
            if any(call.casefold() in {name.casefold() for name in related.procedures} for call in item.calls):
                label += "   ·   appelé"
            entry = QListWidgetItem(label)
            entry.setData(Qt.ItemDataRole.UserRole, related.relative_path)
            self.related_list.addItem(entry)
        self.launch_button.setText("Lancer le script" if item.kind == "Python" else "Script Python uniquement")
        self.launch_button.setEnabled(item.kind == "Python" and self.python_process is None)
        self._refresh_macro_selector()

    def _open_related(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is None:
            return
        relative_path = current.data(Qt.ItemDataRole.UserRole)
        for index in range(self.asset_list.count()):
            candidate = self.asset_list.item(index)
            if candidate.data(Qt.ItemDataRole.UserRole) == relative_path:
                self.asset_list.setCurrentItem(candidate)
                return

    def _edit_description(self) -> None:
        if self.current_item is None:
            return
        text, accepted = QInputDialog.getMultiLineText(
            self, "Note métier", "Complément validé par l’équipe sur l’utilité de cet outil :",
            self.descriptions.get(self.current_item.relative_path, "")
        )
        if not accepted or not text.strip():
            return
        self.descriptions[self.current_item.relative_path] = text.strip()
        DESCRIPTIONS_FILE.write_text(
            json.dumps(self.descriptions, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        self.refresh_catalog()

    def _open_source(self) -> None:
        if self.current_item is not None:
            from PySide6.QtGui import QDesktopServices
            from PySide6.QtCore import QUrl
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.current_item.path)))

    def _launch_python(self) -> None:
        if self.current_item is None or self.current_item.kind != "Python":
            return
        if self.python_process is not None:
            return
        process = QProcess(self)
        process.setWorkingDirectory(str(self.current_item.path.parent))
        process.setProcessChannelMode(QProcess.ProcessChannelMode.ForwardedChannels)
        process.finished.connect(self._python_finished)
        process.errorOccurred.connect(self._python_error)
        self.python_process = process
        self.launch_button.setEnabled(False)
        self.script_status.setText(f"Lancement de {self.current_item.title}…")
        process.start(sys.executable, [str(self.current_item.path)])

    def _python_finished(self, exit_code: int, _exit_status: QProcess.ExitStatus) -> None:
        self.script_status.setText(
            "Script terminé avec succès." if exit_code == 0 else f"Le script s’est terminé avec le code {exit_code}."
        )
        self.python_process = None
        if exit_code != 0:
            QMessageBox.warning(
                self, "Échec du script Python",
                f"Le script s’est terminé avec le code {exit_code}. Consultez la fenêtre ou le terminal du script s’il en ouvre un."
            )
        if self.current_item and self.current_item.kind == "Python":
            self.launch_button.setEnabled(True)

    def _python_error(self, _error: QProcess.ProcessError) -> None:
        if self.python_process is not None:
            self.script_status.setText(f"Impossible de lancer le script : {self.python_process.errorString()}")

    def _browse_workbook(self) -> None:
        if self.workbook_task is not None or self.macro_task is not None:
            return
        dialog = QFileDialog(
            self,
            "Choisir un classeur Excel",
            str(ROOT),
            "Classeurs Excel (*.xlsx *.xlsm *.xltx *.xltm)",
        )
        dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        dialog.setOption(QFileDialog.Option.DontUseNativeDialog)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected_files = dialog.selectedFiles()
        file_path = selected_files[0] if selected_files else ""
        if not file_path:
            return

        self.workbook_preview = None
        self.workbook_path.setText(file_path)
        self.workbook_path.setToolTip("")
        self.workbook_status.setText("Ouverture du classeur et préparation de l’aperçu…")
        self.workbook_status.setToolTip("")
        self.sheet_tabs.clear()
        self.workbook_progress.setRange(0, 0)
        self.workbook_progress.show()
        self.browse_workbook_button.setEnabled(False)
        self._update_macro_enabled()

        task = WorkbookLoadTask(Path(file_path))
        self.workbook_task = task
        task.signals.progress.connect(self._workbook_progress)
        task.signals.loaded.connect(self._workbook_loaded)
        task.signals.failed.connect(self._workbook_failed)
        self.pool.start(task)

    def _workbook_progress(self, message: str) -> None:
        self.workbook_status.setText(message)

    def _workbook_loaded(self, preview: WorkbookPreview) -> None:
        try:
            self._show_workbook_preview(preview)
        except Exception as error:
            self._workbook_failed(f"Impossible d’afficher l’aperçu : {error}")
            return
        self.workbook_preview = preview
        self.workbook_status.setText(
            f"Aperçu prêt · {len(preview.sheets)} feuille(s) · 30 lignes × 16 colonnes maximum · source inchangée"
        )
        self.workbook_status.setToolTip("")
        self._finish_workbook_load()

    def _workbook_failed(self, error: str) -> None:
        self.workbook_preview = None
        self.workbook_status.setText(f"Échec de lecture du classeur : {error}")
        self.workbook_path.setToolTip(error)
        self._finish_workbook_load()
        QMessageBox.critical(self, "Lecture du classeur impossible", error)

    def _finish_workbook_load(self) -> None:
        self.workbook_task = None
        self.workbook_progress.hide()
        self.browse_workbook_button.setEnabled(self.macro_task is None)
        self._update_macro_enabled()

    def _show_workbook_preview(self, preview: WorkbookPreview) -> None:
        self.sheet_tabs.clear()
        for sheet in preview.sheets:
            table = QTableView()
            palette = table.palette()
            palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#f2f6f3"))
            palette.setColor(QPalette.ColorRole.Text, QColor("#293831"))
            palette.setColor(QPalette.ColorRole.Highlight, QColor("#d5e9dc"))
            palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#203a2c"))
            table.setPalette(palette)
            table.setModel(PreviewTableModel(sheet))
            table.setAlternatingRowColors(True)
            table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
            table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
            table.verticalHeader().setDefaultSectionSize(27)
            table.resizeColumnsToContents()
            self.sheet_tabs.addTab(table, sheet.name)

    def _refresh_macro_selector(self) -> None:
        if not hasattr(self, "macro_selector"):
            return
        previous_path = self.current_item.relative_path if self.current_item and self.current_item.kind == "VBA" else None
        self.macro_selector.blockSignals(True)
        self.macro_selector.clear()
        for item in self.items:
            if item.kind == "VBA":
                self.macro_selector.addItem(f"{item.title}   ·   {item.folder}", item.relative_path)
        if previous_path:
            index = self.macro_selector.findData(previous_path)
            if index >= 0:
                self.macro_selector.setCurrentIndex(index)
        self.macro_selector.blockSignals(False)
        self._select_macro_from_workbook()

    def _selected_macro(self) -> CatalogItem | None:
        relative_path = self.macro_selector.currentData()
        return next((item for item in self.items if item.relative_path == relative_path), None)

    def _select_macro_from_workbook(self, *_args: object) -> None:
        item = self._selected_macro()
        self.macro_procedure.clear()
        if item:
            self.macro_procedure.addItems(item.entrypoints)
        self._update_macro_enabled()

    def _update_macro_enabled(self) -> None:
        allowed = (
            self.workbook_preview is not None
            and self._selected_macro() is not None
            and self.macro_procedure.count() > 0
            and self.workbook_task is None
            and self.macro_task is None
        )
        self.run_macro_button.setEnabled(allowed)
        if allowed and os.name != "nt":
            self.run_macro_button.setToolTip(
                "L’exécution VBA nécessite Windows et Microsoft Excel. Cliquez pour afficher les détails."
            )
        elif not allowed:
            self.run_macro_button.setToolTip("Choisissez un classeur Excel et une procédure VBA publique.")
        else:
            self.run_macro_button.setToolTip("")

    def _run_macro(self) -> None:
        item = self._selected_macro()
        if self.workbook_preview is None or item is None or not self.macro_procedure.currentText():
            return
        if os.name != "nt":
            message = (
                "L’exécution VBA n’est pas disponible sous Linux. Elle nécessite un poste Windows avec "
                "Microsoft Excel et pywin32. Aucun fichier n’a été modifié."
            )
            self.workbook_status.setText("Exécution indisponible : Windows et Microsoft Excel sont requis.")
            QMessageBox.information(self, "Excel VBA non disponible", message)
            return
        output_path = self._choose_macro_output_path(self.workbook_preview.path)
        if output_path is None:
            return
        answer = QMessageBox.warning(
            self,
            "Exécuter une macro VBA",
            "Les macros modifient potentiellement le classeur et les fichiers liés. "
            f"Le classeur complet sera enregistré dans « {output_path} ». "
            "N’exécutez que du code que vous avez vérifié. Continuer ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        task = MacroTask(
            self.workbook_preview.path, item.path, self.macro_procedure.currentText(), output_path
        )
        self.macro_task = task
        self.run_macro_button.setEnabled(False)
        self.workbook_status.setText(f"Création et traitement de la copie complète : {output_path}")
        self.workbook_progress.setRange(0, 0)
        self.workbook_progress.show()
        self.browse_workbook_button.setEnabled(False)
        task.signals.progress.connect(self._macro_progress)
        task.signals.finished.connect(self._macro_finished)
        task.signals.failed.connect(self._macro_failed)
        self.pool.start(task)

    def _macro_progress(self, message: str) -> None:
        self.workbook_status.setText(message)

    def _choose_macro_output_path(self, source: Path) -> Path | None:
        suggested_path = source.with_name(f"{source.stem}_avec_macro.xlsm")
        suffix = 2
        while suggested_path.exists():
            suggested_path = source.with_name(f"{source.stem}_avec_macro_{suffix}.xlsm")
            suffix += 1
        dialog = QFileDialog(
            self,
            "Enregistrer la copie complète avec la macro",
            str(suggested_path.parent),
            "Classeur Excel avec macros (*.xlsm)",
        )
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dialog.setFileMode(QFileDialog.FileMode.AnyFile)
        dialog.setDefaultSuffix("xlsm")
        dialog.setOption(QFileDialog.Option.DontUseNativeDialog)
        dialog.setOption(QFileDialog.Option.DontConfirmOverwrite)
        dialog.selectFile(suggested_path.name)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None
        selected_files = dialog.selectedFiles()
        if not selected_files:
            return None

        output_path = Path(selected_files[0]).expanduser().resolve()
        if output_path.suffix.lower() != ".xlsm":
            QMessageBox.warning(self, "Format de sortie invalide", "Choisissez un nom de fichier se terminant par .xlsm.")
            return None
        if output_path == source.expanduser().resolve():
            QMessageBox.warning(
                self, "Choisir un autre fichier", "Le fichier original ne peut pas être utilisé comme destination."
            )
            return None
        if output_path.exists():
            QMessageBox.warning(
                self,
                "Choisir un autre fichier",
                f"Le fichier existe déjà :\n{output_path}\n\nChoisissez un autre nom pour protéger le fichier existant.",
            )
            return None
        return output_path

    def _macro_finished(self, output: str) -> None:
        self.workbook_status.setText(f"Copie créée : {output}")
        self.workbook_progress.hide()
        self.macro_task = None
        self.browse_workbook_button.setEnabled(True)
        self._update_macro_enabled()

    def _macro_failed(self, error: str) -> None:
        self.workbook_status.setText(f"Échec de la macro : {error}")
        self.workbook_status.setToolTip(error)
        self.workbook_progress.hide()
        QMessageBox.critical(self, "Échec de la macro", error)
        self.macro_task = None
        self.browse_workbook_button.setEnabled(True)
        self._update_macro_enabled()


def main() -> int:
    application = QApplication(sys.argv)
    application.setApplicationName("Bibliothèque d’outils")
    window = MainWindow()
    window.show()
    return application.exec()