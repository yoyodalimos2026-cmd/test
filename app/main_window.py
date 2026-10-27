from __future__ import annotations

import json
import os
import shlex
import subprocess
import sys
import tempfile
import tokenize
from contextlib import contextmanager
from pathlib import Path
from typing import Generator

from PySide6.QtCore import (
    QProcess,
    QRunnable,
    QThreadPool,
    Qt,
    Signal,
    QObject,
)
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QStyle,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from app.catalog import CatalogItem, related_items, scan_catalog


ROOT = Path(__file__).resolve().parents[1]
DESCRIPTIONS_FILE = ROOT / "module_descriptions.json"


def _show_message(
    parent: QWidget,
    title: str,
    text: str,
    icon: QMessageBox.Icon,
    buttons: QMessageBox.StandardButton = QMessageBox.StandardButton.Ok,
    default_button: QMessageBox.StandardButton = QMessageBox.StandardButton.NoButton,
) -> QMessageBox.StandardButton:
    screen = parent.screen() or QApplication.primaryScreen()
    available = screen.availableGeometry() if screen else None
    screen_width = available.width() if available else 1024
    screen_height = available.height() if available else 768
    dialog_width = min(680, max(280, int(screen_width * 0.88)))
    dialog_height = min(620, max(220, int(screen_height * 0.82)))
    icon_width = 48
    spacing = 16
    margins = 40
    text_width = max(160, dialog_width - margins - icon_width - spacing - 4)

    dialog = QDialog(parent)
    dialog.setWindowTitle(title)
    dialog.setModal(True)
    dialog.setFixedWidth(dialog_width)

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(20, 18, 20, 16)
    layout.setSpacing(16)

    content = QHBoxLayout()
    content.setSpacing(spacing)
    icon_label = QLabel()
    icon_pixmap = {
        QMessageBox.Icon.Information: QStyle.StandardPixmap.SP_MessageBoxInformation,
        QMessageBox.Icon.Warning: QStyle.StandardPixmap.SP_MessageBoxWarning,
        QMessageBox.Icon.Critical: QStyle.StandardPixmap.SP_MessageBoxCritical,
        QMessageBox.Icon.Question: QStyle.StandardPixmap.SP_MessageBoxQuestion,
    }[icon]
    icon_label.setPixmap(QApplication.style().standardIcon(icon_pixmap).pixmap(icon_width, icon_width))
    icon_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
    icon_label.setFixedWidth(icon_width)
    content.addWidget(icon_label)

    message_label = QLabel(text)
    message_label.setTextFormat(Qt.TextFormat.PlainText)
    message_label.setWordWrap(True)
    message_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
    message_label.setFixedWidth(text_width)
    message_label.adjustSize()

    scroll_area = QScrollArea()
    scroll_area.setWidget(message_label)
    scroll_area.setWidgetResizable(False)
    scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)
    scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    scroll_area.setFixedWidth(text_width + 4)
    scroll_area.setFixedHeight(min(max(72, message_label.sizeHint().height()), dialog_height - 120))
    content.addWidget(scroll_area, 1)
    layout.addLayout(content)

    button_box = QDialogButtonBox(
        standardButtons=QDialogButtonBox.StandardButton(buttons.value),
        orientation=Qt.Orientation.Horizontal,
    )
    for button in button_box.buttons():
        button.setMinimumHeight(36)
        button.setMinimumWidth(max(88, min(button.sizeHint().width(), 170)))
    layout.addWidget(button_box)

    result = QMessageBox.StandardButton.NoButton

    def finish(button: QPushButton) -> None:
        nonlocal result
        result = QMessageBox.StandardButton(button_box.standardButton(button).value)
        dialog.accept()

    def dismiss() -> None:
        nonlocal result
        if button_box.button(QDialogButtonBox.StandardButton.Cancel):
            result = QMessageBox.StandardButton.Cancel

    button_box.clicked.connect(finish)
    dialog.rejected.connect(dismiss)
    if default_button != QMessageBox.StandardButton.NoButton:
        default = button_box.button(QDialogButtonBox.StandardButton(default_button.value))
        if default:
            default.setDefault(True)
            default.setFocus()

    dialog.setStyleSheet("""
        QDialog { background: #ffffff; color: #293831; }
        QLabel { color: #293831; }
        QScrollArea { background: #ffffff; }
        QDialogButtonBox QPushButton {
            background: #ffffff;
            border: 1px solid #d8ded5;
            border-radius: 4px;
            color: #293831;
            padding: 7px 14px;
            font-weight: 600;
        }
        QDialogButtonBox QPushButton:hover { background: #edf4ef; }
    """)
    dialog.exec()
    return result


@contextmanager
def _askpass_environment(username: str, token: str) -> Generator[dict[str, str], None, None]:
    if not username and not token:
        yield {}
        return

    with tempfile.TemporaryDirectory(prefix="git-askpass-") as directory:
        temporary_directory = Path(directory)
        credentials_path = temporary_directory / "credentials.json"
        credentials_path.write_text(
            json.dumps({"username": username, "token": token}),
            encoding="utf-8",
        )
        credentials_path.chmod(0o600)

        helper_path = temporary_directory / "askpass.py"
        helper_path.write_text(
            "import json, os, sys\n"
            "with open(os.environ['GIT_GUI_CREDENTIALS_FILE'], encoding='utf-8') as source:\n"
            "    credentials = json.load(source)\n"
            "prompt = sys.argv[1].casefold() if len(sys.argv) > 1 else ''\n"
            "key = 'token' if 'password' in prompt or 'token' in prompt else 'username'\n"
            "print(credentials[key])\n",
            encoding="utf-8",
        )
        askpass_path = temporary_directory / ("askpass.cmd" if os.name == "nt" else "askpass")
        if os.name == "nt":
            askpass_path.write_text(
                f'@"{sys.executable}" "{helper_path}" %*\r\n',
                encoding="utf-8",
            )
        else:
            askpass_path.write_text(
                f"#!/bin/sh\nexec {shlex.quote(sys.executable)} "
                f"{shlex.quote(str(helper_path))} \"$@\"\n",
                encoding="utf-8",
            )
            askpass_path.chmod(0o700)
        yield {
            "GIT_ASKPASS": str(askpass_path),
            "GIT_GUI_CREDENTIALS_FILE": str(credentials_path),
        }


class GitTaskSignals(QObject):
    finished = Signal()
    failed = Signal(str)


class GitPushTask(QRunnable):
    def __init__(
        self,
        remote: str,
        target_branch: str,
        message: str,
        paths: list[str],
        username: str = "",
        token: str = "",
        repo_path: Path = ROOT,
    ):
        super().__init__()
        self.remote = remote
        self.target_branch = target_branch
        self.message = message
        self.paths = paths
        self.username = username
        self.token = token
        self.repo_path = repo_path
        self.signals = GitTaskSignals()

    def run(self) -> None:
        try:
            if self._git("diff", "--cached", "--name-only", cwd=self.repo_path) != "":
                raise RuntimeError(
                    "Des changements sont déjà indexés dans Git. "
                    "Validez-les ou désindexez-les avant d’utiliser le bouton de publication."
                )
            if self._git("remote", "get-url", "origin", check=False, cwd=self.repo_path) == "":
                self._git("remote", "add", "origin", self.remote, cwd=self.repo_path)
            else:
                self._git("remote", "set-url", "origin", self.remote, cwd=self.repo_path)

            if self._git(
                "show-ref", "--verify", f"refs/heads/{self.target_branch}", check=False, cwd=self.repo_path
            ):
                self._git("switch", self.target_branch, cwd=self.repo_path)
            elif self._git(
                "show-ref",
                "--verify",
                f"refs/remotes/origin/{self.target_branch}",
                check=False,
                cwd=self.repo_path,
            ):
                self._git(
                    "switch",
                    "--track",
                    "-c",
                    self.target_branch,
                    f"origin/{self.target_branch}",
                    cwd=self.repo_path,
                )
            else:
                self._git("switch", "-c", self.target_branch, cwd=self.repo_path)

            self._git("add", "--", *self.paths, cwd=self.repo_path)
            if self._git("diff", "--cached", "--name-only", cwd=self.repo_path):
                try:
                    self._git("commit", "-m", self.message, cwd=self.repo_path)
                except RuntimeError as error:
                    try:
                        self._git("reset", "--", *self.paths, cwd=self.repo_path)
                    except RuntimeError as reset_error:
                        raise RuntimeError(
                            f"{error}\nLes changements n’ont pas pu être désindexés : {reset_error}"
                        ) from error
                    raise
            with _askpass_environment(self.username, self.token) as auth_environment:
                push_args = ["push", "-u", "origin", self.target_branch]
                if self.username and self.token:
                    push_args = ["-c", "credential.helper=", *push_args]
                self._git(
                    *push_args,
                    cwd=self.repo_path,
                    extra_env=auth_environment,
                )
            self.signals.finished.emit()
        except Exception as error:
            message = str(error)
            for secret in (self.username, self.token):
                if secret:
                    message = message.replace(secret, "[masqué]")
            self.signals.failed.emit(message)
        finally:
            self.username = ""
            self.token = ""

    @staticmethod
    def _git(
        *args: str,
        cwd: Path = ROOT,
        check: bool = True,
        extra_env: dict[str, str] | None = None,
    ) -> str:
        try:
            result = subprocess.run(
                ["git", *args],
                cwd=cwd,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env={**os.environ, **(extra_env or {}), "GIT_TERMINAL_PROMPT": "0"},
            )
        except OSError as error:
            raise RuntimeError(f"Impossible de lancer Git : {error}") from error
        if check and result.returncode != 0:
            message = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(message or f"Git a échoué (code {result.returncode}).")
        return result.stdout


class GitBranchSwitchTask(QRunnable):
    def __init__(self, branch: str, repo_path: Path = ROOT):
        super().__init__()
        self.branch = branch
        self.repo_path = repo_path
        self.signals = GitTaskSignals()

    def run(self) -> None:
        try:
            status = GitPushTask._git(
                "status", "--porcelain=v1", "--untracked-files=all", cwd=self.repo_path
            )
            if status:
                raise RuntimeError(
                    "Le dossier contient des modifications non publiées. "
                    "Enregistrez-les ou publiez-les avant de changer de branche."
                )

            if GitPushTask._git(
                "show-ref", "--verify", f"refs/heads/{self.branch}", check=False, cwd=self.repo_path
            ):
                GitPushTask._git("switch", self.branch, cwd=self.repo_path)
            elif GitPushTask._git(
                "show-ref",
                "--verify",
                f"refs/remotes/origin/{self.branch}",
                check=False,
                cwd=self.repo_path,
            ):
                GitPushTask._git(
                    "switch",
                    "--track",
                    "-c",
                    self.branch,
                    f"origin/{self.branch}",
                    cwd=self.repo_path,
                )
            else:
                raise RuntimeError(f"La branche « {self.branch} » n’existe plus dans ce clone.")
            self.signals.finished.emit()
        except Exception as error:
            self.signals.failed.emit(str(error))


class GitFetchTask(QRunnable):
    def __init__(
        self,
        remote: str,
        username: str = "",
        token: str = "",
        repo_path: Path = ROOT,
    ):
        super().__init__()
        self.remote = remote
        self.username = username
        self.token = token
        self.repo_path = repo_path
        self.signals = GitTaskSignals()

    def run(self) -> None:
        try:
            if GitPushTask._git("remote", "get-url", "origin", check=False, cwd=self.repo_path):
                GitPushTask._git("remote", "set-url", "origin", self.remote, cwd=self.repo_path)
            else:
                GitPushTask._git("remote", "add", "origin", self.remote, cwd=self.repo_path)

            with _askpass_environment(self.username, self.token) as auth_environment:
                fetch_args = ["fetch", "--prune", "origin"]
                if self.username and self.token:
                    fetch_args = ["-c", "credential.helper=", *fetch_args]
                GitPushTask._git(
                    *fetch_args,
                    cwd=self.repo_path,
                    extra_env=auth_environment,
                )
            self.signals.finished.emit()
        except Exception as error:
            message = str(error)
            for secret in (self.username, self.token):
                if secret:
                    message = message.replace(secret, "[masqué]")
            self.signals.failed.emit(message)
        finally:
            self.username = ""
            self.token = ""


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
        self.python_process: QProcess | None = None
        self.git_task: GitPushTask | None = None
        self.git_switch_task: GitBranchSwitchTask | None = None
        self.git_fetch_task: GitFetchTask | None = None
        self.current_git_branch = ""
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

    @staticmethod
    def _git(*args: str, check: bool = True) -> str:
        return GitPushTask._git(*args, check=check)

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
             ("Gestion Git", self._build_git_page()),
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
        titles.addWidget(_label("Explorez les modules VBA et scripts disponibles dans vos dossiers.", "muted"))
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

    def _build_git_page(self) -> QWidget:
        page, layout = self._page()
        layout.addWidget(_label("Gestion Git", "page_title"))
        layout.addWidget(
            _label(
                "Modifiez les sources depuis la bibliothèque, puis sélectionnez les fichiers à publier.",
                "muted",
            )
        )

        remote_row = QHBoxLayout()
        remote_row.addWidget(_label("DÉPÔT DISTANT", "section_label"))
        self.git_remote = QLineEdit()
        self.git_remote.setPlaceholderText("URL HTTPS ou SSH du dépôt")
        try:
            remote_url = self._git("remote", "get-url", "origin", check=False).strip()
        except RuntimeError:
            remote_url = ""
        self.git_remote.setText(remote_url)
        remote_row.addWidget(self.git_remote, 1)
        self.refresh_git_button = QPushButton("Actualiser")
        self.refresh_git_button.clicked.connect(self._refresh_git_status)
        remote_row.addWidget(self.refresh_git_button)
        self.fetch_branches_button = QPushButton("Récupérer les branches distantes")
        self.fetch_branches_button.clicked.connect(self._fetch_git_branches)
        remote_row.addWidget(self.fetch_branches_button)
        layout.addLayout(remote_row)

        credentials_row = QHBoxLayout()
        credentials_row.addWidget(_label("COMPTE HTTPS", "section_label"))
        self.git_username = QLineEdit()
        self.git_username.setPlaceholderText("Nom d’utilisateur")
        credentials_row.addWidget(self.git_username)
        self.git_token = QLineEdit()
        self.git_token.setPlaceholderText("Jeton d’accès personnel (pas le mot de passe GitHub)")
        self.git_token.setEchoMode(QLineEdit.EchoMode.Password)
        credentials_row.addWidget(self.git_token, 1)
        layout.addLayout(credentials_row)

        self.git_status = _label("État Git du dépôt courant", "status_line")
        layout.addWidget(self.git_status)
        self.git_files = QListWidget()
        self.git_files.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        layout.addWidget(self.git_files, 1)

        target_row = QHBoxLayout()
        target_row.addWidget(_label("PUBLICATION", "section_label"))
        self.git_target = QComboBox()
        self.git_target.currentIndexChanged.connect(self._update_branch_field)
        target_row.addWidget(self.git_target)
        self.git_branch = QLineEdit("feature/ma-modification")
        self.git_branch.setPlaceholderText("Nom de la nouvelle branche")
        target_row.addWidget(self.git_branch, 1)
        self.git_switch_button = QPushButton("Basculer vers cette branche")
        self.git_switch_button.clicked.connect(self._switch_git_branch)
        target_row.addWidget(self.git_switch_button)
        layout.addLayout(target_row)

        commit_row = QHBoxLayout()
        self.git_commit_message = QLineEdit()
        self.git_commit_message.setPlaceholderText("Résumé des modifications (message de commit)")
        commit_row.addWidget(self.git_commit_message, 1)
        self.git_push_button = QPushButton("Publier les fichiers cochés")
        self.git_push_button.setObjectName("primary_button")
        self.git_push_button.clicked.connect(self._push_git_changes)
        commit_row.addWidget(self.git_push_button)
        layout.addLayout(commit_row)

        notice = _label(
            "Pour HTTPS, saisissez votre nom d’utilisateur et votre jeton : ils servent uniquement à cette publication "
            "et ne sont pas conservés. GitHub exige un jeton personnel, pas le mot de passe du compte. "
            "Laissez les champs vides pour utiliser l’authentification Git déjà configurée, par exemple SSH. "
            "Le bouton crée un commit puis pousse uniquement les fichiers cochés.",
            "notice",
        )
        notice.setWordWrap(True)
        layout.addWidget(notice)
        self._refresh_git_status()
        return page

    def _update_branch_field(self, _index: int) -> None:
        target = self.git_target.currentData()
        is_new_branch = target == "__new_branch__"
        self.git_branch.setEnabled(is_new_branch)
        if hasattr(self, "git_switch_button"):
            self.git_switch_button.setEnabled(
                not is_new_branch
                and target != self.current_git_branch
                and self.git_switch_task is None
                and self.git_fetch_task is None
                and self.git_task is None
            )
        if hasattr(self, "fetch_branches_button"):
            self.fetch_branches_button.setEnabled(
                self.git_fetch_task is None
                and self.git_switch_task is None
                and self.git_task is None
            )

    def _refresh_branch_targets(self, local_branches: set[str], remote_branches: set[str]) -> None:
        selected_target = self.git_target.currentData()
        self.git_target.blockSignals(True)
        self.git_target.clear()
        self.git_target.addItem("Nouvelle branche", "__new_branch__")
        if "main" in local_branches or "main" in remote_branches:
            self.git_target.addItem("main", "main")

        for branch in sorted(local_branches - {"main"}):
            self.git_target.addItem(branch, branch)

        remote_only_branches = remote_branches - local_branches - {"main"}
        if remote_only_branches:
            self.git_target.insertSeparator(self.git_target.count())
            for branch in sorted(remote_only_branches):
                self.git_target.addItem(f"{branch} (distante)", branch)

        selected_index = self.git_target.findData(selected_target)
        self.git_target.setCurrentIndex(selected_index if selected_index >= 0 else 0)
        self.git_target.blockSignals(False)
        self._update_branch_field(self.git_target.currentIndex())

    def _refresh_git_status(self) -> None:
        try:
            entries = self._git("status", "--porcelain=v1", "-z", "--untracked-files=all").split("\0")
            branch = self._git("branch", "--show-current").strip()
            local_branches = set(self._git("branch", "--format=%(refname:short)").splitlines())
            remote_branches = {
                name.removeprefix("origin/")
                for name in self._git(
                    "branch", "-r", "--format=%(refname:short)"
                ).splitlines()
                if name.startswith("origin/") and name != "origin/HEAD"
            }
        except RuntimeError as error:
            self.git_status.setText(f"Impossible de lire l’état Git : {error}")
            _show_message(self, "État Git indisponible", str(error), QMessageBox.Icon.Critical)
            return

        self.current_git_branch = branch
        self._refresh_branch_targets(local_branches, remote_branches)
        self.git_files.clear()
        index = 0
        changed_count = 0
        while index < len(entries):
            entry = entries[index]
            index += 1
            if not entry:
                continue
            status, path = entry[:2], entry[3:]
            if "R" in status or "C" in status:
                index += 1
            item = QListWidgetItem(f"{status}   {path}")
            item.setData(Qt.ItemDataRole.UserRole, path)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.git_files.addItem(item)
            changed_count += 1

        summary = f"Branche actuelle : {branch or '(détachée)'} · {changed_count} fichier(s) modifié(s)"
        if changed_count == 0:
            summary += " · aucun changement à publier"
        self.git_status.setText(summary)

    def _fetch_git_branches(self) -> None:
        remote = self.git_remote.text().strip()
        username = self.git_username.text()
        token = self.git_token.text()
        if not remote:
            _show_message(self, "Dépôt distant requis", "Saisissez l’URL du dépôt distant.", QMessageBox.Icon.Warning)
            return
        if "://" in remote and "@" in remote.partition("://")[2].partition("/")[0]:
            _show_message(
                self,
                "Identifiants dans l’URL",
                "Ne placez pas de mot de passe ou de jeton dans l’URL. Utilisez les champs d’identification.",
                QMessageBox.Icon.Warning,
            )
            return
        if bool(username) != bool(token):
            _show_message(
                self,
                "Identifiants incomplets",
                "Renseignez à la fois le nom d’utilisateur et le jeton, ou laissez les deux champs vides.",
                QMessageBox.Icon.Warning,
            )
            return
        if username and not remote.lower().startswith("https://"):
            _show_message(
                self,
                "Authentification HTTPS requise",
                "Les identifiants saisis dans cette rubrique sont utilisables avec une URL HTTPS. "
                "Pour une URL SSH, laissez ces champs vides et utilisez la clé SSH configurée dans Git.",
                QMessageBox.Icon.Warning,
            )
            return

        self.fetch_branches_button.setEnabled(False)
        self.refresh_git_button.setEnabled(False)
        self.git_push_button.setEnabled(False)
        self.git_switch_button.setEnabled(False)
        self.git_status.setText("Récupération des branches depuis origin…")
        task = GitFetchTask(remote, username=username, token=token)
        self.git_fetch_task = task
        self.git_username.clear()
        self.git_token.clear()
        task.signals.finished.connect(self._git_fetch_finished)
        task.signals.failed.connect(self._git_fetch_failed)
        self.pool.start(task)

    def _git_fetch_finished(self) -> None:
        self.git_fetch_task = None
        self.fetch_branches_button.setEnabled(True)
        self.refresh_git_button.setEnabled(True)
        self.git_push_button.setEnabled(True)
        self._refresh_git_status()
        _show_message(
            self,
            "Branches actualisées",
            "Les branches du dépôt distant ont été récupérées. Elles sont maintenant disponibles dans la liste.",
            QMessageBox.Icon.Information,
        )

    def _git_fetch_failed(self, error: str) -> None:
        self.git_fetch_task = None
        self.fetch_branches_button.setEnabled(True)
        self.refresh_git_button.setEnabled(True)
        self.git_push_button.setEnabled(True)
        self._refresh_git_status()
        self.git_status.setText(f"Échec de la récupération des branches : {error}")
        _show_message(self, "Récupération impossible", error, QMessageBox.Icon.Critical)

    def _switch_git_branch(self) -> None:
        branch = self.git_target.currentData()
        if not branch or branch == "__new_branch__":
            _show_message(
                self, "Branche à sélectionner", "Choisissez une branche existante dans la liste.",
                QMessageBox.Icon.Warning,
            )
            return
        if branch == self.current_git_branch:
            return

        answer = _show_message(
            self,
            "Changer de branche",
            f"Basculer de « {self.current_git_branch or '(détachée)'} » vers « {branch} » ?\n\n"
            "Les changements non publiés empêchent cette opération.",
            QMessageBox.Icon.Question,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        self.git_switch_button.setEnabled(False)
        self.fetch_branches_button.setEnabled(False)
        self.git_push_button.setEnabled(False)
        self.refresh_git_button.setEnabled(False)
        self.git_status.setText(f"Basculement vers « {branch} »…")
        task = GitBranchSwitchTask(branch)
        self.git_switch_task = task
        task.signals.finished.connect(self._git_branch_switched)
        task.signals.failed.connect(self._git_branch_switch_failed)
        self.pool.start(task)

    def _git_branch_switched(self) -> None:
        self.git_switch_task = None
        self.refresh_catalog()
        self.refresh_git_button.setEnabled(True)
        self.git_push_button.setEnabled(True)
        self.fetch_branches_button.setEnabled(True)
        self._refresh_git_status()
        _show_message(
            self, "Branche modifiée", f"Vous êtes maintenant sur « {self.current_git_branch} ».",
            QMessageBox.Icon.Information,
        )

    def _git_branch_switch_failed(self, error: str) -> None:
        self.git_switch_task = None
        self.refresh_git_button.setEnabled(True)
        self.git_push_button.setEnabled(True)
        self.fetch_branches_button.setEnabled(True)
        self._refresh_git_status()
        self.git_status.setText(f"Échec du changement de branche : {error}")
        if "modifications non publiées" in error:
            message = (
                f"{error}\n\n"
                "Pour continuer sans utiliser de commande Git :\n"
                "1. Consultez la liste des fichiers modifiés ci-dessus.\n"
                "2. Cochez les fichiers que vous souhaitez conserver dans le dépôt.\n"
                "3. Saisissez un résumé du changement dans le champ du commit.\n"
                "4. Cliquez sur « Publier les fichiers cochés » et attendez la confirmation.\n"
                "5. Sélectionnez de nouveau la branche voulue, puis cliquez sur "
                "« Basculer vers cette branche ».\n\n"
                "Ne cochez pas et ne publiez pas un fichier si vous voulez abandonner ses changements : "
                "demandez de l’aide avant de le supprimer."
            )
            self.git_files.setFocus()
        else:
            message = error
        _show_message(self, "Changement de branche impossible", message, QMessageBox.Icon.Warning)

    def _push_git_changes(self) -> None:
        remote = self.git_remote.text().strip()
        username = self.git_username.text()
        token = self.git_token.text()
        message = self.git_commit_message.text().strip()
        paths = [
            self.git_files.item(index).data(Qt.ItemDataRole.UserRole)
            for index in range(self.git_files.count())
            if self.git_files.item(index).checkState() == Qt.CheckState.Checked
        ]
        if not remote:
            _show_message(
                self, "Dépôt distant requis", "Saisissez l’URL HTTPS ou SSH du dépôt.",
                QMessageBox.Icon.Warning,
            )
            return
        if "://" in remote and "@" in remote.partition("://")[2].partition("/")[0]:
            _show_message(
                self,
                "Identifiants dans l’URL",
                "Ne placez pas de mot de passe ou de jeton dans l’URL. "
                "Utilisez Git Credential Manager ou une clé SSH.",
                QMessageBox.Icon.Warning,
            )
            return
        if bool(username) != bool(token):
            _show_message(
                self,
                "Identifiants incomplets",
                "Renseignez à la fois le nom d’utilisateur et le jeton, ou laissez les deux champs vides.",
                QMessageBox.Icon.Warning,
            )
            return
        if username and not remote.lower().startswith("https://"):
            _show_message(
                self,
                "Authentification HTTPS requise",
                "Les identifiants saisis dans cette rubrique sont utilisables avec une URL HTTPS. "
                "Pour une URL SSH, laissez les champs vides et utilisez la clé SSH configurée dans Git.",
                QMessageBox.Icon.Warning,
            )
            return
        if not paths:
            _show_message(
                self, "Aucun fichier sélectionné", "Cochez au moins un fichier à publier.",
                QMessageBox.Icon.Warning,
            )
            return
        if not message:
            _show_message(self, "Message requis", "Saisissez un résumé des modifications.", QMessageBox.Icon.Warning)
            return

        target_branch = self.git_target.currentData()
        if target_branch == "__new_branch__":
            target_branch = self.git_branch.text().strip()
            if not target_branch:
                _show_message(
                    self, "Nom de branche requis", "Saisissez le nom de la nouvelle branche.",
                    QMessageBox.Icon.Warning,
                )
                return
            try:
                self._git("check-ref-format", "--branch", target_branch)
            except RuntimeError as error:
                _show_message(self, "Nom de branche invalide", str(error), QMessageBox.Icon.Warning)
                return

        answer = _show_message(
            self,
            "Confirmer la publication",
            f"Créer un commit pour {len(paths)} fichier(s) et le pousser vers « {target_branch} » ?\n\n"
            + "\n".join(paths),
            QMessageBox.Icon.Question,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        self.git_push_button.setEnabled(False)
        self.refresh_git_button.setEnabled(False)
        self.fetch_branches_button.setEnabled(False)
        self.git_switch_button.setEnabled(False)
        self.git_status.setText("Création du commit et publication en cours…")
        task = GitPushTask(remote, target_branch, message, paths, username=username, token=token)
        self.git_task = task
        self.git_username.clear()
        self.git_token.clear()
        task.signals.finished.connect(self._git_push_finished)
        task.signals.failed.connect(self._git_push_failed)
        self.pool.start(task)

    def _git_push_finished(self) -> None:
        self.git_task = None
        self.git_push_button.setEnabled(True)
        self.refresh_git_button.setEnabled(True)
        self.fetch_branches_button.setEnabled(True)
        self.git_commit_message.clear()
        self._refresh_git_status()
        _show_message(
            self, "Publication terminée", "Les fichiers sélectionnés ont été publiés.",
            QMessageBox.Icon.Information,
        )

    def _git_push_failed(self, error: str) -> None:
        self.git_task = None
        self.git_push_button.setEnabled(True)
        self.refresh_git_button.setEnabled(True)
        self.fetch_branches_button.setEnabled(True)
        self.git_status.setText(f"Échec de la publication : {error}")
        _show_message(self, "Publication Git impossible", error, QMessageBox.Icon.Critical)

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
            #asset_list, QListWidget { background: #ffffff; border: 1px solid #e2e6de; border-radius: 5px; }
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
        if self.current_item is None:
            return
        path = self.current_item.path
        try:
            raw = path.read_bytes()
            if path.suffix.lower() == ".py":
                with tokenize.open(path) as source:
                    content = source.read()
                    encoding = source.encoding
            else:
                for encoding in ("utf-8-sig", "cp1252"):
                    try:
                        content = raw.decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise UnicodeError("Le fichier n’est ni en UTF-8 ni en Windows-1252.")
        except (OSError, SyntaxError, UnicodeError) as error:
            _show_message(
                self,
                "Ouverture impossible",
                f"Impossible de lire le fichier source :\n{error}",
                QMessageBox.Icon.Critical,
            )
            return

        newline = "\r\n" if b"\r\n" in raw else "\n"
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Modifier {path.name}")
        dialog.resize(1000, 700)
        layout = QVBoxLayout(dialog)
        layout.addWidget(_label(str(path), "muted"))
        editor = QPlainTextEdit()
        editor.setPlainText(content)
        layout.addWidget(editor, 1)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        save_button = buttons.button(QDialogButtonBox.StandardButton.Save)
        save_button.setText("Enregistrer")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annuler")
        buttons.rejected.connect(dialog.reject)

        def save_source(_checked: bool = False) -> None:
            try:
                with path.open("w", encoding=encoding, newline="") as source:
                    source.write(editor.toPlainText().replace("\n", newline))
            except (OSError, UnicodeError) as error:
                _show_message(
                    dialog, "Enregistrement impossible",
                    f"Impossible d’enregistrer le fichier :\n{error}",
                    QMessageBox.Icon.Critical,
                )
                return
            dialog.accept()

        save_button.clicked.connect(save_source)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        self.script_status.setText(f"Fichier enregistré : {path.name}")
        self._refresh_git_status()

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
            _show_message(
                self, "Échec du script Python",
                f"Le script s’est terminé avec le code {exit_code}. Consultez la fenêtre ou le terminal du script s’il en ouvre un.",
                QMessageBox.Icon.Warning,
            )
        if self.current_item and self.current_item.kind == "Python":
            self.launch_button.setEnabled(True)

    def _python_error(self, _error: QProcess.ProcessError) -> None:
        if self.python_process is not None:
            self.script_status.setText(f"Impossible de lancer le script : {self.python_process.errorString()}")

def main() -> int:
    application = QApplication(sys.argv)
    application.setApplicationName("Bibliothèque d’outils")
    window = MainWindow()
    window.show()
    return application.exec()