from __future__ import annotations

import os
from typing import Optional

from PyQt6.QtCore import Qt, QThread
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QMessageBox, QProgressBar, QFrame, QFileDialog, QTabWidget, QWidget
)

try:
    from src.utils import resource_path
except ModuleNotFoundError:
    from utils import resource_path

try:
    from src.workers import MigrationWorker
except ModuleNotFoundError:
    from workers import MigrationWorker

class PreloadDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Loading Data...")
        self.setFixedSize(420, 240)
        lay = QVBoxLayout(self)
        self.lbl_title = QLabel("Loading Betting Data")
        self.lbl_title.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.lbl_title.setStyleSheet("font-size:18px;font-weight:bold;")
        lay.addWidget(self.lbl_title)
        self.lbl_progress = QLabel("Initializing...")
        lay.addWidget(self.lbl_progress)
        self.bar = QProgressBar(); self.bar.setRange(0,0); lay.addWidget(self.bar)
        self.lbl_status = QLabel("")
        self.lbl_status.setWordWrap(True)
        lay.addWidget(self.lbl_status)
    def update_progress(self, t: str): self.lbl_progress.setText(t)
    def update_status(self, t: str): self.lbl_status.setText(t)


class SettingsDialog(QDialog):
    """Settings dialog with Appearance and Data Migration tabs."""

    def __init__(self, parent: "MainWindow"):
        super().__init__(parent)
        self.main_window = parent
        self.setWindowTitle("Settings")
        self.setFixedSize(650, 560)
        self.setModal(True)
        self._selected_file: Optional[str] = None
        self._migration_thread: Optional[QThread] = None
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # Tab widget
        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)

        # --- Tab 1: Appearance ---
        appearance_tab = QWidget()
        a_layout = QVBoxLayout(appearance_tab)
        a_layout.setContentsMargins(16, 16, 16, 16)
        a_layout.setSpacing(12)

        a_title = QLabel("Theme")
        a_title.setStyleSheet("font-size: 15px; font-weight: bold;")
        a_layout.addWidget(a_title)

        self.lbl_current_theme = QLabel(self._theme_status_text())
        self.lbl_current_theme.setStyleSheet("font-size: 13px;")
        a_layout.addWidget(self.lbl_current_theme)

        self.btn_theme = QPushButton()
        self.btn_theme.setCheckable(True)
        self.btn_theme.setChecked(self.main_window.dark_mode)
        self._sync_theme_button()
        self.btn_theme.toggled.connect(self._on_theme_toggled)
        a_layout.addWidget(self.btn_theme)

        a_layout.addStretch(1)
        self.tabs.addTab(appearance_tab, "Appearance")

        # --- Tab 2: Data Migration ---
        migration_tab = QWidget()
        m_layout = QVBoxLayout(migration_tab)
        m_layout.setContentsMargins(16, 16, 16, 16)
        m_layout.setSpacing(10)

        m_title = QLabel("Import Data from Excel (.xlsx)")
        m_title.setStyleSheet("font-size: 15px; font-weight: bold;")
        m_layout.addWidget(m_title)

        # Format instructions
        fmt_info = QLabel(
            "\U0001f4cb <b>Google Sheet / Excel Format Requirements:</b><br><br>"
            "Your file must have a sheet named <b>MATCHBET</b> with columns (in order):<br>"
            "&nbsp;&nbsp;A: Sport &nbsp; B: Tournament &nbsp; C: Matchup &nbsp; D: Bet<br>"
            "&nbsp;&nbsp;E: Live Status (LIVE or NOT LIVE) &nbsp; F: Odds<br>"
            "&nbsp;&nbsp;G: Bet Amount &nbsp; H: Result (Win/Lose) &nbsp; I: Profit<br>"
            "&nbsp;&nbsp;J: Provider (optional, defaults to BetBy)<br><br>"
            "\u26a0\ufe0f First row = headers (skipped). Empty Sport rows skipped.<br>"
            "\u26a0\ufe0f Bets without Result are imported as pending."
        )
        fmt_info.setWordWrap(True)
        fmt_info.setStyleSheet("font-size: 12px;")
        m_layout.addWidget(fmt_info)

        # Drag-and-drop area
        self.drop_frame = QFrame()
        self.drop_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.drop_frame.setMinimumHeight(80)
        self.drop_frame.setStyleSheet(
            "QFrame { border: 2px dashed #6c7086; border-radius: 10px; }"
            "QFrame:hover { border-color: #89b4fa; }"
        )
        self.drop_frame.setAcceptDrops(True)
        self.drop_frame.dragEnterEvent = self._drag_enter
        self.drop_frame.dropEvent = self._drop_event
        drop_layout = QVBoxLayout(self.drop_frame)
        drop_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_drop = QLabel(
            "Drag and drop your .xlsx file here\nor click Browse below"
        )
        self.lbl_drop.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_drop.setStyleSheet("border: none; color: #6c7086; font-size: 13px;")
        drop_layout.addWidget(self.lbl_drop)
        m_layout.addWidget(self.drop_frame)

        # Browse button
        self.btn_browse = QPushButton("Browse Files")
        self.btn_browse.clicked.connect(self._browse_file)
        m_layout.addWidget(self.btn_browse)

        # Selected file label
        self.lbl_selected_file = QLabel("No file selected")
        self.lbl_selected_file.setStyleSheet("font-size: 12px;")
        self.lbl_selected_file.setWordWrap(True)
        m_layout.addWidget(self.lbl_selected_file)

        # Start Migration button
        self.btn_migrate = QPushButton("Start Migration")
        self.btn_migrate.setEnabled(False)
        self.btn_migrate.clicked.connect(self._start_migration)
        m_layout.addWidget(self.btn_migrate)

        # Progress bar (hidden initially)
        self.migration_progress = QProgressBar()
        self.migration_progress.setRange(0, 100)
        self.migration_progress.hide()
        m_layout.addWidget(self.migration_progress)

        # Status label
        self.lbl_migration_status = QLabel("")
        self.lbl_migration_status.setWordWrap(True)
        m_layout.addWidget(self.lbl_migration_status)

        m_layout.addStretch(1)
        self.tabs.addTab(migration_tab, "Data Migration")

        # --- Tab 3: Database Reset ---
        reset_tab = QWidget()
        r_layout = QVBoxLayout(reset_tab)
        r_layout.setContentsMargins(16, 16, 16, 16)
        r_layout.setSpacing(12)

        r_title = QLabel("Reset Database")
        r_title.setStyleSheet("font-size: 15px; font-weight: bold;")
        r_layout.addWidget(r_title)

        r_desc = QLabel(
            "This will permanently delete <b>all bets</b> from the database "
            "(both settled and pending). This action cannot be undone."
        )
        r_desc.setWordWrap(True)
        r_desc.setStyleSheet("font-size: 13px;")
        r_layout.addWidget(r_desc)

        self.btn_reset_db = QPushButton("Delete All Data")
        self.btn_reset_db.setStyleSheet(
            "QPushButton { background-color: #e74c3c; color: white; font-weight: bold; }"
            "QPushButton:hover { background-color: #c0392b; }"
        )
        self.btn_reset_db.clicked.connect(self._reset_database)
        r_layout.addWidget(self.btn_reset_db)

        self.lbl_reset_status = QLabel("")
        self.lbl_reset_status.setWordWrap(True)
        r_layout.addWidget(self.lbl_reset_status)

        r_layout.addStretch(1)
        self.tabs.addTab(reset_tab, "Database")

        # --- Close button ---
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_row.addWidget(btn_close)
        root.addLayout(close_row)

    # -- helpers --
    def _theme_status_text(self) -> str:
        mode = "Dark Mode" if self.main_window.dark_mode else "Light Mode"
        return f"Current Theme: {mode}"

    def _sync_theme_button(self):
        dark = self.main_window.dark_mode
        self.btn_theme.setText(" Light Mode" if dark else " Dark Mode")
        self.btn_theme.setIcon(
            QIcon(resource_path("icons/moon.svg" if dark else "icons/sun.svg"))
        )
        self.btn_theme.setToolTip("Toggle Dark / Light theme")

    def _on_theme_toggled(self, checked: bool):
        self.main_window.on_theme_toggled(checked)
        self.lbl_current_theme.setText(self._theme_status_text())
        self._sync_theme_button()

    # -- file handling --
    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if path:
            self._set_selected_file(path)

    def _set_selected_file(self, path: str):
        self._selected_file = path
        name = os.path.basename(path)
        self.lbl_selected_file.setText(f"{name}\n{path}")
        self.btn_migrate.setEnabled(True)

    def _drag_enter(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drop_event(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".xlsx"):
                self._set_selected_file(path)
                break

    # -- migration --
    def _start_migration(self):
        if not self._selected_file or not os.path.isfile(self._selected_file):
            QMessageBox.warning(self, "File Error", "Selected file does not exist.")
            return
        self.btn_migrate.setEnabled(False)
        self.btn_browse.setEnabled(False)
        self.migration_progress.setValue(0)
        self.migration_progress.show()
        self.lbl_migration_status.setText("Migrating...")

        self._migration_thread = QThread()
        self._migration_worker = MigrationWorker(self._selected_file, self.main_window.db)
        self._migration_worker.moveToThread(self._migration_thread)
        self._migration_thread.started.connect(self._migration_worker.run)
        self._migration_worker.progress.connect(self._on_migration_progress)
        self._migration_worker.finished.connect(self._on_migration_finished)
        self._migration_thread.start()

    def _on_migration_progress(self, current: int, total: int):
        if total > 0:
            self.migration_progress.setValue(int(current / total * 100))
        self.lbl_migration_status.setText(f"Processing row {current} / {total}...")

    def _on_migration_finished(self, ok: bool, msg: str, total: int, settled: int, pending: int):
        if self._migration_thread:
            self._migration_thread.quit()
            self._migration_thread.wait()
            self._migration_worker.deleteLater()
            self._migration_thread = None
        self.migration_progress.setValue(100 if ok else 0)
        self.btn_browse.setEnabled(True)
        if ok:
            self.lbl_migration_status.setText(
                f"Successfully migrated {total} bets ({settled} settled, {pending} pending)."
            )
            self.btn_migrate.setEnabled(False)
            # Trigger a data refresh on the main window
            self.main_window.refresh_data(force=True, force_network=True)
        else:
            self.lbl_migration_status.setText(f"Migration failed: {msg}")
            self.btn_migrate.setEnabled(True)

    # -- database reset --
    def _reset_database(self):
        reply = QMessageBox.warning(
            self,
            "Confirm Database Reset",
            "Are you sure you want to delete ALL bets from the database?\n\n"
            "This action cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            db = self.main_window.db
            count = db.delete_all_bets()
            self.lbl_reset_status.setText(
                f"Deleted {count} bet(s). Database is now empty."
            )
            self.main_window.refresh_data(force=True)
        except Exception as e:
            self.lbl_reset_status.setText(f"Error: {e}")


