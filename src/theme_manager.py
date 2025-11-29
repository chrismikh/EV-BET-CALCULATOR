#Theme manager for light and dark modes.
#apply_theme(app, dark=True/False) sets palette + stylesheet + matplotlib rcParams
#for future figures. Statistics panel rebuild triggers new styling for plots.


from __future__ import annotations

from PyQt6.QtGui import QPalette, QColor
from PyQt6.QtWidgets import QApplication

try:
    import matplotlib as _mpl  # type: ignore
except Exception:  # pragma: no cover
    _mpl = None  # type: ignore

# Catppuccin Macchiato-inspired dark theme
_DARK_BG = QColor("#1e1e2e")
_DARK_BG_ALT = QColor("#2a2a3c")
_DARK_FG = QColor("#f8f8f2")
_ACCENT = QColor("#89b4fa")
_POSITIVE = QColor("#a6e3a1")
_NEGATIVE = QColor("#f38ba8")
_GRAY = QColor("#6c7086")

# Modern light theme
_LIGHT_BG = QColor("#ffffff")
_LIGHT_BG_ALT = QColor("#f4f4f5")
_LIGHT_FG = QColor("#111827")
_ACCENT_LIGHT = QColor("#1d4ed8")
_POSITIVE_LIGHT = QColor("#16a34a")
_NEGATIVE_LIGHT = QColor("#dc2626")
_GRAY_LIGHT = QColor("#6b7280")


def _build_palette(dark: bool) -> QPalette:
    pal = QPalette()
    if dark:
        pal.setColor(QPalette.ColorRole.Window, _DARK_BG)
        pal.setColor(QPalette.ColorRole.AlternateBase, _DARK_BG_ALT)
        pal.setColor(QPalette.ColorRole.Base, QColor("#242436"))
        pal.setColor(QPalette.ColorRole.Text, _DARK_FG)
        pal.setColor(QPalette.ColorRole.WindowText, _DARK_FG)
        pal.setColor(QPalette.ColorRole.Button, _DARK_BG_ALT)
        pal.setColor(QPalette.ColorRole.ButtonText, _DARK_FG)
        pal.setColor(QPalette.ColorRole.Highlight, _ACCENT)
        pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#1e1e2e"))
        pal.setColor(QPalette.ColorRole.ToolTipBase, _DARK_BG_ALT)
        pal.setColor(QPalette.ColorRole.ToolTipText, _DARK_FG)
        pal.setColor(QPalette.ColorRole.Link, _ACCENT)
    else:
        pal.setColor(QPalette.ColorRole.Window, _LIGHT_BG)
        pal.setColor(QPalette.ColorRole.AlternateBase, _LIGHT_BG_ALT)
        pal.setColor(QPalette.ColorRole.Base, QColor("white"))
        pal.setColor(QPalette.ColorRole.Text, _LIGHT_FG)
        pal.setColor(QPalette.ColorRole.WindowText, _LIGHT_FG)
        pal.setColor(QPalette.ColorRole.Button, _LIGHT_BG_ALT)
        pal.setColor(QPalette.ColorRole.ButtonText, _LIGHT_FG)
        pal.setColor(QPalette.ColorRole.Highlight, _ACCENT_LIGHT)
        pal.setColor(QPalette.ColorRole.HighlightedText, QColor("white"))
        pal.setColor(QPalette.ColorRole.ToolTipBase, QColor("#ffffdc"))
        pal.setColor(QPalette.ColorRole.ToolTipText, _LIGHT_FG)
        pal.setColor(QPalette.ColorRole.Link, _ACCENT_LIGHT)
    return pal

_COMMON = f"""
QToolTip {{ border: 1px solid {_DARK_BG_ALT.name()}; padding: 5px; border-radius: 4px; }}
QStatusBar {{ font-size: 12px; }}
QSplitter::handle {{ background: transparent; }}
QSplitter::handle:horizontal {{ width: 4px; }}
QSplitter::handle:vertical {{ height: 4px; }}
"""

_DARK_SS = _COMMON + f"""
QWidget {{ background-color: {_DARK_BG.name()}; color: {_DARK_FG.name()}; }}
QMainWindow, QDialog {{ background-color: {_DARK_BG.name()}; }}

/* Sidebar */
#Sidebar {{ background-color: {_DARK_BG_ALT.name()}; border-right: 1px solid #36364f; }}

/* Buttons */
QPushButton {{
    background-color: #36364f; padding: 8px 14px;
    border-radius: 10px;
    border: 1px solid #44465e; /* subtle contrast border */
    font-weight: 500;
}}
QPushButton:hover {{ background-color: #41415a; border-color: #4b4d66; }}
QPushButton:pressed {{ background-color: #4b4b69; border-color: #565873; }}
QPushButton:checked {{ background-color: {_ACCENT.name()}; color: {_DARK_BG.name()}; border-color: {_ACCENT.name()}; }}
QPushButton:disabled {{
    background-color: #2b2b3c; color: {_GRAY.name()};
    border: 1px solid #35364a;
}}

/* Inputs */
QLineEdit, QComboBox {{
    background-color: #242436; border: 1px solid #36364f;
    padding: 6px 10px; border-radius: 10px;
}}
/* Ensure sidebar labels don't create square backgrounds */
#Sidebar QLabel {{ background: transparent; }}
QComboBox::drop-down {{ border: none; }}
QComboBox::down-arrow {{ image: url(v); }}

/* Table */
QTableWidget {{
    gridline-color: transparent;
    background-color: {_DARK_BG_ALT.name()};
    border-radius: 12px;
    border: 1px solid #36364f;
}}
/* Round header corners */
QHeaderView::section:horizontal:first {{ border-top-left-radius: 12px; }}
QHeaderView::section:horizontal:last {{ border-top-right-radius: 12px; }}
QTableCornerButton::section {{ border-top-left-radius: 12px; }}
QTableWidget::item {{ padding: 5px; }}
QTableWidget::item:alternate {{ background-color: #303042; }}
QHeaderView::section {{
    background-color: #36364f; font-weight: bold;
    padding: 6px; border: none;
}}
QTableCornerButton::section {{ background-color: #36364f; }}

/* GroupBox / Card */
QGroupBox {{
    font-weight: bold; border: 1px solid #36364f;
    margin-top: 1em; border-radius: 12px;
    padding: 6px;
}}
QGroupBox::title {{
    subcontrol-origin: margin; left: 10px; padding: 0 5px;
}}
#CompareCard {{
    background-color: {_DARK_BG_ALT.name()};
    border-radius: 14px;
    padding: 18px;
}}
#CompareCard QLabel {{ background-color: transparent; }}

/* Positive/Negative/Neutral Colors */
.positive {{ color: {_POSITIVE.name()}; }}
.negative {{ color: {_NEGATIVE.name()}; }}
.neutral {{ color: {_GRAY.name()}; }}
"""

_LIGHT_SS = _COMMON + f"""
QWidget {{ background-color: {_LIGHT_BG.name()}; color: {_LIGHT_FG.name()}; }}
QMainWindow, QDialog {{ background-color: {_LIGHT_BG.name()}; }}

/* Sidebar */
#Sidebar {{ background-color: {_LIGHT_BG_ALT.name()}; border-right: 1px solid #e5e7eb; }}

/* Buttons */
QPushButton {{
    background-color: #e5e7eb; padding: 8px 14px;
    border-radius: 10px; color: #374151;
    border: 1px solid #d4d7dd; /* subtle depth border */
    font-weight: 500;
}}
QPushButton:hover {{ background-color: #d1d5db; border-color: #c7cbd2; }}
QPushButton:pressed {{ background-color: #9ca3af; border-color: #8b939f; }}
QPushButton:checked {{ background-color: {_ACCENT_LIGHT.name()}; color: white; border-color: {_ACCENT_LIGHT.name()}; }}
QPushButton:disabled {{
    background-color: #f1f2f4; color: #9ca3af;
    border: 1px solid #e2e5e9;
}}

/* Inputs */
QLineEdit, QComboBox {{
    background-color: white; border: 1px solid #d1d5db;
    padding: 6px 10px; border-radius: 10px;
}}
#Sidebar QLabel {{ background: transparent; }}
QComboBox::drop-down {{ border: none; }}

/* Table */
QTableWidget {{
    gridline-color: transparent;
    background-color: white;
    border: 1px solid #e5e7eb;
    border-radius: 12px;
}}
QHeaderView::section:horizontal:first {{ border-top-left-radius: 12px; }}
QHeaderView::section:horizontal:last {{ border-top-right-radius: 12px; }}
QTableCornerButton::section {{ border-top-left-radius: 12px; }}
QTableWidget::item {{ padding: 5px; }}
QTableWidget::item:alternate {{ background-color: #f9fafb; }}
QHeaderView::section {{
    background-color: #f9fafb; font-weight: bold;
    padding: 6px; border: none; border-bottom: 1px solid #e5e7eb;
}}
QTableCornerButton::section {{ background-color: #f9fafb; }}

/* GroupBox / Card */
QGroupBox {{
    font-weight: bold; border: 1px solid #e5e7eb;
    margin-top: 1em; border-radius: 12px;
    padding: 6px;
}}
QGroupBox::title {{
    subcontrol-origin: margin; left: 10px; padding: 0 5px;
}}
#CompareCard {{
    background-color: white;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 18px;
}}
#CompareCard QLabel {{ background-color: transparent; }}

/* Positive/Negative/Neutral Colors */
.positive {{ color: {_POSITIVE_LIGHT.name()}; }}
.negative {{ color: {_NEGATIVE_LIGHT.name()}; }}
.neutral {{ color: {_GRAY_LIGHT.name()}; }}
"""


def _apply_matplotlib(dark: bool):  # future figures only
    if _mpl is None:
        return
    if dark:
        _mpl.rcParams.update({
            "figure.facecolor": _DARK_BG.name(),
            "axes.facecolor": _DARK_BG.name(),
            "axes.edgecolor": "#cccccc",
            "axes.labelcolor": _DARK_FG.name(),
            "text.color": _DARK_FG.name(),
            "xtick.color": "#bbbbbb",
            "ytick.color": "#bbbbbb",
            "grid.color": "#444444",
            "savefig.facecolor": _DARK_BG.name(),
        })
    else:
        _mpl.rcParams.update({
            "figure.facecolor": _LIGHT_BG.name(),
            "axes.facecolor": _LIGHT_BG.name(),
            "axes.edgecolor": "#222222",
            "axes.labelcolor": _LIGHT_FG.name(),
            "text.color": _LIGHT_FG.name(),
            "xtick.color": "#333333",
            "ytick.color": "#333333",
            "grid.color": "#d0d0d0",
            "savefig.facecolor": _LIGHT_BG.name(),
        })


def apply_theme(app: QApplication, dark: bool):
    app.setPalette(_build_palette(dark))
    app.setStyleSheet(_DARK_SS if dark else _LIGHT_SS)
    _apply_matplotlib(dark)
    # Store colors for access in main app
    if dark:
        app.setProperty("positiveColor", _POSITIVE)
        app.setProperty("negativeColor", _NEGATIVE)
        app.setProperty("neutralColor", _GRAY)
    else:
        app.setProperty("positiveColor", _POSITIVE_LIGHT)
        app.setProperty("negativeColor", _NEGATIVE_LIGHT)
        app.setProperty("neutralColor", _GRAY_LIGHT)


__all__ = ["apply_theme"]
