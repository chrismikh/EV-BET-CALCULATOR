from __future__ import annotations

import sys, os, re
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QTableWidget, QTableWidgetItem, QComboBox, QPushButton, QGroupBox, QLineEdit,
    QMessageBox, QDialog, QProgressBar, QStatusBar, QHeaderView,
    QTreeWidget, QTreeWidgetItem, QTabWidget, QFrame, QFileDialog,
    QFormLayout, QScrollArea
)

# Ensure project root is in sys.path so we can import from src
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    import src.theme_manager as theme_manager
except ModuleNotFoundError:
    import theme_manager as theme_manager

try:
    from src.database import DatabaseManager
except ModuleNotFoundError:
    from database import DatabaseManager

# RowTuple now stores: (odds, live_wr, prematch_wr, live_bet_count, prematch_bet_count)
RowTuple = Tuple[float, Optional[float], Optional[float], Optional[int], Optional[int]]
MatchBetTuple = Tuple[str, str, str, str, str, float, str]


@dataclass
class SheetCacheEntry:
    rows: List[RowTuple]
    # index maps odds -> (live_wr, prem_wr, live_cnt, prem_cnt)
    index: Dict[float, Tuple[Optional[float], Optional[float], Optional[int], Optional[int]]]


def resource_path(name: str) -> str:
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))  # type: ignore[attr-defined]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)


def round_odds_key(v) -> Optional[float]:
    try:
        return round(float(v), 4)
    except Exception:
        return None


def calc_true_prob(odds: float, wr_raw: float, n: int, k: int = 50):
    """
    Compute the true probability using Bayesian shrinkage:
    p_final = (wr_raw * n + p_market * k) / (n + k)
    """
    if wr_raw is None or n is None:
        return None

    # Market implied probability
    p_market = 1.0 / odds

    # Bayesian shrinkage
    p_final = (wr_raw * n + p_market * k) / (n + k)
    return p_final


def calc_ev(odds: float, wr_raw: float, n: int):
    """
    EV = p_final * odds - 1
    """
    p = calc_true_prob(odds, wr_raw, n)
    if p is None:
        return None
    return p * odds - 1.0


def extract_team_from_bet(matchup: str, bet: str) -> Optional[str]:
    """
    Given a matchup string (e.g. 'Team A vs Team B') and a bet string
    (e.g. '2nd Map: +3.5 Team A'), determine which team the bet was placed on.
    Returns the matched team name or None if no team match is found.

    If the matchup has no 'vs' (e.g. 'Game 6', 'Main Event'), extracts the
    value after the last ': ' in the bet string (e.g. 'Winner: NRG' -> 'NRG',
    'Fastest lap: Verstappen' -> 'Verstappen').
    """
    if not bet:
        return None

    # Check if matchup contains 'vs'
    has_vs = matchup and re.search(r'\s+vs\s+', matchup, flags=re.IGNORECASE)

    if has_vs:
        # Split matchup on ' vs ' to get team names
        parts = re.split(r'\s+vs\s+', matchup, flags=re.IGNORECASE)
        if len(parts) != 2:
            return None

        team_a = parts[0].strip()
        team_b = parts[1].strip()

        if not team_a or not team_b:
            return None

        # Check which team name appears in the bet text
        a_in_bet = team_a.lower() in bet.lower()
        b_in_bet = team_b.lower() in bet.lower()

        if a_in_bet and b_in_bet:
            # Both match — prefer the longer name to avoid false positives
            return team_a if len(team_a) >= len(team_b) else team_b
        elif a_in_bet:
            return team_a
        elif b_in_bet:
            return team_b
        return None
    else:
        # No 'vs' in matchup — extract the value after the last ': ' in the bet
        if ': ' not in bet:
            return None
        after_colon = bet.rsplit(': ', 1)[1].strip()
        if not after_colon:
            return None
        # Strip leading handicap/number prefixes like '+3.5 ', '-1.5 '
        cleaned = re.sub(r'^[+-]?\d+(\.\d+)?\s+', '', after_colon)
        return cleaned if cleaned else None


def fmt_wr(wr: Optional[float]) -> str:
    return "N/A" if wr is None else f"{wr*100:.2f}%"


def fmt_ev(wr: Optional[float], odds: Optional[float], sample_size: Optional[int] = None):
    if wr is None or odds is None or sample_size is None:
        return "N/A", None

    ev = calc_ev(odds, wr, sample_size)
    if ev is None:
        return "N/A", None
    return f"{ev*100:.2f}%", ev


def fetch_matchbet_data_from_db(db: DatabaseManager, force_refresh_network: bool = False) -> List[MatchBetTuple]:
    """Fetch settled bets from the database.

    Returns data in the same MatchBetTuple format used elsewhere:
    (sport, tournament, matchup, bet, live_status, odds, result)
    """
    try:
        return db.fetch_settled_bets(force_refresh=force_refresh_network)
    except Exception as e:
        print(f"Error fetching data from database: {e}")
        return []


def normalize_tournament_name(name: str) -> str:
    """
    Normalizes tournament names by removing years, seasons, and specific suffixes
    to group variations of the same event.
    """
    if not name:
        return ""

    # Keep Formula 1 grouped under a stable root label.
    if re.match(r'^Formula\s*1\b', name, re.IGNORECASE):
        return "Formula 1"

    # Specific fix for StarLadder abbreviations
    name = name.replace("StarLadder SS", "StarLadder StarSeries")

    # --- VALORANT-specific normalizations ---
    # Fix common VCL/VLC typo
    name = re.sub(r'^VLC\b', 'VCL', name)
    # Unify VCT → Champions Tour
    name = re.sub(r'^VCT\b', 'Champions Tour', name)

    # VCT Regional Leagues (Americas, EMEA, Pacific, China)
    match = re.search(r'\bChampions Tour[\s\d:]*\b(Americas|EMEA|Pacific|China)\b', name, re.IGNORECASE)
    if match:
        region = match.group(1)
        if region.lower() == 'emea':
            return "VCT EMEA"
        return f"VCT {region.capitalize()}"

    # Game Changers (VALORANT) – detect before colon split loses the info
    if re.search(r'\bGame\s+Changers\b|:\s*GC\s', name, re.IGNORECASE):
        if re.search(r'\bChampionship\b', name, re.IGNORECASE):
            return "Game Changers Championship"
        return "Game Changers"

    # Valorant Champions (world championship)
    if re.search(r'^Valorant\s+Champions\b', name, re.IGNORECASE):
        return "Valorant Champions"

    # VCT / Champions Tour Masters events
    if (re.search(r'^Champions Tour\b', name) and re.search(r'\bMasters\b', name, re.IGNORECASE)) \
            or re.match(r'^Masters\s+\w+', name, re.IGNORECASE):
        return "VCT Masters"

    # Champions Tour Ascension events
    if re.search(r'^Champions Tour\b', name) and re.search(r'\bAscension\b', name, re.IGNORECASE):
        return "Champions Tour Ascension"

    # VCL (Valorant Challengers League) – group all regions together
    if re.match(r'^VCL\b', name, re.IGNORECASE):
        return "VCL"

    # --- CS2-specific normalizations ---
    # IEM (Intel Extreme Masters) – group all city events together
    if re.match(r'^IEM\b', name):
        return "IEM"

    # PGL events – group all together
    if re.match(r'^PGL\b', name):
        return "PGL"

    # Thunderpick / TP World Championship / TP WC – unify abbreviations
    if re.match(r'^(Thunderpick|TP)\s+(World\s+Championship|WC)\b', name, re.IGNORECASE):
        return "Thunderpick World Championship"

    # ESEA – group all variants (Advanced, divisions, etc.)
    if re.match(r'^ESEA\b', name):
        return "ESEA"

    # European Pro League – group all variants (Regular Season, Series, Division, etc.)
    if re.match(r'^European Pro League\b', name):
        return "European Pro League"

    # CS Asia Championships – preserve "Asia" before region stripping removes it
    if re.match(r'^CS\s+Asia\b', name, re.IGNORECASE):
        return "CS Asia Championships"

    # 1. Remove text after colons or double slashes (e.g., "BLAST: CQ" -> "BLAST", "Galaxy Battle // Phase 4" -> "Galaxy Battle")
    if '//' in name:
        name = name.split('//')[0]
    if ': ' in name:
        name = name.split(': ')[0]

    # 2. Remove years (1990-2029)
    # Matches 4 digits starting with 19 or 20 surrounded by word boundaries
    name = re.sub(r'\b(19|20)\d{2}\b', '', name)

    # 3. Remove common sequential patterns (Case insensitive)
    # "Season 19", "Series 5", "Vol. 2", "Part 1", "#12", "Stage 1", "Phase 2", "Split 1", "Group A"
    patterns = [
        r'\bSeason\s+\d+\b',
        r'\bSeries\s+\d+\b',
        r'\bVol\.?\s*\d+\b',
        r'\bPart\s+\d+\b',
        r'\bStage\s+\d+\b',
        r'\bPhase\s+\d+\b',
        r'\bSplit\s+\d+\b',
        r'\bGroup\s+[A-Za-z0-9]+\b',
        r'#\d+',
        r'\bOS\b',
        r'\b(Asia|Americas|Europe)\s+RMR(\s+[A-Z])?\b',
        r'\bRMR\b',
        r'\bS\d+\b',
        r'(?<!^)\b(Europe|EU|NA|SA|Asia|Americas|Oceania|CIS|European|South American|North American|North America|Pacific|APAC|EMEA|MENA|DACH)\b',
        r'\bLCQ\b',
        r'\b(Play-In|Global Finals|Contenders|CQ|Finals?|Groups?|Playoffs?)\b',
        r'\b\d+(?:st|nd|rd|th)?\s+Division\b',
        r'\bDivision\s+\d+\b',
        r'\bSeries\b',
        r'(?<!^)\b(Atlanta|Katowice|Bangkok|Raleigh|Lisbon)\b',
        r'\b(Spring|Summer|Fall|Winter)\b',
        r'\bWeek\b',
        r'\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b'
    ]
    for pattern in patterns:
        name = re.sub(pattern, '', name, flags=re.IGNORECASE)

    # 4. Remove standalone numbers (1-3 digits) that might be season/edition numbers
    # Only match numbers preceded by whitespace to preserve leading brand numbers (e.g. "500 Casino")
    name = re.sub(r'(?<=\s)\d{1,3}\b', '', name)

    # 5. Clean up extra whitespace and trailing hyphens
    # Replace multiple spaces with single space and trim ends
    name = re.sub(r'\s+', ' ', name).strip(' -')

    return name


def tournament_filter_labels(name: str) -> List[str]:
    """Return all tournament labels a row should match in the tournament filter."""
    base = normalize_tournament_name(name)
    labels: List[str] = [base] if base else []

    # Formula 1 supports both broad and race-specific filtering.
    if re.match(r'^Formula\s*1\b', name, re.IGNORECASE):
        race = re.sub(r'\b(19|20)\d{2}\b', '', name)
        race = re.sub(r'\s+', ' ', race).strip(' -')
        if ': ' in race:
            _, after = race.split(': ', 1)
            race = f"Formula 1: {after.strip()}"
        else:
            race = "Formula 1"

        if "Formula 1" not in labels:
            labels.insert(0, "Formula 1")
        if race and race not in labels:
            labels.append(race)

    return labels


def process_bets_to_cache(bets: List[MatchBetTuple]) -> Dict[str, SheetCacheEntry]:
    agg = {}
    for sport, _, _, _, live_status, odds, result in bets:
        if sport not in agg: agg[sport] = {}
        k = round_odds_key(odds)
        if k is None: continue
        if k not in agg[sport]: agg[sport][k] = {'lw':0,'lt':0,'pw':0,'pt':0}
        
        is_live = "LIVE" in live_status.upper() and "NOT" not in live_status.upper()
        is_win = result.lower() == "win"
        
        if is_live:
            agg[sport][k]['lt'] += 1
            if is_win: agg[sport][k]['lw'] += 1
        else:
            agg[sport][k]['pt'] += 1
            if is_win: agg[sport][k]['pw'] += 1
            
    cache = {}
    for sport, odds_map in agg.items():
        rows = []
        index = {}
        for k in sorted(odds_map.keys()):
            stats = odds_map[k]
            lt, lw = stats['lt'], stats['lw']
            pt, pw = stats['pt'], stats['pw']
            live_wr = (lw / lt) if lt > 0 else None
            prem_wr = (pw / pt) if pt > 0 else None
            tup = (k, live_wr, prem_wr, lt if lt > 0 else None, pt if pt > 0 else None)
            rows.append(tup)
            index[k] = (live_wr, prem_wr, lt if lt > 0 else None, pt if pt > 0 else None)
        cache[sport] = SheetCacheEntry(rows, index)
    return cache


class PreloadWorker(QObject):
    progress = pyqtSignal(str)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db = db
        self.cache: Dict[str, SheetCacheEntry] = {}
        self.matchbet_data: List[MatchBetTuple] = []

    def run(self):
        try:
            self.progress.emit("Loading data from database...")
            self.status.emit("Querying settled bets...")
            
            self.matchbet_data = fetch_matchbet_data_from_db(self.db)
            self.status.emit(f"Loaded {len(self.matchbet_data)} bets. Processing...")
            
            self.cache = process_bets_to_cache(self.matchbet_data)
            
            self.progress.emit("Finalizing...")
            self.finished.emit(True, "")
        except Exception as e:
            self.finished.emit(False, str(e))


class RefreshWorker(QObject):
    finished = pyqtSignal(bool, str, object)
    def __init__(self, db: DatabaseManager, force_refresh_network: bool = False):
        super().__init__()
        self.db = db
        self.force_refresh_network = force_refresh_network

    def run(self):
        try:
            data = fetch_matchbet_data_from_db(self.db, self.force_refresh_network)
            cache = process_bets_to_cache(data)
            self.finished.emit(True, "Refreshed all data", (cache, data))
        except Exception as e:
            self.finished.emit(False, str(e), None)


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


class SortableTableWidgetItem(QTableWidgetItem):
    def __init__(self, text, sort_value):
        super().__init__(text)
        self.sort_value = sort_value

    def __lt__(self, other):
        v1 = self.sort_value
        v2 = other.sort_value
        if v1 is None: v1 = -float('inf')
        if v2 is None: v2 = -float('inf')
        return v1 < v2


class MigrationWorker(QObject):
    """Background worker to import an .xlsx file into the SQLite database."""
    progress = pyqtSignal(int, int)  # current, total
    finished = pyqtSignal(bool, str, int, int, int)  # ok, msg, total, settled, pending

    def __init__(self, file_path: str, db: DatabaseManager):
        super().__init__()
        self.file_path = file_path
        self.db = db

    def run(self):
        try:
            import openpyxl
        except ImportError:
            self.finished.emit(False, "openpyxl is not installed. Run: pip install openpyxl", 0, 0, 0)
            return
        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
            if "MATCHBET" not in wb.sheetnames:
                self.finished.emit(False, "Sheet 'MATCHBET' not found in the workbook.", 0, 0, 0)
                wb.close()
                return
            ws = wb["MATCHBET"]
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            total = len(rows)
            db = self.db
            db.create_tables()
            imported = 0
            settled = 0
            pending = 0
            for i, row in enumerate(rows):
                self.progress.emit(i + 1, total)
                if len(row) < 6:
                    continue
                sport = str(row[0] or "").strip()
                if not sport:
                    continue
                tournament = str(row[1] or "").strip()
                matchup = str(row[2] or "").strip()
                bet = str(row[3] or "").strip()
                live_status = str(row[4] or "").strip() if len(row) > 4 else "NOT LIVE"
                odds_raw = row[5] if len(row) > 5 else None
                bet_amount_raw = row[6] if len(row) > 6 else None
                result_raw = str(row[7] or "").strip() if len(row) > 7 else ""
                profit_raw = row[8] if len(row) > 8 else None

                try:
                    odds = float(str(odds_raw).replace(",", ".")) if odds_raw is not None else None
                except (ValueError, TypeError):
                    odds = None
                if odds is None or odds <= 0:
                    continue

                try:
                    bet_amount = float(str(bet_amount_raw).replace(",", ".")) if bet_amount_raw not in (None, "") else None
                except (ValueError, TypeError):
                    bet_amount = None

                result = result_raw if result_raw else None
                profit = None
                if result:
                    try:
                        profit = float(str(profit_raw).replace(",", ".")) if profit_raw not in (None, "") else None
                    except (ValueError, TypeError):
                        profit = None
                    settled += 1
                else:
                    pending += 1

                db.insert_bet(
                    sport=sport,
                    tournament=tournament,
                    matchup=matchup,
                    bet=bet,
                    live_status=live_status if live_status else "NOT LIVE",
                    odds=odds,
                    bet_amount=bet_amount,
                    result=result,
                    profit=profit,
                )
                imported += 1
            wb.close()
            self.finished.emit(True, "", imported, settled, pending)
        except Exception as e:
            self.finished.emit(False, str(e), 0, 0, 0)


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
            "&nbsp;&nbsp;G: Bet Amount &nbsp; H: Result (Win/Lose) &nbsp; I: Profit<br><br>"
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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.db.create_tables()
        self.setWindowTitle("EV Bet Calculator")
        self.resize(1200, 800)
        self.setMinimumSize(1150, 600)
        self.data_cache: Dict[str, SheetCacheEntry] = {}
        self._previous_data_cache: Dict[str, SheetCacheEntry] = {}
        self.matchbet_data: List[MatchBetTuple] = []
        self._previous_matchbet_data: List[MatchBetTuple] = []
        self.current_rows: List[RowTuple] = []
        self.odds_index: Dict[float, Tuple[Optional[float], Optional[float], Optional[int], Optional[int]]] = {}
        self.dark_mode = True
        self.current_view = "table"  # "table", "statistics", "add_bet", "pending_bets"
        self.editing_bet_id: Optional[str] = None  # None = add mode, str = edit/settle mode
        self._build_ui()
        self._connect()
        theme_manager.apply_theme(QApplication.instance(), dark=self.dark_mode)

    # UI
    def _build_ui(self):
        # Central widget and root layout
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)

        # --- Sidebar (Controls) ---
        sidebar = QWidget(); sidebar.setObjectName("Sidebar"); sidebar.setFixedWidth(280)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(10, 10, 10, 10)
        sidebar_layout.setSpacing(10)
        root_layout.addWidget(sidebar)

        # Sport and Bet Type selectors
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(QLabel("Sport:"))
        self.sport_combo = QComboBox(); controls_layout.addWidget(self.sport_combo)
        controls_layout.addWidget(QLabel("Bet Type:"))
        self.bettype_combo = QComboBox(); self.bettype_combo.addItems(["Live", "Not Live"]); controls_layout.addWidget(self.bettype_combo)
        sidebar_layout.addLayout(controls_layout)

        # Team selector
        team_layout = QHBoxLayout()
        team_layout.addWidget(QLabel("Team:"))
        self.team_combo = QComboBox()
        self.team_combo.addItem("All Teams")
        self.team_combo.setSizePolicy(self.team_combo.sizePolicy())
        team_layout.addWidget(self.team_combo, 1)
        sidebar_layout.addLayout(team_layout)

        # Tournament selector
        tournament_layout = QHBoxLayout()
        tournament_layout.addWidget(QLabel("Tournament:"))
        self.tournament_combo = QComboBox()
        self.tournament_combo.addItem("All Tournaments")
        self.tournament_combo.setSizePolicy(self.tournament_combo.sizePolicy())
        tournament_layout.addWidget(self.tournament_combo, 1)
        sidebar_layout.addLayout(tournament_layout)

        # Comparison GroupBox
        self.compare_group = QGroupBox("Comparison")
        cg = QVBoxLayout(self.compare_group)
        odds_a_layout = QHBoxLayout(); odds_a_layout.addWidget(QLabel("Odds A:")); self.entry_odds_a = QLineEdit(); odds_a_layout.addWidget(self.entry_odds_a); cg.addLayout(odds_a_layout)
        odds_b_layout = QHBoxLayout(); odds_b_layout.addWidget(QLabel("Odds B:")); self.entry_odds_b = QLineEdit(); odds_b_layout.addWidget(self.entry_odds_b); cg.addLayout(odds_b_layout)
        self.btn_compare = QPushButton("Compare"); cg.addWidget(self.btn_compare)
        sidebar_layout.addWidget(self.compare_group)
        
        # Statistics Button
        self.btn_statistics = QPushButton("Statistics")
        self.btn_statistics.setCheckable(True)
        sidebar_layout.addWidget(self.btn_statistics)
        
        # Data Table Button
        self.btn_data_table = QPushButton("Data Table")
        self.btn_data_table.setCheckable(True)
        self.btn_data_table.setChecked(True)  # default view
        sidebar_layout.addWidget(self.btn_data_table)

        # Add New Bet Button
        self.btn_add_bet = QPushButton("Add New Bet")
        self.btn_add_bet.setCheckable(True)
        sidebar_layout.addWidget(self.btn_add_bet)

        # Pending Bets Button
        self.btn_pending_bets = QPushButton("Pending Bets")
        self.btn_pending_bets.setCheckable(True)
        sidebar_layout.addWidget(self.btn_pending_bets)

        self._nav_buttons = [self.btn_statistics, self.btn_data_table, self.btn_add_bet, self.btn_pending_bets]
        
        sidebar_layout.addStretch(1)

        # Bottom buttons
        self.btn_refresh = QPushButton(" Force Refresh"); self.btn_refresh.setIcon(QIcon(resource_path("icons/refresh-cw.svg"))); sidebar_layout.addWidget(self.btn_refresh)

        # --- Main Content (Results) ---
        main_content = QWidget(); main_layout = QVBoxLayout(main_content); main_layout.setContentsMargins(10, 10, 10, 0); root_layout.addWidget(main_content, 1)

        # Comparison Result Card
        compare_card = QGroupBox(); compare_card.setObjectName("CompareCard"); card_layout = QVBoxLayout(compare_card)
        self.compare_result = QLabel("Enter two odds above and click Compare."); self.compare_result.setWordWrap(True); self.compare_result.setAlignment(Qt.AlignmentFlag.AlignCenter); self.compare_result.setMinimumHeight(80); card_layout.addWidget(self.compare_result); main_layout.addWidget(compare_card)

        # Data Table
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Odds", "Live WR %", "Prematch WR %", "EV %"])
        
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(True)

        # Configure Header
        header = self.table.horizontalHeader()
        header.setSortIndicatorShown(True)
        header.setSectionsMovable(False)
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Stretch all columns to fill available space evenly
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        
        main_layout.addWidget(self.table, 1)
        
        # Statistics Panel (initially hidden)
        self.statistics_panel = QWidget()
        statistics_layout = QVBoxLayout(self.statistics_panel)
        statistics_layout.setContentsMargins(10, 10, 10, 10)
        # Empty panel for now
        placeholder_label = QLabel("Statistics panel - coming soon")
        placeholder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        placeholder_label.setStyleSheet("font-size: 16px; color: gray;")
        statistics_layout.addWidget(placeholder_label)
        main_layout.addWidget(self.statistics_panel, 1)
        self.statistics_panel.hide()  # Start hidden

        # --- Add Bet Panel (initially hidden) ---
        self._build_add_bet_panel(main_layout)

        # --- Pending Bets Panel (initially hidden) ---
        self._build_pending_bets_panel(main_layout)

        # --- Change Notification Panel (initially hidden) ---
        self.changes_panel = QWidget()
        changes_layout = QVBoxLayout(self.changes_panel)
        changes_layout.setContentsMargins(10, 10, 10, 10)
        
        # Header with title and dismiss button
        changes_header = QHBoxLayout()
        self.changes_title = QLabel("New Data Added")
        self.changes_title.setStyleSheet("font-size: 14px; font-weight: bold;")
        changes_header.addWidget(self.changes_title)
        changes_header.addStretch(1)
        self.btn_dismiss_changes = QPushButton("Dismiss")
        self.btn_dismiss_changes.setFixedWidth(80)
        changes_header.addWidget(self.btn_dismiss_changes)
        changes_layout.addLayout(changes_header)
        
        # Text area for changes
        self.changes_text = QLabel("No new data")
        self.changes_text.setWordWrap(True)
        self.changes_text.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.changes_text.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        changes_layout.addWidget(self.changes_text)
        
        main_layout.addWidget(self.changes_panel, 0)
        self.changes_panel.hide()  # Start hidden

        # --- Status Bar & Menu ---
        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar); self.set_status("Ready")
        act_settings = QAction("Settings", self); act_settings.triggered.connect(self.open_settings_dialog); self.menuBar().addAction(act_settings)
        act_exit = QAction("Exit", self); act_exit.triggered.connect(self.close); self.menuBar().addAction(act_exit)

    def _connect(self):
        self.btn_refresh.clicked.connect(lambda: self.refresh_data(force=True, force_network=True))
        self.btn_compare.clicked.connect(self.compare_by_odds)
        self.sport_combo.currentTextChanged.connect(self.on_sport_change)
        self.bettype_combo.currentTextChanged.connect(self.on_bet_type_change)
        self.tournament_combo.currentTextChanged.connect(self.on_tournament_change)
        self.entry_odds_a.textChanged.connect(self.recompute_comparison_inline)
        self.entry_odds_b.textChanged.connect(self.recompute_comparison_inline)
        self.team_combo.currentTextChanged.connect(self.on_team_change)
        self.btn_statistics.clicked.connect(self.show_statistics_panel)
        self.btn_data_table.clicked.connect(self.show_data_table_panel)
        self.btn_add_bet.clicked.connect(self.show_add_bet_panel)
        self.btn_pending_bets.clicked.connect(self.show_pending_bets_panel)
        self.btn_dismiss_changes.clicked.connect(self.dismiss_changes_panel)

    # Helpers
    def set_status(self, text: str): self.status_bar.showMessage(text)
    def set_controls_enabled(self, enabled: bool):
        for w in [self.sport_combo,self.bettype_combo,self.team_combo,self.tournament_combo,self.btn_refresh,self.entry_odds_a,self.entry_odds_b,self.btn_compare]:
            w.setEnabled(enabled)

    def get_sorted_sports(self) -> List[str]:
        def count_bets(sport):
            entry = self.data_cache[sport]
            total = 0
            for r in entry.rows:
                # r[3] is live_cnt, r[4] is prem_cnt
                total += (r[3] or 0) + (r[4] or 0)
            return total
        return sorted(self.data_cache.keys(), key=count_bets, reverse=True)
    # Data
    def fill_table(self, rows: List[RowTuple], bet_type: str):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        # Fetch theme colors from application properties (with fallbacks)
        app = QApplication.instance()
        from PyQt6.QtGui import QColor as _QColor  # local alias to avoid confusion
        pos = app.property("positiveColor") if app else None
        neg = app.property("negativeColor") if app else None
        neu = app.property("neutralColor") if app else None
        if not isinstance(pos, _QColor): pos = _QColor('green')
        if not isinstance(neg, _QColor): neg = _QColor('red')
        if not isinstance(neu, _QColor): neu = _QColor('gray')
        for odds, live_wr, prem_wr, _live_cnt, _prem_cnt in rows:
            wr = live_wr if bet_type=="Live" else prem_wr
            cnt = _live_cnt if bet_type=="Live" else _prem_cnt
            ev_str, ev_val = fmt_ev(wr, odds, cnt)
            r = self.table.rowCount(); self.table.insertRow(r)
            
            # Create sortable items
            it_odds = SortableTableWidgetItem(f"{odds:.2f}", odds)
            it_odds.setData(Qt.ItemDataRole.UserRole, (_live_cnt, _prem_cnt))
            it_live = SortableTableWidgetItem(fmt_wr(live_wr), live_wr)
            it_prem = SortableTableWidgetItem(fmt_wr(prem_wr), prem_wr)
            it_ev = SortableTableWidgetItem(ev_str, ev_val)
            
            cells = [it_odds, it_live, it_prem, it_ev]
            for c, it in enumerate(cells):
                it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(r,c,it)
            color = neu if wr is None else pos if ev_val is not None and ev_val>0 else neg
            for c in range(4): self.table.item(r,c).setForeground(color)
        self.table.setSortingEnabled(True)
        
        # Re-apply sort to match the indicator
        header = self.table.horizontalHeader()
        self.table.sortItems(header.sortIndicatorSection(), header.sortIndicatorOrder())

    def update_ev_only(self):
        self.table.setSortingEnabled(False)
        bt = self.bettype_combo.currentText()
        # Fetch theme colors (same logic as fill_table)
        app = QApplication.instance()
        from PyQt6.QtGui import QColor as _QColor
        pos = app.property("positiveColor") if app else None
        neg = app.property("negativeColor") if app else None
        neu = app.property("neutralColor") if app else None
        if not isinstance(pos, _QColor): pos = _QColor('green')
        if not isinstance(neg, _QColor): neg = _QColor('red')
        if not isinstance(neu, _QColor): neu = _QColor('gray')
        for r in range(self.table.rowCount()):
            it_odds = self.table.item(r,0)
            odds = float(it_odds.text())
            counts = it_odds.data(Qt.ItemDataRole.UserRole)
            live_cnt, prem_cnt = counts if counts else (None, None)

            def parse_cell(s:str):
                if s in ("","N/A"): return None
                return float(s.rstrip('%'))/100.0
            live_wr = parse_cell(self.table.item(r,1).text())
            prem_wr = parse_cell(self.table.item(r,2).text())
            wr = live_wr if bt=="Live" else prem_wr
            cnt = live_cnt if bt=="Live" else prem_cnt

            ev_str, ev_val = fmt_ev(wr, odds, cnt)
            
            it = self.table.item(r,3)
            it.setText(ev_str)
            if isinstance(it, SortableTableWidgetItem):
                it.sort_value = ev_val
            
            color = neu if wr is None else pos if ev_val is not None and ev_val>0 else neg
            for c in range(4): self.table.item(r,c).setForeground(color)
        self.table.setSortingEnabled(True)
        
        # Re-apply sort
        header = self.table.horizontalHeader()
        self.table.sortItems(header.sortIndicatorSection(), header.sortIndicatorOrder())
        
        self.recompute_comparison_inline()

    # Comparison
    def render_compare(self, odds_a, odds_b, wr_a, wr_b, ev_a, ev_b, sport, bet_type, count_a, count_b):
        label_type = 'Live' if bet_type=='Live' else 'Prematch'
        if wr_a is None and wr_b is None:
            self.compare_result.setText(f"Both odds are missing {label_type} WR data: {odds_a:.2f} and {odds_b:.2f}."); return
        if wr_a is None or wr_b is None:
            present_odds = odds_b if wr_a is None else odds_a
            present_wr = wr_b if wr_a is None else wr_a
            present_ev = ev_b if wr_a is None else ev_a
            present_count = count_b if wr_a is None else count_a
            missing_odds = odds_a if wr_a is None else odds_b
            cnt = 'N/A' if present_count is None else str(present_count)
            self.compare_result.setText(
                f"Odds {missing_odds:.2f} is missing {label_type} WR.\n\n"
                f"Bet {present_odds:.2f} EV: {present_ev*100:.2f}% (WR: {present_wr*100:.2f}%, Count: {cnt})")
            return
        # Compare EVs directly
        if ev_a is not None and ev_b is not None:
            if ev_a > ev_b: better=f"Bet {odds_a:.2f} is better"
            elif ev_b > ev_a: better=f"Bet {odds_b:.2f} is better"
            else: better="Both bets are equal"
        else:
            better="Cannot compare (missing data)"
        ca = 'N/A' if count_a is None else str(count_a); cb = 'N/A' if count_b is None else str(count_b)
        txt=(f"{better} ({bet_type} - {sport})\n\n"+
             f"Bet {odds_a:.2f} EV: {ev_a*100:.2f}% (WR: {wr_a*100:.2f}%, Count: {ca})\n"+
             f"Bet {odds_b:.2f} EV: {ev_b*100:.2f}% (WR: {wr_b*100:.2f}%, Count: {cb})")
        notes=[]
        if count_a==1: notes.append(f"Note: Odds {odds_a:.2f} has only 1 count.")
        if count_b==1: notes.append(f"Note: Odds {odds_b:.2f} has only 1 count.")
        if notes: txt += "\n\n"+" ".join(notes)
        self.compare_result.setText(txt)

    def recompute_comparison_inline(self):
        try:
            odds_a = float(self.entry_odds_a.text().strip()); odds_b = float(self.entry_odds_b.text().strip())
        except ValueError:
            self.compare_result.setText(""); return
        bet_type = self.bettype_combo.currentText()
        k_a, k_b = round_odds_key(odds_a), round_odds_key(odds_b)
        miss_a = k_a not in self.odds_index; miss_b = k_b not in self.odds_index
        if miss_a and miss_b:
            self.compare_result.setText(f"Both odds not found in data: {odds_a:.2f} and {odds_b:.2f}."); return
        if miss_a or miss_b:
            pk = k_b if miss_a else k_a; po = odds_b if miss_a else odds_a
            wr_live, wr_pre, live_cnt, prem_cnt = self.odds_index[pk]
            wr_p = wr_live if bet_type=="Live" else wr_pre
            cnt_p = live_cnt if bet_type=="Live" else prem_cnt
            if wr_p is None:
                self.compare_result.setText((f"Odd {odds_a:.2f} not found. " if miss_a else f"Odd {odds_b:.2f} not found. ")+f"Odds {po:.2f} is in data but missing {bet_type} WR."); return
            ev_p = calc_ev(po, wr_p, cnt_p)
            if ev_p is None: ev_p = 0.0
            cnt_s='N/A' if cnt_p is None else str(cnt_p)
            missing = f"Odd {odds_a:.2f} not found." if miss_a else f"Odd {odds_b:.2f} not found."
            self.compare_result.setText(f"{missing}\n\nBet {po:.2f} EV: {ev_p*100:.2f}% (WR: {wr_p*100:.2f}%, Count: {cnt_s})")
            return
        wr_a_live, wr_a_pre, live_cnt_a, prem_cnt_a = self.odds_index[k_a]; wr_b_live, wr_b_pre, live_cnt_b, prem_cnt_b = self.odds_index[k_b]
        wr_a = wr_a_live if bet_type=="Live" else wr_a_pre; wr_b = wr_b_live if bet_type=="Live" else wr_b_pre
        cnt_a = live_cnt_a if bet_type=="Live" else prem_cnt_a
        cnt_b = live_cnt_b if bet_type=="Live" else prem_cnt_b
        ev_a = calc_ev(odds_a, wr_a, cnt_a)
        ev_b = calc_ev(odds_b, wr_b, cnt_b)
        self.render_compare(odds_a, odds_b, wr_a, wr_b, ev_a, ev_b, self.sport_combo.currentText(), bet_type, cnt_a, cnt_b)

    def compare_by_odds(self): self.recompute_comparison_inline()
    def on_bet_type_change(self): self.update_ev_only()

    def on_sport_change(self):
        sport = self.sport_combo.currentText()
        self._populate_tournament_combo(sport)
        self._populate_team_combo(sport)
        self.refresh_data(False)

    def on_tournament_change(self):
        self._apply_filters()

    def on_team_change(self):
        self._apply_filters()

    def _get_tournaments_for_sport(self, sport: str) -> List[str]:
        """Return normalized tournament names for the given sport, sorted by bet count descending."""
        counts: Dict[str, int] = {}
        for s, tournament, _, _, _, _, _ in self.matchbet_data:
            if s != sport:
                continue
            for label in tournament_filter_labels(tournament):
                counts[label] = counts.get(label, 0) + 1
        return sorted(counts.keys(), key=lambda t: counts[t], reverse=True)

    def _get_teams_for_sport(self, sport: str) -> List[str]:
        """Return team names for the given sport, sorted alphabetically."""
        teams: set = set()
        for s, _, matchup, bet, _, _, _ in self.matchbet_data:
            if s != sport:
                continue
            team = extract_team_from_bet(matchup, bet)
            if team:
                teams.add(team)
        return sorted(teams, key=str.lower)

    def _populate_tournament_combo(self, sport: str):
        """Repopulate the tournament combo for the given sport."""
        self.tournament_combo.blockSignals(True)
        self.tournament_combo.clear()
        self.tournament_combo.addItem("All Tournaments")
        for t in self._get_tournaments_for_sport(sport):
            self.tournament_combo.addItem(t)
        self.tournament_combo.setCurrentIndex(0)
        self.tournament_combo.blockSignals(False)

    def _populate_team_combo(self, sport: str):
        """Repopulate the team combo for the given sport."""
        self.team_combo.blockSignals(True)
        self.team_combo.clear()
        self.team_combo.addItem("All Teams")
        for t in self._get_teams_for_sport(sport):
            self.team_combo.addItem(t)
        self.team_combo.setCurrentIndex(0)
        self.team_combo.blockSignals(False)

    def _apply_filters(self):
        """Apply all active filters (tournament + team) and update the table and comparison."""
        sport = self.sport_combo.currentText()
        tournament = self.tournament_combo.currentText()
        team = self.team_combo.currentText()
        if not sport or sport not in self.data_cache:
            return

        all_tournaments = (tournament == "All Tournaments" or not tournament)
        all_teams = (team == "All Teams" or not team)

        if all_tournaments and all_teams:
            # No filters — use full sport cache
            entry = self.data_cache[sport]
        else:
            # Filter matchbet_data then rebuild cache
            filtered: List[MatchBetTuple] = []
            for row in self.matchbet_data:
                if row[0] != sport:
                    continue
                # Tournament filter
                if not all_tournaments:
                    labels = tournament_filter_labels(row[1])
                    if tournament not in labels:
                        continue
                # Team filter — only include bets explicitly placed on the selected team
                if not all_teams:
                    bet_team = extract_team_from_bet(row[2], row[3])
                    if bet_team != team:
                        continue
                filtered.append(row)
            if not filtered:
                self.current_rows = []
                self.odds_index = {}
                self.fill_table([], self.bettype_combo.currentText())
                self.recompute_comparison_inline()
                return
            cache = process_bets_to_cache(filtered)
            entry = cache.get(sport)
            if entry is None:
                self.current_rows = []
                self.odds_index = {}
                self.fill_table([], self.bettype_combo.currentText())
                self.recompute_comparison_inline()
                return

        self.current_rows = entry.rows
        self.odds_index = entry.index
        self.fill_table(self.current_rows, self.bettype_combo.currentText())
        self.recompute_comparison_inline()

    def on_theme_toggled(self, checked: bool):
        self.dark_mode = checked
        theme_manager.apply_theme(QApplication.instance(), dark=checked)

    def open_settings_dialog(self):
        dlg = SettingsDialog(self)
        dlg.exec()

    def _update_nav_buttons(self, active):
        """Highlight only the active navigation button."""
        for btn in self._nav_buttons:
            btn.setChecked(btn is active)

    def show_statistics_panel(self):
        """Switch to statistics panel view"""
        self.current_view = "statistics"
        self._update_nav_buttons(self.btn_statistics)
        self.table.hide()
        self.add_bet_panel.hide()
        self.pending_bets_panel.hide()
        self.statistics_panel.show()
        
        # Clear existing layout
        layout = self.statistics_panel.layout()
        if layout is None:
            layout = QVBoxLayout(self.statistics_panel)
            self.statistics_panel.setLayout(layout)
        
        # Remove existing widgets in layout
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        # --- Container 1: Tournament Statistics ---
        tournament_group = QGroupBox("Tournament Statistics")
        t_layout = QVBoxLayout(tournament_group)

        # Group by Sport -> Tournament
        # matchbet_data is list of (sport, tournament, matchup, bet, live_status, odds, result)
        # stats structure: stats[sport][tournament] = {'wins': 0, 'total': 0}
        stats = {}
        for sport, tournament, _, _, _, _, result in self.matchbet_data:
            if not sport: continue
            
            # Normalize sport names
            if sport == "CStwo":
                sport = "CS2"

            if not tournament: continue
            
            norm_tourney = normalize_tournament_name(tournament)
            if not norm_tourney: continue
            
            if sport not in stats:
                stats[sport] = {}
            if norm_tourney not in stats[sport]:
                stats[sport][norm_tourney] = {'wins': 0, 'total': 0}
            
            stats[sport][norm_tourney]['total'] += 1
            if result.lower() == "win":
                stats[sport][norm_tourney]['wins'] += 1
        
        # Create Tree Widget
        tree = QTreeWidget()
        tree.setMinimumWidth(400)
        tree.setHeaderLabels(["Sport / Tournament", "Bet Count", "Winrate %"])
        tree.setAlternatingRowColors(True)
        
        # Configure columns: Stretch first, Fixed width for others
        header = tree.header()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionsMovable(False)
        
        tree.setColumnWidth(1, 100)
        tree.setColumnWidth(2, 100)
        
        # Center align headers for numeric columns
        tree.headerItem().setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
        tree.headerItem().setTextAlignment(2, Qt.AlignmentFlag.AlignCenter)
        
        # Populate Tree
        # Sort sports by total bets descending
        sport_totals = {}
        for s, tourneys in stats.items():
            s_wins = sum(t['wins'] for t in tourneys.values())
            s_total = sum(t['total'] for t in tourneys.values())
            sport_totals[s] = {'wins': s_wins, 'total': s_total}

        sorted_sports = sorted(stats.keys(), key=lambda s: sport_totals[s]['total'], reverse=True)

        for sport in sorted_sports:
            s_data = sport_totals[sport]
            total_sport_bets = s_data['total']
            total_sport_wins = s_data['wins']
            sport_wr = (total_sport_wins / total_sport_bets * 100) if total_sport_bets > 0 else 0.0
            
            sport_item = QTreeWidgetItem(tree)
            sport_item.setText(0, sport)
            sport_item.setText(1, str(total_sport_bets))
            sport_item.setText(2, f"{sport_wr:.1f}%")
            sport_item.setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
            sport_item.setTextAlignment(2, Qt.AlignmentFlag.AlignCenter)
            sport_item.setExpanded(False)
            
            # Sort tournaments by count descending
            tourneys = stats[sport]
            sorted_tourneys = sorted(tourneys.items(), key=lambda item: item[1]['total'], reverse=True)

            for tourney, t_data in sorted_tourneys:
                t_total = t_data['total']
                t_wins = t_data['wins']
                t_wr = (t_wins / t_total * 100) if t_total > 0 else 0.0

                t_item = QTreeWidgetItem(sport_item)
                t_item.setText(0, tourney)
                t_item.setText(1, str(t_total))
                t_item.setText(2, f"{t_wr:.1f}%")
                t_item.setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
                t_item.setTextAlignment(2, Qt.AlignmentFlag.AlignCenter)

        total_visible_bets = sum(d['total'] for d in sport_totals.values())
        t_layout.addWidget(QLabel(f"Total Bets: {total_visible_bets}"))
        t_layout.addWidget(tree)

        # --- Container 2: Empty ---
        second_group = QGroupBox("Additional Statistics")
        QVBoxLayout(second_group)

        # Add containers to main layout
        layout.addWidget(tournament_group, 4)
        layout.addWidget(second_group, 2)
        layout.addStretch(1)

    def show_data_table_panel(self):
        """Switch to data table view"""
        self.current_view = "table"
        self._update_nav_buttons(self.btn_data_table)
        self.statistics_panel.hide()
        self.add_bet_panel.hide()
        self.pending_bets_panel.hide()
        self.table.show()

    # ------------------------------------------------------------------
    # Add Bet Panel
    # ------------------------------------------------------------------
    def _build_add_bet_panel(self, parent_layout):
        """Create the Add Bet / Settle Bet form panel."""
        self.add_bet_panel = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        form_layout = QVBoxLayout(inner)
        form_layout.setContentsMargins(20, 20, 20, 20)
        form_layout.setSpacing(10)

        self.add_bet_title = QLabel("Add New Bet")
        self.add_bet_title.setStyleSheet("font-size: 18px; font-weight: bold;")
        form_layout.addWidget(self.add_bet_title)

        form = QFormLayout()
        form.setSpacing(8)

        # Sport
        self.form_sport = QComboBox()
        self.form_sport.setEditable(True)
        form.addRow("Sport:", self.form_sport)

        # Tournament
        self.form_tournament = QLineEdit()
        form.addRow("Tournament:", self.form_tournament)

        # Matchup
        self.form_matchup = QLineEdit()
        self.form_matchup.setPlaceholderText("e.g. Team A vs Team B")
        form.addRow("Matchup:", self.form_matchup)

        # Bet
        self.form_bet = QLineEdit()
        self.form_bet.setPlaceholderText("e.g. Match: Team A")
        form.addRow("Bet:", self.form_bet)

        # Live Status
        self.form_live_status = QComboBox()
        self.form_live_status.addItems(["LIVE", "NOT LIVE"])
        form.addRow("Live Status:", self.form_live_status)

        # Odds
        self.form_odds = QLineEdit()
        self.form_odds.setPlaceholderText("e.g. 1.75")
        form.addRow("Odds:", self.form_odds)

        # Bet Amount
        self.form_bet_amount = QLineEdit()
        self.form_bet_amount.setPlaceholderText("e.g. 10.00")
        form.addRow("Bet Amount:", self.form_bet_amount)

        # Result
        self.form_result = QComboBox()
        self.form_result.addItems(["(Pending)", "Win", "Lose"])
        form.addRow("Result:", self.form_result)

        # Profit (auto-calculated)
        self.form_profit = QLineEdit()
        self.form_profit.setPlaceholderText("Auto-calculated from odds & amount")
        self.form_profit.setReadOnly(True)
        form.addRow("Profit:", self.form_profit)

        form_layout.addLayout(form)

        # Buttons
        btn_row = QHBoxLayout()
        self.btn_save_bet = QPushButton("Save Bet")
        self.btn_save_bet.clicked.connect(self.save_bet)
        btn_row.addWidget(self.btn_save_bet)

        btn_clear = QPushButton("Clear Form")
        btn_clear.clicked.connect(self.clear_bet_form)
        btn_row.addWidget(btn_clear)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.show_data_table_panel)
        btn_row.addWidget(btn_cancel)

        form_layout.addLayout(btn_row)
        form_layout.addStretch(1)

        scroll.setWidget(inner)
        panel_layout = QVBoxLayout(self.add_bet_panel)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.addWidget(scroll)
        parent_layout.addWidget(self.add_bet_panel, 1)
        self.add_bet_panel.hide()

        # Connect signals for auto-profit calculation
        self.form_result.currentTextChanged.connect(self._auto_calc_profit)
        self.form_odds.textChanged.connect(self._auto_calc_profit)
        self.form_bet_amount.textChanged.connect(self._auto_calc_profit)

    def _auto_calc_profit(self):
        """Recalculate profit when result, odds, or bet amount change."""
        result = self.form_result.currentText()
        if result == "(Pending)":
            self.form_profit.clear()
            return
        try:
            odds = float(self.form_odds.text().strip().replace(",", "."))
            amount = float(self.form_bet_amount.text().strip().replace(",", "."))
        except (ValueError, AttributeError):
            self.form_profit.clear()
            return
        if result == "Win":
            profit = (odds - 1.0) * amount
        else:  # Lose
            profit = -amount
        self.form_profit.setText(f"{profit:.2f}")

    def _populate_form_sports(self):
        """Populate the sport combo in the add-bet form from existing data."""
        self.form_sport.blockSignals(True)
        current = self.form_sport.currentText()
        self.form_sport.clear()
        sports = sorted(set(
            s for s, *_ in self.matchbet_data
        )) if self.matchbet_data else []
        # Also include sports from db in case matchbet_data is empty
        try:
            db_sports = self.db.get_distinct_sports()
            sports = sorted(set(sports) | set(db_sports))
        except Exception:
            pass
        for s in sports:
            self.form_sport.addItem(s)
        if current:
            self.form_sport.setCurrentText(current)
        self.form_sport.blockSignals(False)

    def show_add_bet_panel(self):
        """Switch to add-bet form in add mode."""
        self.current_view = "add_bet"
        self._update_nav_buttons(self.btn_add_bet)
        self.table.hide()
        self.statistics_panel.hide()
        self.pending_bets_panel.hide()
        self.add_bet_panel.show()
        self._reset_form_to_add_mode()

    def _reset_form_to_add_mode(self):
        """Reset the form to add-bet mode (clear fields, enable all)."""
        self.editing_bet_id = None
        self.add_bet_title.setText("Add New Bet")
        self.btn_save_bet.setText("Save Bet")
        self._populate_form_sports()

        # Enable all fields
        for w in [self.form_sport, self.form_tournament, self.form_matchup,
                   self.form_bet, self.form_live_status,
                   self.form_odds, self.form_bet_amount]:
            w.setEnabled(True)

        self.form_result.setEnabled(True)
        self.form_profit.setReadOnly(True)
        self.clear_bet_form()

    def clear_bet_form(self):
        """Clear all form fields to defaults."""
        if self.editing_bet_id is None:
            self.form_sport.setCurrentIndex(0)
        self.form_tournament.clear()
        self.form_matchup.clear()
        self.form_bet.clear()
        self.form_live_status.setCurrentIndex(0)
        self.form_odds.clear()
        self.form_bet_amount.clear()
        self.form_result.setCurrentIndex(0)
        self.form_profit.clear()

    def save_bet(self):
        """Validate and save / update a bet."""
        sport = self.form_sport.currentText().strip()
        tournament = self.form_tournament.text().strip()
        matchup = self.form_matchup.text().strip()
        bet = self.form_bet.text().strip()
        live_status = self.form_live_status.currentText()
        odds_text = self.form_odds.text().strip()
        amount_text = self.form_bet_amount.text().strip()
        result_text = self.form_result.currentText()
        profit_text = self.form_profit.text().strip()

        # --- Validation ---
        errors: List[str] = []
        if not sport:
            errors.append("Sport is required.")
        if not tournament:
            errors.append("Tournament is required.")
        if not matchup:
            errors.append("Matchup is required.")
        if not bet:
            errors.append("Bet is required.")

        # Odds
        odds: Optional[float] = None
        try:
            odds = float(odds_text.replace(",", "."))
            if odds <= 0:
                errors.append("Odds must be > 0.")
        except ValueError:
            errors.append("Odds must be a valid positive number.")

        # Bet amount (optional but must be valid if given)
        bet_amount: Optional[float] = None
        if amount_text:
            try:
                bet_amount = float(amount_text.replace(",", "."))
                if bet_amount < 0:
                    errors.append("Bet Amount must be >= 0.")
            except ValueError:
                errors.append("Bet Amount must be a valid number.")

        # Result / Profit (profit is auto-calculated)
        is_pending = (result_text == "(Pending)")
        result: Optional[str] = None
        profit: Optional[float] = None

        if not is_pending:
            result = result_text
            if not profit_text:
                if not bet_amount:
                    errors.append("Bet Amount is required to calculate profit.")
                else:
                    errors.append("Could not calculate profit – check Odds and Bet Amount.")
            else:
                try:
                    profit = float(profit_text.replace(",", "."))
                except ValueError:
                    errors.append("Profit must be a valid number.")

        if self.editing_bet_id is not None and is_pending:
            errors.append("You must select Win or Lose to settle this bet.")

        if errors:
            QMessageBox.warning(self, "Validation Error", "\n".join(errors))
            return

        # --- Save ---
        try:
            if self.editing_bet_id is not None:
                self.db.update_bet(self.editing_bet_id, result, profit)  # type: ignore[arg-type]
                QMessageBox.information(self, "Success", "Bet settled successfully!")
                self._reset_form_to_add_mode()
                self.refresh_data(force=True)
                self.show_pending_bets_panel()
            else:
                self.db.insert_bet(
                    sport=sport,
                    tournament=tournament,
                    matchup=matchup,
                    bet=bet,
                    live_status=live_status,
                    odds=odds,  # type: ignore[arg-type]
                    bet_amount=bet_amount,
                    result=result,
                    profit=profit,
                )
                msg = "Bet added successfully!" if result else "Pending bet saved!"
                QMessageBox.information(self, "Success", msg)
                self.clear_bet_form()
                self.refresh_data(force=True)
                self.show_data_table_panel()
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Failed to save bet:\n{e}")

    # ------------------------------------------------------------------
    # Pending Bets Panel
    # ------------------------------------------------------------------
    def _build_pending_bets_panel(self, parent_layout):
        """Build the pending-bets view with a table and settle buttons."""
        self.pending_bets_panel = QWidget()
        p_layout = QVBoxLayout(self.pending_bets_panel)
        p_layout.setContentsMargins(10, 10, 10, 10)

        self.lbl_pending_count = QLabel("0 pending bet(s)")
        self.lbl_pending_count.setStyleSheet("font-size: 15px; font-weight: bold;")
        p_layout.addWidget(self.lbl_pending_count)

        self.pending_table = QTableWidget(0, 10)
        self.pending_table.setHorizontalHeaderLabels([
            "ID", "Sport", "Tournament", "Matchup", "Bet",
            "Live Status", "Odds", "Bet Amount", "Date Created", "Actions"
        ])
        self.pending_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.pending_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.pending_table.verticalHeader().setVisible(False)
        ph = self.pending_table.horizontalHeader()
        ph.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        for c in range(1, 9):
            ph.setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)
        ph.setSectionResizeMode(9, QHeaderView.ResizeMode.ResizeToContents)
        p_layout.addWidget(self.pending_table, 1)

        parent_layout.addWidget(self.pending_bets_panel, 1)
        self.pending_bets_panel.hide()

    def show_pending_bets_panel(self):
        """Switch to the pending-bets view."""
        self.current_view = "pending_bets"
        self._update_nav_buttons(self.btn_pending_bets)
        self.table.hide()
        self.statistics_panel.hide()
        self.add_bet_panel.hide()
        self.pending_bets_panel.show()
        self.refresh_pending_bets_table()

    def refresh_pending_bets_table(self):
        """Fetch and display pending bets."""
        try:
            rows = self.db.fetch_pending_bets()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load pending bets:\n{e}")
            return
        self.lbl_pending_count.setText(f"{len(rows)} pending bet(s)")
        self.pending_table.setRowCount(0)
        for row in rows:
            # row: (id, sport, tournament, matchup, bet,
            #        live_status, odds, bet_amount, result, profit, date_created)
            r = self.pending_table.rowCount()
            self.pending_table.insertRow(r)
            bid = row[0]
            items = [
                QTableWidgetItem(str(bid)),
                QTableWidgetItem(str(row[1])),
                QTableWidgetItem(str(row[2])),
                QTableWidgetItem(str(row[3])),
                QTableWidgetItem(str(row[4])),
                QTableWidgetItem(str(row[5])),
                QTableWidgetItem(f"{row[6]:.2f}" if row[6] else ""),
                QTableWidgetItem(f"{row[7]:.2f}" if row[7] else ""),
                QTableWidgetItem(str(row[10] or "")),
            ]
            for c, it in enumerate(items):
                it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.pending_table.setItem(r, c, it)
            btn = QPushButton("Settle Bet")
            btn.setMinimumWidth(90)
            btn.setStyleSheet(
                "QPushButton { background-color: #89b4fa; color: #1e1e2e; font-weight: bold; "
                "padding: 4px 12px; border-radius: 6px; }"
                "QPushButton:hover { background-color: #74a8f7; }"
            )
            btn.clicked.connect(lambda checked, b=row: self._settle_bet(b))
            self.pending_table.setCellWidget(r, 9, btn)
            self.pending_table.setRowHeight(r, 40)

    def _settle_bet(self, row_data):
        """Switch to the add-bet form in edit/settle mode for a pending bet."""
        # row_data: (id, sport, tournament, matchup, bet,
        #            live_status, odds, bet_amount, result, profit, date_created)
        self.editing_bet_id = row_data[0]
        self.current_view = "add_bet"
        self._update_nav_buttons(self.btn_add_bet)
        self.table.hide()
        self.statistics_panel.hide()
        self.pending_bets_panel.hide()
        self.add_bet_panel.show()

        self.add_bet_title.setText("Settle Pending Bet")
        self.btn_save_bet.setText("Update Bet")

        self._populate_form_sports()

        # Fill fields
        self.form_sport.setCurrentText(str(row_data[1]))
        self.form_tournament.setText(str(row_data[2]))
        self.form_matchup.setText(str(row_data[3]))
        self.form_bet.setText(str(row_data[4]))
        idx = self.form_live_status.findText(str(row_data[5]))
        if idx >= 0:
            self.form_live_status.setCurrentIndex(idx)
        self.form_odds.setText(f"{row_data[6]:.2f}" if row_data[6] else "")
        self.form_bet_amount.setText(f"{row_data[7]:.2f}" if row_data[7] else "")
        self.form_result.setCurrentIndex(0)  # (Pending)
        self.form_profit.clear()

        # Disable fields that shouldn't change when settling
        for w in [self.form_sport, self.form_tournament, self.form_matchup,
                   self.form_bet, self.form_live_status,
                   self.form_odds, self.form_bet_amount]:
            w.setEnabled(False)

        # Keep result editable; profit is auto-calculated
        self.form_result.setEnabled(True)
        self.form_profit.setReadOnly(True)

    def dismiss_changes_panel(self):
        """Hide the changes notification panel"""
        self.changes_panel.hide()

    def _detect_odds_changes(self) -> List[Dict]:
        """Detect changes at the odds level showing previous vs new EV, WR, and count"""
        if not self._previous_data_cache:
            return []  # No previous data to compare against
        
        changes = []
        
        # Check each sport in the new data
        for sport, new_entry in self.data_cache.items():
            old_entry = self._previous_data_cache.get(sport)
            
            # Build dict of old odds data for quick lookup
            old_odds_map = {}
            if old_entry:
                for row in old_entry.rows:
                    odds = row[0]  # odds is first element
                    old_odds_map[odds] = {
                        'live_wr': row[1],
                        'prem_wr': row[2], 
                        'live_cnt': row[3],
                        'prem_cnt': row[4]
                    }
            
            # Check each new odds entry
            for row in new_entry.rows:
                odds, live_wr, prem_wr, live_cnt, prem_cnt = row
                
                if odds not in old_odds_map:
                    # New odds entry - check which type has data
                    has_live = live_cnt is not None and live_cnt > 0
                    has_prem = prem_cnt is not None and prem_cnt > 0
                    
                    if has_live:
                        changes.append({
                            'sport': sport,
                            'odds': odds,
                            'bet_type': 'LIVE',
                            'type': 'new',
                            'prev_wr': None,
                            'prev_cnt': None,
                            'new_wr': live_wr,
                            'new_cnt': live_cnt
                        })
                    if has_prem:
                        changes.append({
                            'sport': sport,
                            'odds': odds,
                            'bet_type': 'PREMATCH',
                            'type': 'new',
                            'prev_wr': None,
                            'prev_cnt': None,
                            'new_wr': prem_wr,
                            'new_cnt': prem_cnt
                        })
                else:
                    # Existing odds - check which type changed
                    old = old_odds_map[odds]
                    
                    # Check LIVE changes
                    if old['live_cnt'] != live_cnt or old['live_wr'] != live_wr:
                        changes.append({
                            'sport': sport,
                            'odds': odds,
                            'bet_type': 'LIVE',
                            'type': 'updated',
                            'prev_wr': old['live_wr'],
                            'prev_cnt': old['live_cnt'],
                            'new_wr': live_wr,
                            'new_cnt': live_cnt
                        })
                    
                    # Check PREMATCH changes
                    if old['prem_cnt'] != prem_cnt or old['prem_wr'] != prem_wr:
                        changes.append({
                            'sport': sport,
                            'odds': odds,
                            'bet_type': 'PREMATCH',
                            'type': 'updated',
                            'prev_wr': old['prem_wr'],
                            'prev_cnt': old['prem_cnt'],
                            'new_wr': prem_wr,
                            'new_cnt': prem_cnt
                        })
        
        return changes

    def _calc_ev_display(self, wr: Optional[float], odds: float, cnt: Optional[int]) -> Tuple[str, Optional[float]]:
        """Calculate EV for display"""
        if wr is None or cnt is None:
            return "N/A", None
        ev = calc_ev(odds, wr, cnt)
        if ev is None:
            return "N/A", None
        return f"{ev*100:.2f}%", ev

    def _format_odds_change(self, change: Dict) -> str:
        """Format a single odds change for display - showing only the specific bet type that changed"""
        sport = change['sport']
        odds = change['odds']
        ctype = change['type']
        bet_type = change['bet_type']
        
        icon = "🔴" if bet_type == "LIVE" else "⚫"
        
        prev_ev_str, prev_ev = self._calc_ev_display(change['prev_wr'], odds, change['prev_cnt'])
        new_ev_str, new_ev = self._calc_ev_display(change['new_wr'], odds, change['new_cnt'])
        
        lines = [f"<b>{sport}</b> | Odds: {odds:.2f} | {icon} {bet_type} ({ctype.upper()})<br>"]
        
        if change['prev_cnt'] is not None:
            lines.append(f"&nbsp;&nbsp;Previous: WR={fmt_wr(change['prev_wr'])}, Count={change['prev_cnt']}, EV={prev_ev_str}<br>")
        else:
            lines.append(f"&nbsp;&nbsp;Previous: No data<br>")
        
        lines.append(f"&nbsp;&nbsp;<b>Updated:</b> WR={fmt_wr(change['new_wr'])}, Count={change['new_cnt']}, EV={new_ev_str}<br>")
        
        if change['prev_cnt'] is not None and prev_ev is not None and new_ev is not None:
            diff = new_ev - prev_ev
            diff_str = f"+{diff*100:.2f}%" if diff >= 0 else f"{diff*100:.2f}%"
            color = "green" if diff >= 0 else "red"
            lines.append(f"&nbsp;&nbsp;EV Change: <span style='color:{color}'>{diff_str}</span><br>")
        
        lines.append("<br>")
        return "".join(lines)

    def _format_changes_message(self, changes: List[Dict]) -> str:
        """Format the list of odds changes into a readable message"""
        if not changes:
            return "No new data was added."
        
        lines = [f"<b>{len(changes)} odds value(s) updated:</b><br><br>"]
        
        for change in changes:
            lines.append(self._format_odds_change(change))
        
        return "".join(lines)

    def _display_changes(self, changes: List[Dict]):
        """Display the changes panel with detected changes only"""
        if not changes:
            self.changes_panel.hide()
            return
        
        message = self._format_changes_message(changes)
        count = len(changes)
        self.changes_title.setText(f"📊 {count} Odds Value(s) Changed")
        self.changes_text.setText(message)
        self.changes_panel.show()

    def _format_current_data_message(self) -> str:
        """Format message showing current data state"""
        total_odds = sum(len(entry.rows) for entry in self.data_cache.values())
        total_bets = len(self.matchbet_data)
        
        lines = [f"<b>Data Loaded Successfully</b><br><br>"]
        lines.append(f"Total Sports: {len(self.data_cache)}<br>")
        lines.append(f"Total Odds Values: {total_odds}<br>")
        lines.append(f"Total Bets: {total_bets}<br><br>")
        
        # Show summary by sport
        for sport in sorted(self.data_cache.keys()):
            entry = self.data_cache[sport]
            sport_bets = sum((r[3] or 0) + (r[4] or 0) for r in entry.rows)
            lines.append(f"• <b>{sport}</b>: {len(entry.rows)} odds values ({sport_bets} bets)<br>")
        
        return "".join(lines)

    def refresh_data(self, force: bool=False, force_network: bool=False):
        sheet = self.sport_combo.currentText()
        # If we already have this sheet cached and not forcing, use cache instantly
        if not force and sheet in self.data_cache:
            self.setWindowTitle(f"EV Bet Calculator - {sheet}")
            self.set_status("Loaded from cache")
            self._apply_filters()
            return
        # Otherwise perform an async refresh from database
        # Store previous data for comparison (always track changes)
        self._previous_data_cache = self.data_cache.copy()
        self._previous_matchbet_data = self.matchbet_data.copy()
        self.set_status(f"Refreshing all data..."); self.set_controls_enabled(False)
        t=QThread(); w=RefreshWorker(self.db, force_refresh_network=force_network); w.moveToThread(t)
        t.started.connect(w.run)
        w.finished.connect(lambda ok, info, res: self._on_refresh_done(t,w,ok,info,res))
        t.start()

    def _on_refresh_done(self, thread: QThread, worker: RefreshWorker, ok: bool, info: str, res: object):
        thread.quit(); thread.wait(); worker.deleteLater()
        if ok and res:
            cache, data = res
            self.data_cache = cache
            
            # Detect and display changes at odds level
            changes = self._detect_odds_changes()
            self._display_changes(changes)
            
            self.matchbet_data = data
            
            # Update sport combo if new sports appeared
            current_sport = self.sport_combo.currentText()
            self.sport_combo.blockSignals(True)
            self.sport_combo.clear()
            self.sport_combo.addItems(self.get_sorted_sports())
            if current_sport in self.data_cache:
                self.sport_combo.setCurrentText(current_sport)
            elif self.sport_combo.count() > 0:
                self.sport_combo.setCurrentIndex(0)
            self.sport_combo.blockSignals(False)
            
            # Repopulate tournament and team combos for the current sport
            current_sport = self.sport_combo.currentText()
            self._populate_tournament_combo(current_sport)
            self._populate_team_combo(current_sport)
            
            # Refresh view using filters
            if current_sport in self.data_cache:
                self.setWindowTitle(f"EV Bet Calculator - {current_sport}")
                self._apply_filters()
            else:
                # Cache is empty (e.g. after DB reset) – clear visible tables
                self.setWindowTitle("EV Bet Calculator")
                self.table.setRowCount(0)
                self.refresh_pending_bets_table()
        else:
            QMessageBox.critical(self, "Error", f"Failed to load data: {info}")
        self.set_controls_enabled(True); self.set_status("Ready")

    def set_initial_cache(self, cache: Dict[str, SheetCacheEntry]):
        self.data_cache = cache
        
        # Populate sports combo
        self.sport_combo.blockSignals(True)
        self.sport_combo.clear()
        sports = self.get_sorted_sports()
        self.sport_combo.addItems(sports)
        self.sport_combo.blockSignals(False)
        
        if sports:
            first = sports[0]
            self.sport_combo.setCurrentText(first)
            self._populate_tournament_combo(first)
            self._populate_team_combo(first)
            self._apply_filters()

    def set_matchbet_data(self, data: List[MatchBetTuple]):
        self.matchbet_data = data


def main():
    app = QApplication(sys.argv)

    # Ensure the database is created
    db = DatabaseManager()
    db.create_tables()

    win = MainWindow()
    dlg = PreloadDialog()
    preload_thread = QThread()
    worker = PreloadWorker(win.db)
    worker.moveToThread(preload_thread)
    preload_thread.started.connect(worker.run)
    worker.progress.connect(dlg.update_progress)
    worker.status.connect(dlg.update_status)

    def done(ok: bool, msg: str):
        preload_thread.quit(); preload_thread.wait(); worker.deleteLater(); dlg.accept()
        if not ok:
            QMessageBox.critical(win, "Preload Failed", f"Failed to preload data:\n{msg}")
            win.close()
            return
        win.set_initial_cache(worker.cache)
        win.set_matchbet_data(worker.matchbet_data)
        win.show()
        QTimer.singleShot(0, lambda: win.refresh_data(force=True))

    worker.finished.connect(done)
    QTimer.singleShot(60000, lambda: done(False, "Preloading timed out") if preload_thread.isRunning() else None)
    preload_thread.start()
    dlg.exec()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())