from __future__ import annotations

import sys, os, re, time
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QAction, QColor, QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QTableWidget, QTableWidgetItem, QComboBox, QPushButton, QGroupBox, QLineEdit,
    QMessageBox, QDialog, QProgressBar, QStatusBar, QHeaderView,
    QTreeWidget, QTreeWidgetItem
)

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

# Ensure project root is in sys.path so we can import from src
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

try:
    import src.theme_manager as theme_manager
except ModuleNotFoundError:
    # Fallback if running directly from src without package context
    import theme_manager as theme_manager

load_dotenv()

JSON_KEYFILE = "my-matchbettings-ev-script-878cbe8fa582.json"
SPREADSHEET_NAME = "Match betting"

MATCHBET_SHEET_NAME = "MATCHBET"
MATCHBET_COL_RANGE = "A2:H"

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


def fmt_wr(wr: Optional[float]) -> str:
    return "N/A" if wr is None else f"{wr*100:.2f}%"


def fmt_ev(wr: Optional[float], odds: Optional[float]):
    if wr is None or odds is None:
        return "N/A", None
    ev = (wr * odds) - 1.0
    return f"{ev*100:.2f}%", ev


def authorize_client():
    # Try to get credentials from environment variable first
    env_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if env_path and os.path.exists(env_path):
        path = env_path
    else:
        path = resource_path(JSON_KEYFILE)

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
    return gspread.authorize(creds)


def normalize_tournament_name(name: str) -> str:
    """
    Normalizes tournament names by removing years, seasons, and specific suffixes
    to group variations of the same event.
    """
    if not name:
        return ""

    # Specific fix for StarLadder abbreviations
    name = name.replace("StarLadder SS", "StarLadder StarSeries")

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
        r'(?<!^)\b(Europe|EU|NA|SA|Asia|Americas|Oceania|CIS|European|South American|North American|Pacific|APAC)\b',
        r'\bLCQ\b',
        r'\b(Play-In|Global Finals|Contenders|CQ|Finals?|Groups?|Playoffs?)\b',
        r'\b\d+(?:st|nd|rd|th)?\s+Division\b',
        r'\bDivision\s+\d+\b',
        r'\bSeries\b',
        r'(?<!^)(?<!\bIEM\s)\b(Atlanta|Katowice|Bangkok|Raleigh|Lisbon)\b',
        r'\b(Spring|Summer|Fall|Winter)\b',
        r'\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b'
    ]
    for pattern in patterns:
        name = re.sub(pattern, '', name, flags=re.IGNORECASE)

    # 4. Remove standalone numbers (1-3 digits) that might be season/edition numbers
    # e.g. "ESL Pro League 21" -> "ESL Pro League", "UFC 302" -> "UFC"
    name = re.sub(r'\b\d{1,3}\b', '', name)

    # 5. Clean up extra whitespace and trailing hyphens
    # Replace multiple spaces with single space and trim ends
    name = re.sub(r'\s+', ' ', name).strip(' -')

    return name


def fetch_matchbet_data(spreadsheet) -> List[MatchBetTuple]:
    try:
        ws = spreadsheet.worksheet(MATCHBET_SHEET_NAME)
        try:
            raw = ws.get(MATCHBET_COL_RANGE)
        except Exception:
            raw = ws.batch_get([MATCHBET_COL_RANGE])[0]
        
        data: List[MatchBetTuple] = []
        for r in raw:
            if len(r) < 8:
                r = r + [""] * (8 - len(r))
            
            sport = r[0].strip()
            tournament = r[1].strip()
            matchup = r[2].strip()
            bet = r[3].strip()
            live_status = r[4].strip()
            odds_s = r[5].strip()
            result = r[7].strip()
            
            if not sport: continue
            if not result: continue
            
            try:
                odds = float(odds_s.replace(',', '.'))
            except ValueError:
                continue

            data.append((sport, tournament, matchup, bet, live_status, odds, result))
        return data
    except Exception as e:
        print(f"Error fetching matchbet data: {e}")
        return []


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
    def __init__(self, spreadsheet):
        super().__init__()
        self.spreadsheet = spreadsheet
        self.cache: Dict[str, SheetCacheEntry] = {}
        self.matchbet_data: List[MatchBetTuple] = []

    def run(self):
        try:
            self.progress.emit("Fetching MATCHBET data...")
            self.status.emit("Downloading all bets...")
            
            self.matchbet_data = fetch_matchbet_data(self.spreadsheet)
            self.status.emit(f"Loaded {len(self.matchbet_data)} bets. Processing...")
            
            self.cache = process_bets_to_cache(self.matchbet_data)
            
            self.progress.emit("Finalizing...")
            self.finished.emit(True, "")
        except Exception as e:
            self.finished.emit(False, str(e))


class RefreshWorker(QObject):
    finished = pyqtSignal(bool, str, object)
    def __init__(self, spreadsheet):
        super().__init__()
        self.spreadsheet = spreadsheet

    def run(self):
        try:
            data = fetch_matchbet_data(self.spreadsheet)
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


class MainWindow(QMainWindow):
    def __init__(self, spreadsheet):
        super().__init__()
        self.spreadsheet = spreadsheet
        self.setWindowTitle("Google Sheets Bet EV Viewer")
        self.resize(1200, 800)
        self.setMinimumSize(1150, 600)
        self.data_cache: Dict[str, SheetCacheEntry] = {}
        self.matchbet_data: List[MatchBetTuple] = []
        self.current_rows: List[RowTuple] = []
        self.odds_index: Dict[float, Tuple[Optional[float], Optional[float], Optional[int], Optional[int]]] = {}
        self.dark_mode = True
        self.current_view = "table"  # Track current view: "table" or "statistics"
        self._build_ui()
        self._connect()
        theme_manager.apply_theme(QApplication.instance(), dark=self.dark_mode)
        self.btn_theme.setChecked(True)
        self.btn_theme.setText(" Light Mode")

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

        # Comparison GroupBox
        self.compare_group = QGroupBox("Comparison")
        cg = QVBoxLayout(self.compare_group)
        odds_a_layout = QHBoxLayout(); odds_a_layout.addWidget(QLabel("Odds A:")); self.entry_odds_a = QLineEdit(); odds_a_layout.addWidget(self.entry_odds_a); cg.addLayout(odds_a_layout)
        odds_b_layout = QHBoxLayout(); odds_b_layout.addWidget(QLabel("Odds B:")); self.entry_odds_b = QLineEdit(); odds_b_layout.addWidget(self.entry_odds_b); cg.addLayout(odds_b_layout)
        self.btn_compare = QPushButton("Compare"); cg.addWidget(self.btn_compare)
        sidebar_layout.addWidget(self.compare_group)
        
        # Statistics Button
        self.btn_statistics = QPushButton("Statistics")
        sidebar_layout.addWidget(self.btn_statistics)
        
        # Data Table Button
        self.btn_data_table = QPushButton("Data Table")
        sidebar_layout.addWidget(self.btn_data_table)
        
        sidebar_layout.addStretch(1)

        # Bottom buttons
        self.btn_refresh = QPushButton(" Force Refresh"); self.btn_refresh.setIcon(QIcon(resource_path("icons/refresh-cw.svg"))); sidebar_layout.addWidget(self.btn_refresh)
        self.btn_theme = QPushButton(" Dark Mode"); self.btn_theme.setIcon(QIcon(resource_path("icons/sun.svg"))); self.btn_theme.setCheckable(True); self.btn_theme.setToolTip("Toggle Dark / Light theme"); sidebar_layout.addWidget(self.btn_theme)

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

        # --- Status Bar & Menu ---
        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar); self.set_status("Ready")
        act_exit = QAction("Exit", self); act_exit.triggered.connect(self.close); self.menuBar().addAction(act_exit)

    def _connect(self):
        self.btn_refresh.clicked.connect(lambda: self.refresh_data(force=True))
        self.btn_compare.clicked.connect(self.compare_by_odds)
        self.sport_combo.currentTextChanged.connect(self.on_sport_change)
        self.bettype_combo.currentTextChanged.connect(self.on_bet_type_change)
        self.entry_odds_a.textChanged.connect(self.recompute_comparison_inline)
        self.entry_odds_b.textChanged.connect(self.recompute_comparison_inline)
        self.btn_theme.toggled.connect(self.on_theme_toggled)
        self.btn_statistics.clicked.connect(self.show_statistics_panel)
        self.btn_data_table.clicked.connect(self.show_data_table_panel)

    # Helpers
    def set_status(self, text: str): self.status_bar.showMessage(text)
    def set_controls_enabled(self, enabled: bool):
        for w in [self.sport_combo,self.bettype_combo,self.btn_refresh,self.entry_odds_a,self.entry_odds_b,self.btn_compare,self.btn_theme]:
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
            ev_str, ev_val = fmt_ev(wr, odds)
            r = self.table.rowCount(); self.table.insertRow(r)
            
            # Create sortable items
            it_odds = SortableTableWidgetItem(f"{odds:.2f}", odds)
            it_live = SortableTableWidgetItem(fmt_wr(live_wr), live_wr)
            it_prem = SortableTableWidgetItem(fmt_wr(prem_wr), prem_wr)
            it_ev = SortableTableWidgetItem(ev_str, ev_val)
            
            cells = [it_odds, it_live, it_prem, it_ev]
            for c, it in enumerate(cells):
                it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(r,c,it)
            color = neu if wr is None else pos if ev_val and ev_val>0 else neg
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
            odds = float(self.table.item(r,0).text())
            def parse_cell(s:str):
                if s in ("","N/A"): return None
                return float(s.rstrip('%'))/100.0
            live_wr = parse_cell(self.table.item(r,1).text())
            prem_wr = parse_cell(self.table.item(r,2).text())
            wr = live_wr if bt=="Live" else prem_wr
            ev_str, ev_val = fmt_ev(wr, odds)
            
            it = self.table.item(r,3)
            it.setText(ev_str)
            if isinstance(it, SortableTableWidgetItem):
                it.sort_value = ev_val
            
            color = neu if wr is None else pos if ev_val and ev_val>0 else neg
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
        score_a, score_b = wr_a*odds_a, wr_b*odds_b
        if score_a>score_b: better=f"Bet {odds_a:.2f} is better"
        elif score_b>score_a: better=f"Bet {odds_b:.2f} is better"
        else: better="Both bets are equal"
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
            ev_p = (wr_p*po)-1.0; cnt_s='N/A' if cnt_p is None else str(cnt_p)
            missing = f"Odd {odds_a:.2f} not found." if miss_a else f"Odd {odds_b:.2f} not found."
            self.compare_result.setText(f"{missing}\n\nBet {po:.2f} EV: {ev_p*100:.2f}% (WR: {wr_p*100:.2f}%, Count: {cnt_s})")
            return
        wr_a_live, wr_a_pre, live_cnt_a, prem_cnt_a = self.odds_index[k_a]; wr_b_live, wr_b_pre, live_cnt_b, prem_cnt_b = self.odds_index[k_b]
        wr_a = wr_a_live if bet_type=="Live" else wr_a_pre; wr_b = wr_b_live if bet_type=="Live" else wr_b_pre
        ev_a = None if wr_a is None else (wr_a*odds_a)-1.0; ev_b = None if wr_b is None else (wr_b*odds_b)-1.0
        cnt_a = live_cnt_a if bet_type=="Live" else prem_cnt_a
        cnt_b = live_cnt_b if bet_type=="Live" else prem_cnt_b
        self.render_compare(odds_a, odds_b, wr_a, wr_b, ev_a, ev_b, self.sport_combo.currentText(), bet_type, cnt_a, cnt_b)

    def compare_by_odds(self): self.recompute_comparison_inline()
    def on_bet_type_change(self): self.update_ev_only()
    def on_sport_change(self): self.refresh_data(False)

    def on_theme_toggled(self, checked: bool):
        self.dark_mode = checked
        theme_manager.apply_theme(QApplication.instance(), dark=checked)
        self.btn_theme.setText(" Light Mode" if checked else " Dark Mode")
        self.btn_theme.setIcon(QIcon(resource_path("icons/moon.svg" if checked else "icons/sun.svg")))

    def show_statistics_panel(self):
        """Switch to statistics panel view"""
        self.current_view = "statistics"
        self.table.hide()
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
        self.statistics_panel.hide()
        self.table.show()

    def refresh_data(self, force: bool=False):
        sheet = self.sport_combo.currentText()
        # If we already have this sheet cached and not forcing, use cache instantly
        if not force and sheet in self.data_cache:
            entry = self.data_cache[sheet]
            self.current_rows = entry.rows
            self.odds_index = entry.index
            self.fill_table(self.current_rows, self.bettype_combo.currentText())
            self.recompute_comparison_inline()
            self.setWindowTitle(f"Google Sheets Bet EV Viewer - {sheet}")
            self.set_status("Loaded from cache")
            return
        # Otherwise perform an async refresh from Google Sheets
        self.set_status(f"Refreshing all data..."); self.set_controls_enabled(False)
        t=QThread(); w=RefreshWorker(self.spreadsheet); w.moveToThread(t)
        t.started.connect(w.run)
        w.finished.connect(lambda ok, info, res: self._on_refresh_done(t,w,ok,info,res))
        t.start()

    def _on_refresh_done(self, thread: QThread, worker: RefreshWorker, ok: bool, info: str, res: object):
        thread.quit(); thread.wait(); worker.deleteLater()
        if ok and res:
            cache, data = res
            self.data_cache = cache
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
            
            # Refresh view
            current_sport = self.sport_combo.currentText()
            if current_sport in self.data_cache:
                entry = self.data_cache[current_sport]
                self.current_rows=entry.rows; self.odds_index=entry.index
                self.fill_table(self.current_rows, self.bettype_combo.currentText())
                self.setWindowTitle(f"Google Sheets Bet EV Viewer - {current_sport}")
                self.recompute_comparison_inline()
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
            entry = self.data_cache[first]
            self.current_rows=entry.rows; self.odds_index=entry.index
            self.fill_table(self.current_rows, self.bettype_combo.currentText()); self.recompute_comparison_inline()

    def set_matchbet_data(self, data: List[MatchBetTuple]):
        self.matchbet_data = data


def main():
    app = QApplication(sys.argv)
    
    client = None
    spreadsheet = None
    last_error = None
    max_retries = 3

    for attempt in range(max_retries):
        try:
            client = authorize_client()
            spreadsheet = client.open(SPREADSHEET_NAME)
            break
        except Exception as e:
            last_error = e
            print(f"Connection attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)

    if spreadsheet is None:
        QMessageBox.critical(None, "Startup Error", f"Failed to initialize Google Sheets after {max_retries} attempts: {last_error}")
        return 1

    win = MainWindow(spreadsheet)
    dlg = PreloadDialog(); preload_thread = QThread(); worker = PreloadWorker(spreadsheet); worker.moveToThread(preload_thread)
    preload_thread.started.connect(worker.run)
    worker.progress.connect(dlg.update_progress); worker.status.connect(dlg.update_status)
    def done(ok: bool, msg: str):
        preload_thread.quit(); preload_thread.wait(); worker.deleteLater(); dlg.accept()
        if not ok:
            QMessageBox.critical(win, "Preload Failed", f"Failed to preload data:\n{msg}"); win.close(); return
        win.set_initial_cache(worker.cache)
        win.set_matchbet_data(worker.matchbet_data)
        win.show()
    worker.finished.connect(done)
    QTimer.singleShot(60000, lambda: done(False, "Preloading timed out") if preload_thread.isRunning() else None)
    preload_thread.start(); dlg.exec()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())