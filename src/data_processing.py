from __future__ import annotations

from typing import List, Dict

try:
    from src.database import DatabaseManager
except ModuleNotFoundError:
    from database import DatabaseManager

try:
    from src.models import MatchBetTuple, SheetCacheEntry
except ModuleNotFoundError:
    from models import MatchBetTuple, SheetCacheEntry

try:
    from src.utils import round_odds_key
except ModuleNotFoundError:
    from utils import round_odds_key

def fetch_matchbet_data_from_db(db: DatabaseManager, force_refresh_network: bool = False) -> List[MatchBetTuple]:
    """Fetch settled bets from the database.

    Returns data in the same MatchBetTuple format used elsewhere:
    (sport, tournament, matchup, bet, live_status, odds, result, provider)
    """
    try:
        return db.fetch_settled_bets(force_refresh=force_refresh_network)
    except Exception as e:
        print(f"Error fetching data from database: {e}")
        return []


def process_bets_to_cache(bets: List[MatchBetTuple]) -> Dict[str, SheetCacheEntry]:
    agg = {}
    for sport, _, _, _, live_status, odds, result, _provider in bets:
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


