from __future__ import annotations

from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

# RowTuple now stores: (odds, live_wr, prematch_wr, live_bet_count, prematch_bet_count)
RowTuple = Tuple[float, Optional[float], Optional[float], Optional[int], Optional[int]]
MatchBetTuple = Tuple[str, str, str, str, str, float, str, str]


@dataclass
class SheetCacheEntry:
    rows: List[RowTuple]
    # index maps odds -> (live_wr, prem_wr, live_cnt, prem_cnt)
    index: Dict[float, Tuple[Optional[float], Optional[float], Optional[int], Optional[int]]]


