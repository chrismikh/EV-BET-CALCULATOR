from __future__ import annotations

import sys, os, re
from typing import Optional

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


