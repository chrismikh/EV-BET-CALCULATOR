"""
DatabaseManager – SQLite backend for the EV Bet Calculator.

Stores all bet records locally, replacing the previous Google-Sheets dependency.
"""

from __future__ import annotations

import os
import sqlite3
from contextlib import contextmanager
from typing import List, Optional, Tuple


# Default database path: in the same directory as this file (src/)
_DEFAULT_DB_DIR = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_DB_PATH = os.path.join(_DEFAULT_DB_DIR, "matchbet.db")


class DatabaseManager:
    """Manages all SQLite interactions for the bet database."""

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = db_path or _DEFAULT_DB_PATH

    # ------------------------------------------------------------------
    # Connection helpers
    # ------------------------------------------------------------------
    @contextmanager
    def get_connection(self):
        """Yield a sqlite3 Connection inside a managed context."""
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA journal_mode=WAL")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # ------------------------------------------------------------------
    # Schema
    # ------------------------------------------------------------------
    def create_tables(self) -> None:
        """Create the bets table if it does not already exist.

        Also handles migration from the old schema that included a game_map column.
        """
        with self.get_connection() as conn:
            # Check if table already exists with old game_map column
            cur = conn.execute("PRAGMA table_info(bets)")
            columns = [row[1] for row in cur.fetchall()]
            if columns and "game_map" in columns:
                # Migrate: recreate table without game_map
                conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS bets_new (
                        id            INTEGER PRIMARY KEY AUTOINCREMENT,
                        sport         TEXT    NOT NULL,
                        tournament    TEXT    NOT NULL,
                        matchup       TEXT    NOT NULL,
                        bet           TEXT    NOT NULL,
                        live_status   TEXT    NOT NULL,
                        odds          REAL    NOT NULL,
                        bet_amount    REAL,
                        result        TEXT,
                        profit        REAL,
                        date_created  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                    """
                )
                conn.execute(
                    """
                    INSERT INTO bets_new (id, sport, tournament, matchup, bet,
                                          live_status, odds, bet_amount, result,
                                          profit, date_created)
                    SELECT id, sport, tournament, matchup, bet,
                           live_status, odds, bet_amount, result,
                           profit, date_created
                    FROM bets
                    """
                )
                conn.execute("DROP TABLE bets")
                conn.execute("ALTER TABLE bets_new RENAME TO bets")
                return

            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS bets (
                    id            INTEGER PRIMARY KEY AUTOINCREMENT,
                    sport         TEXT    NOT NULL,
                    tournament    TEXT    NOT NULL,
                    matchup       TEXT    NOT NULL,
                    bet           TEXT    NOT NULL,
                    live_status   TEXT    NOT NULL,
                    odds          REAL    NOT NULL,
                    bet_amount    REAL,
                    result        TEXT,
                    profit        REAL,
                    date_created  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )

    # ------------------------------------------------------------------
    # CRUD helpers
    # ------------------------------------------------------------------
    def insert_bet(
        self,
        sport: str,
        tournament: str,
        matchup: str,
        bet: str,
        live_status: str,
        odds: float,
        bet_amount: Optional[float] = None,
        result: Optional[str] = None,
        profit: Optional[float] = None,
    ) -> int:
        """Insert a new bet row and return its id."""
        with self.get_connection() as conn:
            cur = conn.execute(
                """
                INSERT INTO bets (sport, tournament, matchup, bet,
                                  live_status, odds, bet_amount, result, profit)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sport,
                    tournament,
                    matchup,
                    bet,
                    live_status,
                    odds,
                    bet_amount,
                    result,
                    profit,
                ),
            )
            return cur.lastrowid  # type: ignore[return-value]

    def update_bet(self, bet_id: int, result: str, profit: float) -> None:
        """Update result and profit for a specific bet."""
        with self.get_connection() as conn:
            conn.execute(
                "UPDATE bets SET result = ?, profit = ? WHERE id = ?",
                (result, profit, bet_id),
            )

    # ------------------------------------------------------------------
    # Query helpers
    # ------------------------------------------------------------------
    def fetch_all_bets(self) -> List[Tuple]:
        """Return every row in the bets table (all columns)."""
        with self.get_connection() as conn:
            cur = conn.execute(
                "SELECT id, sport, tournament, matchup, bet, "
                "live_status, odds, bet_amount, result, profit, date_created "
                "FROM bets ORDER BY id"
            )
            return cur.fetchall()

    def fetch_settled_bets(self) -> List[Tuple[str, str, str, str, str, float, str]]:
        """
        Return settled bets (result IS NOT NULL AND result != '') in the
        same MatchBetTuple shape used by the rest of the app:
        (sport, tournament, matchup, bet, live_status, odds, result)
        """
        with self.get_connection() as conn:
            cur = conn.execute(
                "SELECT sport, tournament, matchup, bet, live_status, odds, result "
                "FROM bets "
                "WHERE result IS NOT NULL AND result != '' "
                "ORDER BY id"
            )
            return cur.fetchall()

    def fetch_pending_bets(self) -> List[Tuple]:
        """
        Return bets where the result has not been settled yet.
        Each row: (id, sport, tournament, matchup, bet,
                   live_status, odds, bet_amount, result, profit, date_created)
        """
        with self.get_connection() as conn:
            cur = conn.execute(
                "SELECT id, sport, tournament, matchup, bet, "
                "live_status, odds, bet_amount, result, profit, date_created "
                "FROM bets "
                "WHERE result IS NULL OR result = '' "
                "ORDER BY id DESC"
            )
            return cur.fetchall()

    def get_distinct_sports(self) -> List[str]:
        """Return a sorted list of unique sport names."""
        with self.get_connection() as conn:
            cur = conn.execute(
                "SELECT DISTINCT sport FROM bets ORDER BY sport"
            )
            return [row[0] for row in cur.fetchall()]

    def delete_all_bets(self) -> int:
        """Delete every row from the bets table. Returns the number of rows deleted."""
        with self.get_connection() as conn:
            cur = conn.execute("DELETE FROM bets")
            return cur.rowcount
