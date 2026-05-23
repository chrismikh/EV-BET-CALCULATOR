"""
DatabaseManager – Firebase Firestore backend for the EV Bet Calculator.

Stores all bet records to the Cloud using Firebase REST API.
"""

from __future__ import annotations

import os
import time
import requests
from typing import List, Optional, Tuple, Dict, Any
from datetime import datetime

try:
    from src.firebase_manager import FirebaseAuthManager
except ModuleNotFoundError:
    from firebase_manager import FirebaseAuthManager

# Default client secrets path
_DEFAULT_DIR = os.path.dirname(os.path.abspath(__file__))
_SECRETS_PATH = os.path.join(_DEFAULT_DIR, "client_secrets.json")

_GLOBAL_AUTH = None

class DatabaseManager:
    """Manages all Firestore interactions for the bet database."""

    def __init__(self, db_path: Optional[str] = None):
        global _GLOBAL_AUTH
        if _GLOBAL_AUTH is None:
            _GLOBAL_AUTH = FirebaseAuthManager(_SECRETS_PATH)
            try:
                print("Starting Google Login...")
                _GLOBAL_AUTH.login()
                print(f"Logged in as {_GLOBAL_AUTH.email}")
            except Exception as e:
                print(f"Failed to login: {e}")
        
        self.auth = _GLOBAL_AUTH
        self._cached_docs: Optional[List[Dict[str, Any]]] = None
        self._cache_timestamp: float = 0
        self._cache_ttl: float = 14400

    def _request_with_backoff(self, method: str, url: str, **kwargs):
        for attempt in range(5):
            res = requests.request(method, url, headers=self.auth.auth_headers, **kwargs)
            if res.status_code in [429, 503]:
                if attempt == 4:
                    res.raise_for_status()
                time.sleep((2 ** attempt) + 0.5)
            else:
                res.raise_for_status()
                return res

    # ------------------------------------------------------------------
    # Core Converters to map Python Types <-> Firestore JSON 
    # ------------------------------------------------------------------
    def _to_firestore_fields(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Convert a standard dictionary to the verbose Firestore JSON shape."""
        fields = {}
        for k, v in data.items():
            if isinstance(v, str):
                fields[k] = {"stringValue": v}
            elif isinstance(v, bool):
                fields[k] = {"booleanValue": v}
            elif isinstance(v, int):
                fields[k] = {"integerValue": str(v)}
            elif isinstance(v, float):
                # Firestore sometimes converts ending floats to integers, handle safely
                fields[k] = {"doubleValue": float(v)}
            elif v is None:
                fields[k] = {"nullValue": None}
        return {"fields": fields}

    def _from_firestore_fields(self, doc_name: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        """Convert the verbose Firestore JSON shape back to a standard dictionary."""
        res = {"id": doc_name.split("/")[-1]} # Extract document ID from its path
        for k, v in fields.items():
            if "stringValue" in v:
                res[k] = v["stringValue"]
            elif "doubleValue" in v:
                res[k] = float(v["doubleValue"])
            elif "integerValue" in v:
                res[k] = int(v["integerValue"])
            elif "booleanValue" in v:
                res[k] = bool(v["booleanValue"])
            elif "nullValue" in v:
                res[k] = None
        return res

    # ------------------------------------------------------------------
    # Schema
    # ------------------------------------------------------------------
    def create_tables(self) -> None:
        """NoSchema needed for Firestore, so this is a no-op."""
        pass

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
    ) -> str:
        """Insert a new bet row and return its custom string id."""
        # Clean nulls
        if result is None: result = ""
        if profit is None: profit = 0.0
        if bet_amount is None: bet_amount = 0.0

        data = {
            "sport": sport,
            "tournament": tournament,
            "matchup": matchup,
            "bet": bet,
            "live_status": live_status,
            "odds": odds,
            "bet_amount": bet_amount,
            "result": result,
            "profit": profit,
            "date_created": datetime.utcnow().isoformat()
        }
        
        payload = self._to_firestore_fields(data)
        url = self.auth.get_bets_url()
        
        res = self._request_with_backoff("POST", url, json=payload)
        
        # API returns the created document metadata
        doc = res.json()
        doc_name = doc.get("name", "")
        
        # Update cache directly instead of invalidating to prevent 429 errors
        if self._cached_docs is not None:
            new_doc = self._from_firestore_fields(doc_name, doc.get("fields", {}))
            self._cached_docs.append(new_doc)
            
        return doc_name.split("/")[-1]

    def update_bet(self, bet_id: str, result: str, profit: float) -> None:
        """Update result and profit for a specific bet."""
        # Using PATCH requires appending ?updateMask.fieldPaths=... for each field
        url = self.auth.get_bets_url(bet_id) + "?updateMask.fieldPaths=result&updateMask.fieldPaths=profit"
        
        data = {"result": result, "profit": float(profit)}
        payload = self._to_firestore_fields(data)
        
        self._request_with_backoff("PATCH", url, json=payload)
        
        # Update cache directly instead of invalidating
        if self._cached_docs is not None:
            for doc in self._cached_docs:
                if doc.get("id") == bet_id:
                    doc["result"] = result
                    doc["profit"] = float(profit)
                    break

    def update_bet_fields(self, bet_id: str, fields: Dict[str, Any]) -> None:
        """Update arbitrary fields for a specific bet (used for editing pending bets).

        `fields` is a map of key->value where keys are top-level bet fields
        (sport, tournament, matchup, bet, live_status, odds, bet_amount, etc.).
        This method constructs the appropriate updateMask query and issues a
        PATCH request to Firestore then applies changes to the local cache.
        """
        if not fields:
            return

        # Build updateMask parameters
        mask_parts = []
        for k in fields.keys():
            mask_parts.append(f"updateMask.fieldPaths={k}")

        url = self.auth.get_bets_url(bet_id)
        if mask_parts:
            url = url + "?" + "&".join(mask_parts)

        payload = self._to_firestore_fields(fields)
        self._request_with_backoff("PATCH", url, json=payload)

        # Update cache directly
        if self._cached_docs is not None:
            for doc in self._cached_docs:
                if doc.get("id") == bet_id:
                    for k, v in fields.items():
                        doc[k] = v
                    break

    # ------------------------------------------------------------------
    # Query helpers
    # ------------------------------------------------------------------
    def _fetch_docs_as_dicts(self, force_refresh: bool = False) -> List[Dict[str, Any]]:
        """Helper to get all user bet documents parsed into dicts."""
        cache_expired = (time.time() - self._cache_timestamp) > self._cache_ttl
        if self._cached_docs is not None and not force_refresh and not cache_expired:
            return self._cached_docs

        base_url = self.auth.get_bets_url() + "?pageSize=1000"
        all_docs = []
        page_token = None
        
        while True:
            url = base_url
            if page_token:
                url += f"&pageToken={page_token}"
                
            try:
                res = self._request_with_backoff("GET", url)
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    break
                raise
            
            data = res.json()
            docs = data.get("documents", [])
            all_docs.extend(docs)
            
            page_token = data.get("nextPageToken")
            if not page_token:
                break
                
        self._cached_docs = [self._from_firestore_fields(d.get("name", ""), d.get("fields", {})) for d in all_docs]
        self._cache_timestamp = time.time()
        return self._cached_docs

    def fetch_all_bets(self, force_refresh: bool = False) -> List[Tuple]:
        """Return every row in the bets table (all columns) mapped to SQLite tuple structure."""
        docs = self._fetch_docs_as_dicts(force_refresh=force_refresh)
        # Ensure ordering
        docs.sort(key=lambda x: x.get("date_created", ""))
        
        # Same shape as fetch_all_bets in SQLite
        return [
            (
                d["id"],
                d.get("sport", ""),
                d.get("tournament", ""),
                d.get("matchup", ""),
                d.get("bet", ""),
                d.get("live_status", ""),
                d.get("odds", 0.0),
                d.get("bet_amount", 0.0),
                d.get("result", ""),
                d.get("profit", 0.0),
                d.get("date_created", "")
            )
            for d in docs
        ]

    def fetch_settled_bets(self, force_refresh: bool = False) -> List[Tuple[str, str, str, str, str, float, str]]:
        docs = self._fetch_docs_as_dicts(force_refresh=force_refresh)
        settled = [d for d in docs if d.get("result")]
        settled.sort(key=lambda x: x.get("date_created", ""))
        
        return [
            (
                d.get("sport", ""),
                d.get("tournament", ""),
                d.get("matchup", ""),
                d.get("bet", ""),
                d.get("live_status", ""),
                float(d.get("odds", 0.0)),
                d.get("result", "")
            )
            for d in settled
        ]

    def fetch_pending_bets(self, force_refresh: bool = False) -> List[Tuple]:
        docs = self._fetch_docs_as_dicts(force_refresh=force_refresh)
        pending = [d for d in docs if not d.get("result")]
        # Sort descending (similar to order by id desc)
        pending.sort(key=lambda x: x.get("date_created", ""), reverse=True)
        
        return [
            (
                d["id"],
                d.get("sport", ""),
                d.get("tournament", ""),
                d.get("matchup", ""),
                d.get("bet", ""),
                d.get("live_status", ""),
                d.get("odds", 0.0),
                d.get("bet_amount", 0.0),
                d.get("result", ""),
                d.get("profit", 0.0),
                d.get("date_created", "")
            )
            for d in pending
        ]

    def get_distinct_sports(self) -> List[str]:
        docs = self._fetch_docs_as_dicts()
        sports = {d.get("sport", "") for d in docs if d.get("sport")}
        return sorted(list(sports))

    def delete_all_bets(self) -> int:
        docs = self._fetch_docs_as_dicts(force_refresh=True)
        for i, d in enumerate(docs):
            url = self.auth.get_bets_url(d["id"])
            self._request_with_backoff("DELETE", url)
            if (i + 1) % 20 == 0:
                time.sleep(0.5)
        self._cached_docs = []
        return len(docs)
