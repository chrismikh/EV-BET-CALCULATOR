from __future__ import annotations

import re
from typing import List

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

    # Map Esports World Cup / EWC variants to a single label
    if re.search(r'\b(Esports\s*World\s*Cup|EWC)\b', name, re.IGNORECASE):
        return "EWC"

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
        return "VCT GC"

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

    # R6S BLAST Major events – group all host-city variants together
    if re.match(r'^BLAST\s+R6\s+Major\b', name, re.IGNORECASE):
        return "BLAST R6 Major"

    # Thunderpick / TP World Championship / TP WC – unify abbreviations
    if re.match(r'^(Thunderpick|TP)\s+(World\s+Championship|WC)\b', name, re.IGNORECASE):
        return "Thunderpick World Championship"

    # ESEA – group all variants (Advanced, divisions, etc.)
    if re.match(r'^ESEA\b', name):
        return "ESEA"

    # European Pro League – group all variants (Regular Season, Series, Division, etc.)
    if re.match(r'^European Pro League\b', name):
        return "European Pro League"

    # LCK – group all variants (Season, Challengers, etc.)
    if re.match(r'^LCK\b', name, re.IGNORECASE):
        return "LCK"

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


