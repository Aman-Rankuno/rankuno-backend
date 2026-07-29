import os
import re
import pandas as pd
from urllib.parse import urlparse
from app.config import settings


def get_rulebook_path(domain: str) -> str:
    parsed = urlparse(domain)
    hostname = parsed.hostname or domain
    hostname = hostname.replace("www.", "", 1)
    return os.path.join(settings.RULEBOOKS_DIR, f"{hostname}.xlsx")


def _find_header_row(path: str, sheet_name: str) -> int:
    """Auto-detect which row actually holds the real column headers,
    rather than assuming a fixed row index. Looks for a row containing
    the literal text "URL Pattern" in any cell, scanning the first 5 rows.
    Returns the 0-indexed row to pass as pandas' header= argument, or 0
    if nothing is found (safest fallback: assume no offset rather than
    guessing an offset that might skip real data)."""
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=5)
    for i in range(len(raw)):
        row_values = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if "url pattern" in row_values:
            return i
    return 0


def load_rulebook(domain: str) -> dict:
    path = get_rulebook_path(domain)
    if not os.path.exists(path):
        return None
    try:
        header_row = _find_header_row(path, "Rulebook")
        df = pd.read_excel(path, sheet_name="Rulebook", header=header_row)
        df = df.dropna(subset=["URL Pattern"])

        rules = []
        fallback = {"theme1": "Others", "theme2": "", "language": "", "priority": "N/A"}

        for _, row in df.iterrows():
            # Rule Type column may be missing or empty - default to Contains
            rule_type = str(row.get("Rule Type", "Contains")).strip()
            if not rule_type or rule_type == "nan":
                rule_type = "Contains"

            pattern = str(row.get("URL Pattern", "")).strip()
            theme1 = str(row.get("Theme Name 1", "")).strip()
            theme2 = str(row.get("Theme Name 2", "")).strip()

            # Handle multiple possible Language column spellings
            language = ""
            for lang_col in ["Language", "Languuage", "Lang"]:
                val = row.get(lang_col, "")
                if val and str(val) != "nan":
                    language = str(val).strip()
                    break

            # Priority hardcoded per issue type - rulebook priority only used for Table 2 sorting
            priority = str(row.get("Priority", "")).strip()
            if not priority or priority == "nan":
                priority = "N/A"

            if not pattern or pattern == "nan":
                continue

            # Detect fallback row. \u2014 is the em dash (—) written as an
            # explicit Unicode escape so it can never be corrupted again by
            # an editor guessing the wrong file encoding (this line was
            # previously mangled into mojibake, breaking fallback detection
            # for any rulebook using an em dash as its fallback marker).
            if rule_type in ("\u2014", "-", "Fallback") or "no match" in pattern.lower():
                fallback = {
                    "theme1": theme1 or "Others",
                    "theme2": theme2 if theme2 and theme2 != "nan" else "",
                    "language": language,
                    "priority": priority or "N/A",
                }
                continue

            rules.append({
                "rule_type": rule_type,
                "pattern": pattern,
                "theme1": theme1,
                "theme2": theme2 if theme2 and theme2 != "nan" else "",
                "language": language,
                "priority": priority or "N/A",
            })

        return {"rules": rules, "fallback": fallback}

    except Exception as e:
        print(f"Error loading rulebook for {domain}: {e}")
        return None


def classify_url(url: str, rulebook: dict) -> tuple:
    """Returns (theme1, theme2, language, priority)

    Picks the MOST SPECIFIC matching rule, not the first one in file
    order. Rulebooks commonly have a broad parent rule (e.g. Contains
    "/services") followed by many narrower sub-path rules (e.g. Contains
    "/services/annotation-services-for-ai-ml") that carry real Theme 2
    data the parent rule doesn't have. "First match wins" silently picks
    the broad parent rule for every sub-path URL, since it's listed
    first and Contains-matches just as well - Theme 1 still comes out
    right (parent and child usually share it), but Theme 2 is always
    lost. Fixed by collecting every matching rule and picking the one
    with the longest pattern, since a longer matched substring reliably
    indicates a more specific rule regardless of row order. Exact Match
    rules are prioritized over Contains/StartsWith/EndsWith when both
    match, since an exact match is maximally specific by definition.
    """
    if not rulebook:
        return "", "", "", "Low"

    EXACT_TYPES = {"exact match", "exactmatch", "exact"}
    matches = []  # (is_exact, pattern_length, rule)

    for rule in rulebook["rules"]:
        rule_type = rule["rule_type"].strip().lower()
        pattern = rule["pattern"]
        matched = False
        try:
            if rule_type == "contains":
                matched = pattern.lower() in url.lower()
            elif rule_type in ("starts with", "startswith"):
                matched = url.lower().startswith(pattern.lower())
            elif rule_type in ("ends with", "endswith"):
                matched = url.lower().endswith(pattern.lower())
            elif rule_type in EXACT_TYPES:
                matched = url.lower() == pattern.lower()
            elif rule_type == "regex":
                matched = bool(re.search(pattern, url))
        except Exception:
            pass
        if matched:
            matches.append((rule_type in EXACT_TYPES, len(pattern), rule))

    if matches:
        # Highest specificity wins: exact match first, then longest pattern.
        matches.sort(key=lambda m: (m[0], m[1]), reverse=True)
        best_rule = matches[0][2]
        return best_rule["theme1"], best_rule["theme2"], best_rule["language"], best_rule["priority"]

    fb = rulebook["fallback"]
    return fb["theme1"], fb["theme2"], fb["language"], fb["priority"]
