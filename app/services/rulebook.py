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


def load_rulebook(domain: str) -> dict:
    path = get_rulebook_path(domain)
    if not os.path.exists(path):
        return None
    try:
        df = pd.read_excel(path, sheet_name="Rulebook", header=1)
        df = df.dropna(subset=["URL Pattern"])

        rules = []
        fallback = {"theme1": "Others", "theme2": "", "language": "", "priority": "Low"}

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
            priority = str(row.get("Priority Weightage", "Low")).strip()
            if not priority or priority == "nan":
                priority = "Low"

            if not pattern or pattern == "nan":
                continue

            # Detect fallback row
            if rule_type in ("—", "-", "Fallback") or "no match" in pattern.lower():
                fallback = {
                    "theme1": theme1 or "Others",
                    "theme2": theme2 if theme2 and theme2 != "nan" else "",
                    "language": language,
                    "priority": priority or "Low",
                }
                continue

            rules.append({
                "rule_type": rule_type,
                "pattern": pattern,
                "theme1": theme1,
                "theme2": theme2 if theme2 and theme2 != "nan" else "",
                "language": language,
                "priority": priority or "Low",
            })

        return {"rules": rules, "fallback": fallback}

    except Exception as e:
        print(f"Error loading rulebook for {domain}: {e}")
        return None


def classify_url(url: str, rulebook: dict) -> tuple:
    """Returns (theme1, theme2, language, priority)"""
    if not rulebook:
        return "", "", "", "Low"

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
            elif rule_type in ("exact match", "exactmatch", "exact"):
                matched = url.lower() == pattern.lower()
            elif rule_type == "regex":
                matched = bool(re.search(pattern, url))
        except Exception:
            pass
        if matched:
            return rule["theme1"], rule["theme2"], rule["language"], rule["priority"]

    fb = rulebook["fallback"]
    return fb["theme1"], fb["theme2"], fb["language"], fb["priority"]