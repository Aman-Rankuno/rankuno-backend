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
        page_type_df = pd.read_excel(path, sheet_name="Page Type Rules", header=2)
        page_type_df = page_type_df.dropna(subset=["Rule Type", "Rule / Pattern", "Page Type"])
        page_type_rules = []
        fallback_page_type = "Others"
        for _, row in page_type_df.iterrows():
            rule_type = str(row["Rule Type"]).strip()
            pattern = str(row["Rule / Pattern"]).strip()
            page_type = str(row["Page Type"]).strip()
            if rule_type.lower() == "fallback":
                fallback_page_type = page_type
            else:
                page_type_rules.append({
                    "rule_type": rule_type,
                    "pattern": pattern,
                    "page_type": page_type,
                })

        language_df = pd.read_excel(path, sheet_name="Language Rules", header=2)
        language_df = language_df.dropna(subset=["URL Pattern", "Language"])
        language_rules = []
        for _, row in language_df.iterrows():
            language_rules.append({
                "pattern": str(row["URL Pattern"]).strip(),
                "language": str(row["Language"]).strip(),
            })

        return {
            "page_type_rules": page_type_rules,
            "fallback_page_type": fallback_page_type,
            "language_rules": language_rules,
        }
    except Exception as e:
        print(f"Error loading rulebook for {domain}: {e}")
        return None


def classify_url(url: str, rulebook: dict) -> tuple:
    if not rulebook:
        return "", ""

    page_type = rulebook["fallback_page_type"]
    for rule in rulebook["page_type_rules"]:
        rule_type = rule["rule_type"].lower()
        pattern = rule["pattern"]
        matched = False
        try:
            if rule_type == "contains":
                matched = pattern.lower() in url.lower()
            elif rule_type == "starts with":
                matched = url.lower().startswith(pattern.lower())
            elif rule_type == "ends with":
                matched = url.lower().endswith(pattern.lower())
            elif rule_type == "exact match":
                matched = url.lower() == pattern.lower()
            elif rule_type == "regex":
                matched = bool(re.search(pattern, url))
        except Exception:
            pass
        if matched:
            page_type = rule["page_type"]
            break

    language = "Unknown"
    for rule in rulebook["language_rules"]:
        if rule["pattern"].lower() in url.lower():
            language = rule["language"]
            break

    return page_type, language