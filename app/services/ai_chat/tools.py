"""
AI Chat tool layer.

These are the ONLY ways the AI can get real numbers about a crawl - it
never sees raw CSVs or is asked to "read" data itself. Every tool here
runs a concrete query against the crawl's actual CSV files and returns a
small, precise JSON result. This is intentional: LLMs are unreliable at
counting/aggregating from raw text, so all counting happens in Python,
not in the model.

Reuses ISSUES and load_url_set from masterfile_overview_report.py rather
than duplicating that logic - if the issue list or CSV-matching behavior
changes there, this stays in sync automatically.
"""
import os
from app.services.masterfile_overview_report import ISSUES, load_url_set


def _find_issue(issue_label: str):
    """Word-overlap match against category+issue combined, since ISSUES
    labels are often short and ambiguous alone (e.g. "Missing" appears
    under Page Titles, Meta Description, H1, Canonicals, Hreflang, and
    Sitemaps). A caller like an LLM naturally asks for something like
    "missing meta description" - longer than the bare issue label - so a
    plain substring check (needle in issue) can never match; this instead
    scores every ISSUES entry by how many of the caller's words appear in
    that entry's "category + issue" text and returns the best match.
    Returns (category, issue, severity, priority, csvs) or None if no
    entry shares any word with the query."""
    needle_words = set(issue_label.lower().replace("-", " ").split())
    if not needle_words:
        return None
    best_match = None
    best_score = 0
    for cat, issue, sev, pri, csvs in ISSUES:
        combined_words = set((cat + " " + issue).lower().replace("-", " ").split())
        score = len(needle_words & combined_words)
        if score > best_score:
            best_score = score
            best_match = (cat, issue, sev, pri, csvs)
    return best_match


def get_crawl_summary(crawl) -> dict:
    """Basic facts about the crawl: domain, status, dates, page count."""
    return {
        "domain": crawl.domain,
        "crawl_type": crawl.crawl_type,
        "status": crawl.status,
        "pages_crawled": crawl.pages_crawled,
        "created_at": crawl.created_at.isoformat() if crawl.created_at else None,
        "completed_at": crawl.completed_at.isoformat() if crawl.completed_at else None,
        "report_path": crawl.report_path,
    }


def get_issue_count(crawl, issue_label: str) -> dict:
    """Real affected-URL count for one issue type, matched by label
    (case-insensitive substring, e.g. 'meta description missing' or just
    'missing meta')."""
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        return {"error": "Crawl report folder not found on disk."}

    match = _find_issue(issue_label)
    if not match:
        return {
            "error": f"No issue type matching '{issue_label}' found.",
            "hint": "Try list_top_issues to see available issue names.",
        }
    cat, issue, sev, pri, csvs = match
    if not csvs:
        return {
            "category": cat,
            "issue": issue,
            "severity": sev,
            "priority": pri,
            "count": None,
            "note": "This issue type has no backing CSV data in this crawl (not currently tracked).",
        }
    urls = load_url_set(crawl.report_path, csvs)
    return {
        "category": cat,
        "issue": issue,
        "severity": sev,
        "priority": pri,
        "affected_url_count": len(urls),
    }


def list_top_issues(crawl, limit: int = 10) -> dict:
    """Top N issue types ranked by affected-URL count, across all issue
    types with real CSV backing (skips issue types with no csv_files
    defined, since those have no real data to count)."""
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        return {"error": "Crawl report folder not found on disk."}

    limit = max(1, min(int(limit), 50))
    results = []
    for cat, issue, sev, pri, csvs in ISSUES:
        if not csvs:
            continue
        urls = load_url_set(crawl.report_path, csvs)
        if urls:
            results.append({
                "category": cat,
                "issue": issue,
                "severity": sev,
                "priority": pri,
                "affected_url_count": len(urls),
            })
    results.sort(key=lambda r: r["affected_url_count"], reverse=True)
    return {"top_issues": results[:limit], "total_issue_types_checked": len(results)}


def list_affected_urls(crawl, issue_label: str, limit: int = 20) -> dict:
    """Sample of actual URLs affected by a specific issue type."""
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        return {"error": "Crawl report folder not found on disk."}

    match = _find_issue(issue_label)
    if not match:
        return {
            "error": f"No issue type matching '{issue_label}' found.",
            "hint": "Try list_top_issues to see available issue names.",
        }
    cat, issue, sev, pri, csvs = match
    if not csvs:
        return {
            "category": cat,
            "issue": issue,
            "urls": [],
            "note": "This issue type has no backing CSV data in this crawl.",
        }
    limit = max(1, min(int(limit), 100))
    urls = load_url_set(crawl.report_path, csvs)
    sample = sorted(urls)[:limit]
    return {
        "category": cat,
        "issue": issue,
        "total_affected": len(urls),
        "sample_shown": len(sample),
        "urls": sample,
    }


# Provider-neutral tool schema (name, description, JSON-schema parameters).
# providers.py converts this into each provider's specific function-calling
# format (Anthropic uses "input_schema", OpenAI wraps it under "function").
TOOL_SCHEMAS = [
    {
        "name": "get_crawl_summary",
        "description": "Get basic facts about the crawl: domain, status, crawl type, dates, and total pages crawled.",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_issue_count",
        "description": "Get the real affected-URL count for one specific issue type (e.g. 'missing meta description', 'H1 missing', 'canonicals multiple').",
        "parameters": {
            "type": "object",
            "properties": {
                "issue_label": {
                    "type": "string",
                    "description": "The issue type to look up, e.g. 'missing meta description' or 'H1 duplicate'. Matched case-insensitively as a substring.",
                }
            },
            "required": ["issue_label"],
        },
    },
    {
        "name": "list_top_issues",
        "description": "List the top N issue types in this crawl ranked by number of affected URLs, across all tracked issue categories (Response Codes, Page Titles, Canonicals, etc).",
        "parameters": {
            "type": "object",
            "properties": {
                "limit": {
                    "type": "integer",
                    "description": "How many top issues to return (default 10, max 50).",
                }
            },
            "required": [],
        },
    },
    {
        "name": "list_affected_urls",
        "description": "Get a sample list of actual URLs affected by a specific issue type.",
        "parameters": {
            "type": "object",
            "properties": {
                "issue_label": {
                    "type": "string",
                    "description": "The issue type to look up, e.g. 'missing meta description'.",
                },
                "limit": {
                    "type": "integer",
                    "description": "Max number of URLs to return (default 20, max 100).",
                },
            },
            "required": ["issue_label"],
        },
    },
]

TOOL_FUNCTIONS = {
    "get_crawl_summary": get_crawl_summary,
    "get_issue_count": get_issue_count,
    "list_top_issues": list_top_issues,
    "list_affected_urls": list_affected_urls,
}


def execute_tool(name: str, crawl, arguments: dict) -> dict:
    """Dispatch a tool call by name to its implementation, passing the
    bound crawl object plus whatever arguments the model supplied."""
    fn = TOOL_FUNCTIONS.get(name)
    if not fn:
        return {"error": f"Unknown tool: {name}"}
    try:
        return fn(crawl, **arguments)
    except TypeError as e:
        return {"error": f"Invalid arguments for {name}: {e}"}
    except Exception as e:
        return {"error": f"Tool {name} failed: {e}"}
