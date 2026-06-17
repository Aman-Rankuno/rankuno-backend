import os
import io
import math
import pandas as pd
import xlsxwriter
from collections import defaultdict
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}
FONT = "Rockwell"
RED = "#FF0000"
WHITE = "#FFFFFF"
BLACK = "#000000"

ISSUE_CONFIGS = [
    {"csv": "canonicals_canonicalised.csv",        "label": "Canonicalised",            "priority": "High",   "severity": "Issue"},
    {"csv": "canonicals_missing.csv",               "label": "Missing",                  "priority": "High",   "severity": "Issue"},
    {"csv": "canonicals_multiple.csv",              "label": "Multiple",                 "priority": "High",   "severity": "Issue"},
    {"csv": "canonicals_nonindexable_canonical.csv","label": "Non-Indexable Canonical",  "priority": "High",   "severity": "Issue"},
    {"csv": "canonicals_canonical_is_relative.csv", "label": "Relative Canonical",       "priority": "High",   "severity": "Warning"},
    {"csv": "canonicals_unlinked.csv",              "label": "Unlinked",                 "priority": "Medium", "severity": "Warning"},
    {"csv": "canonicals_multiple_conflicting.csv",  "label": "Multiple Conflicting",     "priority": "High",   "severity": "Issue"},
    {"csv": "canonicals_outside_head.csv",          "label": "Canonical Outside <Head>", "priority": "High",   "severity": "Issue"},
]

ISSUE_LABELS = [c["label"] for c in ISSUE_CONFIGS]
ISSUE_PRIORITIES = {c["label"]: c["priority"] for c in ISSUE_CONFIGS}

ISSUE_SUMMARY = (
    "Issue Summary:\n"
    "1. Canonicalised - URLs that are canonicalised to another URL, indicating search engines may not index the reported URL.\n"
    "2. Missing - URLs missing a rel=\"canonical\" tag, making it unclear which version should be treated as preferred.\n"
    "3. Multiple - URLs containing multiple rel=\"canonical\" tags, sending conflicting signals to search engines.\n"
    "4. Non-Indexable Canonical - URLs pointing their canonical tag to a non-indexable page.\n"
    "5. Relative Canonical - URLs using relative paths in the canonical tag instead of absolute URLs.\n"
    "6. Canonical Outside <Head> - URLs where the canonical tag is placed outside the <head> section.\n"
    "7. Unlinked - URLs discovered only through canonical annotations and not linked internally.\n"
    "8. Multiple Conflicting Canonicals - URLs containing multiple canonical declarations referencing different targets."
)


def safe_str(v):
    if v is None:
        return ""
    s = str(v)
    return "" if s in ("nan", "None", "NaN") else s


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else f
    except Exception:
        return None


def _gc(df, *names):
    for n in names:
        m = next((c for c in df.columns if c.lower() == n.lower()), None)
        if m:
            return m
    return None


def load_df(report_path, filename):
    p = os.path.join(report_path, filename)
    if os.path.exists(p) and os.path.getsize(p) > 10:
        try:
            return pd.read_csv(p, encoding="utf-8", low_memory=False)
        except Exception:
            pass
    return pd.DataFrame()


def build_canonicals_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "Canonical Analysis.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    rulebook = load_rulebook(domain)

    # Load internal_all for enrichment
    df_int = load_df(report_path, "internal_all.csv")
    internal_map = {}
    total_crawled = 0
    if not df_int.empty:
        total_crawled = len(df_int)
        a_col = _gc(df_int, "Address")
        pt_col = _gc(df_int, "Content Type", "Page Type")
        sc_col = _gc(df_int, "Status Code")
        idx_col = _gc(df_int, "Indexability")
        inlink_col = _gc(df_int, "Inlinks")
        can1_col = _gc(df_int, "Canonical Link Element 1")
        can2_col = _gc(df_int, "Canonical Link Element 2")
        for _, r in df_int.iterrows():
            url = safe_str(r.get(a_col, "")) if a_col else ""
            internal_map[url] = {
                "page_type": safe_str(r.get(pt_col, "")) if pt_col else "",
                "status_code": safe_str(r.get(sc_col, "")) if sc_col else "",
                "indexability": safe_str(r.get(idx_col, "")) if idx_col else "",
                "inlinks": safe_num(r.get(inlink_col, 0)) if inlink_col else None,
                "can1": safe_str(r.get(can1_col, "")) if can1_col else "",
                "can2": safe_str(r.get(can2_col, "")) if can2_col else "",
            }

    # Load GSC and GA4
    gsc_map = {}
    df_gsc = load_df(report_path, "search_console_all.csv")
    if not df_gsc.empty:
        a = _gc(df_gsc, "Address")
        imp = next((c for c in df_gsc.columns if "impression" in c.lower()), None)
        clk = next((c for c in df_gsc.columns if "click" in c.lower()), None)
        if a:
            for _, r in df_gsc.iterrows():
                gsc_map[safe_str(r[a])] = {
                    "impressions": safe_num(r.get(imp, 0)) if imp else None,
                    "clicks": safe_num(r.get(clk, 0)) if clk else None,
                }

    ga_map = {}
    df_ga = load_df(report_path, "analytics_all.csv")
    if not df_ga.empty:
        a = _gc(df_ga, "Address")
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[safe_str(r[a])] = safe_num(r.get(s, 0))

    # Process all issue CSVs
    all_rows = []
    issue_url_sets = {lbl: set() for lbl in ISSUE_LABELS}

    for cfg in ISSUE_CONFIGS:
        df = load_df(report_path, cfg["csv"])
        label = cfg["label"]
        if df.empty:
            continue
        addr_col = _gc(df, "Address")
        occ_col = _gc(df, "Occurrences")
        can1_csv = _gc(df, "Canonical Link Element 1")
        if not addr_col:
            continue
        for _, row in df.iterrows():
            url = safe_str(row.get(addr_col, ""))
            if not url:
                continue
            info = internal_map.get(url, {})
            # Filter: indexable, 200, HTML
            if info.get("indexability", "").lower() != "indexable":
                continue
            if info.get("status_code", "") != "200":
                continue
            if "html" not in info.get("page_type", "").lower():
                continue

            theme1, theme2, _, priority = classify_url(url, rulebook)
            gsc = gsc_map.get(url, {})
            can1 = safe_str(row.get(can1_csv, "")) if can1_csv else info.get("can1", "")
            can2 = info.get("can2", "")

            # Canonical 1 enrichment
            can1_info = internal_map.get(can1, {})
            can1_sc = can1_info.get("status_code", "-") if can1 else "-"
            can1_idx = can1_info.get("indexability", "-") if can1 else "-"
            can1_inlinks = can1_info.get("inlinks") if can1 else None

            # Canonical 2 enrichment
            can2_info = internal_map.get(can2, {})
            can2_sc = can2_info.get("status_code", "-") if can2 else "-"
            can2_idx = can2_info.get("indexability", "-") if can2 else "-"
            can2_inlinks = can2_info.get("inlinks") if can2 else None

            issue_url_sets[label].add(url)
            all_rows.append({
                "error_type": label,
                "address": url,
                "theme1": theme1 or "-",
                "theme2": theme2 or "-",
                "page_type": info.get("page_type", ""),
                "status_code": info.get("status_code", "200"),
                "indexability": info.get("indexability", "Indexable"),
                "occurrence": safe_num(row.get(occ_col, "")) if occ_col else None,
                "can1": can1 or "-",
                "can1_sc": can1_sc,
                "can1_idx": can1_idx,
                "inlinks": info.get("inlinks"),
                "can2": can2 or "-",
                "can2_sc": can2_sc,
                "can2_idx": can2_idx,
                "can2_inlinks": can2_inlinks,
                "impressions": gsc.get("impressions"),
                "clicks": gsc.get("clicks"),
                "sessions": ga_map.get(url),
                "priority": priority or "N/A",
            })

    # Sort Table 3 by impressions descending
    all_rows.sort(key=lambda x: (x.get("impressions") or 0), reverse=True)

    # Table 2: theme counts per issue type
    theme_counts = defaultdict(lambda: defaultdict(int))
    theme_total = defaultdict(int)
    theme_priority_map = {}
    for r in all_rows:
        theme = r["theme1"] if r["theme1"] != "-" else "Others"
        theme_counts[theme][r["error_type"]] += 1
        theme_total[theme] += 1
        if theme not in theme_priority_map:
            theme_priority_map[theme] = r["priority"]

    sorted_themes = sorted(theme_priority_map.items(),
                           key=lambda x: PRIORITY_ORDER.get(x[1], 3))

    # Build workbook
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})

    def f(**kw):
        return wb.add_format(kw)

    f_summary = f(font_name=FONT, font_size=8, font_color=BLACK, text_wrap=True, valign="top", border=1)
    f_red_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
    f_red_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED, border=1, align="left", valign="vcenter", text_wrap=True)
    f_lbl = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK)
    f_col_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", text_wrap=True)
    f_col_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="left", valign="vcenter", text_wrap=True)
    f_bold_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_cell = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_cell_lft = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0.00%")
    f_rgt = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="right", valign="vcenter")

    ws = wb.add_worksheet("Canonical")
    ws.set_column("A:A", 60); ws.set_column("B:B", 20); ws.set_column("C:D", 20)
    ws.set_column("E:E", 25); ws.set_column("F:G", 15); ws.set_column("H:H", 12)
    ws.set_column("I:I", 50); ws.set_column("J:K", 20); ws.set_column("L:L", 12)
    ws.set_column("M:M", 50); ws.set_column("N:O", 20); ws.set_column("P:P", 12)
    ws.set_column("Q:S", 18)

    # Rows 0-3: Issue Summary
    ws.merge_range(0, 0, 3, 18, ISSUE_SUMMARY, f_summary)
    ws.set_row(0, 80)

    # Row 5: blank label area
    ws.write(5, 0, "Summary Table", f_lbl)

    # Row 6: Table 1 label
    ws.write(6, 0, "Table 1", f_lbl)

    # Row 7: Table 1 headers
    ws.set_row(7, 36)
    ws.write(7, 0, "Canonical Issues", f_red_lft)
    for i, lbl in enumerate(ISSUE_LABELS):
        ws.write(7, i + 1, lbl, f_red_hdr)

    # Row 8: Issue Priority
    ws.write(8, 0, "Issue Priority", f_bold_lft)
    for i, lbl in enumerate(ISSUE_LABELS):
        ws.write(8, i + 1, ISSUE_PRIORITIES[lbl], f_cell)

    # Row 9: #Affected URLs
    ws.write(9, 0, "#Affected URLs", f_bold_lft)
    for i, lbl in enumerate(ISSUE_LABELS):
        ws.write(9, i + 1, len(issue_url_sets[lbl]), f_num)

    # Row 10: % Share
    ws.write(10, 0, "% share against Total URLs Crawled", f_bold_lft)
    for i, lbl in enumerate(ISSUE_LABELS):
        pct = len(issue_url_sets[lbl]) / total_crawled if total_crawled > 0 else 0
        ws.write(10, i + 1, pct, f_pct)

    # Row 12: Table 2 label
    ws.write(12, 0, "Table 2", f_lbl)

    # Row 13: Table 2 title
    ws.merge_range(13, 0, 13, len(ISSUE_LABELS) + 2, "Page Theme Wise Issue", f_red_lft)

    # Row 14: Table 2 headers
    ws.set_row(14, 36)
    ws.write(14, 0, "Page Theme 1", f_col_lft)
    ws.write(14, 1, "Priority Basis Page Theme 1", f_col_hdr)
    ws.write(14, 2, "Total Pages", f_col_hdr)
    for i, lbl in enumerate(ISSUE_LABELS):
        ws.write(14, i + 3, lbl, f_col_hdr)

    # Table 2 data
    T2_DATA_START = 15
    for rof, (theme, priority) in enumerate(sorted_themes):
        r = T2_DATA_START + rof
        ws.write(r, 0, theme, f_bold_lft)
        ws.write(r, 1, priority, f_cell)
        ws.write(r, 2, theme_total.get(theme, 0), f_num)
        for i, lbl in enumerate(ISSUE_LABELS):
            ws.write(r, i + 3, theme_counts[theme].get(lbl, 0), f_num)

    # Table 3
    T3_LABEL_ROW = T2_DATA_START + len(sorted_themes) + 1
    ws.write(T3_LABEL_ROW, 0, "Table 3", f_lbl)

    T3_HDR_ROW = T3_LABEL_ROW + 1
    T3_DATA_START = T3_HDR_ROW + 1

    ws.set_row(T3_HDR_ROW, 31.5)
    t3_headers = [
        "Error Type", "Address", "Page Theme 1", "Page Theme 2",
        "Content Type", "Status Code", "Indexability", "Occurance",
        "Canonical Link Element 1", "Canonical 1 Status Code", "Canonical 1 Indexability",
        "Inlink Count",
        "Canonical Link Element 2", "Canonical 2 Status Code", "Canonical 2 Indexability",
        "Inlink Count",
        "Impressions", "Clicks", "Organic Sessions"
    ]
    for i, h in enumerate(t3_headers):
        ws.write(T3_HDR_ROW, i, h, f_red_lft if i == 0 else f_red_hdr)

    for rof, row in enumerate(all_rows):
        r = T3_DATA_START + rof
        ws.set_row(r, 14.5)
        ws.write(r, 0, row["error_type"], f_cell_lft)
        ws.write(r, 1, row["address"], f_cell_lft)
        ws.write(r, 2, row["theme1"], f_cell)
        ws.write(r, 3, row["theme2"], f_cell)
        ws.write(r, 4, row["page_type"], f_cell)
        ws.write(r, 5, row["status_code"], f_cell)
        ws.write(r, 6, row["indexability"], f_cell)
        ws.write(r, 7, safe_num(row["occurrence"]), f_num)
        ws.write(r, 8, row["can1"], f_cell_lft)
        ws.write(r, 9, row["can1_sc"], f_cell)
        ws.write(r, 10, row["can1_idx"], f_cell)
        ws.write(r, 11, safe_num(row["inlinks"]), f_num)
        ws.write(r, 12, row["can2"], f_cell_lft)
        ws.write(r, 13, row["can2_sc"], f_cell)
        ws.write(r, 14, row["can2_idx"], f_cell)
        ws.write(r, 15, safe_num(row["can2_inlinks"]), f_num)
        ws.write(r, 16, safe_num(row["impressions"]), f_rgt)
        ws.write(r, 17, safe_num(row["clicks"]), f_rgt)
        ws.write(r, 18, safe_num(row["sessions"]), f_rgt)

    wb.close()
    buf.seek(0)
    return buf.read()
