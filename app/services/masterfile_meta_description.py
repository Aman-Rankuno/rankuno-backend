import os
import io
import math
import pandas as pd
import xlsxwriter
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

FONT = "Rockwell"
RED = "#FF0000"
WHITE = "#FFFFFF"
BLACK = "#000000"

SHEET_NAME = "Meta Description Issues"

ISSUE_KEYS = ["Short", "Long", "Missing", "Duplicate", "Multiple", "Outside Head"]
ISSUE_PRIORITIES = {
    "Short": "Low",
    "Long": "Low",
    "Missing": "High",
    "Duplicate": "High",
    "Multiple": "Medium",
    "Outside Head": "High",
}
ISSUE_CSVS = {
    "Short":        "meta_description_below_400_pixels.csv",
    "Long":         "meta_description_over_985_pixels.csv",
    "Missing":      "meta_description_missing.csv",
    "Duplicate":    "meta_description_duplicate.csv",
    "Multiple":     "meta_description_multiple.csv",
    "Outside Head": "meta_description_outside_head.csv",
}
T3_COLS = [
    "Address", "Page Theme 1", "Page Theme 2", "Content Type",
    "Status Code", "Indexability", "Meta Description",
    "Meta Description Pixel Width",
    "Short", "Long", "Missing", "Duplicate", "Multiple", "Outside Head",
]

ISSUE_SUMMARY_TEXT = (
    "Issue Summary: Meta descriptions are HTML attributes that provide a brief "
    "summary of webpage content for search engines and users. Meta description "
    "issues include missing, duplicate, overly long, short, or irrelevant "
    "descriptions that may reduce search visibility and CTR."
)

# Excel hard limit is 1,048,576 rows; cap detail rows below it
MAX_T3_ROWS = 1_000_000


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


def build_meta_description_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    rulebook = load_rulebook(domain)

    def load_csv(filename):
        path = os.path.join(report_path, filename)
        if os.path.exists(path):
            try:
                return pd.read_csv(path, encoding="utf-8", low_memory=False)
            except Exception:
                return pd.DataFrame()
        return pd.DataFrame()

    df_internal_all = load_csv("internal_all.csv")
    df_gsc          = load_csv("search_console_all.csv")
    df_ga           = load_csv("analytics_all.csv")
    if df_internal_all.empty:
        raise ValueError("internal_all.csv not found or empty")

    def url_set(filename):
        df = load_csv(filename)
        if df.empty:
            return set()
        col = next((c for c in df.columns if c.lower() == "address"), None)
        return set(df[col].dropna().astype(str)) if col else set()

    issue_url_sets = {issue: url_set(fname) for issue, fname in ISSUE_CSVS.items()}

    gsc_map = {}
    if not df_gsc.empty:
        a   = next((c for c in df_gsc.columns if c.lower() == "address"), None)
        imp = next((c for c in df_gsc.columns if "impression" in c.lower()), None)
        clk = next((c for c in df_gsc.columns if "click" in c.lower()), None)
        if a:
            for _, r in df_gsc.iterrows():
                gsc_map[str(r[a])] = {
                    "impressions": r.get(imp, 0) if imp else 0,
                    "clicks":      r.get(clk, 0) if clk else 0,
                }
    ga_map = {}
    if not df_ga.empty:
        a = next((c for c in df_ga.columns if c.lower() == "address"), None)
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[str(r[a])] = r.get(s, 0)

    def get_col(df, *names):
        for n in names:
            m = next((c for c in df.columns if c.lower() == n.lower()), None)
            if m:
                return m
        return None

    ia_addr   = get_col(df_internal_all, "Address")
    ia_status = get_col(df_internal_all, "Status Code")
    ia_idx    = get_col(df_internal_all, "Indexability")
    ia_ct     = get_col(df_internal_all, "Content Type")
    ia_meta   = get_col(df_internal_all, "Meta Description 1", "Meta Description")
    ia_pixel  = get_col(df_internal_all, "Meta Description 1 Pixel Width", "Meta Description Pixel Width")

    mask = (
        (df_internal_all[ia_status].astype(str) == "200") &
        (df_internal_all[ia_idx].astype(str).str.lower() == "indexable") &
        (df_internal_all[ia_ct].astype(str).str.lower().str.contains("text/html", na=False))
    )
    total_indexable = int(mask.sum())
    df_base = df_internal_all[mask].copy()
    df_base["_url"] = df_base[ia_addr].astype(str)

    # Classification cached per unique URL
    _cls_cache = {}

    def classify_cached(url):
        hit = _cls_cache.get(url)
        if hit is None:
            hit = classify_url(url, rulebook)
            _cls_cache[url] = hit
        return hit

    classifications = [classify_cached(url) for url in df_base["_url"]]
    df_base["Page Theme 1"] = [c[0] if c[0] else "Others" for c in classifications]
    df_base["Page Theme 2"] = [c[1] if c[1] else "-" for c in classifications]
    df_base["_priority"]    = [c[3] for c in classifications]
    for issue in ISSUE_KEYS:
        df_base[issue] = df_base["_url"].isin(issue_url_sets[issue]).map({True: "Yes", False: "No"})
    df_base["Impressions"]      = df_base["_url"].map(lambda u: gsc_map.get(u, {}).get("impressions", 0))
    df_base["Clicks"]           = df_base["_url"].map(lambda u: gsc_map.get(u, {}).get("clicks", 0))
    df_base["Organic Sessions"] = df_base["_url"].map(lambda u: ga_map.get(u, 0))

    rename_map = {}
    if ia_addr:   rename_map[ia_addr]   = "Address"
    if ia_ct:     rename_map[ia_ct]     = "Content Type"
    if ia_status: rename_map[ia_status] = "Status Code"
    if ia_idx:    rename_map[ia_idx]    = "Indexability"
    if ia_meta:   rename_map[ia_meta]   = "Meta Description"
    if ia_pixel:  rename_map[ia_pixel]  = "Meta Description Pixel Width"
    df_table3 = df_base.rename(columns=rename_map)
    for col in T3_COLS:
        if col not in df_table3.columns:
            df_table3[col] = None
    df_table3 = df_table3[T3_COLS + ["_priority", "Impressions", "Clicks", "Organic Sessions"]].copy()
    # Team decision: the detail table lists only pages with at least one issue;
    # clean pages are excluded (they still count in Table 1's Total Pages and
    # the theme table's Total Pages, which describe the full crawl)
    issue_mask = df_table3[ISSUE_KEYS].eq("Yes").any(axis=1)
    df_table3 = df_table3[issue_mask]
    df_table3 = df_table3.sort_values("Impressions", ascending=False).reset_index(drop=True)
    # Unique pages carrying at least one issue (the row total for #Affected URLs)
    total_affected = int(df_table3["Address"].nunique())

    issue_counts = {issue: len(issue_url_sets[issue]) for issue in ISSUE_KEYS}

    # Theme aggregation uses the UNFILTERED page set so Total Pages per theme
    # and the issue percentages describe the whole crawl, not just issue pages
    theme_data = {}
    for _, row in df_base.iterrows():
        theme    = row["Page Theme 1"] or "Others"
        priority = row.get("_priority", "Low")
        if theme not in theme_data:
            theme_data[theme] = {"total": 0, "priority": priority, **{k: 0 for k in ISSUE_KEYS}}
        else:
            if PRIORITY_ORDER.get(priority, 2) < PRIORITY_ORDER.get(theme_data[theme]["priority"], 2):
                theme_data[theme]["priority"] = priority
        theme_data[theme]["total"] += 1
        for issue in ISSUE_KEYS:
            if row[issue] == "Yes":
                theme_data[theme][issue] += 1
    sorted_themes = sorted(
        theme_data.items(),
        key=lambda x: PRIORITY_ORDER.get(x[1]["priority"], 2)
    )

    # ── Build workbook entirely in code ───────────────────────────────────────
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {
        "in_memory": True,
        "nan_inf_to_errors": True,
        "strings_to_urls": False,
    })

    def f(**kw):
        return wb.add_format(kw)

    f_summary = f(font_name=FONT, font_size=10, font_color=BLACK,
                  text_wrap=True, valign="top", border=1)
    f_section = f(bold=True, font_name=FONT, font_size=14, font_color=BLACK,
                  border=1, valign="vcenter")
    f_lbl = f(bold=True, font_name=FONT, font_size=10, font_color=BLACK)
    f_red_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="left", valign="vcenter", text_wrap=True)
    f_red_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
    f_cell = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_cell_lft = f(font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_bold_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="#,##0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0%")
    f_rgt = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="right", valign="vcenter", num_format="#,##0")

    ws = wb.add_worksheet(SHEET_NAME)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 55)
    ws.set_column("B:H", 18)
    ws.set_column("I:N", 14)
    ws.set_column("O:Q", 16)

    # ── Issue Summary box (rows 1-4, merged and bordered) ────────────────────
    ws.merge_range(0, 0, 3, 7, ISSUE_SUMMARY_TEXT, f_summary)
    ws.set_row(0, 18)

    # ── "Summary Table" section header (row 6) ───────────────────────────────
    ws.merge_range(5, 0, 6, 7, "Summary Table", f_section)

    # ── Table 1 (label row 8; headers 9; priority 10; counts 11; % 12) ───────
    ws.write(7, 0, "Table 1", f_lbl)

    ws.write(8, 0, "Meta Description Issue Types", f_red_lft)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(8, i + 1, issue, f_red_ctr)
    ws.write(8, 7, "Total Affected URLs", f_red_ctr)

    ws.write(9, 0, "Issue Priority", f_red_lft)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(9, i + 1, ISSUE_PRIORITIES[issue], f_cell)
    ws.write(9, 7, "", f_cell)

    ws.write(10, 0, "#Affected URLs", f_red_lft)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(10, i + 1, issue_counts[issue], f_num)
    ws.write(10, 7, total_affected, f_num)

    ws.write(11, 0, "% Share against Total URLs Crawled", f_red_lft)
    # (column H on this row shows the crawl-wide denominator itself)
    # Live formulas dividing each count cell by the Total Pages cell, with the
    # computed value cached so Excel shows the number immediately
    denom = total_indexable if total_indexable > 0 else 1
    for i, issue in enumerate(ISSUE_KEYS):
        col_letter = chr(ord("B") + i)
        pct = (issue_counts[issue] / total_indexable) if total_indexable > 0 else 0
        ws.write_formula(11, i + 1, f"={col_letter}11/{denom}", f_pct, pct)
    ws.write(11, 7, total_indexable, f_num)

    # ── Page Theme Wise table (Table 2: label 15; title 16; headers 17) ──────
    T2_LABEL_ROW = 14
    T2_TITLE_ROW = 15
    T2_HEADER_ROW = 16
    T2_DATA_START = 17
    ws.write(T2_LABEL_ROW, 0, "Table 2", f_lbl)
    ws.merge_range(T2_TITLE_ROW, 0, T2_TITLE_ROW, 8,
                   "Page Theme Wise Meta Description Analysis", f_red_lft)
    t2_headers = ["Page Theme", "Total Pages", "Priority"] + ISSUE_KEYS
    for i, h in enumerate(t2_headers):
        ws.write(T2_HEADER_ROW, i, h, f_red_lft if i == 0 else f_red_ctr)

    for row_offset, (theme, counts) in enumerate(sorted_themes):
        r = T2_DATA_START + row_offset
        total_theme = counts["total"]
        ws.write(r, 0, theme, f_bold_ctr)
        ws.write(r, 1, total_theme, f_num)
        ws.write(r, 2, counts["priority"], f_cell)
        for i, issue in enumerate(ISSUE_KEYS):
            cnt = counts[issue]
            if cnt > 0 and total_theme > 0:
                ws.write(r, 3 + i, f"{cnt} ({round(cnt / total_theme * 100)}%)", f_cell)
            else:
                ws.write(r, 3 + i, 0, f_num)

    # ── Detail table ("Table 2" per house convention) ────────────────────────
    t2_last_row   = T2_DATA_START + max(len(sorted_themes), 1) - 1
    t3_label_row  = t2_last_row + 1
    t3_header_row = t3_label_row + 1
    t3_data_start = t3_header_row + 1

    ws.write(t3_label_row, 0, "Table 3", f_lbl)
    for i, h in enumerate(T3_COLS):
        ws.write(t3_header_row, i, h, f_red_lft if i == 0 else f_red_ctr)
    extra_cols = ["Impressions", "Clicks", "Organic Sessions"]
    for j, h in enumerate(extra_cols):
        ws.write(t3_header_row, len(T3_COLS) + j, h, f_red_ctr)

    t3_rows = df_table3.head(MAX_T3_ROWS)
    for row_offset, (_, row) in enumerate(t3_rows.iterrows()):
        r = t3_data_start + row_offset
        ws.write(r, 0, safe_str(row.get("Address")), f_cell_lft)
        ws.write(r, 1, safe_str(row.get("Page Theme 1")), f_cell)
        ws.write(r, 2, safe_str(row.get("Page Theme 2")), f_cell)
        ws.write(r, 3, safe_str(row.get("Content Type")), f_cell)
        ws.write(r, 4, safe_str(row.get("Status Code")), f_cell)
        ws.write(r, 5, safe_str(row.get("Indexability")), f_cell)
        ws.write(r, 6, safe_str(row.get("Meta Description")), f_cell_lft)
        ws.write(r, 7, safe_num(row.get("Meta Description Pixel Width")), f_rgt)
        for i, issue in enumerate(ISSUE_KEYS):
            ws.write(r, 8 + i, safe_str(row.get(issue)), f_cell)
        ws.write(r, 14, safe_num(row.get("Impressions")), f_rgt)
        ws.write(r, 15, safe_num(row.get("Clicks")), f_rgt)
        ws.write(r, 16, safe_num(row.get("Organic Sessions")), f_rgt)

    truncated = len(df_table3) - len(t3_rows)
    if truncated > 0:
        ws.write(t3_data_start + len(t3_rows), 0,
                 f"Note: {truncated:,} additional rows omitted due to the Excel row limit; "
                 "full detail available in the source CSV exports", f_lbl)

    last_data_row = t3_data_start + max(len(t3_rows), 1) - 1
    ws.autofilter(t3_header_row, 0, last_data_row, len(T3_COLS) + len(extra_cols) - 1)

    wb.close()
    buf.seek(0)
    return buf.read()