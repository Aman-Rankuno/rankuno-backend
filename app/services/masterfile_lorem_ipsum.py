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
NAVY = "#1F3864"

ISSUE_SUMMARY = (
    "Issue Summary:\n"
    "Lorem Ipsum Placeholder - Pages containing placeholder text such as \"Lorem Ipsum,\" "
    "indicating incomplete or unfinished content that should be replaced before publication."
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

def build_lorem_ipsum_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "Lorem Ipsum Placeholder.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    rulebook = load_rulebook(domain)

    # Load source CSV
    df_lorem = load_df(report_path, "content_lorem_ipsum_placeholder.csv")

    # Load internal_all for page type, canonical, sitemap, status
    df_int = load_df(report_path, "internal_all.csv")
    internal_map = {}
    total_crawled = 0
    if not df_int.empty:
        total_crawled = len(df_int)
        a_col = _gc(df_int, "Address")
        pt_col = _gc(df_int, "Content Type", "Page Type")
        sc_col = _gc(df_int, "Status Code")
        idx_col = _gc(df_int, "Indexability")
        can_col = _gc(df_int, "Canonical Link Element 1", "Canonical URL")
        sit_col = _gc(df_int, "Sitemap 1", "In Sitemap")
        for _, r in df_int.iterrows():
            url = safe_str(r.get(a_col, "")) if a_col else ""
            internal_map[url] = {
                "page_type": safe_str(r.get(pt_col, "")) if pt_col else "",
                "status_code": safe_str(r.get(sc_col, "")) if sc_col else "",
                "indexability": safe_str(r.get(idx_col, "")) if idx_col else "",
                "canonical": safe_str(r.get(can_col, "")) if can_col else "",
                "in_sitemap": safe_str(r.get(sit_col, "")) if sit_col else "",
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

    # Filter: only indexable HTML 200 URLs
    rows = []
    addr_col = _gc(df_lorem, "Address")
    occ_col = _gc(df_lorem, "Occurrences", "Occurrence")
    idx_col_l = _gc(df_lorem, "Indexability")

    if not df_lorem.empty and addr_col:
        for _, row in df_lorem.iterrows():
            url = safe_str(row.get(addr_col, ""))
            if not url:
                continue
            info = internal_map.get(url, {})
            # Filter: indexable, 200, HTML
            if info.get("indexability", "").lower() != "indexable":
                continue
            if info.get("status_code", "") != "200":
                continue
            pt = info.get("page_type", "").lower()
            if "html" not in pt:
                continue

            theme1, theme2, _, priority = classify_url(url, rulebook)
            gsc = gsc_map.get(url, {})

            rows.append({
                "error_type": "Lorem Ipsum",
                "address": url,
                "page_type": info.get("page_type", ""),
                "theme1": theme1 or "-",
                "theme2": theme2 or "-",
                "status_code": info.get("status_code", "200"),
                "indexability": info.get("indexability", "Indexable"),
                "occurrence": safe_num(row.get(occ_col, "")) if occ_col else None,
                "canonical": info.get("canonical", ""),
                "in_sitemap": info.get("in_sitemap", ""),
                "impressions": gsc.get("impressions"),
                "clicks": gsc.get("clicks"),
                "sessions": ga_map.get(url),
                "priority": priority or "N/A",
            })

    # Sort Table 3 by impressions descending
    rows.sort(key=lambda x: (x.get("impressions") or 0), reverse=True)

    # Table 2: theme counts
    theme_counts = defaultdict(int)
    theme_priority_map = {}
    for r in rows:
        theme = r["theme1"] if r["theme1"] != "-" else "Others"
        theme_counts[theme] += 1
        if theme not in theme_priority_map:
            theme_priority_map[theme] = r["priority"]

    sorted_themes = sorted(theme_priority_map.items(),
                           key=lambda x: PRIORITY_ORDER.get(x[1], 3))

    affected = len(rows)
    pct_share = affected / total_crawled if total_crawled > 0 else 0

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

    ws = wb.add_worksheet("Lorem Ipsum Placeholder")
    ws.set_column("A:A", 60)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 15)
    ws.set_column("D:E", 20)
    ws.set_column("F:G", 15)
    ws.set_column("H:H", 12)
    ws.set_column("I:I", 50)
    ws.set_column("J:J", 20)
    ws.set_column("K:M", 18)

    # Rows 0-3: Issue Summary
    ws.merge_range(0, 0, 3, 12, ISSUE_SUMMARY, f_summary)
    ws.set_row(0, 60)

    # Row 4: blank
    # Row 5: Table 1 label
    ws.write(5, 0, "Table 1", f_lbl)

    # Row 6: Table 1 headers
    ws.set_row(6, 36)
    ws.write(6, 0, "Content Issue Types", f_red_lft)
    ws.write(6, 1, "Lorem Ipsum Placeholder", f_red_hdr)

    # Row 7: Issue Priority
    ws.write(7, 0, "Issue Priority", f_red_lft)
    ws.write(7, 1, "Medium", f_cell)

    # Row 8: #Affected URLs
    ws.write(8, 0, "#Affected URLs", f_red_lft)
    ws.write(8, 1, affected, f_num)

    # Row 9: % Share
    ws.write(9, 0, "% Share against Total URLs Crawled", f_red_lft)
    ws.write(9, 1, pct_share, f_pct)

    # Row 11: Table 2 label
    ws.write(11, 0, "Table 2", f_lbl)

    # Row 12: Table 2 title
    ws.merge_range(12, 0, 12, 2, "Page Theme Wise URL Analysis ", f_red_lft)

    # Row 13: Table 2 headers
    ws.set_row(13, 36)
    ws.write(13, 0, "Page Theme 1", f_col_lft)
    ws.write(13, 1, "Priority Basis Page Theme 1", f_col_hdr)
    ws.write(13, 2, "Lorem Ipsum Placeholder", f_col_hdr)

    # Table 2 data
    T2_DATA_START = 14
    for i, (theme, priority) in enumerate(sorted_themes):
        r = T2_DATA_START + i
        ws.write(r, 0, theme, f_bold_lft)
        ws.write(r, 1, priority, f_cell)
        ws.write(r, 2, theme_counts.get(theme, 0), f_num)

    # Table 3
    T3_LABEL_ROW = T2_DATA_START + len(sorted_themes) + 1
    ws.write(T3_LABEL_ROW, 0, "Table 3", f_lbl)

    T3_HDR_ROW = T3_LABEL_ROW + 1
    T3_DATA_START = T3_HDR_ROW + 1

    ws.set_row(T3_HDR_ROW, 31.5)
    t3_headers = [
        "Error Type", "Address", "Page Type", "Page Theme 1", "Page Theme 2",
        "Status Code", "Indexability", "Occurance", "Canonical URL",
        "Avaialble in Sitemap", "Impressions", "Clicks", "Organic Sessions"
    ]
    for i, h in enumerate(t3_headers):
        ws.write(T3_HDR_ROW, i, h, f_red_lft if i == 0 else f_red_hdr)

    for rof, row in enumerate(rows):
        r = T3_DATA_START + rof
        ws.set_row(r, 14.5)
        ws.write(r, 0, row["error_type"], f_cell_lft)
        ws.write(r, 1, row["address"], f_cell_lft)
        ws.write(r, 2, row["page_type"], f_cell)
        ws.write(r, 3, row["theme1"], f_cell)
        ws.write(r, 4, row["theme2"], f_cell)
        ws.write(r, 5, row["status_code"], f_cell)
        ws.write(r, 6, row["indexability"], f_cell)
        ws.write(r, 7, safe_num(row["occurrence"]), f_num)
        ws.write(r, 8, row["canonical"], f_cell_lft)
        ws.write(r, 9, row["in_sitemap"], f_cell)
        ws.write(r, 10, safe_num(row["impressions"]), f_rgt)
        ws.write(r, 11, safe_num(row["clicks"]), f_rgt)
        ws.write(r, 12, safe_num(row["sessions"]), f_rgt)

    wb.close()
    buf.seek(0)
    return buf.read()
