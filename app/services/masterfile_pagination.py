import os
import io
import math
import pandas as pd
import xlsxwriter
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

ISSUE_KEYS = [
    "Pagination URL Not in Anchor Tag",
    "Non-200 Pagination URLs",
    "Unlinked Pagination URLs",
    "Non-Indexable",
    "Multiple Pagination URLs",
    "Pagination Loop",
    "Sequence Error",
]

ISSUE_CSV = {
    "Pagination URL Not in Anchor Tag": "pagination_pagination_url_not_in_anchor_tag.csv",
    "Non-200 Pagination URLs": "pagination_non200_pagination_urls.csv",
    "Unlinked Pagination URLs": "pagination_unlinked_pagination_urls.csv",
    "Non-Indexable": "pagination_nonindexable.csv",
    "Multiple Pagination URLs": "pagination_multiple_pagination_urls.csv",
    "Pagination Loop": "pagination_pagination_loop.csv",
    "Sequence Error": "pagination_sequence_error.csv",
}

ISSUE_PRIORITY = {
    "Pagination URL Not in Anchor Tag": "Low",
    "Non-200 Pagination URLs": "High",
    "Unlinked Pagination URLs": "High",
    "Non-Indexable": "High",
    "Multiple Pagination URLs": "Low",
    "Pagination Loop": "Low",
    "Sequence Error": "Medium",
}

SUMMARY_TEXT = (
    "Issue Summary:\n"
    "1. Pagination URL Not in Anchor Tags - Pagination URLs are not properly linked in the page navigation, "
    "making it harder for crawlers to discover deeper paginated pages and their content.\n"
    "2. Non-200 Pagination URLs - Pagination links point to URLs returning errors or redirects instead of a "
    "200 OK status, which disrupts crawl paths and can prevent search engines from accessing paginated content.\n"
    "3. Unlinked Pagination URLs - Paginated pages exist but are not internally linked through the pagination "
    "sequence, reducing discoverability and risking orphaned content.\n"
    "4. Non-Indexable - A paginated URL is blocked from indexing (e.g. noindex, canonicalized, blocked in "
    "robots), which may prevent search engines from accessing or consolidating content signals correctly.\n"
    "5. Multiple Pagination URLs - A page contains more than one conflicting pagination reference "
    "(e.g. multiple next/prev paths), creating ambiguity for crawlers and weakening pagination signals.\n"
    "6. Pagination Loop - Pagination links create a circular path by pointing back to previously visited "
    "pages instead of progressing sequentially, causing crawl inefficiencies.\n"
    "7. Sequence Error - Pagination numbering or next/prev relationships are broken or out of order, "
    "confusing crawlers and disrupting proper content discovery across paginated series."
)


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f) or f == 0) else f
    except Exception:
        return None


def _get_page_type(content_type: str) -> str:
    ct = str(content_type).lower()
    if "html" in ct:
        return "HTML"
    if "pdf" in ct:
        return "PDF"
    if ct.startswith("image/"):
        return "Image"
    if "css" in ct:
        return "CSS"
    if "javascript" in ct:
        return "JavaScript"
    if ct and ct != "nan":
        return "Other"
    return ""


def _gc(df, *names):
    for n in names:
        m = next((c for c in df.columns if c.lower() == n.lower()), None)
        if m:
            return m
    return None


def build_pagination_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "Pagination.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

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
    df_gsc = load_csv("search_console_all.csv")
    df_ga = load_csv("analytics_all.csv")

    # GSC lookup
    gsc_map = {}
    if not df_gsc.empty:
        a = _gc(df_gsc, "Address")
        imp = next((c for c in df_gsc.columns if "impression" in c.lower()), None)
        clk = next((c for c in df_gsc.columns if "click" in c.lower()), None)
        if a:
            for _, r in df_gsc.iterrows():
                gsc_map[str(r[a])] = {
                    "impressions": safe_num(r.get(imp, 0)) if imp else None,
                    "clicks": safe_num(r.get(clk, 0)) if clk else None,
                }

    # GA lookup
    ga_map = {}
    if not df_ga.empty:
        a = _gc(df_ga, "Address")
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[str(r[a])] = safe_num(r.get(s, 0))

    # Count total HTML pages from internal_all (all HTML, not just indexable)
    total_html = 0
    internal_map = {}
    if not df_internal_all.empty:
        a_col = _gc(df_internal_all, "Address")
        ct_col = _gc(df_internal_all, "Content Type")
        sc_col = _gc(df_internal_all, "Status Code")
        idx_col = _gc(df_internal_all, "Indexability")

        for _, r in df_internal_all.iterrows():
            url = str(r[a_col]) if a_col else ""
            ct = str(r.get(ct_col, "")) if ct_col else ""
            sc = str(r.get(sc_col, "")) if sc_col else ""
            idx = str(r.get(idx_col, "")) if idx_col else ""
            pt = _get_page_type(ct)
            if pt == "HTML":
                total_html += 1
            internal_map[url] = {
                "content_type": ct, "status_code": sc,
                "indexability": idx, "page_type": pt,
            }

    # Build Table 3 rows from each issue CSV
    all_rows = []

    for issue_key, csv_file in ISSUE_CSV.items():
        df = load_csv(csv_file)
        if df.empty:
            continue

        addr_col = _gc(df, "Address")
        sc_col2 = _gc(df, "Status Code")
        idx_col2 = _gc(df, "Indexability")
        idx_status_col = _gc(df, "Indexability Status")
        rel_next_col = _gc(df, "rel=\"next\" 1", "rel=next 1")
        rel_prev_col = _gc(df, "rel=\"prev\" 1", "rel=prev 1")
        http_next_col = _gc(df, "HTTP rel=\"next\" 1", "HTTP rel=next 1")
        http_prev_col = _gc(df, "HTTP rel=\"prev\" 1", "HTTP rel=prev 1")
        canon_col = _gc(df, "Canonical Link Element 1")
        http_canon_col = _gc(df, "HTTP Canonical")
        meta_robots_col = _gc(df, "Meta Robots 1")
        x_robots_col = _gc(df, "X-Robots-Tag 1")

        for _, row in df.iterrows():
            url = str(row.get(addr_col, "")) if addr_col else ""
            int_data = internal_map.get(url, {})
            status_code = str(row.get(sc_col2, "")) if sc_col2 else int_data.get("status_code", "")
            indexability = str(row.get(idx_col2, "")) if idx_col2 else int_data.get("indexability", "")

            page_theme1, page_theme2, _, _ = classify_url(url, rulebook)
            gsc = gsc_map.get(url, {})

            def gv(col):
                return str(row.get(col, "")) if col and col in row.index else ""

            all_rows.append({
                "Error Type": issue_key,
                "Address": url,
                "Page Theme 1": page_theme1 or "-",
                "Page Theme 2": page_theme2 if page_theme2 else "-",
                "Status Code": status_code,
                "Indexability": indexability,
                "Indexability Status": gv(idx_status_col),
                "rel=\"next\" 1": gv(rel_next_col),
                "rel=\"prev\" 1": gv(rel_prev_col),
                "HTTP rel=\"next\" 1": gv(http_next_col),
                "HTTP rel=\"prev\" 1": gv(http_prev_col),
                "Canonical Link Element 1": gv(canon_col),
                "HTTP Canonical": gv(http_canon_col),
                "Meta Robots 1": gv(meta_robots_col),
                "X-Robots-Tag 1": gv(x_robots_col),
                "Impressions": gsc.get("impressions"),
                "Clicks": gsc.get("clicks"),
                "Organic Sessions": ga_map.get(url),
            })

    df_table3 = pd.DataFrame(all_rows) if all_rows else pd.DataFrame(columns=[
        "Error Type", "Address", "Page Theme 1", "Page Theme 2",
        "Status Code", "Indexability", "Indexability Status",
        "rel=\"next\" 1", "rel=\"prev\" 1", "HTTP rel=\"next\" 1", "HTTP rel=\"prev\" 1",
        "Canonical Link Element 1", "HTTP Canonical", "Meta Robots 1", "X-Robots-Tag 1",
        "Impressions", "Clicks", "Organic Sessions",
    ])
    if not df_table3.empty:
        df_table3 = df_table3.sort_values(
            "Impressions", ascending=False, na_position="last"
        ).reset_index(drop=True)

    # Build unique themes for Table 2 sorted by priority
    theme_priority = {}
    for _, row in df_table3.iterrows():
        theme = row["Page Theme 1"] if row["Page Theme 1"] != "-" else "Others"
        _, _, _, priority = classify_url(row["Address"], rulebook)
        if theme not in theme_priority:
            theme_priority[theme] = priority
        else:
            if PRIORITY_ORDER.get(priority, 3) < PRIORITY_ORDER.get(theme_priority[theme], 3):
                theme_priority[theme] = priority

    sorted_themes = sorted(
        theme_priority.items(),
        key=lambda x: PRIORITY_ORDER.get(x[1], 3),
    )
    num_themes = len(sorted_themes)

    # Layout (0-indexed):
    # Rows 0-10 (Excel 1-11): blank
    # Row 11 (Excel 12): Issue Summary merged box (height 52.5)
    # Rows 12-13 (Excel 13-14): blank
    # Row 14 (Excel 15): Table 1 label
    # Row 15 (Excel 16): Table 1 headers (height 42)
    # Row 16 (Excel 17): Issue Priority
    # Row 17 (Excel 18): #Affected URLs
    # Row 18 (Excel 19): % Share (height 21)
    # Rows 19-20 (Excel 20-21): blank
    # Row 21 (Excel 22): Table 2 label
    # Row 22 (Excel 23): Table 2 title merged A23:I23
    # Row 23 (Excel 24): Table 2 headers (height 42)
    # Rows 24..24+num_themes-1: Table 2 data
    # 1 blank row
    # Table 3 label
    # 1 blank row
    # Table 3 headers (height 31.5)
    # Table 3 data

    T1_HDR_ROW_IDX = 15
    T1_HDR_ROW_XL = 16

    T2_HDR_ROW_IDX = 23
    T2_HDR_ROW_XL = 24
    T2_DATA_START_IDX = 24
    T2_DATA_END_IDX = T2_DATA_START_IDX + max(num_themes - 1, 0)

    T3_LABEL_IDX = T2_DATA_END_IDX + 2
    T3_HDR_IDX = T3_LABEL_IDX + 2
    T3_DATA_START_IDX = T3_HDR_IDX + 1
    T3_DATA_START_XL = T3_DATA_START_IDX + 1

    # BUILD WORKBOOK
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws = wb.add_worksheet("Pagination")

    RED = "#FF0000"
    WHITE = "#FFFFFF"
    BLACK = "#000000"
    DARK = "#404040"
    FONT = "Rockwell"

    def f(**kw):
        return wb.add_format(kw)

    f_red_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
    f_red_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="left", valign="vcenter", text_wrap=True)
    f_dark_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                   bg_color=DARK, border=1, align="center", valign="vcenter", text_wrap=True)
    f_section = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK)
    f_ctr = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_lft = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_rgt = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="right", valign="vcenter")
    f_bold_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0.00%")
    f_t2_merged = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                    bg_color=RED, border=1, align="center", valign="vcenter")
    f_summary = f(font_name=FONT, font_size=8, font_color=BLACK,
                  text_wrap=True, valign="top", border=1)

    # Column widths (match template)
    ws.set_column("A:A", 17.0)
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
                "L", "M", "N", "O", "P", "Q", "R"]:
        ws.set_column(f"{col}:{col}", 13.0)

    # Row 12 (idx 11): Issue Summary merged box
    ws.set_row(11, 52.5)
    ws.merge_range(11, 0, 11, 17, SUMMARY_TEXT, f_summary)

    # Table 1 label
    ws.write(14, 0, "Table 1", f_section)

    # Table 1 headers
    ws.set_row(T1_HDR_ROW_IDX, 42.0)
    ws.write(T1_HDR_ROW_IDX, 0, "Content Issue Types", f_red_ctr)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(T1_HDR_ROW_IDX, i + 1, issue, f_red_ctr)

    # Issue Priority
    ws.write(16, 0, "Issue Priority", f_red_ctr)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(16, i + 1, ISSUE_PRIORITY[issue], f_ctr)

    # #Affected URLs
    ws.write(17, 0, "#Affected URLs", f_red_lft)
    for i, col_l in enumerate(["B", "C", "D", "E", "F", "G", "H"]):
        hdr_ref = col_l + str(T1_HDR_ROW_XL)
        formula = f"=COUNTIF(A{T3_DATA_START_XL}:A1048576,{hdr_ref})"
        ws.write_formula(17, i + 1, formula, f_ctr)

    # % Share
    ws.set_row(18, 21.0)
    ws.write(18, 0, "% Share against Total  URLs Crawled", f_red_lft)
    for i, col_l in enumerate(["B", "C", "D", "E", "F", "G", "H"]):
        affected_ref = col_l + "18"
        if total_html > 0:
            formula = f"={affected_ref}/{total_html}"
        else:
            formula = '=IF(FALSE,0,"")'
        ws.write_formula(18, i + 1, formula, f_pct)

    # Table 2 label
    ws.write(21, 0, "Table 2", f_section)

    # Table 2 title merged A23:I23
    ws.merge_range(22, 0, 22, 8, "Page Theme Wise URL Analysis ", f_t2_merged)

    # Table 2 headers
    ws.set_row(T2_HDR_ROW_IDX, 42.0)
    ws.write(T2_HDR_ROW_IDX, 0, "Page Theme 1", f_dark_ctr)
    ws.write(T2_HDR_ROW_IDX, 1, "Priority Basis Page Theme 1", f_dark_ctr)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(T2_HDR_ROW_IDX, i + 2, issue, f_dark_ctr)

    # Table 2 data rows
    for row_offset, (theme, priority) in enumerate(sorted_themes):
        r = T2_DATA_START_IDX + row_offset
        ws.set_row(r, 14.5)
        theme_cell = "A" + str(r + 1)
        ws.write(r, 0, theme, f_bold_ctr)
        ws.write(r, 1, priority, f_ctr)
        for i, col_l in enumerate(["C", "D", "E", "F", "G", "H", "I"]):
            hdr_ref = col_l + str(T2_HDR_ROW_XL)
            formula = (
                f"=COUNTIFS("
                f"A{T3_DATA_START_XL}:A1048576,{hdr_ref},"
                f"C{T3_DATA_START_XL}:C1048576,{theme_cell})"
            )
            ws.write_formula(r, i + 2, formula, f_ctr)

    # Table 3 label
    ws.write(T3_LABEL_IDX, 0, "Table 3", f_section)

    # Table 3 headers
    ws.set_row(T3_HDR_IDX, 31.5)
    t3_headers = [
        "Error Type", "Address", "Page Theme 1", "Page Theme 2",
        "Status Code", "Indexability", "Indexability Status",
        "rel=\"next\" 1", "rel=\"prev\" 1", "HTTP rel=\"next\" 1", "HTTP rel=\"prev\" 1",
        "Canonical Link Element 1", "HTTP Canonical", "Meta Robots 1", "X-Robots-Tag 1",
        "Impressions", "Clicks", "Organic Sessions ",
    ]
    for i, h in enumerate(t3_headers):
        ws.write(T3_HDR_IDX, i, h, f_red_lft if i == 0 else f_red_ctr)

    # Table 3 data
    for row_offset, (_, row) in enumerate(df_table3.iterrows()):
        r = T3_DATA_START_IDX + row_offset
        ws.set_row(r, 14.5)
        ws.write(r, 0, row.get("Error Type", "") or "", f_lft)
        ws.write(r, 1, row.get("Address", "") or "", f_lft)
        ws.write(r, 2, row.get("Page Theme 1", "") or "-", f_ctr)
        ws.write(r, 3, row.get("Page Theme 2", "-") or "-", f_ctr)
        ws.write(r, 4, row.get("Status Code", "") or "", f_ctr)
        ws.write(r, 5, row.get("Indexability", "") or "", f_ctr)
        ws.write(r, 6, row.get("Indexability Status", "") or "", f_ctr)
        ws.write(r, 7, row.get("rel=\"next\" 1", "") or "", f_lft)
        ws.write(r, 8, row.get("rel=\"prev\" 1", "") or "", f_lft)
        ws.write(r, 9, row.get("HTTP rel=\"next\" 1", "") or "", f_lft)
        ws.write(r, 10, row.get("HTTP rel=\"prev\" 1", "") or "", f_lft)
        ws.write(r, 11, row.get("Canonical Link Element 1", "") or "", f_lft)
        ws.write(r, 12, row.get("HTTP Canonical", "") or "", f_lft)
        ws.write(r, 13, row.get("Meta Robots 1", "") or "", f_ctr)
        ws.write(r, 14, row.get("X-Robots-Tag 1", "") or "", f_ctr)
        ws.write(r, 15, safe_num(row.get("Impressions")), f_rgt)
        ws.write(r, 16, safe_num(row.get("Clicks")), f_rgt)
        ws.write(r, 17, safe_num(row.get("Organic Sessions")), f_rgt)

    wb.close()
    buf.seek(0)
    return buf.read()
