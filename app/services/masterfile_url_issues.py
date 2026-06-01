import os
import io
import math
import pandas as pd
import xlsxwriter
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

ISSUE_KEYS = [
    "Uppercase", "Underscores", "Parameters",
    "Multiple Slashes", "Repetative Path", "Contains Space"
]

ISSUE_CSV = {
    "Uppercase": "url_uppercase.csv",
    "Underscores": "url_underscores.csv",
    "Parameters": "url_parameters.csv",
    "Multiple Slashes": "url_multiple_slashes.csv",
    "Repetative Path": "url_repetitive_path.csv",
    "Contains Space": "url_contains_space.csv",
}


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f) or f == 0) else f
    except Exception:
        return None


def _get_page_type(content_type: str) -> str:
    ct = str(content_type).lower()
    if "html" in ct: return "HTML"
    if "pdf" in ct: return "PDF"
    if ct.startswith("image/"): return "Image"
    if "css" in ct: return "CSS"
    if "javascript" in ct: return "JavaScript"
    if ct and ct != "nan": return "Other"
    return ""


def build_url_issues_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "URL Issues.xlsx")
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
    df_overview = load_csv("crawl_overview.csv")

    # GSC lookup
    gsc_map = {}
    if not df_gsc.empty:
        a = next((c for c in df_gsc.columns if c.lower() == "address"), None)
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
        a = next((c for c in df_ga.columns if c.lower() == "address"), None)
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[str(r[a])] = safe_num(r.get(s, 0))

    # internal_all: content type counts
    internal_map = {}
    total_all = total_html = total_js = total_css = 0
    total_images = total_pdf = total_other = total_unknown = 0

    if not df_internal_all.empty:
        def gc(df, *names):
            for n in names:
                m = next((c for c in df.columns if c.lower() == n.lower()), None)
                if m: return m
            return None

        a_col = gc(df_internal_all, "Address")
        ct_col = gc(df_internal_all, "Content Type")
        sc_col = gc(df_internal_all, "Status Code")
        idx_col = gc(df_internal_all, "Indexability")
        total_all = len(df_internal_all)

        for _, r in df_internal_all.iterrows():
            url = str(r[a_col]) if a_col else ""
            ct = str(r.get(ct_col, "")) if ct_col else ""
            sc = str(r.get(sc_col, "")) if sc_col else ""
            idx = str(r.get(idx_col, "")) if idx_col else ""
            pt = _get_page_type(ct)
            if pt == "HTML": total_html += 1
            elif pt == "JavaScript": total_js += 1
            elif pt == "CSS": total_css += 1
            elif pt == "Image": total_images += 1
            elif pt == "PDF": total_pdf += 1
            elif pt == "Other": total_other += 1
            else: total_unknown += 1
            internal_map[url] = {"content_type": ct, "status_code": sc,
                                  "indexability": idx, "page_type": pt}

    # Try to get overview data from crawl_overview.csv
    overview_data = {
        "All": total_all, "HTML": total_html, "JavaScript": total_js,
        "CSS": total_css, "Images": total_images, "PDF": total_pdf,
        "Flash": 0, "Other": total_other, "Unknown": total_unknown,
    }

    if not df_overview.empty:
        def gc_ov(df, *names):
            for n in names:
                m = next((c for c in df.columns if c.lower() == n.lower()), None)
                if m: return m
            return None
        type_col = gc_ov(df_overview, "Type", "Category", "Content Type")
        count_col = gc_ov(df_overview, "Total", "Count", "URLs")
        if type_col and count_col:
            for _, r in df_overview.iterrows():
                t = str(r.get(type_col, "")).strip()
                try:
                    cnt = int(float(r.get(count_col, 0)))
                except Exception:
                    cnt = 0
                if t in overview_data:
                    overview_data[t] = cnt

    # Build issue rows
    all_rows = []

    for issue_key, csv_file in ISSUE_CSV.items():
        df = load_csv(csv_file)
        if df.empty:
            continue

        def gc2(df, *names):
            for n in names:
                m = next((c for c in df.columns if c.lower() == n.lower()), None)
                if m: return m
            return None

        addr_col = gc2(df, "Address")
        ct_col2 = gc2(df, "Content Type")
        sc_col2 = gc2(df, "Status Code")
        idx_col2 = gc2(df, "Indexability")

        for _, row in df.iterrows():
            url = str(row.get(addr_col, "")) if addr_col else ""
            int_data = internal_map.get(url, {})
            content_type = str(row.get(ct_col2, "")) if ct_col2 else int_data.get("content_type", "")
            status_code = str(row.get(sc_col2, "")) if sc_col2 else int_data.get("status_code", "")
            indexability = str(row.get(idx_col2, "")) if idx_col2 else int_data.get("indexability", "")
            page_type = _get_page_type(content_type)

            if status_code != "200": continue
            if indexability.lower() != "indexable": continue
            if page_type != "HTML": continue

            page_theme1, page_theme2, _, _ = classify_url(url, rulebook)
            gsc = gsc_map.get(url, {})

            all_rows.append({
                "Error Type": issue_key,
                "Address": url,
                "Page Theme 1": page_theme1 or "",
                "Page Theme 2": page_theme2 if page_theme2 else "-",
                "Content Type": content_type,
                "Status Code": 200,
                "Indexability": indexability,
                "Impressions": gsc.get("impressions"),
                "Clicks": gsc.get("clicks"),
                "Organic Sessions": ga_map.get(url),
            })

    df_table3 = pd.DataFrame(all_rows) if all_rows else pd.DataFrame(columns=[
        "Error Type", "Address", "Page Theme 1", "Page Theme 2",
        "Content Type", "Status Code", "Indexability",
        "Impressions", "Clicks", "Organic Sessions"
    ])
    if not df_table3.empty:
        df_table3 = df_table3.sort_values(
            "Impressions", ascending=False, na_position='last'
        ).reset_index(drop=True)

    # Build unique themes for Table 2 sorted by priority
    theme_priority = {}
    for _, row in df_table3.iterrows():
        theme = row["Page Theme 1"] or "Others"
        _, _, _, priority = classify_url(row["Address"], rulebook)
        if theme not in theme_priority:
            theme_priority[theme] = priority
        else:
            if PRIORITY_ORDER.get(priority, 3) < PRIORITY_ORDER.get(theme_priority[theme], 3):
                theme_priority[theme] = priority

    sorted_themes = sorted(
        theme_priority.items(),
        key=lambda x: PRIORITY_ORDER.get(x[1], 3)
    )
    num_themes = len(sorted_themes)

    # Fixed layout rows (0-indexed)
    # Row 1-8: blank (rows 1-8 in Excel)
    # Row 9 (idx 8): Overview Table label
    # Row 10 (idx 9): Overview headers
    # Row 11 (idx 10): Internal label
    # Rows 12-20 (idx 11-19): Overview data (All, HTML, JS, CSS, Images, PDF, Flash, Other, Unknown)
    # Rows 21-24: blank
    # Row 25 (idx 24): Table 1 label
    # Row 26 (idx 25): Table 1 headers
    # Row 27 (idx 26): Issue Priority
    # Row 28 (idx 27): #Affected URLs (formulas)
    # Row 29 (idx 28): % Share (formulas)
    # Rows 30-31: blank
    # Row 32 (idx 31): Table 2 label
    # Row 33 (idx 32): Table 2 title (merged)
    # Row 34 (idx 33): Table 2 headers
    # Rows 35+ (idx 34+): Table 2 data (dynamic per themes)
    # 1 blank row
    # Table 3 label
    # Table 3 headers
    # Table 3 data

    T2_HDR_ROW_IDX = 33       # 0-indexed = Excel row 34
    T2_DATA_START_IDX = 34    # 0-indexed = Excel row 35
    T2_DATA_END_IDX = T2_DATA_START_IDX + num_themes - 1

    T3_LABEL_IDX = T2_DATA_END_IDX + 2       # blank row then label
    T3_HDR_IDX = T3_LABEL_IDX + 1
    T3_DATA_START_IDX = T3_HDR_IDX + 1

    # Excel 1-indexed rows for formulas
    T3_DATA_START_XL = T3_DATA_START_IDX + 1  # e.g. 47
    T1_HDR_ROW_XL = 26                         # Excel row 26 = Table 1 headers
    T2_HDR_ROW_XL = 34                         # Excel row 34 = Table 2 headers
    HTML_COUNT_REF = "B13"                      # HTML count in overview table

    # BUILD WORKBOOK
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'in_memory': True, 'nan_inf_to_errors': True})
    ws = wb.add_worksheet('URL Issues')

    # FORMATS
    RED = '#FF0000'
    WHITE = '#FFFFFF'
    BLACK = '#000000'
    DARK = '#404040'
    FONT = 'Rockwell'
    FONT_OV = 'Aptos Narrow'
    SZ = 8
    SZ_OV = 11

    def f(**kw):
        return wb.add_format(kw)

    f_red = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
               bg_color=RED, border=1, align='center', valign='vcenter', text_wrap=True)
    f_red_lft = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
                  bg_color=RED, border=1, align='left', valign='vcenter', text_wrap=True)
    f_dark = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
               bg_color=DARK, border=1, align='center', valign='vcenter', text_wrap=True)
    f_dark_lft = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
                   bg_color=DARK, border=1, align='left', valign='vcenter')
    f_section = f(bold=True, font_name=FONT, font_size=SZ, font_color=BLACK)
    f_ov_hdr = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
                 bg_color=RED, border=1, align='center', valign='vcenter')
    f_ov_lft = f(font_name=FONT_OV, font_size=SZ_OV, font_color=BLACK,
                 border=1, align='left', valign='vcenter')
    f_ov_data = f(font_name=FONT_OV, font_size=SZ_OV, font_color=BLACK,
                  border=1, align='center', valign='vcenter')
    f_ov_pct = f(font_name=FONT_OV, font_size=SZ_OV, font_color=BLACK,
                 border=1, align='center', valign='vcenter', num_format='0.00%')
    f_ctr = f(font_name=FONT, font_size=SZ, font_color=BLACK,
               bg_color=WHITE, border=1, align='center', valign='vcenter')
    f_lft = f(font_name=FONT, font_size=SZ, font_color=BLACK,
               bg_color=WHITE, border=1, align='left', valign='vcenter')
    f_rgt = f(font_name=FONT, font_size=SZ, font_color=BLACK,
               bg_color=WHITE, border=1, align='right', valign='vcenter')
    f_bold_lft = f(bold=True, font_name=FONT, font_size=SZ, font_color=BLACK,
                   bg_color=WHITE, border=1, align='left', valign='vcenter')
    f_bold_ctr = f(bold=True, font_name=FONT, font_size=SZ, font_color=BLACK,
                   bg_color=WHITE, border=1, align='center', valign='vcenter')
    f_num = f(font_name=FONT, font_size=SZ, font_color=BLACK,
               bg_color=WHITE, border=1, align='center', valign='vcenter', num_format='0')
    f_pct = f(font_name=FONT, font_size=SZ, font_color=BLACK,
               bg_color=WHITE, border=1, align='center', valign='vcenter', num_format='0.00%')
    f_t2_merged = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
                    bg_color=RED, border=1, align='left', valign='vcenter')

    # COLUMN WIDTHS (from template)
    ws.set_column('A:A', 24.27)
    ws.set_column('B:B', 22.45)
    ws.set_column('C:C', 29.82)
    ws.set_column('D:D', 19.18)
    ws.set_column('E:E', 20.45)
    ws.set_column('F:F', 18.54)
    ws.set_column('G:G', 18.18)
    ws.set_column('H:H', 18.54)
    ws.set_column('I:I', 15.82)
    ws.set_column('J:J', 18.54)

    # ROWS 2-5: Issue Summary box
    f_summary_box = f(font_name=FONT, font_size=8, font_color=BLACK,
                      text_wrap=True, valign='top', border=1)
    f_summary_title = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                        text_wrap=True, valign='top', border=1)
    summary_text = ('Issue Summary:  URLs are the web addresses of webpages that help users '
                    'and search engines understand page content and site structure. URL issues include '
                    'problems related to structure, readability, duplication, parameters, indexing, or '
                    'crawlability that may impact SEO performance and user experience.')
    ws.merge_range(1, 0, 4, 9, summary_text, f_summary_box)
    ws.set_row(1, 50)
    ws.set_row(2, 15)
    ws.set_row(3, 15)
    ws.set_row(4, 15)

    # ROW 7 (idx 6): Summary Table label
    f_summary_lbl = f(bold=True, font_name=FONT, font_size=10, font_color=BLACK)
    ws.write(6, 0, 'Summary Table', f_summary_lbl)

    # ROW 9 (idx 8): Overview Table label
    ws.set_row(8, 15)
    ws.write(8, 0, 'Internal URLs Overview Table', f_section)

    # ROW 10 (idx 9): Overview headers
    ws.set_row(9, 18)
    ws.write(9, 0, 'Summary', f_ov_hdr)
    ws.write(9, 1, 'URLs', f_ov_hdr)
    ws.write(9, 2, '% of Total', f_ov_hdr)
    ws.write(9, 3, 'Total URLs', f_ov_hdr)

    # ROW 11 (idx 10): Internal label
    ws.set_row(10, 14.5)
    ws.write(10, 0, 'Internal', f_ov_lft)

    # ROWS 12-20 (idx 11-19): Overview data
    overview_rows = [
        ('All', overview_data['All']),
        ('HTML', overview_data['HTML']),
        ('JavaScript', overview_data['JavaScript']),
        ('CSS', overview_data['CSS']),
        ('Images', overview_data['Images']),
        ('PDF', overview_data['PDF']),
        ('Flash', overview_data['Flash']),
        ('Other', overview_data['Other']),
        ('Unknown', overview_data['Unknown']),
    ]
    for i, (label, count) in enumerate(overview_rows):
        r = 11 + i
        ws.set_row(r, 14.5)
        ws.write(r, 0, label, f_ov_lft)
        ws.write(r, 1, count, f_ov_data)
        pct = count / overview_data['All'] if overview_data['All'] > 0 else 0
        ws.write(r, 2, pct, f_ov_pct)
        ws.write(r, 3, overview_data['All'], f_ov_data)

    # ROW 25 (idx 24): Table 1 label
    ws.write(24, 0, 'Table 1', f_section)

    # ROW 26 (idx 25): Table 1 headers
    ws.write(25, 0, 'URL Issue Types', f_red_lft)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(25, i + 1, issue, f_red)

    # ROW 27 (idx 26): Issue Priority (hardcoded per spec)
    priorities = ['Low', 'Low', 'Medium', 'Medium', 'High', 'Medium']
    ws.write(26, 0, 'Issue Priority', f_red_lft)
    for i, p in enumerate(priorities):
        ws.write(26, i + 1, p, f_ctr)

    # ROW 28 (idx 27): #Affected URLs
    # Dynamic formula: COUNTIF from T3_DATA_START to end of sheet
    ws.write(27, 0, '#Affected URLs', f_red_lft)
    for i, col_l in enumerate(['B', 'C', 'D', 'E', 'F', 'G']):
        hdr_ref = col_l + str(T1_HDR_ROW_XL)   # e.g. B26
        formula = f'=COUNTIF(A{T3_DATA_START_XL}:A1048576,{hdr_ref})'
        ws.write_formula(27, i + 1, formula, f_ctr)

    # ROW 29 (idx 28): % Share = #Affected / HTML count
    ws.set_row(28, 31.5)
    ws.write(28, 0, '% Share against Total  Indexable and HTML URLs Crawled', f_red_lft)
    for i, col_l in enumerate(['B', 'C', 'D', 'E', 'F', 'G']):
        affected_ref = col_l + '28'
        formula = f'=IF({HTML_COUNT_REF}>0,{affected_ref}/{HTML_COUNT_REF},"")'
        ws.write_formula(28, i + 1, formula, f_pct)

    # ROW 32 (idx 31): Table 2 label
    ws.write(31, 0, 'Table 2', f_section)

    # ROW 33 (idx 32): Table 2 title merged A33:H33
    ws.set_row(32, 10.5)
    ws.merge_range(32, 0, 32, 7, 'Page Theme Wise URL Analysis ', f_t2_merged)

    # ROW 34 (idx 33): Table 2 headers
    ws.write(T2_HDR_ROW_IDX, 0, 'Page Theme', f_dark_lft)
    ws.write(T2_HDR_ROW_IDX, 1, 'Priority', f_dark)
    for i, issue in enumerate(ISSUE_KEYS):
        ws.write(T2_HDR_ROW_IDX, i + 2, issue, f_dark)

    # TABLE 2 DATA: theme rows with dynamic COUNTIFS formulas
    for row_offset, (theme, priority) in enumerate(sorted_themes):
        r = T2_DATA_START_IDX + row_offset
        ws.set_row(r, 14.5)
        theme_cell = 'A' + str(r + 1)   # e.g. A35
        ws.write(r, 0, theme, f_bold_lft)
        ws.write(r, 1, priority, f_bold_ctr)
        for i, col_l in enumerate(['C', 'D', 'E', 'F', 'G', 'H']):
            hdr_ref = col_l + str(T2_HDR_ROW_XL)   # e.g. C34
            formula = (
                f'=COUNTIFS('
                f'A{T3_DATA_START_XL}:A1048576,{hdr_ref},'
                f'C{T3_DATA_START_XL}:C1048576,{theme_cell})'
            )
            ws.write_formula(r, i + 2, formula, f_ctr)

    # TABLE 3 label
    ws.write(T3_LABEL_IDX, 0, 'Table 3', f_section)

    # TABLE 3 headers
    t3_headers = [
        'Error Type', 'Address', 'Page Theme 1', 'Page Theme 2',
        'Content Type', 'Status Code', 'Indexability',
        'Impressions', 'Clicks', 'Organic Sessions '
    ]
    for i, h in enumerate(t3_headers):
        ws.write(T3_HDR_IDX, i, h, f_red if i > 0 else f_red_lft)

    # TABLE 3 DATA
    for row_offset, (_, row) in enumerate(df_table3.iterrows()):
        r = T3_DATA_START_IDX + row_offset
        ws.set_row(r, 14.5)
        ws.write(r, 0, row.get("Error Type", "") or "", f_lft)
        ws.write(r, 1, row.get("Address", "") or "", f_lft)
        ws.write(r, 2, row.get("Page Theme 1", "") or "", f_ctr)
        pt2 = row.get("Page Theme 2", "-") or "-"
        ws.write(r, 3, pt2, f_ctr)
        ws.write(r, 4, row.get("Content Type", "") or "", f_lft)
        ws.write(r, 5, 200, f_num)
        ws.write(r, 6, row.get("Indexability", "") or "", f_ctr)
        ws.write(r, 7, safe_num(row.get("Impressions")), f_rgt)
        ws.write(r, 8, safe_num(row.get("Clicks")), f_rgt)
        ws.write(r, 9, safe_num(row.get("Organic Sessions")), f_rgt)

    wb.close()

    buf.seek(0)
    return buf.read()