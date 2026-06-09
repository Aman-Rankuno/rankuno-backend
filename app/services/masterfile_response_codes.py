import os
import io
import math
import pandas as pd
import xlsxwriter
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

SC_KEYS = ["301", "302", "304", "307", "308",
           "400", "401", "403", "404",
           "500", "502", "503", "504", "No response code"]

SC_PRIORITY = {
    "301": "Medium", "302": "High", "304": "Medium",
    "307": "High", "308": "Medium",
    "400": "High", "401": "High", "403": "High", "404": "High",
    "500": "High", "502": "High", "503": "High", "504": "High",
    "No response code": "High",
}


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f) or f == 0) else f
    except Exception:
        return None


def _clean(v):
    if v is None:
        return "-"
    import math as _math
    try:
        if _math.isnan(float(v)):
            return "-"
    except Exception:
        pass
    s = str(v).strip()
    return "-" if (not s or s == "nan") else s


def _get_page_type(content_type: str) -> str:
    ct = str(content_type).lower()
    if "html" in ct: return "HTML"
    if "pdf" in ct: return "PDF"
    if ct.startswith("image/"): return "Image"
    if "css" in ct: return "CSS"
    if "javascript" in ct: return "JavaScript"
    if ct and ct != "nan": return "Other"
    return ""


def build_response_codes_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "Response_codes-Internal.xlsx")
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

    df_3xx = load_csv("response_codes_internal_redirection_(3xx).csv")
    df_4xx = load_csv("response_codes_internal_client_error_(4xx).csv")
    df_5xx = load_csv("response_codes_internal_server_error_(5xx).csv")
    df_no_response = load_csv("response_codes_internal_no_response.csv")
    df_redirect_chain = load_csv("response_codes_internal_redirect_chain.csv")
    df_redirect_loop = load_csv("response_codes_internal_redirect_loop.csv")
    df_blocked = load_csv("response_codes_internal_blocked_by_robots_txt.csv")
    df_internal_all = load_csv("internal_all.csv")
    df_gsc = load_csv("search_console_all.csv")
    df_ga = load_csv("analytics_all.csv")

    # Internal all counts
    total_all = len(df_internal_all) if not df_internal_all.empty else 0
    total_html = total_js = total_css = 0
    total_images = total_pdf = total_other = total_unknown = 0

    if not df_internal_all.empty:
        def gc(df, *names):
            for n in names:
                m = next((c for c in df.columns if c.lower() == n.lower()), None)
                if m: return m
            return None
        ct_col = gc(df_internal_all, "Content Type")
        if ct_col:
            for ct in df_internal_all[ct_col].astype(str):
                pt = _get_page_type(ct)
                if pt == "HTML": total_html += 1
                elif pt == "JavaScript": total_js += 1
                elif pt == "CSS": total_css += 1
                elif pt == "Image": total_images += 1
                elif pt == "PDF": total_pdf += 1
                elif pt == "Other": total_other += 1
                else: total_unknown += 1

    # Redirect chain/loop lookup sets
    chain_urls = set()
    loop_urls = set()
    if not df_redirect_chain.empty:
        col = next((c for c in df_redirect_chain.columns if c.lower() == "address"), None)
        if col:
            chain_urls = set(df_redirect_chain[col].dropna().astype(str))
    if not df_redirect_loop.empty:
        col = next((c for c in df_redirect_loop.columns if c.lower() == "address"), None)
        if col:
            loop_urls = set(df_redirect_loop[col].dropna().astype(str))

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

    # Tag and combine frames
    def tag(df, label):
        if df.empty: return df
        df = df.copy()
        df["_error_type"] = label
        return df

    frames = [df for df in [
        tag(df_3xx, "3xx"), tag(df_4xx, "4xx"), tag(df_5xx, "5xx"),
        tag(df_no_response, "No response code"), tag(df_blocked, "Blocked by robots")
    ] if not df.empty]

    if not frames:
        # No data - return empty workbook
        buf = io.BytesIO()
        wb = xlsxwriter.Workbook(buf, {'in_memory': True})
        ws = wb.add_worksheet('Response codes-Internal')
        ws.write(0, 0, 'No response code data found for this crawl.')
        wb.close()
        buf.seek(0)
        return buf.read()

    df_combined = pd.concat(frames, ignore_index=True)

    def get_col(df, *names):
        for n in names:
            m = next((c for c in df.columns if c.lower() == n.lower()), None)
            if m: return m
        return None

    addr_col = get_col(df_combined, "Address")
    sc_col = get_col(df_combined, "Status Code")
    st_col = get_col(df_combined, "Status")
    idx_col = get_col(df_combined, "Indexability")
    idx_st_col = get_col(df_combined, "Indexability Status")
    ct_col = get_col(df_combined, "Content Type")
    inl_col = get_col(df_combined, "Inlinks")
    ru_col = get_col(df_combined, "Redirect URL")
    rt_col = get_col(df_combined, "Redirect Type")

    # Build Table 3 rows
    rows = []
    for _, row in df_combined.iterrows():
        url = str(row.get(addr_col, "")) if addr_col else ""
        error_type = str(row.get("_error_type", ""))
        sc = str(row.get(sc_col, "")) if sc_col else ""
        page_theme1, page_theme2, _, _ = classify_url(url, rulebook)
        gsc = gsc_map.get(url, {})
        content_type = _clean(row.get(ct_col, "")) if ct_col else "-"

        redirect_chain = "TRUE" if (error_type == "3xx" and url in chain_urls) else ""
        redirect_loop = "TRUE" if (error_type == "3xx" and url in loop_urls) else ""

        # Normalize error type label
        if error_type == "3xx":
            err_label = "3xx"
        elif error_type == "4xx":
            err_label = "4xx"
        elif error_type == "5xx":
            err_label = "5xx"
        elif error_type == "No response code":
            err_label = "No response code"
        else:
            err_label = error_type

        rows.append({
            "Error type": err_label,
            "Address": url,
            "Page Theme 1": page_theme1 or "",
            "Page Theme 2": page_theme2 if page_theme2 else "-",
            "Content Type": content_type,
            "Status Code": safe_num(sc) or sc,
            "Status": _clean(row.get(st_col, "")) if st_col else "-",
            "Indexability": _clean(row.get(idx_col, "")) if idx_col else "-",
            "Indexability Status": _clean(row.get(idx_st_col, "")) if idx_st_col else "-",
            "Inlinks": safe_num(row.get(inl_col, 0)) if inl_col else None,
            "Redirect URL": _clean(row.get(ru_col, "")) if ru_col else "-",
            "Redirect Type": _clean(row.get(rt_col, "")) if rt_col else "-",
            "Redirect chain": redirect_chain,
            "Redirect loop": redirect_loop,
            "Impressions": gsc.get("impressions"),
            "Clicks": gsc.get("clicks"),
            "Organic Sessions": ga_map.get(url),
        })

    df_table3 = pd.DataFrame(rows)
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

    # Redirect chain/loop counts for Table 4
    chain_count = len(df_redirect_chain) if not df_redirect_chain.empty else 0
    loop_count = len(df_redirect_loop) if not df_redirect_loop.empty else 0
    total_3xx = len(df_3xx) if not df_3xx.empty else 0

    # Fixed row positions (0-indexed)
    # Rows 0-2: Issue Summary (rows 1-3 in Excel)
    # Row 3 (idx): Overview Table label = Excel row 4
    # Row 4 (idx): Overview headers = Excel row 5
    # Row 5 (idx): Internal label = Excel row 6
    # Rows 6-14 (idx): Overview data = Excel rows 7-15
    # Row 15 (idx): Table 1 label = Excel row 16
    # Row 16 (idx): Table 1 headers = Excel row 17
    # Row 17 (idx): Issue Priority = Excel row 18
    # Row 18 (idx): #Affected URLs = Excel row 19
    # Row 19 (idx): % Share = Excel row 20
    # Row 20 (idx): Table 2 label = Excel row 21
    # Row 21 (idx): Table 2 title = Excel row 22
    # Row 22 (idx): Table 2 headers = Excel row 23
    # Rows 23+ (idx): Table 2 data = Excel rows 24+

    T2_HDR_IDX = 22
    T2_DATA_START_IDX = 23
    T2_DATA_END_IDX = T2_DATA_START_IDX + num_themes - 1

    T3_LABEL_IDX = T2_DATA_END_IDX + 2
    T3_HDR_IDX = T3_LABEL_IDX + 1
    T3_DATA_START_IDX = T3_HDR_IDX + 1

    # Excel 1-indexed rows for formulas
    T3_DATA_START_XL = T3_DATA_START_IDX + 1
    T2_HDR_XL = T2_HDR_IDX + 1           # row 23
    T1_HDR_XL = 17                         # Table 1 header row in Excel
    INTERNAL_ALL_REF = "B7"               # Internal All count in overview table

    # BUILD WORKBOOK
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'in_memory': True, 'nan_inf_to_errors': True})
    ws = wb.add_worksheet('Response codes-Internal')

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
    f_summary = f(font_name=FONT, font_size=SZ, font_color=BLACK,
                  text_wrap=True, valign='top', border=1)
    f_t2_title = f(bold=True, font_name=FONT, font_size=SZ, font_color=WHITE,
                   bg_color=RED, border=1, align='left', valign='vcenter')

    # COLUMN WIDTHS
    ws.set_column('A:A', 19.90)
    ws.set_column('B:B', 21.90)
    ws.set_column('C:C', 11.63)
    ws.set_column('D:D', 11.90)
    ws.set_column('E:E', 9.63)
    ws.set_column('F:F', 10.45)
    ws.set_column('G:G', 13.0)
    ws.set_column('H:H', 13.18)
    ws.set_column('I:I', 9.63)
    ws.set_column('J:J', 10.73)
    ws.set_column('K:K', 11.09)
    ws.set_column('L:L', 10.27)
    ws.set_column('M:M', 9.73)
    ws.set_column('N:N', 13.82)
    ws.set_column('O:O', 13.18)
    ws.set_column('P:P', 13.0)
    ws.set_column('Q:Q', 13.0)

    # ROWS 1-2 (idx 0-1): Issue Summary merged A:F
    summary_text = (
        'Issue Summary '
        'Response code issues occur when URLs return redirects, errors, server issues, or no response '
        'instead of loading properly. These issues can affect user experience, crawling, indexing, and '
        'website performance. This report helps identify affected URLs and prioritize fixes basis '
        'page importance and performance history.'
    )
    ws.merge_range(0, 0, 1, 5, summary_text, f_summary)
    ws.set_row(0, 56)
    ws.set_row(1, 56)

    # ROW 4 (idx 3): Overview Table label
    ws.set_row(3, 14.5)
    ws.write(3, 0, 'Internal URLs Overview Table', f_section)

    # ROW 5 (idx 4): Overview headers
    ws.set_row(4, 14.5)
    ws.write(4, 0, 'Summary', f_ov_hdr)
    ws.write(4, 1, 'URLs', f_ov_hdr)
    ws.write(4, 2, '% of Total', f_ov_hdr)
    ws.write(4, 3, 'Total URLs', f_ov_hdr)

    # ROW 6 (idx 5): Internal label
    ws.set_row(5, 14.5)
    ws.write(5, 0, 'Internal', f_ov_lft)

    # ROWS 7-15 (idx 6-14): Overview data
    ov_rows = [
        ('All', total_all),
        ('HTML', total_html),
        ('JavaScript', total_js),
        ('CSS', total_css),
        ('Images', total_images),
        ('PDF', total_pdf),
        ('Flash', 0),
        ('Other', total_other),
        ('Unknown', total_unknown),
    ]
    for i, (label, count) in enumerate(ov_rows):
        r = 6 + i
        ws.set_row(r, 14.5)
        ws.write(r, 0, label, f_ov_lft)
        ws.write(r, 1, count, f_ov_data)
        pct = count / total_all if total_all > 0 else 0
        ws.write(r, 2, pct, f_ov_pct)
        ws.write(r, 3, total_all, f_ov_data)

    # ROW 16 (idx 15): Table 1 label
    ws.set_row(15, 56)
    ws.write(15, 0, 'Table 1', f_section)

    # ROW 17 (idx 16): Table 1 headers
    ws.set_row(16, 14.5)
    ws.write(16, 1, 'URL Issue Types', f_red_lft)
    for i, sc in enumerate(SC_KEYS):
        ws.write(16, i + 2, sc, f_red)
    # Table 4 headers (redirect chain/loop) at col V(21), W(22), X(23)
    ws.write(16, 21, 'URL Issue Types', f_red_lft)
    ws.write(16, 22, 'Redirect chain', f_red)
    ws.write(16, 23, 'Redirect loop', f_red)

    # ROW 18 (idx 17): Issue Priority
    ws.set_row(17, 14.5)
    ws.write(17, 1, 'Issue Priority', f_red_lft)
    for i, sc in enumerate(SC_KEYS):
        ws.write(17, i + 2, SC_PRIORITY.get(sc, 'High'), f_ctr)
    ws.write(17, 21, 'Issue Priority', f_red_lft)
    ws.write(17, 22, 'High', f_ctr)
    ws.write(17, 23, 'High', f_ctr)

    # ROW 19 (idx 18): #Affected URLs
    # Table 1 sums from Table 2 rows (same as template: =SUM(C24:C33))
    # Dynamic: sum from T2_DATA_START_XL to T2_DATA_END_XL
    T2_DATA_START_XL = T2_DATA_START_IDX + 1
    T2_DATA_END_XL = T2_DATA_END_IDX + 1

    ws.set_row(18, 14.5)
    ws.write(18, 1, '#Affected URLs', f_red_lft)
    for i, sc in enumerate(SC_KEYS):
        col_l = xlsxwriter.utility.xl_col_to_name(i + 2)
        formula = f'=SUM({col_l}{T2_DATA_START_XL}:{col_l}{T2_DATA_END_XL})'
        ws.write_formula(18, i + 2, formula, f_ctr)
    # Table 4
    ws.write(18, 21, '#Affected URLs', f_red_lft)
    ws.write(18, 22, chain_count if chain_count > 0 else None, f_ctr)
    ws.write(18, 23, loop_count if loop_count > 0 else None, f_ctr)

    # ROW 20 (idx 19): % Share
    ws.set_row(19, 31.5)
    ws.write(19, 1, '% Share against Total  URLs Crawled', f_red_lft)
    for i, sc in enumerate(SC_KEYS):
        col_l = xlsxwriter.utility.xl_col_to_name(i + 2)
        formula = f'=IF({INTERNAL_ALL_REF}>0,{col_l}19/{INTERNAL_ALL_REF},"")'
        ws.write_formula(19, i + 2, formula, f_pct)
    # Table 4 % share against total 3xx
    ws.write(19, 21, '% share against Total  3xx URLs Crawled', f_red_lft)
    pct_chain = chain_count / total_3xx if total_3xx > 0 and chain_count > 0 else None
    pct_loop = loop_count / total_3xx if total_3xx > 0 and loop_count > 0 else None
    ws.write(19, 22, pct_chain, f_pct)
    ws.write(19, 23, pct_loop, f_pct)

    # ROW 21 (idx 20): Table 2 label
    ws.set_row(20, 56)
    ws.write(20, 0, 'Table 2', f_section)

    # ROW 22 (idx 21): Table 2 title
    ws.set_row(21, 14.5)
    ws.merge_range(21, 0, 21, 16, 'Page Theme Wise URL Analysis ', f_t2_title)

    # ROW 23 (idx 22): Table 2 headers
    ws.set_row(22, 14.5)
    ws.write(22, 0, 'Page Theme 1', f_dark_lft)
    ws.write(22, 1, 'Priority Basis Page Theme 1', f_dark)
    for i, sc in enumerate(SC_KEYS):
        ws.write(22, i + 2, sc, f_dark)

    # TABLE 2 DATA with dynamic COUNTIFS
    for row_offset, (theme, priority) in enumerate(sorted_themes):
        r = T2_DATA_START_IDX + row_offset
        ws.set_row(r, 14.5)
        theme_cell = 'A' + str(r + 1)
        ws.write(r, 0, theme, f_bold_lft)
        ws.write(r, 1, priority, f_bold_ctr)
        for i, sc in enumerate(SC_KEYS):
            col_l = xlsxwriter.utility.xl_col_to_name(i + 2)
            hdr_ref = col_l + str(T2_HDR_XL)
            if sc == "No response code":
                formula = (
                    f'=COUNTIFS('
                    f'A{T3_DATA_START_XL}:A1048576,{hdr_ref},'
                    f'C{T3_DATA_START_XL}:C1048576,{theme_cell})'
                )
            else:
                formula = (
                    f'=COUNTIFS('
                    f'F{T3_DATA_START_XL}:F1048576,{hdr_ref},'
                    f'C{T3_DATA_START_XL}:C1048576,{theme_cell})'
                )
            ws.write_formula(r, i + 2, formula, f_ctr)

    # TABLE 3 label
    ws.set_row(T3_LABEL_IDX, 23)
    ws.write(T3_LABEL_IDX, 0, 'Table 3', f_section)

    # TABLE 3 headers
    t3_headers = [
        'Error type', 'Address', 'Page Theme 1', 'Page Theme 2',
        'Content Type', 'Status Code', 'Status', 'Indexability',
        'Indexability Status', 'Inlinks', 'Redirect URL', 'Redirect Type',
        'Redirect chain', 'Redirect loop', 'Impressions', 'Clicks', 'Organic Sessions'
    ]
    ws.set_row(T3_HDR_IDX, 14.5)
    for i, h in enumerate(t3_headers):
        ws.write(T3_HDR_IDX, i, h, f_red if i > 0 else f_red_lft)

    # TABLE 3 DATA
    t3_cols = [
        "Error type", "Address", "Page Theme 1", "Page Theme 2",
        "Content Type", "Status Code", "Status", "Indexability",
        "Indexability Status", "Inlinks", "Redirect URL", "Redirect Type",
        "Redirect chain", "Redirect loop", "Impressions", "Clicks", "Organic Sessions"
    ]
    for row_offset, (_, row) in enumerate(df_table3.iterrows()):
        r = T3_DATA_START_IDX + row_offset
        ws.set_row(r, 14.5)
        for ci, col_name in enumerate(t3_cols):
            val = row.get(col_name, "")
            if val is None or (isinstance(val, float) and math.isnan(val)):
                val = ""
            if ci == 5:  # Status Code - number format
                try:
                    sc_val = int(float(val)) if val and val != "" else None
                except Exception:
                    sc_val = val
                ws.write(r, ci, sc_val, f_num)
            elif ci in [9]:  # Inlinks - number
                ws.write(r, ci, safe_num(val), f_rgt)
            elif ci in [14, 15, 16]:  # Impressions, Clicks, Sessions
                ws.write(r, ci, safe_num(val), f_rgt)
            elif ci in [0, 1, 4, 6, 7, 8, 10, 11]:  # left aligned
                ws.write(r, ci, str(val) if val else "", f_lft)
            else:  # center aligned
                ws.write(r, ci, str(val) if val else "", f_ctr)

    wb.close()
    buf.seek(0)
    return buf.read()