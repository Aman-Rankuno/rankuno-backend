import io
import math
import os
import pandas as pd
import xlsxwriter
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

H1_ISSUES = [
    ("Missing",           "h1_missing.csv",           "High"),
    ("Duplicate",         "h1_duplicate.csv",          "Low"),
    ("Over 70 Characters","h1_over_70_characters.csv", "Low"),
    ("Multiple",          "h1_multiple.csv",           "Low"),
]


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f) or f == 0) else f
    except Exception:
        return None


def _clean(v):
    s = str(v).strip()
    return "-" if (not s or s == "nan") else s


def _load_csv(report_dir, filename):
    path = os.path.join(report_dir, filename)
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_csv(path, encoding="utf-8", low_memory=False)
    except UnicodeDecodeError:
        df = pd.read_csv(path, encoding="latin-1", low_memory=False)
    return df


def _filter_indexable_200_html(df):
    if df.empty:
        return df
    out = df.copy()
    if "Indexability" in out.columns:
        out = out[out["Indexability"].astype(str).str.strip() == "Indexable"]
    if "Status Code" in out.columns:
        out = out[out["Status Code"].astype(str).str.strip() == "200"]
    if "Content Type" in out.columns:
        out = out[out["Content Type"].astype(str).str.contains("text/html", na=False)]
    return out


def generate(report_dir: str, domain: str) -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws = wb.add_worksheet("H1")

    RED = "#FF0000"
    GREY_HDR = "#F2F2F2"
    WHITE = "#FFFFFF"

    def fmt(**kw):
        return wb.add_format(kw)

    f_issue_summary = fmt(bold=True, font_size=10, valign="top", text_wrap=True, border=1)
    f_section_label = fmt(bold=True, font_size=10)
    f_t1_hdr = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter", text_wrap=True)
    f_t1_lbl = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, valign="vcenter")
    f_t1_val = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter")
    f_t1_pct = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter", num_format="0.00%")
    f_t2_title = fmt(bold=True, font_color=WHITE, bg_color=RED,
                     border=1, valign="vcenter")
    f_t2_hdr = fmt(bold=True, bg_color=GREY_HDR, border=1,
                   valign="vcenter", text_wrap=True)
    f_t2_cell = fmt(border=1, valign="vcenter")
    f_t2_num = fmt(border=1, valign="vcenter", align="center")
    f_t3_hdr = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, valign="vcenter", text_wrap=True)
    f_t3_cell = fmt(border=1, valign="vcenter", text_wrap=True)
    f_t3_num = fmt(border=1, valign="vcenter", align="center")

    ws.set_column(0, 0, 20)
    ws.set_column(1, 1, 45)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 18)
    ws.set_column(4, 4, 25)
    ws.set_column(5, 5, 12)
    ws.set_column(6, 6, 14)
    ws.set_column(7, 7, 18)
    ws.set_column(8, 8, 10)
    ws.set_column(9, 9, 35)
    ws.set_column(10, 10, 12)
    ws.set_column(11, 11, 35)
    ws.set_column(12, 12, 12)
    ws.set_column(13, 13, 13)
    ws.set_column(14, 14, 10)
    ws.set_column(15, 15, 15)

    internal_df = _load_csv(report_dir, "internal_all.csv")
    gsc_df = _load_csv(report_dir, "search_console_all.csv")
    ga4_df = _load_csv(report_dir, "analytics_all.csv")

    gsc_map = {}
    if not gsc_df.empty and "Address" in gsc_df.columns:
        for _, row in gsc_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if addr:
                gsc_map[addr] = {
                    "Impressions": safe_num(row.get("Impressions")),
                    "Clicks": safe_num(row.get("Clicks")),
                }

    ga4_map = {}
    if not ga4_df.empty and "Address" in ga4_df.columns:
        for _, row in ga4_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if addr:
                ga4_map[addr] = safe_num(
                    row.get("GA4 Sessions") or row.get("Sessions") or
                    row.get("Organic Sessions") or row.get("Organic sessions")
                )

    internal_lookup = {}
    if not internal_df.empty and "Address" in internal_df.columns:
        for _, row in internal_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if addr:
                internal_lookup[addr] = {
                    "content_type": _clean(row.get("Content Type", "")),
                    "status_code": _clean(row.get("Status Code", "")),
                    "indexability": _clean(row.get("Indexability", "")),
                    "indexability_status": _clean(row.get("Indexability Status", "")),
                }

    try:
        rulebook = load_rulebook(domain)
    except Exception:
        rulebook = []

    internal_filtered = _filter_indexable_200_html(internal_df)
    total_indexable = len(internal_filtered) if not internal_filtered.empty else 0

    theme_totals = {}
    if not internal_filtered.empty and "Address" in internal_filtered.columns:
        for _, row in internal_filtered.iterrows():
            addr = str(row.get("Address", "")).strip()
            t1, t2, lang, pri = classify_url(addr, rulebook)
            theme = t1 if t1 else "-"
            theme_totals[theme] = theme_totals.get(theme, 0) + 1

    # Build T3 rows
    t3_rows = []
    for issue_label, csv_name, _ in H1_ISSUES:
        df = _load_csv(report_dir, csv_name)
        df = _filter_indexable_200_html(df)
        if df.empty:
            continue
        for _, row in df.iterrows():
            addr = str(row.get("Address", "")).strip()
            t1, t2, lang, pri = classify_url(addr, rulebook)
            _il = internal_lookup.get(addr, {})
            content_type = _il.get("content_type") or _clean(row.get("Content Type", ""))
            status_code = _il.get("status_code") or _clean(row.get("Status Code", ""))
            indexability = _il.get("indexability") or _clean(row.get("Indexability", ""))
            idx_status = _il.get("indexability_status") or _clean(row.get("Indexability Status", ""))
            gsc = gsc_map.get(addr, {})
            sessions = ga4_map.get(addr)
            t3_rows.append({
                "error_type": issue_label,
                "address": addr,
                "theme1": t1 if t1 else "-",
                "theme2": t2 if t2 else "-",
                "content_type": content_type,
                "status_code": status_code,
                "indexability": indexability,
                "idx_status": idx_status,
                "occurrence": safe_num(row.get("Occurrences")) or safe_num(row.get("Occurrence")) or "-",
                "h1_1": _clean(row.get("H1-1", "")),
                "h1_1_len": safe_num(row.get("H1-1 Length")) or "-",
                "h1_2": _clean(row.get("H1-2", "")),
                "h1_2_len": safe_num(row.get("H1-2 Length")) or "-",
                "impressions": gsc.get("Impressions"),
                "clicks": gsc.get("Clicks"),
                "sessions": sessions,
            })

    t3_rows.sort(key=lambda r: (r["impressions"] is None, -(r["impressions"] or 0)))

    # Build T2 theme data
    R_T2_DATA_START = 17
    theme_data = {}
    for r in t3_rows:
        th = r["theme1"]
        if th not in theme_data:
            theme_data[th] = {"Missing": 0, "Duplicate": 0,
                               "Over 70 Characters": 0, "Multiple": 0, "priority_basis": "N/A"}
        theme_data[th][r["error_type"]] += 1

    for _, row in (internal_filtered.iterrows() if not internal_filtered.empty else iter([])):
        addr = str(row.get("Address", "")).strip()
        t1, t2, lang, pri = classify_url(addr, rulebook)
        theme = t1 if t1 else "-"
        if theme in theme_data:
            existing = theme_data[theme]["priority_basis"]
            if existing == "N/A" or (pri and PRIORITY_ORDER.get(pri, 99) < PRIORITY_ORDER.get(existing, 99)):
                theme_data[theme]["priority_basis"] = pri if pri else "N/A"

    t2_rows_sorted = sorted(
        theme_data.items(),
        key=lambda x: PRIORITY_ORDER.get(x[1]["priority_basis"], 3)
    )

    t2_count = len(t2_rows_sorted)
    R_TABLE3_LABEL = R_T2_DATA_START + t2_count + 1
    R_T3_HDR = R_TABLE3_LABEL + 1
    R_T3_DATA_START = R_T3_HDR + 1
    t3_data_excel_row = R_T3_DATA_START + 1

    # Issue Summary rows 2-4
    ws.set_row(0, 15)
    ws.set_row(1, 55)
    ws.set_row(2, 40)
    ws.set_row(3, 40)
    summary_text = (
        "Issue Summary:\n"
        "1. Missing - URLs missing a <title> tag, resulting in search engines generating their own title for "
        "search results, which may reduce relevance and click-through rates.\n"
        "2. Duplicate - URLs sharing the same meta title as other pages on the website, making it difficult "
        "for search engines to differentiate page relevance and uniqueness.\n"
        "3. Over 70 Characters - URLs with meta titles exceeding the recommended character limit, which may "
        "result in truncation within search engine results pages (SERPs).\n"
        "4. Multiple - URLs containing multiple <title> tags, potentially causing search engines to "
        "misinterpret the preferred page title."
    )
    ws.merge_range(1, 0, 3, 15, summary_text, f_issue_summary)

    ws.write(7, 0, "Summary Table ", f_section_label)
    ws.write(8, 0, "Table 1")

    # Table 1 headers (5 cols, NO Total Pages)
    t1_headers = ["H1 Issue Types", "Missing", "Duplicate", "Over 70 Characters", "Multiple"]
    for ci, h in enumerate(t1_headers):
        ws.write(9, ci, h, f_t1_hdr)

    t1_priorities = ["Issue Priority", "High", "Low", "Low", "Low"]
    for ci, v in enumerate(t1_priorities):
        ws.write(10, ci, v, f_t1_hdr)

    ws.write(11, 0, "#Affected URLs", f_t1_lbl)
    col_a = "A{}:A1048576".format(t3_data_excel_row)
    for ci, label in enumerate(["Missing", "Duplicate", "Over 70 Characters", "Multiple"], 1):
        ws.write_formula(11, ci, '=COUNTIF({},"{}") '.format(col_a, label), f_t1_val)

    ws.set_row(12, 24)
    ws.write(12, 0, "% share against Total  URLs Crawled", f_t1_lbl)
    total_ref = total_indexable if total_indexable > 0 else 1
    for ci, col_letter in enumerate(["B", "C", "D", "E"], 1):
        ws.write_formula(12, ci, "={}12/{}".format(col_letter, total_ref), f_t1_pct)

    # Table 2
    ws.write(14, 0, "Table 2")
    ws.set_row(15, 16)
    ws.merge_range(15, 0, 15, 6, "Page Theme Wise H1 Analysis ", f_t2_title)

    t2_headers = ["Page Theme 1", "Priority Basis Page Theme 1", "Total Pages",
                  "Missing", "Duplicate", "Over 70 Characters", "Multiple"]
    for ci, h in enumerate(t2_headers):
        ws.write(16, ci, h, f_t2_hdr)

    for ri, (theme, counts) in enumerate(t2_rows_sorted):
        rr = R_T2_DATA_START + ri
        pri_basis = counts["priority_basis"]
        total_pages = theme_totals.get(theme, 0)
        t3_a = "A{}:A1048576".format(t3_data_excel_row)
        t3_c = "C{}:C1048576".format(t3_data_excel_row)
        theme_cell = '"{}"'.format(theme)
        ws.write(rr, 0, theme, f_t2_cell)
        ws.write(rr, 1, pri_basis, f_t2_cell)
        ws.write(rr, 2, total_pages, f_t2_num)
        for ci, label in enumerate(["Missing", "Duplicate", "Over 70 Characters", "Multiple"], 3):
            ws.write_formula(rr, ci,
                             '=COUNTIFS({},"{}", {},{})'.format(t3_a, label, t3_c, theme_cell),
                             f_t2_num)

    # Table 3
    ws.write(R_TABLE3_LABEL, 0, "Table 3")

    t3_headers = [
        "Error Type", "Address", "Page Theme 1", "Page Theme 2",
        "Content Type", "Status Code", "Indexability", "Indexability Status",
        "Occurance", "H1-1", "H1-1 Length", "H1-2", "H1-2 Length",
        "Impressions", "Clicks", "Sessions",
    ]
    for ci, h in enumerate(t3_headers):
        ws.write(R_T3_HDR, ci, h, f_t3_hdr)

    for ri, row in enumerate(t3_rows):
        rr = R_T3_DATA_START + ri
        ws.write(rr, 0,  row["error_type"],  f_t3_cell)
        ws.write(rr, 1,  row["address"],     f_t3_cell)
        ws.write(rr, 2,  row["theme1"],      f_t3_cell)
        ws.write(rr, 3,  row["theme2"],      f_t3_cell)
        ws.write(rr, 4,  row["content_type"],f_t3_cell)
        ws.write(rr, 5,  row["status_code"], f_t3_num)
        ws.write(rr, 6,  row["indexability"],f_t3_cell)
        ws.write(rr, 7,  row["idx_status"],  f_t3_cell)
        ws.write(rr, 8,  row["occurrence"],  f_t3_num)
        ws.write(rr, 9,  row["h1_1"],        f_t3_cell)
        ws.write(rr, 10, row["h1_1_len"],    f_t3_num)
        ws.write(rr, 11, row["h1_2"],        f_t3_cell)
        ws.write(rr, 12, row["h1_2_len"],    f_t3_num)
        ws.write(rr, 13, row["impressions"], f_t3_num)
        ws.write(rr, 14, row["clicks"],      f_t3_num)
        ws.write(rr, 15, row["sessions"],    f_t3_num)

    wb.close()
    buf.seek(0)
    return buf.read()
