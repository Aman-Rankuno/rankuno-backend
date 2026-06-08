import io
import math
import os
import pandas as pd
import xlsxwriter
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

DIRECTIVES_ISSUES = [
    ("Noindex",  "directives_noindex.csv",  "High"),
    ("Nofollow", "directives_nofollow.csv", "Medium"),
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
    ws = wb.add_worksheet("Directives")

    RED = "#FF0000"
    GREY_HDR = "#F2F2F2"
    WHITE = "#FFFFFF"

    def fmt(**kw):
        return wb.add_format(kw)

    f_issue_summary = fmt(bold=True, font_size=10, valign="top", text_wrap=True, border=1)
    f_section_label = fmt(bold=True, font_size=10)
    # T1 header row (red bg)
    f_t1_hdr = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter", text_wrap=True)
    # T1 data rows (white bg, black text)
    f_t1_lbl = fmt(bold=True, border=1, valign="vcenter")
    f_t1_val = fmt(bold=True, border=1, align="center", valign="vcenter")
    f_t1_pct = fmt(bold=True, border=1, align="center", valign="vcenter", num_format="0.00%")
    f_t2_title = fmt(bold=True, font_color=WHITE, bg_color=RED, border=1, valign="vcenter")
    f_t2_hdr = fmt(bold=True, bg_color=GREY_HDR, border=1, valign="vcenter", text_wrap=True)
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
    ws.set_column(6, 6, 15)
    ws.set_column(7, 7, 13)
    ws.set_column(8, 8, 10)
    ws.set_column(9, 9, 15)

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
                }

    try:
        rulebook = load_rulebook(domain)
    except Exception:
        rulebook = []

    internal_filtered = _filter_indexable_200_html(internal_df)
    total_indexable = len(internal_filtered) if not internal_filtered.empty else 0

    # Build T3 rows -- filter: self-canonical OR canonical is blank
    t3_rows = []
    for issue_label, csv_name, _ in DIRECTIVES_ISSUES:
        df = _load_csv(report_dir, csv_name)
        if df.empty:
            continue
        if "Canonical Link Element 1" in df.columns:
            addr_col = df["Address"].astype(str).str.strip()
            canon_col = df["Canonical Link Element 1"].astype(str).str.strip()
            mask = (canon_col == addr_col) | (canon_col == "") | (canon_col == "nan") | df["Canonical Link Element 1"].isnull()
            df = df[mask]
        if df.empty:
            continue
        for _, row in df.iterrows():
            addr = str(row.get("Address", "")).strip()
            t1, t2, lang, pri = classify_url(addr, rulebook)
            _il = internal_lookup.get(addr, {})
            content_type = _il.get("content_type") or "-"
            status_code = _il.get("status_code") or "-"
            indexability = _il.get("indexability") or "-"
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
                "impressions": gsc.get("Impressions"),
                "clicks": gsc.get("Clicks"),
                "sessions": sessions,
            })

    t3_rows.sort(key=lambda r: (r["impressions"] is None, -(r["impressions"] or 0)))

    # Build T2 theme data
    R_T2_DATA_START = 15
    theme_data = {}
    for r in t3_rows:
        th = r["theme1"]
        if th not in theme_data:
            theme_data[th] = {"Noindex": 0, "Nofollow": 0, "priority_basis": "N/A"}
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
        "1. Noindex - URLs containing a noindex directive, instructing search engines not to include "
        "the page in search engine results pages (SERPs).\n"
        "2. Nofollow - URLs containing a nofollow directive, instructing search engines not to follow "
        "links present on the page or pass link equity through them."
    )
    ws.merge_range(1, 0, 3, 9, summary_text, f_issue_summary)

    ws.write(5, 0, "Summary Table ", f_section_label)
    ws.write(6, 0, "Table 1")

    # Table 1 -- only header row is red, data rows are white
    t1_headers = ["Meta Robot Tags", "No Index", "No Follow"]
    for ci, h in enumerate(t1_headers):
        ws.write(7, ci, h, f_t1_hdr)

    t1_priorities = ["Issue Priority", "High", "Medium"]
    for ci, v in enumerate(t1_priorities):
        ws.write(8, ci, v, f_t1_lbl)

    ws.write(9, 0, "#Affected URLs", f_t1_lbl)
    col_a = "A{}:A1048576".format(t3_data_excel_row)
    ws.write_formula(9, 1, '=COUNTIF({},"Noindex")'.format(col_a), f_t1_val)
    ws.write_formula(9, 2, '=COUNTIF({},"Nofollow")'.format(col_a), f_t1_val)

    ws.set_row(10, 24)
    ws.write(10, 0, "% share against Total  URLs Crawled", f_t1_lbl)
    total_ref = total_indexable if total_indexable > 0 else 1
    ws.write_formula(10, 1, "=B10/{}".format(total_ref), f_t1_pct)
    ws.write_formula(10, 2, "=C10/{}".format(total_ref), f_t1_pct)

    # Table 2 -- 5 cols: Page Theme 1 | Priority Basis | Priority | No Index | No Follow
    ws.write(12, 0, "Table 2")
    ws.set_row(13, 16)
    ws.merge_range(13, 0, 13, 4, "Page Theme Wise Meta Robot Analysis ", f_t2_title)

    t2_headers = ["Page Theme 1", "Priority Basis Page Theme 1", "Priority", "No Index", "No Follow"]
    for ci, h in enumerate(t2_headers):
        ws.write(14, ci, h, f_t2_hdr)

    for ri, (theme, counts) in enumerate(t2_rows_sorted):
        rr = R_T2_DATA_START + ri
        pri_basis = counts["priority_basis"]
        t3_a = "A{}:A1048576".format(t3_data_excel_row)
        t3_c = "C{}:C1048576".format(t3_data_excel_row)
        theme_cell = '"{}"'.format(theme)
        ws.write(rr, 0, theme, f_t2_cell)
        ws.write(rr, 1, pri_basis, f_t2_cell)
        ws.write(rr, 2, pri_basis, f_t2_cell)
        ws.write_formula(rr, 3, '=COUNTIFS({},"{}", {},{})'.format(t3_a, "Noindex", t3_c, theme_cell), f_t2_num)
        ws.write_formula(rr, 4, '=COUNTIFS({},"{}", {},{})'.format(t3_a, "Nofollow", t3_c, theme_cell), f_t2_num)

    # Table 3
    ws.write(R_TABLE3_LABEL, 0, "Table 3")

    t3_headers = [
        "Error Type", "Address", "Page Theme 1", "Page Theme 2",
        "Content Type", "Status Code", "Indexability",
        "Impressions", "Clicks", "Organic Sessions",
    ]
    for ci, h in enumerate(t3_headers):
        ws.write(R_T3_HDR, ci, h, f_t3_hdr)

    for ri, row in enumerate(t3_rows):
        rr = R_T3_DATA_START + ri
        ws.write(rr, 0, row["error_type"],   f_t3_cell)
        ws.write(rr, 1, row["address"],      f_t3_cell)
        ws.write(rr, 2, row["theme1"],       f_t3_cell)
        ws.write(rr, 3, row["theme2"],       f_t3_cell)
        ws.write(rr, 4, row["content_type"], f_t3_cell)
        ws.write(rr, 5, row["status_code"],  f_t3_num)
        ws.write(rr, 6, row["indexability"], f_t3_cell)
        ws.write(rr, 7, row["impressions"],  f_t3_num)
        ws.write(rr, 8, row["clicks"],       f_t3_num)
        ws.write(rr, 9, row["sessions"],     f_t3_num)

    wb.close()
    buf.seek(0)
    return buf.read()
