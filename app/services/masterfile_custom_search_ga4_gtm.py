import io
import math
import os
import pandas as pd
import xlsxwriter
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

# Custom extractor column name -> friendly label
EXTRACTOR_COLS = [
    ("GA4", "GA4 Available"),
    ("GTM Head", "GTM Available in Head?"),
    ("GTM Body", "GTM Available in Body?"),
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


def _yes_no(v):
    """Yes if extractor value is non-empty, No otherwise."""
    if v is None:
        return "No"
    s = str(v).strip()
    if not s or s.lower() == "nan" or s == "-":
        return "No"
    return "Yes"


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
    ws = wb.add_worksheet("Custom Search GA4 GTM")

    RED = "#FF0000"
    GREY_HDR = "#F2F2F2"
    WHITE = "#FFFFFF"

    def fmt(**kw):
        return wb.add_format(kw)

    f_issue_summary = fmt(bold=True, font_size=10, valign="top", text_wrap=True, border=1)
    f_t1_hdr = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter", text_wrap=True)
    f_t1_lbl = fmt(bold=True, font_color=WHITE, bg_color=RED, border=1, valign="vcenter")
    f_t1_val = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter")
    f_t1_pct = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, align="center", valign="vcenter", num_format="0.00%")
    f_t2_title = fmt(bold=True, font_color=WHITE, bg_color=RED, border=1, valign="vcenter")
    f_t2_hdr = fmt(bold=True, bg_color=GREY_HDR, border=1, valign="vcenter", text_wrap=True)
    f_t2_cell = fmt(border=1, valign="vcenter")
    f_t2_num = fmt(border=1, valign="vcenter", align="center")
    f_t3_hdr = fmt(bold=True, font_color=WHITE, bg_color=RED,
                   border=1, valign="vcenter", text_wrap=True)
    f_t3_cell = fmt(border=1, valign="vcenter", text_wrap=True)
    f_t3_num = fmt(border=1, valign="vcenter", align="center")

    ws.set_column(0, 0, 45)
    ws.set_column(1, 1, 18)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 25)
    ws.set_column(4, 4, 12)
    ws.set_column(5, 5, 14)
    ws.set_column(6, 6, 16)
    ws.set_column(7, 7, 22)
    ws.set_column(8, 8, 22)
    ws.set_column(9, 9, 13)
    ws.set_column(10, 10, 10)
    ws.set_column(11, 11, 16)

    internal_df = _load_csv(report_dir, "internal_all.csv")
    gsc_df = _load_csv(report_dir, "search_console_all.csv")
    ga4_df = _load_csv(report_dir, "analytics_all.csv")
    extraction_df = _load_csv(report_dir, "custom_extraction_all.csv")

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

    # internal_all lookup for Content Type, Status Code, Indexability
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

    theme_totals = {}
    if not internal_filtered.empty and "Address" in internal_filtered.columns:
        for _, row in internal_filtered.iterrows():
            addr = str(row.get("Address", "")).strip()
            t1, t2, lang, pri = classify_url(addr, rulebook)
            theme = t1 if t1 else "-"
            theme_totals[theme] = theme_totals.get(theme, 0) + 1

    # Filter custom_extraction_all to indexable 200 HTML only
    # The CSV may not have Indexability/Status/Content Type cols, so we filter via internal_lookup
    valid_addrs = set(internal_filtered["Address"].astype(str).str.strip().tolist()) if not internal_filtered.empty else set()

    # Build T3 rows from custom_extraction_all
    # Only build if at least one configured extractor column actually exists in the CSV
    any_extractor_present = any(col in extraction_df.columns for col in ("GA4", "GTM Head", "GTM Body"))
    t3_rows = []
    if any_extractor_present and not extraction_df.empty and "Address" in extraction_df.columns:
        for _, row in extraction_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if not addr or addr not in valid_addrs:
                continue
            t1, t2, lang, pri = classify_url(addr, rulebook)
            il = internal_lookup.get(addr, {})
            ga4_avail = _yes_no(row.get("GA4")) if "GA4" in extraction_df.columns else "-"
            gtm_head = _yes_no(row.get("GTM Head")) if "GTM Head" in extraction_df.columns else "-"
            gtm_body = _yes_no(row.get("GTM Body")) if "GTM Body" in extraction_df.columns else "-"
            gsc = gsc_map.get(addr, {})
            sessions = ga4_map.get(addr)
            t3_rows.append({
                "address": addr,
                "theme1": t1 if t1 else "-",
                "theme2": t2 if t2 else "-",
                "content_type": il.get("content_type", "-"),
                "status_code": il.get("status_code", "-"),
                "indexability": il.get("indexability", "-"),
                "ga4_avail": ga4_avail,
                "gtm_head": gtm_head,
                "gtm_body": gtm_body,
                "impressions": gsc.get("Impressions"),
                "clicks": gsc.get("Clicks"),
                "sessions": sessions,
            })

    t3_rows.sort(key=lambda r: (r["impressions"] is None, -(r["impressions"] or 0)))

    # Build T2 theme data — affected = "No" or "-" (missing tag)
    R_T2_DATA_START = 16
    theme_data = {}
    for r in t3_rows:
        th = r["theme1"]
        if th not in theme_data:
            theme_data[th] = {"GA4 Available": 0, "GTM Head": 0, "GTM Body": 0, "priority_basis": "N/A"}
        if r["ga4_avail"] in ("No", "-"):
            theme_data[th]["GA4 Available"] += 1
        if r["gtm_head"] in ("No", "-"):
            theme_data[th]["GTM Head"] += 1
        if r["gtm_body"] in ("No", "-"):
            theme_data[th]["GTM Body"] += 1

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
        "Issue Summary: GA4 and GTM code identification checks help verify whether Google Analytics 4 (GA4) and "
        "Google Tag Manager (GTM) tracking codes are properly implemented across website pages. Issues may include "
        "missing, duplicate, incorrectly placed, inconsistent, or non-functional tracking codes that can impact "
        "data collection accuracy, analytics reporting, and marketing performance measurement."
    )
    ws.merge_range(1, 0, 3, 11, summary_text, f_issue_summary)

    # Table 1
    ws.write(7, 0, "Table 1")
    t1_headers = ["Custom Extraction Type", "GA4 Available", "GTM Available - Head", "GTM Available - Body"]
    for ci, h in enumerate(t1_headers):
        ws.write(8, ci, h, f_t1_hdr)

    t1_priorities = ["Issue Priority", "High", "High", "High"]
    for ci, v in enumerate(t1_priorities):
        ws.write(9, ci, v, f_t1_hdr)

    ws.write(10, 0, "#Affected URLs", f_t1_lbl)
    # Affected = "No" or "-" in cols G/H/I of T3
    col_g = "G{}:G1048576".format(t3_data_excel_row)
    col_h = "H{}:H1048576".format(t3_data_excel_row)
    col_i = "I{}:I1048576".format(t3_data_excel_row)
    ws.write_formula(10, 1, '=COUNTIF({},"No")+COUNTIF({},"-")'.format(col_g, col_g), f_t1_val)
    ws.write_formula(10, 2, '=COUNTIF({},"No")+COUNTIF({},"-")'.format(col_h, col_h), f_t1_val)
    ws.write_formula(10, 3, '=COUNTIF({},"No")+COUNTIF({},"-")'.format(col_i, col_i), f_t1_val)

    ws.set_row(11, 24)
    ws.write(11, 0, "% share against Total  URLs Crawled", f_t1_lbl)
    total_ref = total_indexable if total_indexable > 0 else 1
    ws.write_formula(11, 1, "=B11/{}".format(total_ref), f_t1_pct)
    ws.write_formula(11, 2, "=C11/{}".format(total_ref), f_t1_pct)
    ws.write_formula(11, 3, "=D11/{}".format(total_ref), f_t1_pct)

    # Table 2
    ws.write(13, 0, "Table 2")
    ws.set_row(14, 16)
    ws.merge_range(14, 0, 14, 4, "Page Theme Wise URL Analysis ", f_t2_title)

    t2_headers = ["Page Theme 1", "Priority Basis Page Theme 1",
                  "GA4 Available", "GTM Available - Head", "GTM Available - Body"]
    for ci, h in enumerate(t2_headers):
        ws.write(15, ci, h, f_t2_hdr)

    for ri, (theme, counts) in enumerate(t2_rows_sorted):
        rr = R_T2_DATA_START + ri
        pri_basis = counts["priority_basis"]
        t3_b = "B{}:B1048576".format(t3_data_excel_row)
        t3_g = "G{}:G1048576".format(t3_data_excel_row)
        t3_h = "H{}:H1048576".format(t3_data_excel_row)
        t3_i = "I{}:I1048576".format(t3_data_excel_row)
        theme_cell = '"{}"'.format(theme)
        ws.write(rr, 0, theme, f_t2_cell)
        ws.write(rr, 1, pri_basis, f_t2_cell)
        # COUNTIFS for GA4 affected (No or -) per theme
        ws.write_formula(rr, 2,
                         '=COUNTIFS({},"No",{},{})+COUNTIFS({},"-",{},{})'.format(t3_g, t3_b, theme_cell, t3_g, t3_b, theme_cell),
                         f_t2_num)
        ws.write_formula(rr, 3,
                         '=COUNTIFS({},"No",{},{})+COUNTIFS({},"-",{},{})'.format(t3_h, t3_b, theme_cell, t3_h, t3_b, theme_cell),
                         f_t2_num)
        ws.write_formula(rr, 4,
                         '=COUNTIFS({},"No",{},{})+COUNTIFS({},"-",{},{})'.format(t3_i, t3_b, theme_cell, t3_i, t3_b, theme_cell),
                         f_t2_num)

    # Table 3
    ws.write(R_TABLE3_LABEL, 0, "Table 3")
    t3_headers = [
        "Address", "Page Theme 1", "Page Theme 2", "Content Type",
        "Status Code", "Indexability", "GA4 Available",
        "GTM Available in Head?", "GTM Available in Body?",
        "Impressions", "Clicks", "Organic Sessions",
    ]
    for ci, h in enumerate(t3_headers):
        ws.write(R_T3_HDR, ci, h, f_t3_hdr)

    for ri, row in enumerate(t3_rows):
        rr = R_T3_DATA_START + ri
        ws.write(rr, 0,  row["address"],      f_t3_cell)
        ws.write(rr, 1,  row["theme1"],       f_t3_cell)
        ws.write(rr, 2,  row["theme2"],       f_t3_cell)
        ws.write(rr, 3,  row["content_type"], f_t3_cell)
        ws.write(rr, 4,  row["status_code"],  f_t3_num)
        ws.write(rr, 5,  row["indexability"], f_t3_cell)
        ws.write(rr, 6,  row["ga4_avail"],    f_t3_cell)
        ws.write(rr, 7,  row["gtm_head"],     f_t3_cell)
        ws.write(rr, 8,  row["gtm_body"],     f_t3_cell)
        ws.write(rr, 9,  row["impressions"],  f_t3_num)
        ws.write(rr, 10, row["clicks"],       f_t3_num)
        ws.write(rr, 11, row["sessions"],     f_t3_num)

    wb.close()
    buf.seek(0)
    return buf.read()
