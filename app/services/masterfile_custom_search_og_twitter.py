import io
import math
import os
import pandas as pd
import xlsxwriter
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

# 12 extractor columns. Tuple = (SF custom_extraction_all CSV column name, T1/T2 display label, T3 display label)
EXTRACTORS = [
    ("OG Type",             "OG Type",             "OG Type"),
    ("OG Title",            "OG Title",            "OG Title"),
    ("OG Image",            "OG Image",            "OG Image"),
    ("OG Image Width",      "OG Image Width",      "OG Image Width"),
    ("OG Image Height",     "OG Image Height",     "OG Image Height"),
    ("OG Description",      "OG Description",      "OG Description"),
    ("OG Sitename",         "OG Sitename",         "OG Sitename"),
    ("Twitter Card",        "Twitter Card",        "Twitter Card"),
    ("Twitter Title",       "Twitter Title",       "Twitter Title"),
    ("Twitter Site",        "Twitter Site",        "Twitter Site"),
    ("Twitter Description", "Twitter Description", "Twitter Description"),
    ("Twitter Image",       "Twitter Image",       "Twitter Image"),
]
NUM_EXTRACTORS = len(EXTRACTORS)  # 12


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


def _col_letter(idx_zero):
    # 0 -> A, 25 -> Z, 26 -> AA, etc.
    n = idx_zero
    s = ""
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def generate(report_dir: str, domain: str) -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    ws = wb.add_worksheet("Custom Search OG Twitter")

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

    # Column widths
    ws.set_column(0, 0, 45)   # Address
    ws.set_column(1, 1, 18)   # Page Theme 1
    ws.set_column(2, 2, 18)   # Page Theme 2
    ws.set_column(3, 3, 25)   # Content Type
    ws.set_column(4, 4, 12)   # Status Code
    ws.set_column(5, 5, 14)   # Indexability
    ws.set_column(6, 17, 16)  # 12 extractor cols
    ws.set_column(18, 18, 13) # Impressions
    ws.set_column(19, 19, 10) # Clicks
    ws.set_column(20, 20, 16) # Organic Sessions

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

    valid_addrs = set(internal_filtered["Address"].astype(str).str.strip().tolist()) if not internal_filtered.empty else set()

    # Build T3 rows
    # Only build if at least one configured extractor column actually exists in the CSV
    any_extractor_present = any(csv_col in extraction_df.columns for csv_col, _, _ in EXTRACTORS)
    t3_rows = []
    if any_extractor_present and not extraction_df.empty and "Address" in extraction_df.columns:
        for _, row in extraction_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if not addr or addr not in valid_addrs:
                continue
            t1, t2, lang, pri = classify_url(addr, rulebook)
            il = internal_lookup.get(addr, {})
            extractor_vals = []
            for csv_col, _, _ in EXTRACTORS:
                if csv_col in extraction_df.columns:
                    extractor_vals.append(_yes_no(row.get(csv_col)))
                else:
                    extractor_vals.append("-")
            gsc = gsc_map.get(addr, {})
            sessions = ga4_map.get(addr)
            t3_rows.append({
                "address": addr,
                "theme1": t1 if t1 else "-",
                "theme2": t2 if t2 else "-",
                "content_type": il.get("content_type", "-"),
                "status_code": il.get("status_code", "-"),
                "indexability": il.get("indexability", "-"),
                "extractors": extractor_vals,
                "impressions": gsc.get("Impressions"),
                "clicks": gsc.get("Clicks"),
                "sessions": sessions,
            })

    t3_rows.sort(key=lambda r: (r["impressions"] is None, -(r["impressions"] or 0)))

    # T2 theme aggregation
    R_T2_DATA_START = 15  # Excel row 16
    theme_data = {}
    for r in t3_rows:
        th = r["theme1"]
        if th not in theme_data:
            theme_data[th] = {
                "counts": [0] * NUM_EXTRACTORS,
                "priority_basis": "N/A",
            }
        for i, ev in enumerate(r["extractors"]):
            if ev in ("No", "-"):
                theme_data[th]["counts"][i] += 1

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
    R_TABLE3_LABEL = R_T2_DATA_START + t2_count + 1  # blank row + label
    R_T3_HDR = R_TABLE3_LABEL + 1
    R_T3_DATA_START = R_T3_HDR + 1
    t3_data_excel_row = R_T3_DATA_START + 1  # 1-indexed Excel row

    # Issue Summary rows 2-4 merged A:U
    ws.set_row(0, 15)
    ws.set_row(1, 55)
    ws.set_row(2, 40)
    ws.set_row(3, 40)
    summary_text = (
        "Issue Summary: Open Graph (OG) and Twitter tags are social metadata elements that control how webpages appear "
        "when shared on social media platforms. Issues include missing, duplicate, incomplete, incorrect, or improperly "
        "configured OG and Twitter tags that can affect content previews, social sharing appearance, "
        "engagement, and click-through rates."
    )
    ws.merge_range(1, 0, 3, 20, summary_text, f_issue_summary)

    # Table 1 (rows 7-11, Excel rows 8-12)
    ws.write(6, 0, "Table 1")
    # T1 col A label + 12 extractor headers in cols B..M (idx 1..12)
    ws.write(7, 0, "Custom Extraction Type", f_t1_hdr)
    for i, (_, t1_label, _) in enumerate(EXTRACTORS):
        ws.write(7, i + 1, t1_label, f_t1_hdr)

    # Row 9 (idx 8): Issue Priority
    ws.write(8, 0, "Issue Priority", f_t1_hdr)
    for i in range(NUM_EXTRACTORS):
        ws.write(8, i + 1, "Low", f_t1_hdr)

    # Row 10 (idx 9): #Affected URLs formulas
    ws.write(9, 0, "#Affected URLs", f_t1_lbl)
    for i in range(NUM_EXTRACTORS):
        # T3 col index for this extractor is 6 + i (G + i)
        t3_col_letter = _col_letter(6 + i)
        rng = "{0}{1}:{0}1048576".format(t3_col_letter, t3_data_excel_row)
        ws.write_formula(9, i + 1,
                         '=COUNTIF({0},"No")+COUNTIF({0},"-")'.format(rng),
                         f_t1_val)

    # Row 11 (idx 10): % share
    ws.set_row(10, 24)
    ws.write(10, 0, "% share against Total  URLs Crawled", f_t1_lbl)
    total_ref = total_indexable if total_indexable > 0 else 1
    for i in range(NUM_EXTRACTORS):
        t1_col_letter = _col_letter(i + 1)  # B, C, D, ... M
        ws.write_formula(10, i + 1,
                         "={0}10/{1}".format(t1_col_letter, total_ref),
                         f_t1_pct)

    # Table 2 (row 13, Excel row 14)
    ws.write(12, 0, "Table 2")
    ws.set_row(13, 16)
    # T2 spans cols A..N (idx 0..13) = theme + priority + 12 extractors
    ws.merge_range(13, 0, 13, 13, "Page Theme Wise URL Analysis ", f_t2_title)

    # T2 headers (row 14, idx 14)
    ws.write(14, 0, "Page Theme 1", f_t2_hdr)
    ws.write(14, 1, "Priority Basis Page Theme 1", f_t2_hdr)
    for i, (_, t2_label, _) in enumerate(EXTRACTORS):
        ws.write(14, i + 2, t2_label, f_t2_hdr)

    # T2 data rows
    for ri, (theme, info) in enumerate(t2_rows_sorted):
        rr = R_T2_DATA_START + ri
        pri_basis = info["priority_basis"]
        ws.write(rr, 0, theme, f_t2_cell)
        ws.write(rr, 1, pri_basis, f_t2_cell)
        t3_b = "B{0}:B1048576".format(t3_data_excel_row)
        theme_cell = '"{0}"'.format(theme)
        for i in range(NUM_EXTRACTORS):
            t3_col_letter = _col_letter(6 + i)
            t3_rng = "{0}{1}:{0}1048576".format(t3_col_letter, t3_data_excel_row)
            ws.write_formula(
                rr, i + 2,
                '=COUNTIFS({0},"No",{1},{2})+COUNTIFS({0},"-",{1},{2})'.format(t3_rng, t3_b, theme_cell),
                f_t2_num,
            )

    # Table 3 label and headers
    ws.write(R_TABLE3_LABEL, 0, "Table 3")
    t3_headers = [
        "Address", "Page Theme 1", "Page Theme 2", "Content Type",
        "Status Code", "Indexability",
    ] + [t3_label for _, _, t3_label in EXTRACTORS] + [
        "Impressions", "Clicks", "Organic Sessions",
    ]
    for ci, h in enumerate(t3_headers):
        ws.write(R_T3_HDR, ci, h, f_t3_hdr)

    # T3 data
    for ri, row in enumerate(t3_rows):
        rr = R_T3_DATA_START + ri
        ws.write(rr, 0, row["address"],      f_t3_cell)
        ws.write(rr, 1, row["theme1"],       f_t3_cell)
        ws.write(rr, 2, row["theme2"],       f_t3_cell)
        ws.write(rr, 3, row["content_type"], f_t3_cell)
        ws.write(rr, 4, row["status_code"],  f_t3_num)
        ws.write(rr, 5, row["indexability"], f_t3_cell)
        for i, ev in enumerate(row["extractors"]):
            ws.write(rr, 6 + i, ev, f_t3_cell)
        ws.write(rr, 18, row["impressions"], f_t3_num)
        ws.write(rr, 19, row["clicks"],      f_t3_num)
        ws.write(rr, 20, row["sessions"],    f_t3_num)

    wb.close()
    buf.seek(0)
    return buf.read()
