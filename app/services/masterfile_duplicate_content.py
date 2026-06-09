import io
import math
import os
import pandas as pd
import xlsxwriter
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

DUPLICATE_ISSUES = [
    ("Exact Duplicates", "content_exact_duplicates.csv", "High"),
    ("Near Duplicates", "content_near_duplicates.csv", "Medium"),
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
    ws = wb.add_worksheet("Duplicate Content")

    RED = "#FF0000"
    GREY_HDR = "#F2F2F2"
    WHITE = "#FFFFFF"

    def fmt(**kw):
        return wb.add_format(kw)

    f_issue_summary = fmt(bold=True, font_size=10, valign="top", text_wrap=True, border=1)
    f_section_label = fmt(bold=True, font_size=10)
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

    ws.set_column(0, 0, 20)
    ws.set_column(1, 1, 45)
    ws.set_column(2, 2, 10)
    ws.set_column(3, 3, 18)
    ws.set_column(4, 4, 18)
    ws.set_column(5, 5, 12)
    ws.set_column(6, 6, 14)
    ws.set_column(7, 7, 35)
    ws.set_column(8, 8, 45)
    ws.set_column(9, 9, 12)
    ws.set_column(10, 10, 18)
    ws.set_column(11, 11, 45)
    ws.set_column(12, 12, 13)
    ws.set_column(13, 13, 10)
    ws.set_column(14, 14, 15)

    internal_df = _load_csv(report_dir, "internal_all.csv")
    gsc_df = _load_csv(report_dir, "search_console_all.csv")
    ga4_df = _load_csv(report_dir, "analytics_all.csv")
    content_all_df = _load_csv(report_dir, "content_all.csv")
    sitemaps_all_df = _load_csv(report_dir, "sitemaps_all.csv")

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

    # content_all lookup for Canonical Link Element 1 and Closest Similarity Match
    content_lookup = {}
    if not content_all_df.empty and "Address" in content_all_df.columns:
        for _, row in content_all_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if addr:
                content_lookup[addr] = {
                    "canonical": _clean(row.get("Canonical Link Element 1", "")),
                    "closest_match": _clean(row.get("Closest Similarity Match", "")),
                }

    # internal_all lookup for Status Code
    internal_lookup = {}
    if not internal_df.empty and "Address" in internal_df.columns:
        for _, row in internal_df.iterrows():
            addr = str(row.get("Address", "")).strip()
            if addr:
                internal_lookup[addr] = _clean(row.get("Status Code", ""))

    # sitemaps_all set for "Available in Sitemap?" check
    sitemap_urls = set()
    if not sitemaps_all_df.empty and "Address" in sitemaps_all_df.columns:
        sitemap_urls = set(sitemaps_all_df["Address"].astype(str).str.strip().tolist())

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
    for issue_label, csv_name, _ in DUPLICATE_ISSUES:
        df = _load_csv(report_dir, csv_name)
        df = _filter_indexable_200_html(df)
        if df.empty:
            continue
        for _, row in df.iterrows():
            addr = str(row.get("Address", "")).strip()
            t1, t2, lang, pri = classify_url(addr, rulebook)
            cl = content_lookup.get(addr, {})
            status_code = internal_lookup.get(addr, "-")
            indexability = _clean(row.get("Indexability", ""))
            hash_val = _clean(row.get("Hash", "")) if "Hash" in df.columns else "-"
            near_dup_addr = _clean(row.get("Closest Similarity Match", "")) if "Closest Similarity Match" in df.columns else cl.get("closest_match", "-")
            similarity = 100 if issue_label == "Exact Duplicates" else "-"
            available_in_sitemap = "Yes" if addr in sitemap_urls else "No"
            canonical = _clean(row.get("Canonical Link Element 1", "")) if "Canonical Link Element 1" in df.columns else cl.get("canonical", "-")
            gsc = gsc_map.get(addr, {})
            sessions = ga4_map.get(addr)
            t3_rows.append({
                "error_type": issue_label,
                "address": addr,
                "page_type": "HTML",
                "theme1": t1 if t1 else "-",
                "theme2": t2 if t2 else "-",
                "status_code": status_code,
                "indexability": indexability,
                "hash": hash_val,
                "near_dup_addr": near_dup_addr,
                "similarity": similarity,
                "available_in_sitemap": available_in_sitemap,
                "canonical": canonical,
                "impressions": gsc.get("Impressions"),
                "clicks": gsc.get("Clicks"),
                "sessions": sessions,
            })

    t3_rows.sort(key=lambda r: (r["impressions"] is None, -(r["impressions"] or 0)))

    # Build T2 theme data
    R_T2_DATA_START = 16
    theme_data = {}
    for r in t3_rows:
        th = r["theme1"]
        if th not in theme_data:
            theme_data[th] = {"Near Duplicates": 0, "Exact Duplicates": 0, "priority_basis": "N/A"}
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
        "1. Near Duplicates - URLs with highly similar content and only minor variations, potentially diluting "
        "ranking signals and reducing overall content uniqueness.\n"
        "2. Exact Duplicates - URLs containing identical or substantially identical page content, which can "
        "create duplication issues and make it difficult for search engines to determine the preferred "
        "version to index."
    )
    ws.merge_range(1, 0, 3, 14, summary_text, f_issue_summary)

    # Table 1
    ws.write(7, 0, "Table 1")
    t1_headers = ["Content Issue Types", "Exact Duplicates", "Near Duplicates"]
    for ci, h in enumerate(t1_headers):
        ws.write(8, ci, h, f_t1_hdr)

    t1_priorities = ["Issue Priority", "High", "Medium"]
    for ci, v in enumerate(t1_priorities):
        ws.write(9, ci, v, f_t1_hdr)

    ws.write(10, 0, "#Affected URLs", f_t1_lbl)
    col_a = "A{}:A1048576".format(t3_data_excel_row)
    ws.write_formula(10, 1, '=COUNTIF({},"Exact Duplicates")'.format(col_a), f_t1_val)
    ws.write_formula(10, 2, '=COUNTIF({},"Near Duplicates")'.format(col_a), f_t1_val)

    ws.set_row(11, 24)
    ws.write(11, 0, "% Share against Total  URLs Crawled", f_t1_lbl)
    total_ref = total_indexable if total_indexable > 0 else 1
    ws.write_formula(11, 1, "=B11/{}".format(total_ref), f_t1_pct)
    ws.write_formula(11, 2, "=C11/{}".format(total_ref), f_t1_pct)

    # Table 2
    ws.write(13, 0, "Table 2")
    ws.set_row(14, 16)
    ws.merge_range(14, 0, 14, 3, "Page Theme Wise URL Analysis ", f_t2_title)

    t2_headers = ["Page Theme 1", "Priority Basis Page Theme 1", "Near Duplicates", "Exact Duplicates"]
    for ci, h in enumerate(t2_headers):
        ws.write(15, ci, h, f_t2_hdr)

    for ri, (theme, counts) in enumerate(t2_rows_sorted):
        rr = R_T2_DATA_START + ri
        pri_basis = counts["priority_basis"]
        t3_a = "A{}:A1048576".format(t3_data_excel_row)
        t3_d = "D{}:D1048576".format(t3_data_excel_row)
        theme_cell = '"{}"'.format(theme)
        ws.write(rr, 0, theme, f_t2_cell)
        ws.write(rr, 1, pri_basis, f_t2_cell)
        ws.write_formula(rr, 2,
                         '=COUNTIFS({},"Near Duplicates", {},{})'.format(t3_a, t3_d, theme_cell),
                         f_t2_num)
        ws.write_formula(rr, 3,
                         '=COUNTIFS({},"Exact Duplicates", {},{})'.format(t3_a, t3_d, theme_cell),
                         f_t2_num)

    # Table 3
    ws.write(R_TABLE3_LABEL, 0, "Table 3")
    t3_headers = [
        "Error Type", "Address", "Page Type", "Page Theme 1", "Page Theme 2",
        "Status Code", "Indexability", "Hash", "Near Duplicate Address",
        "Similarity", "Available in Sitemap?", "Canonical URL",
        "Impressions", "Clicks", "Organic Sessions",
    ]
    for ci, h in enumerate(t3_headers):
        ws.write(R_T3_HDR, ci, h, f_t3_hdr)

    for ri, row in enumerate(t3_rows):
        rr = R_T3_DATA_START + ri
        ws.write(rr, 0,  row["error_type"],          f_t3_cell)
        ws.write(rr, 1,  row["address"],             f_t3_cell)
        ws.write(rr, 2,  row["page_type"],           f_t3_cell)
        ws.write(rr, 3,  row["theme1"],              f_t3_cell)
        ws.write(rr, 4,  row["theme2"],              f_t3_cell)
        ws.write(rr, 5,  row["status_code"],         f_t3_num)
        ws.write(rr, 6,  row["indexability"],        f_t3_cell)
        ws.write(rr, 7,  row["hash"],                f_t3_cell)
        ws.write(rr, 8,  row["near_dup_addr"],       f_t3_cell)
        ws.write(rr, 9,  row["similarity"],          f_t3_num)
        ws.write(rr, 10, row["available_in_sitemap"],f_t3_cell)
        ws.write(rr, 11, row["canonical"],           f_t3_cell)
        ws.write(rr, 12, row["impressions"],         f_t3_num)
        ws.write(rr, 13, row["clicks"],              f_t3_num)
        ws.write(rr, 14, row["sessions"],            f_t3_num)

    wb.close()
    buf.seek(0)
    return buf.read()
