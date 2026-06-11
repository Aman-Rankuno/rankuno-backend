import os
import io
import math
import pandas as pd
import xlsxwriter
from collections import defaultdict
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

INLINK_BUCKETS = [
    "No incoming internal link",
    "Only one incoming internal link",
    "Incoming internal links between 2-5",
    "Incoming internal links between 6-15",
    "Incoming internal links between 16-50",
    "Incoming internal links between 51-100",
    "Incoming internal links beyond 100",
]

CSV_NAMES = [
    "success_(2xx)_inlinks.csv",
    "internal_success_(2xx)_inlinks.csv",
]


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f) or f == 0) else f
    except Exception:
        return None


def _gc(df, *names):
    for n in names:
        m = next((c for c in df.columns if c.lower() == n.lower()), None)
        if m:
            return m
    return None


def _inlink_zone(count: int) -> str:
    if count == 0:
        return INLINK_BUCKETS[0]
    if count == 1:
        return INLINK_BUCKETS[1]
    if count <= 5:
        return INLINK_BUCKETS[2]
    if count <= 15:
        return INLINK_BUCKETS[3]
    if count <= 50:
        return INLINK_BUCKETS[4]
    if count <= 100:
        return INLINK_BUCKETS[5]
    return INLINK_BUCKETS[6]

def build_functional_internal_links_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "Functional Internal Links Analysis.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    rulebook = load_rulebook(domain)

    # Find the inlinks CSV
    csv_path = None
    for name in CSV_NAMES:
        p = os.path.join(report_path, name)
        if os.path.exists(p):
            csv_path = p
            break
    if not csv_path:
        raise FileNotFoundError("success_(2xx)_inlinks.csv not found in report folder")

    # Load internal_all for source indexability lookup
    internal_map = {}
    internal_all_path = os.path.join(report_path, "internal_all.csv")
    if os.path.exists(internal_all_path):
        df_int = pd.read_csv(internal_all_path, encoding="utf-8", low_memory=False,
                             usecols=lambda c: c.lower() in ("address", "indexability", "status code", "status"))
        a_col = _gc(df_int, "Address")
        idx_col = _gc(df_int, "Indexability")
        sc_col = _gc(df_int, "Status Code")
        st_col = _gc(df_int, "Status")
        for _, r in df_int.iterrows():
            url = str(r[a_col]) if a_col else ""
            internal_map[url] = {
                "indexability": str(r.get(idx_col, "")) if idx_col else "",
                "status_code": str(r.get(sc_col, "")) if sc_col else "",
                "status": str(r.get(st_col, "")) if st_col else "",
            }

    # Load GSC and GA
    def load_csv_safe(filename, usecols=None):
        path = os.path.join(report_path, filename)
        if not os.path.exists(path):
            return pd.DataFrame()
        try:
            if usecols:
                return pd.read_csv(path, encoding="utf-8", low_memory=False, usecols=usecols)
            return pd.read_csv(path, encoding="utf-8", low_memory=False)
        except Exception:
            return pd.DataFrame()

    gsc_map = {}
    df_gsc = load_csv_safe("search_console_all.csv")
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

    ga_map = {}
    df_ga = load_csv_safe("analytics_all.csv")
    if not df_ga.empty:
        a = _gc(df_ga, "Address")
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[str(r[a])] = safe_num(r.get(s, 0))

    inlink_counts = defaultdict(int)
    CHUNK = 50000
    for chunk in pd.read_csv(csv_path, encoding="utf-8", low_memory=False, chunksize=CHUNK):
        lp_col = _gc(chunk, "Link Position")
        src_col = _gc(chunk, "Source")
        dst_col = _gc(chunk, "Destination")
        if not (lp_col and src_col and dst_col):
            break
        mask = chunk[lp_col].str.lower() == "content"
        chunk = chunk[mask]
        for _, row in chunk.iterrows():
            src = str(row[src_col])
            dst = str(row[dst_col])
            if src == dst:
                continue
            src_info = internal_map.get(src, {})
            if src_info.get("indexability", "").lower() != "indexable":
                continue
            inlink_counts[dst] += 1

    all_rows = []
    for chunk in pd.read_csv(csv_path, encoding="utf-8", low_memory=False, chunksize=CHUNK):
        lp_col = _gc(chunk, "Link Position")
        src_col = _gc(chunk, "Source")
        dst_col = _gc(chunk, "Destination")
        type_col = _gc(chunk, "Type")
        alt_col = _gc(chunk, "Alt Text")
        anc_col = _gc(chunk, "Anchor")
        sc_col = _gc(chunk, "Status Code")
        st_col = _gc(chunk, "Status")
        fol_col = _gc(chunk, "Follow")
        lo_col = _gc(chunk, "Link Origin")
        if not (lp_col and src_col and dst_col):
            break
        mask = chunk[lp_col].str.lower() == "content"
        chunk = chunk[mask]
        for _, row in chunk.iterrows():
            src = str(row.get(src_col, ""))
            dst = str(row.get(dst_col, ""))
            is_self = src == dst
            src_info = internal_map.get(src, {})
            src_idx = src_info.get("indexability", "")
            if src_idx.lower() != "indexable":
                continue
            src_theme1, src_theme2, _, _ = classify_url(src, rulebook)
            dst_theme1, dst_theme2, _, _ = classify_url(dst, rulebook)
            gsc_src = gsc_map.get(src, {})
            gsc_dst = gsc_map.get(dst, {})
            inlink_count = inlink_counts.get(dst, 0)
            all_rows.append({
                "Source": src,
                "Source Page Theme 1": src_theme1 or "-",
                "Source Page Theme 2": src_theme2 if src_theme2 else "-",
                "Source Indexability": src_idx,
                "Source Status Code": src_info.get("status_code", ""),
                "Source Status": src_info.get("status", ""),
                "Destination": dst,
                "Dest Page Theme 1": dst_theme1 or "-",
                "Dest Page Theme 2": dst_theme2 if dst_theme2 else "-",
                "Alt Text": "" if str(row.get(alt_col, "")) in ("nan", "None", "") else str(row.get(alt_col, "")),
                "Anchor": "" if str(row.get(anc_col, "")) in ("nan", "None", "") else str(row.get(anc_col, "")),
                "Dest Status Code": str(row.get(sc_col, "")) if sc_col else "",
                "Dest Status": str(row.get(st_col, "")) if st_col else "",
                "Follow": str(row.get(fol_col, "")) if fol_col else "",
                "Type": str(row.get(type_col, "")) if type_col else "",
                "Link Position": str(row.get(lp_col, "")) if lp_col else "",
                "Link Origin": str(row.get(lo_col, "")) if lo_col else "",
                "Impressions Source": gsc_src.get("impressions"),
                "Clicks Source": gsc_src.get("clicks"),
                "Sessions Source": ga_map.get(src),
                "Impressions Dest": gsc_dst.get("impressions"),
                "Clicks Dest": gsc_dst.get("clicks"),
                "Sessions Dest": ga_map.get(dst),
                "Inlinks": inlink_count,
                "Inlinks Zone": _inlink_zone(inlink_count),
                "Is Self Link": "True" if is_self else "False",
            })

    df_t2 = pd.DataFrame(all_rows) if all_rows else pd.DataFrame(columns=[
        "Source", "Source Page Theme 1", "Source Page Theme 2",
        "Source Indexability", "Source Status Code", "Source Status",
        "Destination", "Dest Page Theme 1", "Dest Page Theme 2",
        "Alt Text", "Anchor", "Dest Status Code", "Dest Status",
        "Follow", "Type", "Link Position", "Link Origin",
        "Impressions Source", "Clicks Source", "Sessions Source",
        "Impressions Dest", "Clicks Dest", "Sessions Dest",
        "Inlinks", "Inlinks Zone", "Is Self Link",
    ])
    if not df_t2.empty:
        df_t2 = df_t2.sort_values("Impressions Dest", ascending=False, na_position="last").reset_index(drop=True)
    total_links = len(df_t2)
    theme_priority = {}
    for _, row in df_t2.iterrows():
        theme = row["Source Page Theme 1"] if row["Source Page Theme 1"] != "-" else "Others"
        _, _, _, priority = classify_url(row["Source"], rulebook)
        if theme not in theme_priority:
            theme_priority[theme] = priority
        else:
            if PRIORITY_ORDER.get(priority, 3) < PRIORITY_ORDER.get(theme_priority[theme], 3):
                theme_priority[theme] = priority
    sorted_themes = sorted(theme_priority.items(), key=lambda x: PRIORITY_ORDER.get(x[1], 3))
    link_types = ["Hyperlink", "JavaScript", "Iframe"]
    t1_data = {}
    for theme, _ in sorted_themes:
        t1_data[theme] = {"Hyperlink": 0, "JavaScript": 0, "Iframe": 0}
    if not df_t2.empty:
        for _, row in df_t2.iterrows():
            theme = row["Source Page Theme 1"] if row["Source Page Theme 1"] != "-" else "Others"
            lt = str(row.get("Type", "")).strip()
            if lt in t1_data.get(theme, {}):
                t1_data[theme][lt] += 1
    dest_info = {}
    if not df_t2.empty:
        for _, row in df_t2.iterrows():
            dst = row["Destination"]
            if dst not in dest_info:
                dest_info[dst] = {"theme": row["Dest Page Theme 1"] if row["Dest Page Theme 1"] != "-" else "Others", "inlinks": row["Inlinks"]}
    dash_data = defaultdict(lambda: defaultdict(int))
    for dst, info in dest_info.items():
        dash_data[info["theme"]][_inlink_zone(info["inlinks"])] += 1
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True})
    RED = "#FF0000"
    WHITE = "#FFFFFF"
    BLACK = "#000000"
    DARK = "#404040"
    FONT = "Rockwell"
    def f(**kw):
        return wb.add_format(kw)
    f_title = f(bold=True, font_name=FONT, font_size=12, font_color=BLACK)
    f_red_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
    f_red_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED, border=1, align="left", valign="vcenter", text_wrap=True)
    f_dark_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=DARK, border=1, align="center", valign="vcenter", text_wrap=True)
    f_dark_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=DARK, border=1, align="left", valign="vcenter", text_wrap=True)
    f_section = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK)
    f_ctr = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_lft = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_rgt = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="right", valign="vcenter")
    f_bold_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_bold_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0.0%")
    f_summary_box = f(font_name=FONT, font_size=8, font_color=BLACK, text_wrap=True, valign="top", border=1)
    f_merged_red = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED, border=1, align="center", valign="vcenter")
    ws_dash = wb.add_worksheet("Dashboard")
    ws_dash.set_column("A:A", 22)
    ws_dash.set_column("B:B", 22)
    for col in ["C", "D", "E", "F", "G", "H", "I"]:
        ws_dash.set_column(f"{col}:{col}", 18)
    ws_dash.set_column("J:J", 12)
    ws_dash.merge_range(0, 0, 0, 9, "Functional Internal Links Summary", f_title)
    ws_dash.merge_range(3, 0, 3, 1, "Total qualifying links analysed", f_bold_lft)
    ws_dash.write(3, 2, total_links, f_num)
    ws_dash.set_row(6, 42)
    ws_dash.write(6, 0, "Page Theme 1", f_red_lft)
    ws_dash.write(6, 1, "Priority Basis Page Theme 1", f_red_ctr)
    for i, bucket in enumerate(INLINK_BUCKETS):
        ws_dash.write(6, i + 2, bucket, f_red_ctr)
    ws_dash.write(6, 9, "TOTAL", f_red_ctr)
    theme_row_start = 7
    for row_offset, (theme, priority) in enumerate(sorted_themes):
        r = theme_row_start + row_offset
        ws_dash.write(r, 0, theme, f_bold_lft)
        ws_dash.write(r, 1, priority, f_ctr)
        row_total = 0
        for i, bucket in enumerate(INLINK_BUCKETS):
            val = dash_data[theme].get(bucket, 0)
            ws_dash.write(r, i + 2, val, f_num)
            row_total += val
        ws_dash.write(r, 9, row_total, f_num)
    total_row = theme_row_start + len(sorted_themes)
    ws_dash.write(total_row, 0, "Total Pages", f_bold_lft)
    ws_dash.write(total_row, 1, "-", f_ctr)
    for i, bucket in enumerate(INLINK_BUCKETS):
        col_letter = chr(ord('C') + i)
        ws_dash.write_formula(total_row, i + 2, f"=SUM({col_letter}{theme_row_start + 1}:{col_letter}{total_row})", f_num)
    ws_dash.write_formula(total_row, 9, f"=SUM(J{theme_row_start + 1}:J{total_row})", f_num)
    pct_row = total_row + 1
    ws_dash.write(pct_row, 0, "% of Master List", f_bold_lft)
    ws_dash.write(pct_row, 1, "-", f_ctr)
    for i in range(len(INLINK_BUCKETS)):
        col_letter = chr(ord('C') + i)
        ws_dash.write_formula(pct_row, i + 2, f"=IF(J{total_row + 1}>0,{col_letter}{total_row + 1}/J{total_row + 1},\"\")", f_pct)
    ws_dash.write(pct_row, 9, "100.0%", f_pct)
    ws_fl = wb.add_worksheet("Functional links")
    ws_fl.set_column("A:A", 50)
    ws_fl.set_column("B:B", 18)
    ws_fl.set_column("C:C", 18)
    ws_fl.set_column("D:D", 15)
    ws_fl.set_column("E:E", 15)
    ws_fl.set_column("F:F", 15)
    ws_fl.set_column("G:G", 50)
    ws_fl.set_column("H:H", 18)
    ws_fl.set_column("I:I", 18)
    ws_fl.set_column("J:J", 18)
    ws_fl.set_column("K:K", 20)
    ws_fl.set_column("L:L", 15)
    ws_fl.set_column("M:M", 15)
    ws_fl.set_column("N:N", 10)
    ws_fl.set_column("O:O", 15)
    ws_fl.set_column("P:P", 15)
    ws_fl.set_column("Q:Q", 15)
    for col in ["R", "S", "T", "U", "V", "W"]:
        ws_fl.set_column(f"{col}:{col}", 15)
    ws_fl.set_column("X:X", 10)
    ws_fl.set_column("Y:Y", 30)
    ws_fl.set_column("Z:Z", 25)
    ws_fl.merge_range(0, 0, 3, 25, "Issue Summary: Functional Internal Links Analysis evaluates internal links pointing to 200 OK destination URLs. Only links with Link Position = Content are included. Source URLs must be indexable. Self-links (source = destination) are excluded from inlink counts.", f_summary_box)
    ws_fl.set_row(0, 50)
    ws_fl.write(4, 0, "Table 1", f_section)
    ws_fl.merge_range(5, 0, 5, 4, "Page Theme Wise URL Analysis ", f_merged_red)
    ws_fl.set_row(6, 42)
    ws_fl.write(6, 0, "Page Theme 1", f_dark_lft)
    ws_fl.write(6, 1, "Priority Basis Page Theme 1", f_dark_ctr)
    ws_fl.write(6, 2, "# Links\nLink Type - Hyperlink", f_dark_ctr)
    ws_fl.write(6, 3, "# Links\nLink Type - Javascript", f_dark_ctr)
    ws_fl.write(6, 4, "# Links\nLink Type - Iframe", f_dark_ctr)
    T1_DATA_START = 7
    for row_offset, (theme, priority) in enumerate(sorted_themes):
        r = T1_DATA_START + row_offset
        ws_fl.write(r, 0, theme, f_bold_lft)
        ws_fl.write(r, 1, priority, f_ctr)
        ws_fl.write(r, 2, t1_data.get(theme, {}).get("Hyperlink", 0), f_num)
        ws_fl.write(r, 3, t1_data.get(theme, {}).get("JavaScript", 0), f_num)
        ws_fl.write(r, 4, t1_data.get(theme, {}).get("Iframe", 0), f_num)
    T2_LABEL_ROW = T1_DATA_START + len(sorted_themes) + 1
    ws_fl.write(T2_LABEL_ROW, 0, "Table 2", f_section)
    T2_HDR_ROW = T2_LABEL_ROW + 1
    T2_DATA_START = T2_HDR_ROW + 1
    ws_fl.set_row(T2_HDR_ROW, 31.5)
    t2_headers = ["Source", "Page Theme 1", "Page Theme 2", "Source - Indexability", "Source - Status Code", "Source - Status", "Destination", "Page Theme 1", "Page Theme 2", "Alt Text", "Anchor", "Destination - Status Code", "Destination - Status", "Follow", "Type", "Link Position", "Link Origin", "Impressions - Source", "Clicks - Source", "Organic Sessions - Source", "Impressions - Destination", "Clicks - Destination", "Organic Sessions - Destination", "# Inlinks", "# Inlinks - Zones", "Is source URL = Destination URL?"]
    for i, h in enumerate(t2_headers):
        ws_fl.write(T2_HDR_ROW, i, h, f_red_lft if i == 0 else f_red_ctr)
    for row_offset, (_, row) in enumerate(df_t2.iterrows()):
        r = T2_DATA_START + row_offset
        ws_fl.set_row(r, 14.5)
        ws_fl.write(r, 0, row.get("Source", "") or "", f_lft)
        ws_fl.write(r, 1, row.get("Source Page Theme 1", "") or "-", f_ctr)
        ws_fl.write(r, 2, row.get("Source Page Theme 2", "") or "-", f_ctr)
        ws_fl.write(r, 3, row.get("Source Indexability", "") or "", f_ctr)
        ws_fl.write(r, 4, row.get("Source Status Code", "") or "", f_ctr)
        ws_fl.write(r, 5, row.get("Source Status", "") or "", f_ctr)
        ws_fl.write(r, 6, row.get("Destination", "") or "", f_lft)
        ws_fl.write(r, 7, row.get("Dest Page Theme 1", "") or "-", f_ctr)
        ws_fl.write(r, 8, row.get("Dest Page Theme 2", "") or "-", f_ctr)
        ws_fl.write(r, 9, row.get("Alt Text", "") or "", f_lft)
        ws_fl.write(r, 10, row.get("Anchor", "") or "", f_lft)
        ws_fl.write(r, 11, row.get("Dest Status Code", "") or "", f_ctr)
        ws_fl.write(r, 12, row.get("Dest Status", "") or "", f_ctr)
        ws_fl.write(r, 13, row.get("Follow", "") or "", f_ctr)
        ws_fl.write(r, 14, row.get("Type", "") or "", f_ctr)
        ws_fl.write(r, 15, row.get("Link Position", "") or "", f_ctr)
        ws_fl.write(r, 16, row.get("Link Origin", "") or "", f_ctr)
        ws_fl.write(r, 17, safe_num(row.get("Impressions Source")), f_rgt)
        ws_fl.write(r, 18, safe_num(row.get("Clicks Source")), f_rgt)
        ws_fl.write(r, 19, safe_num(row.get("Sessions Source")), f_rgt)
        ws_fl.write(r, 20, safe_num(row.get("Impressions Dest")), f_rgt)
        ws_fl.write(r, 21, safe_num(row.get("Clicks Dest")), f_rgt)
        ws_fl.write(r, 22, safe_num(row.get("Sessions Dest")), f_rgt)
        ws_fl.write(r, 23, row.get("Inlinks", 0), f_num)
        ws_fl.write(r, 24, row.get("Inlinks Zone", "") or "", f_ctr)
        ws_fl.write_formula(r, 25, "=IF(A"+str(r+1)+"=G"+str(r+1)+",True,False)", f_ctr)
    wb.close()
    buf.seek(0)
    return buf.read()
