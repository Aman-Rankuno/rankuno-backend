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

CHUNK = 200_000

# Excel hard limit is 1,048,576 rows. Headers and tables sit above Table 2,
# so cap Table 2 at 1,000,000 rows. Rows are sorted by Impressions Dest
# descending before the cap, so the highest value links are always kept.
MAX_T2_ROWS = 1_000_000


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


def _clean_text(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    return s.where(~s.isin(["nan", "None", "<NA>"]), "")


def build_functional_internal_links_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    rulebook = load_rulebook(domain)

    # ── Locate the inlinks CSV ────────────────────────────────────────────
    csv_path = None
    for name in CSV_NAMES:
        p = os.path.join(report_path, name)
        if os.path.exists(p):
            csv_path = p
            break
    if not csv_path:
        raise FileNotFoundError("success_(2xx)_inlinks.csv not found in report folder")

    # ── internal_all: indexability and status lookups (vectorized) ───────
    internal_idx, internal_sc, internal_st = {}, {}, {}
    internal_all_path = os.path.join(report_path, "internal_all.csv")
    if os.path.exists(internal_all_path):
        df_int = pd.read_csv(
            internal_all_path, encoding="utf-8", low_memory=False,
            usecols=lambda c: c.lower() in ("address", "indexability", "status code", "status"),
        )
        a_col = _gc(df_int, "Address")
        idx_col = _gc(df_int, "Indexability")
        sc_col = _gc(df_int, "Status Code")
        st_col = _gc(df_int, "Status")
        if a_col:
            addr = df_int[a_col].astype(str)
            if idx_col:
                internal_idx = dict(zip(addr, df_int[idx_col].astype(str)))
            if sc_col:
                internal_sc = dict(zip(addr, df_int[sc_col].astype(str)))
            if st_col:
                internal_st = dict(zip(addr, df_int[st_col].astype(str)))

    indexable_set = {u for u, v in internal_idx.items() if v.strip().lower() == "indexable"}

    # ── GSC and GA lookups (vectorized) ───────────────────────────────────
    def load_csv_safe(filename):
        path = os.path.join(report_path, filename)
        if not os.path.exists(path):
            return pd.DataFrame()
        try:
            return pd.read_csv(path, encoding="utf-8", low_memory=False)
        except Exception:
            return pd.DataFrame()

    gsc_imp, gsc_clk = {}, {}
    df_gsc = load_csv_safe("search_console_all.csv")
    if not df_gsc.empty:
        a = _gc(df_gsc, "Address")
        imp = next((c for c in df_gsc.columns if "impression" in c.lower()), None)
        clk = next((c for c in df_gsc.columns if "click" in c.lower()), None)
        if a:
            addr = df_gsc[a].astype(str)
            if imp:
                gsc_imp = {u: safe_num(v) for u, v in zip(addr, df_gsc[imp])}
            if clk:
                gsc_clk = {u: safe_num(v) for u, v in zip(addr, df_gsc[clk])}

    ga_map = {}
    df_ga = load_csv_safe("analytics_all.csv")
    if not df_ga.empty:
        a = _gc(df_ga, "Address")
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            ga_map = {u: safe_num(v) for u, v in zip(df_ga[a].astype(str), df_ga[s])}

    # ── Single pass over the inlinks CSV ──────────────────────────────────
    # Filter each chunk down to Link Position = Content with an indexable
    # source, then concatenate. This replaces the previous two full reads.
    parts = []
    col_names = {}
    for chunk in pd.read_csv(csv_path, encoding="utf-8", low_memory=False, chunksize=CHUNK):
        if not col_names:
            col_names = {
                "lp": _gc(chunk, "Link Position"),
                "src": _gc(chunk, "Source"),
                "dst": _gc(chunk, "Destination"),
                "type": _gc(chunk, "Type"),
                "alt": _gc(chunk, "Alt Text"),
                "anc": _gc(chunk, "Anchor"),
                "sc": _gc(chunk, "Status Code"),
                "st": _gc(chunk, "Status"),
                "fol": _gc(chunk, "Follow"),
                "lo": _gc(chunk, "Link Origin"),
            }
        if not (col_names["lp"] and col_names["src"] and col_names["dst"]):
            break
        mask = chunk[col_names["lp"]].astype(str).str.lower() == "content"
        chunk = chunk[mask]
        if chunk.empty:
            continue
        chunk = chunk[chunk[col_names["src"]].astype(str).isin(indexable_set)]
        if not chunk.empty:
            parts.append(chunk)

    df = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
    del parts

    t2_columns = [
        "Source", "Source Page Theme 1", "Source Page Theme 2",
        "Source Indexability", "Source Status Code", "Source Status",
        "Destination", "Dest Page Theme 1", "Dest Page Theme 2",
        "Alt Text", "Anchor", "Dest Status Code", "Dest Status",
        "Follow", "Type", "Link Position", "Link Origin",
        "Impressions Source", "Clicks Source", "Sessions Source",
        "Impressions Dest", "Clicks Dest", "Sessions Dest",
        "Inlinks", "Inlinks Zone", "Is Self Link",
    ]

    if not df.empty:
        src = df[col_names["src"]].astype(str)
        dst = df[col_names["dst"]].astype(str)

        # Inlink counts: non-self links only, indexable sources already enforced
        inlink_counts = dst[src.values != dst.values].value_counts().to_dict()

        # Classify each unique URL exactly once
        cls_cache = {}
        for u in set(src).union(set(dst)):
            cls_cache[u] = classify_url(u, rulebook)

        def _t1(u):
            return cls_cache[u][0] or "-"

        def _t2c(u):
            return cls_cache[u][1] or "-"

        def _col(key, default=""):
            c = col_names.get(key)
            return _clean_text(df[c]) if c else pd.Series(default, index=df.index)

        df_t2 = pd.DataFrame({
            "Source": src.values,
            "Source Page Theme 1": src.map(_t1).values,
            "Source Page Theme 2": src.map(_t2c).values,
            "Source Indexability": src.map(lambda u: internal_idx.get(u, "")).values,
            "Source Status Code": src.map(lambda u: internal_sc.get(u, "")).values,
            "Source Status": src.map(lambda u: internal_st.get(u, "")).values,
            "Destination": dst.values,
            "Dest Page Theme 1": dst.map(_t1).values,
            "Dest Page Theme 2": dst.map(_t2c).values,
            "Alt Text": _col("alt").values,
            "Anchor": _col("anc").values,
            "Dest Status Code": (df[col_names["sc"]].astype(str) if col_names["sc"] else pd.Series("", index=df.index)).values,
            "Dest Status": (df[col_names["st"]].astype(str) if col_names["st"] else pd.Series("", index=df.index)).values,
            "Follow": (df[col_names["fol"]].astype(str) if col_names["fol"] else pd.Series("", index=df.index)).values,
            "Type": (df[col_names["type"]].astype(str) if col_names["type"] else pd.Series("", index=df.index)).values,
            "Link Position": df[col_names["lp"]].astype(str).values,
            "Link Origin": (df[col_names["lo"]].astype(str) if col_names["lo"] else pd.Series("", index=df.index)).values,
            "Impressions Source": src.map(lambda u: gsc_imp.get(u)).values,
            "Clicks Source": src.map(lambda u: gsc_clk.get(u)).values,
            "Sessions Source": src.map(lambda u: ga_map.get(u)).values,
            "Impressions Dest": dst.map(lambda u: gsc_imp.get(u)).values,
            "Clicks Dest": dst.map(lambda u: gsc_clk.get(u)).values,
            "Sessions Dest": dst.map(lambda u: ga_map.get(u)).values,
            "Inlinks": dst.map(lambda u: inlink_counts.get(u, 0)).values,
            "Is Self Link": (src.values == dst.values),
        })
        df_t2["Inlinks Zone"] = df_t2["Inlinks"].map(_inlink_zone)
        df_t2 = df_t2.sort_values("Impressions Dest", ascending=False, na_position="last").reset_index(drop=True)
        if len(df_t2) > MAX_T2_ROWS:
            df_t2 = df_t2.head(MAX_T2_ROWS)
        del df, src, dst
    else:
        df_t2 = pd.DataFrame(columns=t2_columns)
        cls_cache = {}

    total_links = len(df_t2)

    # ── Theme priority (one classify lookup per unique source) ───────────
    theme_priority = {}
    if not df_t2.empty:
        for u, t in zip(df_t2["Source"], df_t2["Source Page Theme 1"]):
            theme = t if t != "-" else "Others"
            priority = cls_cache[u][3]
            current = theme_priority.get(theme)
            if current is None or PRIORITY_ORDER.get(priority, 3) < PRIORITY_ORDER.get(current, 3):
                theme_priority[theme] = priority
    sorted_themes = sorted(theme_priority.items(), key=lambda x: PRIORITY_ORDER.get(x[1], 3))

    # ── Table 1: theme x link type counts (crosstab) ──────────────────────
    link_types_ordered = []
    t1_data = {}
    if not df_t2.empty:
        theme_series = df_t2["Source Page Theme 1"].where(df_t2["Source Page Theme 1"] != "-", "Others")
        type_series = df_t2["Type"].astype(str).str.strip()
        ct = pd.crosstab(theme_series, type_series)
        link_types_ordered = sorted(list(ct.columns), key=lambda x: (0 if x == "Hyperlink" else 1, x))
        for theme, _ in sorted_themes:
            t1_data[theme] = {lt: int(ct.at[theme, lt]) if theme in ct.index and lt in ct.columns else 0 for lt in link_types_ordered}
    else:
        link_types_ordered = ["Hyperlink", "JavaScript", "Iframe"]
        t1_data = {theme: {lt: 0 for lt in link_types_ordered} for theme, _ in sorted_themes}

    # Re-sort for Table 1: priority then Hyperlink descending
    sorted_themes = sorted(
        sorted_themes,
        key=lambda x: (PRIORITY_ORDER.get(x[1], 3), -t1_data.get(x[0], {}).get('Hyperlink', 0))
    )

    # ── Dashboard: destination theme x inlink zone ────────────────────────
    dash_data = defaultdict(lambda: defaultdict(int))
    if not df_t2.empty:
        dd = df_t2.drop_duplicates(subset=["Destination"], keep="first")
        for theme, zone in zip(dd["Dest Page Theme 1"], dd["Inlinks Zone"]):
            theme = theme if theme != "-" else "Others"
            dash_data[theme][zone] += 1

    # Sort dashboard: priority then total descending
    dash_sorted_themes = sorted(
        sorted_themes,
        key=lambda x: (PRIORITY_ORDER.get(x[1], 3), -sum(dash_data.get(x[0], {}).values()))
    )

    # ── Workbook ──────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {
        "in_memory": False,
        "nan_inf_to_errors": True,
        "strings_to_urls": False,
    })

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
    f_bold_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="#,##0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK, bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0.0%")
    f_summary_box = f(font_name=FONT, font_size=8, font_color=BLACK, text_wrap=True, valign="top", border=1)
    f_merged_red = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED, border=1, align="center", valign="vcenter")

    # ── Dashboard sheet ───────────────────────────────────────────────────
    ws_dash = wb.add_worksheet("Dashboard")
    ws_dash.set_column("A:A", 22)
    ws_dash.set_column("B:B", 22)
    for col in ["C", "D", "E", "F", "G", "H", "I"]:
        ws_dash.set_column(f"{col}:{col}", 18)
    ws_dash.set_column("J:J", 12)

    ws_dash.merge_range(0, 0, 0, 9, "Functional Internal Links Summary", f_title)
    ws_dash.merge_range(1, 0, 1, 9, "Internal Link Audit & Opportunity Analysis\nEvaluate internal link distribution across key pages to uncover opportunities for improving crawlability, authority flow, and organic visibility.", f_summary_box)
    ws_dash.set_row(1, 36)
    ws_dash.merge_range(3, 0, 3, 1, "Total qualifying links analysed", f_bold_lft)
    ws_dash.write(3, 2, total_links, f_num)

    ws_dash.set_row(6, 42)
    ws_dash.write(6, 0, "Page Theme 1", f_red_lft)
    ws_dash.write(6, 1, "Priority Basis Page Theme 1", f_red_ctr)
    for i, bucket in enumerate(INLINK_BUCKETS):
        ws_dash.write(6, i + 2, bucket, f_red_ctr)
    ws_dash.write(6, 9, "TOTAL", f_red_ctr)

    theme_row_start = 7
    for row_offset, (theme, priority) in enumerate(dash_sorted_themes):
        r = theme_row_start + row_offset
        ws_dash.write(r, 0, theme, f_bold_lft)
        ws_dash.write(r, 1, priority, f_ctr)
        for i, bucket in enumerate(INLINK_BUCKETS):
            val = dash_data.get(theme, {}).get(bucket, 0)
            ws_dash.write(r, i + 2, val, f_num)
        col_start = chr(ord("C"))
        col_end = chr(ord("C") + len(INLINK_BUCKETS) - 1)
        row_total = sum(dash_data.get(theme, {}).values())
        ws_dash.write_formula(r, 9, f"=SUM({col_start}{r+1}:{col_end}{r+1})", f_num, row_total)

    total_row = theme_row_start + len(dash_sorted_themes)
    ws_dash.write(total_row, 0, "Total Pages", f_bold_lft)
    ws_dash.write(total_row, 1, "-", f_ctr)
    col_totals = []
    for i, bucket in enumerate(INLINK_BUCKETS):
        col_letter = chr(ord('C') + i)
        col_sum = sum(dash_data[t].get(bucket, 0) for t, _ in dash_sorted_themes)
        col_totals.append(col_sum)
        ws_dash.write_formula(
            total_row, i + 2,
            f"=SUM({col_letter}{theme_row_start + 1}:{col_letter}{total_row})",
            f_num, col_sum,
        )
    grand_total = sum(col_totals)
    ws_dash.write_formula(
        total_row, 9,
        f"=SUM(J{theme_row_start + 1}:J{total_row})",
        f_num, grand_total,
    )

    pct_row = total_row + 1
    ws_dash.write(pct_row, 0, "% of Total URLs analysed", f_bold_lft)
    ws_dash.write(pct_row, 1, "-", f_ctr)
    for i in range(len(INLINK_BUCKETS)):
        col_letter = chr(ord('C') + i)
        cached_pct = (col_totals[i] / grand_total) if grand_total else ""
        ws_dash.write_formula(
            pct_row, i + 2,
            f"=IF(J{total_row + 1}>0,{col_letter}{total_row + 1}/J{total_row + 1},\"\")",
            f_pct, cached_pct,
        )
    ws_dash.write(pct_row, 9, "100.0%", f_pct)

    # ── Functional links sheet ────────────────────────────────────────────
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
    ws_fl.set_default_row(14.5)

    ws_fl.merge_range(0, 0, 3, 4, "Internal Link Audit & Opportunity Analysis\nEvaluate internal link distribution across key pages to uncover opportunities for improving crawlability, authority flow, and organic visibility.", f_summary_box)
    ws_fl.set_row(0, 50)

    ws_fl.write(4, 0, "Table 1", f_section)
    ws_fl.merge_range(5, 0, 5, 4, "Page Theme Wise Internal Links Analysis", f_merged_red)
    ws_fl.set_row(6, 42)
    ws_fl.write(6, 0, "Page Theme 1", f_red_lft)
    ws_fl.write(6, 1, "Priority Basis Page Theme 1", f_red_ctr)
    for lt_i, lt in enumerate(link_types_ordered):
        ws_fl.write(6, lt_i + 2, f"# Links\nLink Type - {lt}", f_dark_ctr)

    T1_DATA_START = 7
    for row_offset, (theme, priority) in enumerate(sorted_themes):
        r = T1_DATA_START + row_offset
        ws_fl.write(r, 0, theme, f_bold_lft)
        ws_fl.write(r, 1, priority, f_ctr)
        for lt_i, lt in enumerate(link_types_ordered):
            ws_fl.write(r, lt_i + 2, t1_data.get(theme, {}).get(lt, 0), f_num)

    T2_LABEL_ROW = T1_DATA_START + len(sorted_themes) + 1
    ws_fl.write(T2_LABEL_ROW, 0, "Table 2", f_section)
    T2_HDR_ROW = T2_LABEL_ROW + 1
    T2_DATA_START = T2_HDR_ROW + 1
    ws_fl.set_row(T2_HDR_ROW, 31.5)

    t2_headers = ["Source", "Page Theme 1", "Page Theme 2", "Source - Indexability", "Source - Status Code", "Source - Status", "Destination", "Page Theme 1", "Page Theme 2", "Alt Text", "Anchor", "Destination - Status Code", "Destination - Status", "Follow", "Type", "Link Position", "Link Origin", "Impressions - Source", "Clicks - Source", "Organic Sessions - Source", "Impressions - Destination", "Clicks - Destination", "Organic Sessions - Destination", "# Inlinks to the Destination Page", "# Inlinks - Zones"]
    for i, h in enumerate(t2_headers):
        ws_fl.write(T2_HDR_ROW, i, h, f_red_lft if i == 0 else f_red_ctr)

    # Pull columns out as Python lists once, then write row by row.
    # This avoids the per-row pandas overhead of iterrows().
    if not df_t2.empty:
        c_src = df_t2["Source"].tolist()
        c_spt1 = df_t2["Source Page Theme 1"].tolist()
        c_spt2 = df_t2["Source Page Theme 2"].tolist()
        c_sidx = df_t2["Source Indexability"].tolist()
        c_ssc = df_t2["Source Status Code"].tolist()
        c_sst = df_t2["Source Status"].tolist()
        c_dst = df_t2["Destination"].tolist()
        c_dpt1 = df_t2["Dest Page Theme 1"].tolist()
        c_dpt2 = df_t2["Dest Page Theme 2"].tolist()
        c_alt = df_t2["Alt Text"].tolist()
        c_anc = df_t2["Anchor"].tolist()
        c_dsc = df_t2["Dest Status Code"].tolist()
        c_dst_st = df_t2["Dest Status"].tolist()
        c_fol = df_t2["Follow"].tolist()
        c_typ = df_t2["Type"].tolist()
        c_lp = df_t2["Link Position"].tolist()
        c_lo = df_t2["Link Origin"].tolist()
        c_imp_s = df_t2["Impressions Source"].tolist()
        c_clk_s = df_t2["Clicks Source"].tolist()
        c_ses_s = df_t2["Sessions Source"].tolist()
        c_imp_d = df_t2["Impressions Dest"].tolist()
        c_clk_d = df_t2["Clicks Dest"].tolist()
        c_ses_d = df_t2["Sessions Dest"].tolist()
        c_inl = df_t2["Inlinks"].tolist()
        c_zone = df_t2["Inlinks Zone"].tolist()
        c_self = df_t2["Is Self Link"].tolist()

        for idx in range(total_links):
            r = T2_DATA_START + idx
            ws_fl.write(r, 0, c_src[idx] or "", f_lft)
            ws_fl.write(r, 1, c_spt1[idx] or "-", f_ctr)
            ws_fl.write(r, 2, c_spt2[idx] or "-", f_ctr)
            ws_fl.write(r, 3, c_sidx[idx] or "", f_ctr)
            ws_fl.write(r, 4, c_ssc[idx] or "", f_ctr)
            ws_fl.write(r, 5, c_sst[idx] or "", f_ctr)
            ws_fl.write(r, 6, c_dst[idx] or "", f_lft)
            ws_fl.write(r, 7, c_dpt1[idx] or "-", f_ctr)
            ws_fl.write(r, 8, c_dpt2[idx] or "-", f_ctr)
            ws_fl.write(r, 9, c_alt[idx] or "", f_lft)
            ws_fl.write(r, 10, c_anc[idx] or "", f_lft)
            ws_fl.write(r, 11, c_dsc[idx] or "", f_ctr)
            ws_fl.write(r, 12, c_dst_st[idx] or "", f_ctr)
            ws_fl.write(r, 13, c_fol[idx] or "", f_ctr)
            ws_fl.write(r, 14, c_typ[idx] or "", f_ctr)
            ws_fl.write(r, 15, c_lp[idx] or "", f_ctr)
            ws_fl.write(r, 16, c_lo[idx] or "", f_ctr)
            ws_fl.write(r, 17, safe_num(c_imp_s[idx]), f_rgt)
            ws_fl.write(r, 18, safe_num(c_clk_s[idx]), f_rgt)
            ws_fl.write(r, 19, safe_num(c_ses_s[idx]), f_rgt)
            ws_fl.write(r, 20, safe_num(c_imp_d[idx]), f_rgt)
            ws_fl.write(r, 21, safe_num(c_clk_d[idx]), f_rgt)
            ws_fl.write(r, 22, safe_num(c_ses_d[idx]), f_rgt)
            ws_fl.write(r, 23, int(c_inl[idx]) if c_inl[idx] else 0, f_num)
            zone_val = c_zone[idx] if not c_self[idx] else ""
            ws_fl.write(r, 24, zone_val or "", f_ctr)

    wb.close()
    buf.seek(0)
    return buf.read()