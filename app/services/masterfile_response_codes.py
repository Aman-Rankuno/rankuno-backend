import os
import io
import shutil
import zipfile
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}

# Template layout constants
T1_DATA_START_ROW = 15   # #Affected URLs row
T1_PCT_ROW = 16          # % Share row
T2_DATA_START_ROW = 21   # First theme data row
T2_DATA_END_ROW = 30     # Last sample theme row
T3_DATA_START_ROW = 34   # First data row in Table 3

SC_KEYS = ["301", "302", "307", "308", "400", "401", "403", "404",
           "500", "502", "503", "504", "No response code"]

# Column positions in Table 1 (B=2 is label, C=3 is first sc)
T1_LABEL_COL = 2
T1_SC_START_COL = 3  # C

# Column positions in Table 4
T4_LABEL_COL = 17   # Q
T4_CHAIN_COL = 18   # R
T4_LOOP_COL = 19    # S

# Table 2 columns
T2_THEME_COL = 1    # A
T2_PRIORITY_COL = 2 # B
T2_SC_START_COL = 3 # C

# Table 3 columns
T3_COLS = [
    "Error type", "Address", "Page Theme 1", "Page Theme 2",
    "Content Type", "Status Code", "Status", "Indexability",
    "Indexability Status", "Inlinks", "Redirect URL", "Redirect Type",
    "Redirect chain", "Redirect loop", "Impressions", "Clicks", "Organic Sessions"
]


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

    # Total URLs from internal_all (all content types)
    total_crawled = len(df_internal_all) if not df_internal_all.empty else 0

    def tag(df, label):
        if df.empty:
            return df
        df = df.copy()
        df["_error_type"] = label
        return df

    df_3xx_t = tag(df_3xx, "3xx")
    df_4xx_t = tag(df_4xx, "4xx")
    df_5xx_t = tag(df_5xx, "5xx")
    df_no_t = tag(df_no_response, "No response code")
    df_blocked_t = tag(df_blocked, "Blocked by robots")

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
                    "impressions": r.get(imp, 0) if imp else 0,
                    "clicks": r.get(clk, 0) if clk else 0,
                }

    # GA lookup
    ga_map = {}
    if not df_ga.empty:
        a = next((c for c in df_ga.columns if c.lower() == "address"), None)
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[str(r[a])] = r.get(s, 0)

    # Combine frames
    frames = [df for df in [df_3xx_t, df_4xx_t, df_5xx_t, df_no_t, df_blocked_t] if not df.empty]
    if not frames:
        # Return template as-is if no data
        with open(template_path, "rb") as f:
            return f.read()

    df_combined = pd.concat(frames, ignore_index=True)

    def get_col(df, *names):
        for n in names:
            m = next((c for c in df.columns if c.lower() == n.lower()), None)
            if m:
                return m
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
        page_theme1, page_theme2 = classify_url(url, rulebook)
        gsc = gsc_map.get(url, {})

        if error_type == "3xx":
            redirect_chain = "TRUE" if url in chain_urls else ""
            redirect_loop = "TRUE" if url in loop_urls else ""
        else:
            redirect_chain = ""
            redirect_loop = ""

        rows.append({
            "Error type": error_type,
            "Address": url,
            "Page Theme 1": page_theme1,
            "Page Theme 2": page_theme2,
            "Content Type": row.get(ct_col, "") if ct_col else "",
            "Status Code": row.get(sc_col, "") if sc_col else "",
            "Status": row.get(st_col, "") if st_col else "",
            "Indexability": row.get(idx_col, "") if idx_col else "",
            "Indexability Status": row.get(idx_st_col, "") if idx_st_col else "",
            "Inlinks": row.get(inl_col, 0) if inl_col else 0,
            "Redirect URL": row.get(ru_col, "") if ru_col else "",
            "Redirect Type": row.get(rt_col, "") if rt_col else "",
            "Redirect chain": redirect_chain,
            "Redirect loop": redirect_loop,
            "Impressions": gsc.get("impressions", 0),
            "Clicks": gsc.get("clicks", 0),
            "Organic Sessions": ga_map.get(url, 0),
        })

    df_table3 = pd.DataFrame(rows)
    df_table3 = df_table3.sort_values("Impressions", ascending=False).reset_index(drop=True)

    # Table 1 summary
    sc_counts = {k: 0 for k in SC_KEYS}
    for _, row in df_table3.iterrows():
        sc = str(row["Status Code"])
        et = str(row["Error type"])
        if et == "No response code":
            sc_counts["No response code"] += 1
        elif sc in sc_counts:
            sc_counts[sc] += 1

    # Table 2 - page theme wise
    theme_counts = {}
    for _, row in df_table3.iterrows():
        theme = row["Page Theme 1"] or "Others"
        sc = str(row["Status Code"])
        et = str(row["Error type"])
        if theme not in theme_counts:
            theme_counts[theme] = {"total": 0, "_priority": "Low"}
            for k in SC_KEYS:
                theme_counts[theme][k] = 0
        theme_counts[theme]["total"] += 1
        if et == "No response code":
            theme_counts[theme]["No response code"] += 1
        elif sc in theme_counts[theme]:
            theme_counts[theme][sc] += 1

    sorted_themes = sorted(
        theme_counts.items(),
        key=lambda x: PRIORITY_ORDER.get(x[1].get("_priority", "Low"), 2)
    )

    # Table 4
    chain_count = len(df_redirect_chain) if not df_redirect_chain.empty else 0
    loop_count = len(df_redirect_loop) if not df_redirect_loop.empty else 0
    total_3xx = len(df_3xx) if not df_3xx.empty else 0

    # Load template
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # ── Fill Table 1: #Affected URLs (row 15) and % Share (row 16)
    for i, sc in enumerate(SC_KEYS):
        col = T1_SC_START_COL + i
        count = sc_counts[sc]
        ws.cell(T1_DATA_START_ROW, col).value = count if count > 0 else None
        pct = f"{round(count / total_crawled * 100, 1)}%" if total_crawled > 0 and count > 0 else None
        ws.cell(T1_PCT_ROW, col).value = pct

    # ── Fill Table 4: #Affected URLs (row 15) and % Share (row 16)
    ws.cell(T1_DATA_START_ROW, T4_CHAIN_COL).value = chain_count if chain_count > 0 else None
    ws.cell(T1_DATA_START_ROW, T4_LOOP_COL).value = loop_count if loop_count > 0 else None
    pct_chain = f"{round(chain_count / total_3xx * 100, 1)}%" if total_3xx > 0 and chain_count > 0 else None
    pct_loop = f"{round(loop_count / total_3xx * 100, 1)}%" if total_3xx > 0 and loop_count > 0 else None
    ws.cell(T1_PCT_ROW, T4_CHAIN_COL).value = pct_chain
    ws.cell(T1_PCT_ROW, T4_LOOP_COL).value = pct_loop

    # ── Clear Table 2 sample data rows (21-30)
    for r in range(T2_DATA_START_ROW, T2_DATA_END_ROW + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    # ── Fill Table 2
    for row_offset, (theme, counts) in enumerate(sorted_themes):
        r = T2_DATA_START_ROW + row_offset
        total_theme = counts["total"]
        ws.cell(r, T2_THEME_COL).value = theme
        ws.cell(r, T2_PRIORITY_COL).value = counts.get("_priority", "High")
        for i, sc in enumerate(SC_KEYS):
            cnt = counts.get(sc, 0)
            if cnt > 0 and total_theme > 0:
                pct = round(cnt / total_theme * 100)
                val = f"{cnt} ({pct}%)"
            else:
                val = None
            ws.cell(r, T2_SC_START_COL + i).value = val

    # ── Clear Table 3 sample data rows (34 onwards)
    for r in range(T3_DATA_START_ROW, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    # ── Fill Table 3
    for row_offset, (_, row) in enumerate(df_table3.iterrows()):
        r = T3_DATA_START_ROW + row_offset
        for col_offset, col_name in enumerate(T3_COLS):
            val = row.get(col_name, "")
            if pd.isna(val) or val == "":
                val = None
            ws.cell(r, 1 + col_offset).value = val

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()