import os
import io
import zipfile
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

BLACK_FONT = Font(name="Arial", size=9, color="FF000000")
PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}

T1_PRIORITY_ROW = 13
T1_DATA_START_ROW = 15
T1_PCT_ROW = 16
T2_DATA_START_ROW = 21
T2_DATA_END_ROW = 30
T3_DATA_START_ROW = 34

SC_KEYS = ["301", "302", "307", "308", "400", "401", "403", "404",
           "500", "502", "503", "504", "No response code"]

SC_PRIORITY = {
    "301": "Medium", "302": "High", "307": "High", "308": "Medium",
    "400": "High", "401": "High", "403": "High", "404": "High",
    "500": "High", "502": "High", "503": "High", "504": "High",
    "No response code": "High",
}

T1_SC_START_COL = 3
T4_CHAIN_COL = 18
T4_LOOP_COL = 19
T2_THEME_COL = 1
T2_PRIORITY_COL = 2
T2_SC_START_COL = 3

T3_COLS = [
    "Error type", "Address", "Page Theme 1", "Page Theme 2",
    "Content Type", "Status Code", "Status", "Indexability",
    "Indexability Status", "Inlinks", "Redirect URL", "Redirect Type",
    "Redirect chain", "Redirect loop", "Impressions", "Clicks", "Organic Sessions"
]

KEEP_FROM_TEMPLATE = {
    'xl/drawings/drawing1.xml',
    'xl/drawings/vmlDrawing1.vml',
    'xl/comments1.xml',
    'xl/persons/person.xml',
    'xl/documenttasks/documenttask1.xml',
    'xl/threadedComments/threadedComment1.xml',
}

DRAWING_CT = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
DRAWING_REL = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
VML_REL = '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'


def _merge_with_drawings(template_path: str, openpyxl_buf: io.BytesIO) -> bytes:
    openpyxl_buf.seek(0)

    # Step 1: merge openpyxl output + template drawing files
    step1_buf = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as t_zip:
        with zipfile.ZipFile(openpyxl_buf, 'r') as o_zip:
            with zipfile.ZipFile(step1_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
                for item in o_zip.namelist():
                    out_zip.writestr(item, o_zip.read(item))
                for item in KEEP_FROM_TEMPLATE:
                    if item in t_zip.namelist():
                        out_zip.writestr(item, t_zip.read(item))
                for item in t_zip.namelist():
                    if 'drawings/_rels' in item and item not in o_zip.namelist():
                        out_zip.writestr(item, t_zip.read(item))

    # Step 2: patch [Content_Types].xml and sheet1.xml.rels
    step1_buf.seek(0)
    step1_bytes = step1_buf.read()

    with zipfile.ZipFile(io.BytesIO(step1_bytes), 'r') as z:
        ct = z.read('[Content_Types].xml').decode('utf-8')
        sheet_rels = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')

    if 'drawing1' not in ct:
        ct = ct.replace('</Types>', DRAWING_CT + '</Types>')
    if 'drawing1' not in sheet_rels:
        sheet_rels = sheet_rels.replace('</Relationships>', DRAWING_REL + VML_REL + '</Relationships>')

    final_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(step1_bytes), 'r') as z:
        with zipfile.ZipFile(final_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for item in z.namelist():
                if item == '[Content_Types].xml':
                    out_zip.writestr(item, ct.encode('utf-8'))
                elif item == 'xl/worksheets/_rels/sheet1.xml.rels':
                    out_zip.writestr(item, sheet_rels.encode('utf-8'))
                else:
                    out_zip.writestr(item, z.read(item))

    final_buf.seek(0)
    return final_buf.read()


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

    ga_map = {}
    if not df_ga.empty:
        a = next((c for c in df_ga.columns if c.lower() == "address"), None)
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[str(r[a])] = r.get(s, 0)

    frames = [df for df in [df_3xx_t, df_4xx_t, df_5xx_t, df_no_t, df_blocked_t] if not df.empty]
    if not frames:
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

    rows = []
    for _, row in df_combined.iterrows():
        url = str(row.get(addr_col, "")) if addr_col else ""
        error_type = str(row.get("_error_type", ""))
        page_theme1, page_theme2, _, _ = classify_url(url, rulebook)
        gsc = gsc_map.get(url, {})
        redirect_chain = ("TRUE" if url in chain_urls else "") if error_type == "3xx" else ""
        redirect_loop = ("TRUE" if url in loop_urls else "") if error_type == "3xx" else ""

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

    sc_counts = {k: 0 for k in SC_KEYS}
    for _, row in df_table3.iterrows():
        sc = str(row["Status Code"])
        et = str(row["Error type"])
        if et == "No response code":
            sc_counts["No response code"] += 1
        elif sc in sc_counts:
            sc_counts[sc] += 1

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

    chain_count = len(df_redirect_chain) if not df_redirect_chain.empty else 0
    loop_count = len(df_redirect_loop) if not df_redirect_loop.empty else 0
    total_3xx = len(df_3xx) if not df_3xx.empty else 0

    wb = openpyxl.load_workbook(template_path, data_only=False)
    ws = wb.active

    for r in list(range(T1_DATA_START_ROW, T1_PCT_ROW + 1)) + list(range(T2_DATA_START_ROW, T2_DATA_END_ROW + 1)) + list(range(T3_DATA_START_ROW, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).font = BLACK_FONT

    for i, sc in enumerate(SC_KEYS):
        col = T1_SC_START_COL + i
        count = sc_counts[sc]
        ws.cell(T1_DATA_START_ROW, col).value = count if count > 0 else None
        ws.cell(T1_DATA_START_ROW, col).font = BLACK_FONT
        pct = f"{round(count / total_crawled * 100, 1)}%" if total_crawled > 0 and count > 0 else None
        ws.cell(T1_PCT_ROW, col).value = pct
        ws.cell(T1_PCT_ROW, col).font = BLACK_FONT

    
    ws.cell(T1_DATA_START_ROW, T4_CHAIN_COL).value = chain_count if chain_count > 0 else None
    ws.cell(T1_DATA_START_ROW, T4_CHAIN_COL).font = BLACK_FONT
    ws.cell(T1_DATA_START_ROW, T4_LOOP_COL).value = loop_count if loop_count > 0 else None
    ws.cell(T1_DATA_START_ROW, T4_LOOP_COL).font = BLACK_FONT
    pct_chain = f"{round(chain_count / total_3xx * 100, 1)}%" if total_3xx > 0 and chain_count > 0 else None
    pct_loop = f"{round(loop_count / total_3xx * 100, 1)}%" if total_3xx > 0 and loop_count > 0 else None
    ws.cell(T1_PCT_ROW, T4_CHAIN_COL).value = pct_chain
    ws.cell(T1_PCT_ROW, T4_CHAIN_COL).font = BLACK_FONT
    ws.cell(T1_PCT_ROW, T4_LOOP_COL).value = pct_loop
    ws.cell(T1_PCT_ROW, T4_LOOP_COL).font = BLACK_FONT

    for r in range(T2_DATA_START_ROW, T2_DATA_END_ROW + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    for row_offset, (theme, counts) in enumerate(sorted_themes):
        r = T2_DATA_START_ROW + row_offset
        total_theme = counts["total"]
        ws.cell(r, T2_THEME_COL).value = theme
        ws.cell(r, T2_THEME_COL).font = BLACK_FONT
        ws.cell(r, T2_PRIORITY_COL).value = counts.get("_priority", "High")
        ws.cell(r, T2_PRIORITY_COL).font = BLACK_FONT
        for i, sc in enumerate(SC_KEYS):
            cnt = counts.get(sc, 0)
            val = f"{cnt} ({round(cnt / total_theme * 100)}%)" if cnt > 0 and total_theme > 0 else None
            cell = ws.cell(r, T2_SC_START_COL + i)
            cell.value = val
            cell.font = BLACK_FONT

    for r in range(T3_DATA_START_ROW, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    for row_offset, (_, row) in enumerate(df_table3.iterrows()):
        r = T3_DATA_START_ROW + row_offset
        for col_offset, col_name in enumerate(T3_COLS):
            val = row.get(col_name, "")
            if pd.isna(val) or val == "":
                val = None
            cell = ws.cell(r, 1 + col_offset)
            cell.value = val
            cell.font = BLACK_FONT

    openpyxl_buf = io.BytesIO()
    wb.save(openpyxl_buf)
    return _merge_with_drawings(template_path, openpyxl_buf)