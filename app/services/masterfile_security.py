import os
import io
import math
import pandas as pd
import xlsxwriter
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

FONT = "Rockwell"
RED = "#FF0000"
WHITE = "#FFFFFF"
BLACK = "#000000"
DARK = "#1F3864"
GREY_FILL = "#D9D9D9"

# ============================================================
# Tab G1: issues sourced from Screaming Frog's per-issue
# --export-tabs CSVs (each already has Address/Content Type/
# Status Code/Indexability/Indexability Status columns natively
# - no VLOOKUP from internal_all needed here).
#
# page_scope: "all" = no content-type filter beyond Status 200;
#             "html" = additionally require Content Type starts
#             with "text/html" (per instruction #1).
# ============================================================
TAB1_ISSUES = [
    ("HTTP URLs", "security_http_urls.csv", "Issue", "High", "html"),
    ("Bad Content Type", "security_bad_content_type.csv", "Warning", "Low", "all"),
    ("Missing Secure Referrer-Policy", "security_missing_secure_referrerpolicy_header.csv", "Warning", "Low", "all"),
    ("Missing Content-Security-Policy", "security_missing_contentsecuritypolicy_header.csv", "Warning", "Low", "html"),
    ("Missing HSTS Header", "security_missing_hsts_header.csv", "Warning", "Low", "html"),
    ("Missing X-Content-Type-Options", "security_missing_xcontenttypeoptions_header.csv", "Warning", "Low", "all"),
    ("Missing X-Frames-Options Header", "security_missing_xframeoptions_header.csv", "Warning", "Low", "all"),
    ("Form On HTTP URL", "security_form_on_http_url.csv", "Issue", "High", "html"),
]

# ============================================================
# Tab G2: issues sourced from --bulk-export CSVs (Source/
# Destination link pairs). These CSVs have NO Content Type or
# Status Code columns - both are looked up from internal_all.csv
# by matching on the Source Page address, per instruction #2.
# ============================================================
TAB2_ISSUES = [
    ("Mixed Content", "mixed_content.csv", "Issue", "High"),
    ("Form URL Insecure", "form_url_insecure.csv", "Issue", "High"),
    ("Unsafe Cross-Origin Links", "unsafe_crossorigin_links.csv", "Warning", "Low"),
    ("Protocol-Relative Resource Link", "protocolrelative_outlinks.csv", "Warning", "Low"),
]

TAB1_SUMMARY = [
    ("security_http_urls", "Analyzing URLs using the insecure HTTP protocol instead of HTTPS, ensuring all crawlable assets are encrypted and secure for users."),
    ("security_bad_content_type", "Identifying URLs where the returned Content-Type header is invalid, missing, or mismatched with the actual asset type, which can trigger browser security warnings or rendering issues."),
    ("security_missing_secure_referrerpolicy_header", "Checking for the absence or misconfiguration of the Referrer-Policy header on HTTPS pages, preventing the accidental leakage of sensitive user data or internal URL structures to third-party sites."),
    ("security_missing_contentsecuritypolicy_header", "Identifying pages missing the Content-Security-Policy (CSP) header, leaving the site vulnerable to Cross-Site Scripting (XSS) and data injection attacks."),
    ("security_missing_hsts_header", "Pinpointing HTTPS pages missing the Strict-Transport-Security (HSTS) header, which fails to force browsers to interact with the site exclusively over secure connections."),
    ("security_missing_xcontenttypeoptions_header", "Flagging responses that lack the X-Content-Type-Options: nosniff header, exposing users to drive-by download attacks and MIME-sniffing vulnerabilities."),
    ("security_missing_xframeoptions_header", "Detecting URLs missing the X-Frame-Options (or CSP frame-ancestors) header, leaving the page susceptible to clickjacking attacks where malicious sites overlay content."),
    ("security_form_on_http_url", "Auditing forms hosted on insecure HTTP pages, which risk exposing sensitive user input (like login credentials or personal data) to interception via man-in-the-middle attacks."),
]

TAB2_SUMMARY = [
    ("mixed_content", "Identifying secure HTTPS pages that load resources (such as images, scripts, or stylesheets) over an insecure HTTP connection, which degrades browser security and triggers \"Not Secure\" warnings."),
    ("unsafe_crossorigin_links", "Auditing internal or external links that open in a new tab (target=\"_blank\") but lack the rel=\"noopener\" or rel=\"noreferrer\" attributes, exposing the source page to performance issues or malicious tab-jacking exploits."),
    ("protocolrelative_outlinks", "Detecting resource links that use the legacy //domain.com format instead of explicitly defining https://, which can result in assets failing to load or loading insecurely depending on the user's connection environment."),
    ("form_url_insecure", "Pinpointing forms hosted on secure HTTPS pages that submit user data (action attribute) to an insecure HTTP destination URL, rendering the submitted data vulnerable to interception."),
]


def _load_internal_all(report_path):
    """Load internal_all.csv into a lookup dict keyed by Address, plus
    the two crawl-wide totals both tabs' % share formulas reference."""
    path = os.path.join(report_path, "internal_all.csv")
    df = pd.read_csv(path, encoding="utf-8", low_memory=False,
                      usecols=lambda c: c.lower() in (
                          "address", "content type", "status code", "indexability", "indexability status"
                      ))
    cols = {c.lower(): c for c in df.columns}
    a, ct, sc = cols.get("address"), cols.get("content type"), cols.get("status code")
    idx, idxs = cols.get("indexability"), cols.get("indexability status")

    lookup = {}
    total_all_200 = 0
    total_html_200 = 0
    for _, row in df.iterrows():
        addr = str(row[a]) if a else None
        content_type = str(row[ct]) if ct and pd.notna(row[ct]) else ""
        status = str(row[sc]) if sc and pd.notna(row[sc]) else ""
        indexability = str(row[idx]) if idx and pd.notna(row[idx]) else ""
        indexability_status = str(row[idxs]) if idxs and pd.notna(row[idxs]) else ""
        if addr:
            lookup[addr] = (content_type, status, indexability, indexability_status)
        if status == "200":
            total_all_200 += 1
            if content_type.lower().startswith("text/html"):
                total_html_200 += 1
    return lookup, total_all_200, total_html_200


def _load_gsc_ga4(report_path):
    gsc_imp, gsc_clk, ga_sess = {}, {}, {}
    gsc_path = os.path.join(report_path, "search_console_all.csv")
    if os.path.exists(gsc_path):
        try:
            df = pd.read_csv(gsc_path, encoding="utf-8", low_memory=False)
            cols = {c.lower(): c for c in df.columns}
            ac = cols.get("address")
            ic = next((cols[c] for c in cols if "impression" in c), None)
            cc = next((cols[c] for c in cols if "click" in c), None)
            if ac:
                addr = df[ac].astype(str)
                if ic:
                    gsc_imp = dict(zip(addr, df[ic]))
                if cc:
                    gsc_clk = dict(zip(addr, df[cc]))
        except Exception:
            pass
    ga_path = os.path.join(report_path, "analytics_all.csv")
    if os.path.exists(ga_path):
        try:
            df = pd.read_csv(ga_path, encoding="utf-8", low_memory=False)
            cols = {c.lower(): c for c in df.columns}
            ac = cols.get("address")
            sc = next((cols[c] for c in cols if "session" in c), None)
            if ac and sc:
                ga_sess = dict(zip(df[ac].astype(str), df[sc]))
        except Exception:
            pass
    return gsc_imp, gsc_clk, ga_sess


def _safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f) or f == 0) else f
    except Exception:
        return None


def _read_issue_csv(report_path, filename):
    path = os.path.join(report_path, filename)
    if not os.path.exists(path) or os.path.getsize(path) < 20:
        return pd.DataFrame()
    try:
        return pd.read_csv(path, encoding="utf-8", low_memory=False)
    except Exception:
        return pd.DataFrame()


def build_security_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    rulebook = load_rulebook(domain)
    internal_lookup, total_all_200, total_html_200 = _load_internal_all(report_path)
    gsc_imp, gsc_clk, ga_sess = _load_gsc_ga4(report_path)

    cls_cache = {}

    def classify(url):
        if url not in cls_cache:
            cls_cache[url] = classify_url(url, rulebook)
        return cls_cache[url]

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True, "strings_to_urls": False})

    def f(**kw):
        return wb.add_format(kw)

    f_title = f(bold=True, font_name=FONT, font_size=13, font_color=BLACK)
    f_summary_hdr = f(bold=True, font_name=FONT, font_size=10, font_color=BLACK)
    f_summary_body = f(font_name=FONT, font_size=9, font_color=BLACK, text_wrap=True, valign="top")
    f_section = f(bold=True, font_name=FONT, font_size=11, font_color=WHITE, bg_color=DARK, border=1)
    f_red_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=RED,
                  border=1, align="center", valign="vcenter", text_wrap=True)
    f_dark_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE, bg_color=DARK,
                   border=1, align="center", valign="vcenter", text_wrap=True)
    f_grey_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK, bg_color=GREY_FILL,
                   border=1, align="center", valign="vcenter", text_wrap=True)
    f_cell = f(font_name=FONT, font_size=8, font_color=BLACK, border=1, align="center", valign="vcenter")
    f_cell_lft = f(font_name=FONT, font_size=8, font_color=BLACK, border=1, align="left", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK, border=1, align="center", valign="vcenter", num_format="#,##0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK, border=1, align="center", valign="vcenter", num_format="0.00%")
    f_note = f(italic=True, font_name=FONT, font_size=8, font_color="595959", text_wrap=True, valign="top")

    # ============================================================
    # Shared helper: writes the "Issue Summary" + "Instructions" text
    # block at the top of a sheet, returns the next free row.
    # ============================================================
    def write_summary_block(ws, title_text, summary_items, instructions, denom_note):
        ws.merge_range(0, 0, 0, 10, title_text, f_title)
        r = 2
        ws.write(r, 0, "Issue Summary:", f_summary_hdr)
        r += 1
        ws.write(r, 0, f"We are analysing {len(summary_items)} reports here", f_summary_body)
        r += 1
        for i, (name, desc) in enumerate(summary_items, start=1):
            ws.merge_range(r, 0, r, 10, f"{i}. {name} - {desc}", f_summary_body)
            ws.set_row(r, 28)
            r += 1
        r += 1
        ws.write(r, 0, "Instructions:", f_summary_hdr)
        r += 1
        for i, instr in enumerate(instructions, start=1):
            ws.merge_range(r, 0, r, 10, f"{i}. {instr}", f_summary_body)
            ws.set_row(r, 24)
            r += 1
        r += 1
        ws.merge_range(r, 0, r, 10, denom_note, f_note)
        ws.set_row(r, 30)
        r += 2
        return r

    # ============================================================
    # TAB G1
    # ============================================================
    ws1 = wb.add_worksheet("G1")
    ws1.set_column("A:A", 34)
    ws1.set_column("B:I", 14)

    tab1_instructions = [
        "The Address column should include only 200 status code, All page types for specific security issues - Bad content type, Missing Secure Referrer Policy, Missing X-Content-Type Options Header, Missing X-Frame-Options Header and HTML page type for rest of the issues.",
        "Export these issues from security all.",
        "Sort Table 3 By impressions",
        "Sort Table 2 basis page theme priority",
        "% share against Total URLs Crawled - considers all crawled URLs from internal_all (including CSS, PDF, images etc); if an issue requires just HTML pages, % share is against total HTML URLs crawled instead.",
        "If any page theme is not provided in the rulebook then keep it as -",
        "#Affected URLs (Table 1) - number of affected URLs per issue, from Table 3.",
        "% Share against Total URLs Crawled (Table 1) - affected URLs divided by the applicable total (see note below).",
        "Page Theme Wise URL Analysis (Table 2) - number of affected URLs per issue per Page Theme 1.",
        "Priority for Page Theme comes from the rulebook.",
    ]
    denom_note_1 = (
        f"Denominators used below: Total URLs Crawled (Status 200, all content types) = {total_all_200} | "
        f"Total HTML URLs Crawled (Status 200, text/html) = {total_html_200}. "
        f"Source: internal_all.csv for this crawl. Bad Content Type, Missing Secure Referrer-Policy, "
        f"Missing X-Content-Type-Options, and Missing X-Frame-Options use the All-page-types total; "
        f"the remaining 4 issues use the HTML-only total, per instruction #1/#6."
    )
    r = write_summary_block(ws1, "Security Issues - Tab 1 (Export-Tabs Sourced)", TAB1_SUMMARY, tab1_instructions, denom_note_1)

    # Pre-fetch + classify affected URLs per issue, honoring the
    # all-page-types vs HTML-only page_scope and the 200-status filter.
    tab1_data = []  # list of dicts per issue: {label, sev, pri, rows: [...]}
    for label, csv_name, sev, pri, scope in TAB1_ISSUES:
        df = _read_issue_csv(report_path, csv_name)
        rows = []
        if not df.empty:
            cols = {c.lower(): c for c in df.columns}
            a_col = cols.get("address")
            ct_col = cols.get("content type")
            sc_col = cols.get("status code")
            idx_col = cols.get("indexability")
            idxs_col = cols.get("indexability status")
            for _, row in df.iterrows():
                addr = str(row[a_col]) if a_col else None
                if not addr:
                    continue
                status = str(row[sc_col]) if sc_col and pd.notna(row[sc_col]) else ""
                content_type = str(row[ct_col]) if ct_col and pd.notna(row[ct_col]) else ""
                if status != "200":
                    continue
                if scope == "html" and not content_type.lower().startswith("text/html"):
                    continue
                t1, t2, _, tpri = classify(addr)
                rows.append({
                    "address": addr,
                    "theme1": t1 or "-",
                    "theme2": t2 or "-",
                    "theme_priority": tpri or "-",
                    "content_type": content_type,
                    "status": status,
                    "indexability": str(row[idx_col]) if idx_col and pd.notna(row[idx_col]) else "",
                    "indexability_status": str(row[idxs_col]) if idxs_col and pd.notna(row[idxs_col]) else "",
                    "impressions": _safe_num(gsc_imp.get(addr)),
                    "clicks": _safe_num(gsc_clk.get(addr)),
                    "sessions": _safe_num(ga_sess.get(addr)),
                })
        # Sort Table 3 by impressions descending (instruction #3), blanks last
        rows.sort(key=lambda x: (x["impressions"] is None, -(x["impressions"] or 0)))
        tab1_data.append({"label": label, "sev": sev, "pri": pri, "scope": scope, "rows": rows})

    themes_seen = {}
    for d in tab1_data:
        for row in d["rows"]:
            th, thp = row["theme1"], row["theme_priority"]
            if th != "-" and (th not in themes_seen or thp == "High"):
                themes_seen.setdefault(th, thp)
    theme_order = sorted(themes_seen.keys(), key=lambda t: (themes_seen[t] != "High", themes_seen[t] != "Medium", t))
    if not theme_order:
        theme_order = ["-"]

    # -- Table 1 --
    t1_hdr_row = r
    ws1.write(r, 0, "Table 1", f_section)
    for ci in range(1, 9):
        ws1.write(r, ci, "", f_section)
    r += 1
    ws1.write(r, 0, "Security Error", f_dark_hdr)
    for ci, d in enumerate(tab1_data):
        ws1.write(r, 1 + ci, d["label"], f_red_hdr)
    r += 1
    ws1.write(r, 0, "Issue Priority", f_dark_hdr)
    for ci, d in enumerate(tab1_data):
        ws1.write(r, 1 + ci, d["pri"], f_cell)
    r += 1

    affected_row = r
    ws1.write(r, 0, "#Affected URLs", f_dark_hdr)
    r += 1
    pct_row = r
    ws1.write(r, 0, "% share against Total  URLs Crawled", f_dark_hdr)
    r += 2

    # -- Table 2 --
    ws1.write(r, 0, "Table 2", f_section)
    for ci in range(1, 9):
        ws1.write(r, ci, "", f_section)
    r += 1
    ws1.merge_range(r, 0, r, 8, "Page Theme Wise URL Analysis", f_section)
    r += 1
    ws1.write(r, 0, "Page Theme 1", f_dark_hdr)
    for ci, d in enumerate(tab1_data):
        ws1.write(r, 1 + ci, d["label"], f_red_hdr)
    r += 1
    table2_start_row = r
    for theme in theme_order:
        ws1.write(r, 0, theme, f_cell_lft)
        r += 1
    r += 1

    # -- Table 3 (one block per issue) --
    ws1.write(r, 0, "Table 3", f_section)
    for ci in range(1, 9):
        ws1.write(r, ci, "", f_section)
    r += 1
    t3_headers = ["Address (Main issue)", "Page Theme 1", "Page Theme 2", "Content Type", "Status Code",
                  "Indexability", "Indexability Status", "Impressions", "Clicks", "Organic Sessions"]

    t3_blocks = {}  # label -> (data_start_row, data_end_row) 1-indexed for formulas
    for d in tab1_data:
        ws1.merge_range(r, 0, r, 9, d["label"], f_section)
        r += 1
        ws1.write(r, 0, "Error Type", f_grey_hdr)
        for ci, h in enumerate(t3_headers):
            ws1.write(r, 1 + ci, h, f_grey_hdr)
        r += 1
        data_start = r + 1  # 1-indexed
        for row in d["rows"]:
            ws1.write(r, 0, d["label"], f_cell_lft)
            ws1.write(r, 1, row["address"], f_cell_lft)
            ws1.write(r, 2, row["theme1"], f_cell)
            ws1.write(r, 3, row["theme2"], f_cell)
            ws1.write(r, 4, row["content_type"], f_cell)
            ws1.write(r, 5, row["status"], f_cell)
            ws1.write(r, 6, row["indexability"], f_cell)
            ws1.write(r, 7, row["indexability_status"], f_cell)
            ws1.write(r, 8, row["impressions"], f_num)
            ws1.write(r, 9, row["clicks"], f_num)
            ws1.write(r, 10, row["sessions"], f_num)
            r += 1
        data_end = r  # 1-indexed, inclusive
        if data_end < data_start:
            data_end = data_start  # empty block: formulas still reference a valid (empty) range
            r += 1
        t3_blocks[d["label"]] = (data_start, data_end)
        r += 1

    # -- Now backfill Table 1 / Table 2 formulas referencing the Table 3 blocks --
    # NOTE: Table 3 stacks each issue's rows vertically in its own block, always
    # in the same columns (Address = col B, Page Theme 1 = col C). Table 1/2 lay
    # issues out horizontally instead, so the per-issue column index used for
    # writing INTO Table 1/2 (1+ci) must never be reused as a column reference
    # INTO Table 3 - the two are unrelated axes.
    addr_col_letter = xlsxwriter.utility.xl_col_to_name(1)   # Table 3's Address column (fixed)
    theme_col_letter = xlsxwriter.utility.xl_col_to_name(2)  # Table 3's Page Theme 1 column (fixed)
    for ci, d in enumerate(tab1_data):
        start, end = t3_blocks[d["label"]]
        ws1.write_formula(affected_row, 1 + ci,
                           f'=COUNTA({addr_col_letter}{start}:{addr_col_letter}{end})', f_num)
        denom_ref = total_html_200 if d["scope"] == "html" else total_all_200
        ws1.write_formula(pct_row, 1 + ci,
                           f'=IF({denom_ref}=0,0,{xlsxwriter.utility.xl_col_to_name(1 + ci)}{affected_row+1}/{denom_ref})', f_pct)
        for ti, theme in enumerate(theme_order):
            ws1.write_formula(
                table2_start_row + ti, 1 + ci,
                f'=COUNTIFS({theme_col_letter}{start}:{theme_col_letter}{end},$A{table2_start_row + ti + 1},'
                f'{addr_col_letter}{start}:{addr_col_letter}{end},"<>")',
                f_num,
            )

    # ============================================================
    # TAB G2
    # ============================================================
    ws2 = wb.add_worksheet("G2")
    ws2.set_column("A:A", 34)
    ws2.set_column("B:M", 16)

    tab2_instructions = [
        "The Address column should include only 200 status code and an HTML page type.",
        "Export these issues from Bulk Export. Content Type and Status Code are not present in these exports and are looked up from internal_all by Source Page address.",
        "Sort Table 3 By impressions",
        "Sort Table 2 basis page theme priority",
        "Slicer for Page Theme 2 was requested in the original spec but is skipped - xlsxwriter cannot create native Excel slicers. Add manually in Excel if needed (Insert > Table > Insert Slicer) after opening this file.",
        "% share against Total URLs Crawled - considers all crawled URLs from internal_all (including CSS, PDF, images etc); if an issue requires just HTML pages, % share is against total HTML URLs crawled instead.",
        "If any page theme is not provided in the rulebook then keep it as -",
        "#Affected URLs (Table 1) - number of affected URLs per issue, from Table 3.",
        "% Share against Total URLs Crawled (Table 1) - affected URLs divided by the applicable total (see note below).",
        "Page Theme Wise URL Analysis (Table 2) - number of affected URLs per issue per Page Theme 1 (Source Page theme).",
        "Priority for Page Theme comes from the rulebook.",
    ]
    denom_note_2 = (
        f"Denominator used below: Total HTML URLs Crawled (Status 200, text/html) = {total_html_200}, "
        f"since all 4 issues on this tab are filtered to HTML pages per instruction #1. "
        f"Source: internal_all.csv for this crawl. Slicer for Page Theme 2 skipped (xlsxwriter limitation) - "
        f"add manually in Excel if needed."
    )
    r = write_summary_block(ws2, "Security Issues - Tab 2 (Bulk-Export Sourced)", TAB2_SUMMARY, tab2_instructions, denom_note_2)

    tab2_data = []
    for label, csv_name, sev, pri in TAB2_ISSUES:
        df = _read_issue_csv(report_path, csv_name)
        rows = []
        if not df.empty:
            cols = {c.lower(): c for c in df.columns}
            src_col = cols.get("source")
            dst_col = cols.get("destination") or cols.get("form action link")
            for _, row in df.iterrows():
                src = str(row[src_col]) if src_col and pd.notna(row[src_col]) else None
                if not src:
                    continue
                ct, status, idxb, idxbs = internal_lookup.get(src, ("", "", "", ""))
                if status != "200" or not ct.lower().startswith("text/html"):
                    continue
                dst = str(row[dst_col]) if dst_col and pd.notna(row[dst_col]) else ""
                t1s, t2s, _, tpri_s = classify(src)
                t1d, t2d, _, _ = classify(dst) if dst else ("-", "-", None, "-")
                rows.append({
                    "source": src,
                    "src_theme1": t1s or "-",
                    "src_theme2": t2s or "-",
                    "theme_priority": tpri_s or "-",
                    "destination": dst,
                    "dst_theme1": t1d or "-",
                    "dst_theme2": t2d or "-",
                    "src_impressions": _safe_num(gsc_imp.get(src)),
                    "src_clicks": _safe_num(gsc_clk.get(src)),
                    "src_sessions": _safe_num(ga_sess.get(src)),
                    "dst_impressions": _safe_num(gsc_imp.get(dst)),
                    "dst_clicks": _safe_num(gsc_clk.get(dst)),
                    "dst_sessions": _safe_num(ga_sess.get(dst)),
                })
        rows.sort(key=lambda x: (x["src_impressions"] is None, -(x["src_impressions"] or 0)))
        tab2_data.append({"label": label, "sev": sev, "pri": pri, "rows": rows})

    themes_seen2 = {}
    for d in tab2_data:
        for row in d["rows"]:
            th, thp = row["src_theme1"], row["theme_priority"]
            if th != "-" and (th not in themes_seen2 or thp == "High"):
                themes_seen2.setdefault(th, thp)
    theme_order2 = sorted(themes_seen2.keys(), key=lambda t: (themes_seen2[t] != "High", themes_seen2[t] != "Medium", t))
    if not theme_order2:
        theme_order2 = ["-"]

    # -- Table 1 --
    ws2.write(r, 0, "Table 1", f_section)
    for ci in range(1, 5):
        ws2.write(r, ci, "", f_section)
    r += 1
    ws2.write(r, 0, "Security Error", f_dark_hdr)
    for ci, d in enumerate(tab2_data):
        ws2.write(r, 1 + ci, d["label"], f_red_hdr)
    r += 1
    ws2.write(r, 0, "Issue Priority", f_dark_hdr)
    for ci, d in enumerate(tab2_data):
        ws2.write(r, 1 + ci, d["pri"], f_cell)
    r += 1
    affected_row2 = r
    ws2.write(r, 0, "#Affected URLs", f_dark_hdr)
    r += 1
    pct_row2 = r
    ws2.write(r, 0, "% share against Total  URLs Crawled", f_dark_hdr)
    r += 2

    # -- Table 2 --
    ws2.write(r, 0, "Table 2", f_section)
    for ci in range(1, 5):
        ws2.write(r, ci, "", f_section)
    r += 1
    ws2.merge_range(r, 0, r, 4, "Page Theme Wise URL Analysis", f_section)
    r += 1
    ws2.write(r, 0, "Page Theme 1", f_dark_hdr)
    for ci, d in enumerate(tab2_data):
        ws2.write(r, 1 + ci, d["label"], f_red_hdr)
    r += 1
    table2_start_row2 = r
    for theme in theme_order2:
        ws2.write(r, 0, theme, f_cell_lft)
        r += 1
    r += 1

    # -- Table 3 --
    ws2.write(r, 0, "Table 3", f_section)
    for ci in range(1, 5):
        ws2.write(r, ci, "", f_section)
    r += 1
    t3_headers2 = ["Source Page", "Page Theme 1", "Page Theme 2", "Destination Link", "Page Theme 1",
                   "Page Theme 2", "Impressions - Source Page", "Clicks - Source Page", "Organic Sessions - Source Page",
                   "Impressions - Destination Page", "Clicks - Destination Page", "Organic Sessions - Destination Page"]

    t3_blocks2 = {}
    for d in tab2_data:
        ws2.merge_range(r, 0, r, 12, d["label"], f_section)
        r += 1
        ws2.write(r, 0, "Error Type", f_grey_hdr)
        for ci, h in enumerate(t3_headers2):
            ws2.write(r, 1 + ci, h, f_grey_hdr)
        r += 1
        data_start = r + 1
        for row in d["rows"]:
            ws2.write(r, 0, d["label"], f_cell_lft)
            ws2.write(r, 1, row["source"], f_cell_lft)
            ws2.write(r, 2, row["src_theme1"], f_cell)
            ws2.write(r, 3, row["src_theme2"], f_cell)
            ws2.write(r, 4, row["destination"], f_cell_lft)
            ws2.write(r, 5, row["dst_theme1"], f_cell)
            ws2.write(r, 6, row["dst_theme2"], f_cell)
            ws2.write(r, 7, row["src_impressions"], f_num)
            ws2.write(r, 8, row["src_clicks"], f_num)
            ws2.write(r, 9, row["src_sessions"], f_num)
            ws2.write(r, 10, row["dst_impressions"], f_num)
            ws2.write(r, 11, row["dst_clicks"], f_num)
            ws2.write(r, 12, row["dst_sessions"], f_num)
            r += 1
        data_end = r
        if data_end < data_start:
            data_end = data_start
            r += 1
        t3_blocks2[d["label"]] = (data_start, data_end)
        r += 1

    src_col_letter = xlsxwriter.utility.xl_col_to_name(1)   # Table 3's Source Page column (fixed)
    theme_col_letter2 = xlsxwriter.utility.xl_col_to_name(2)  # Table 3's Page Theme 1 (Source) column (fixed)
    for ci, d in enumerate(tab2_data):
        start, end = t3_blocks2[d["label"]]
        ws2.write_formula(affected_row2, 1 + ci,
                           f'=COUNTA({src_col_letter}{start}:{src_col_letter}{end})', f_num)
        ws2.write_formula(pct_row2, 1 + ci,
                           f'=IF({total_html_200}=0,0,{xlsxwriter.utility.xl_col_to_name(1 + ci)}{affected_row2+1}/{total_html_200})', f_pct)
        for ti, theme in enumerate(theme_order2):
            ws2.write_formula(
                table2_start_row2 + ti, 1 + ci,
                f'=COUNTIFS({theme_col_letter2}{start}:{theme_col_letter2}{end},$A{table2_start_row2 + ti + 1},'
                f'{src_col_letter}{start}:{src_col_letter}{end},"<>")',
                f_num,
            )

    wb.close()
    buf.seek(0)
    return buf.read()
