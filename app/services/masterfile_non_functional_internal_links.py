import os
import io
import glob
import math
import pandas as pd
import xlsxwriter
from collections import defaultdict
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}

# Excel hard limit is 1,048,576 rows per sheet; cap Table 3 detail rows
# below it, leaving headroom for the summary tables above
MAX_T3_ROWS = 1_000_000

FONT = "Rockwell"
RED = "#FF0000"
WHITE = "#FFFFFF"
BLACK = "#000000"
DARK = "#1F3864"
LIGHT_GREY = "#F2F2F2"

# CSV name mappings: (csv_name_variants, sheet_name, issue_label, severity, priority)
# NOTE: load_df now matches by suffix as well, so SF export prefixes like
# "internal_" or "response_codes_internal_" are handled automatically.
ISSUE_CONFIGS = [
    {
        "csvs": ["canonicalised_inlinks.csv", "nonindexable_canonical_inlinks.csv"],
        "labels": ["Canonicalised Inlinks", "Nonindexable Canonical Inlinks"],
        "sheet": "Canonical Issue - Inlinks",
        "severity": "Issue",
        "priority": "High",
        "summary": (
            "Issue Summary: We are analysing 2 reports here\n"
            "1. canonicals_canonicalised_inlinks - Analysing internal links pointing to URLs that "
            "canonicalise to another URL, to identify non-preferred URLs being linked internally.\n"
            "2. canonicals_nonindexable_canonical_inlinks - Analysing internal links pointing to URLs "
            "whose canonical targets are non-indexable, to identify broken or invalid canonical implementations."
        ),
    },
    {
        "csvs": ["http_urls_inlinks.csv"],
        "labels": ["Security Http URLs Inlinks"],
        "sheet": "Security http URLs - Inlinks",
        "severity": "Issue",
        "priority": "Medium",
        "summary": (
            "Issue Summary: We are analysing 1 report here\n"
            "1. security_http_urls_inlinks - Analysing internal links pointing to http protocol URLs, "
            "to identify non-preferred URLs being linked internally."
        ),
    },
    {
        "csvs": ["client_error_(4xx)_inlinks.csv"],
        "labels": ["4xx Inlinks"],
        "sheet": "Internal 4xx - Inlinks",
        "severity": "Issue",
        "priority": "High",
        "summary": (
            "Issue Summary: We are analysing 1 report here\n"
            "1. response_codes_internal_client_error_(4xx)_inlinks - Analysing internal links pointing to "
            "4xx URLs, to identify non-preferred URLs being linked internally."
        ),
    },
    {
        "csvs": ["server_error_(5xx)_inlinks.csv"],
        "labels": ["5xx Inlinks"],
        "sheet": "Internal 5xx - Inlinks",
        "severity": "Issue",
        "priority": "Medium",
        "summary": (
            "Issue Summary: We are analysing 1 report here\n"
            "1. response_codes_internal_server_error_(5xx)_inlinks - Analysing internal links pointing to "
            "5xx URLs, to identify non-preferred URLs being linked internally."
        ),
    },
    {
        "csvs": ["redirection_(3xx)_inlinks.csv"],
        "labels": ["3xx Inlinks"],
        "sheet": "Internal 3xx - Inlinks",
        "severity": "Issue",
        "priority": "Medium",
        "has_redirect_chain": True,
        "summary": (
            "Issue Summary: We are analysing 2 reports here\n"
            "1. response_codes_internal_redirection_(3xx)_inlinks and redirect_chains - Analysing internal "
            "links pointing to 3xx URLs, to identify non-preferred URLs being linked internally."
        ),
    },
    {
        "csvs": ["no_response_inlinks.csv"],
        "labels": ["No Response Inlinks"],
        "sheet": "Internal No Response - Inlinks",
        "severity": "Issue",
        "priority": "High",
        "summary": (
            "Issue Summary: We are analysing 1 report here\n"
            "1. response_codes_internal_no_response_inlinks - Analysing internal links pointing to No Response "
            "URLs, to identify non-preferred URLs being linked internally."
        ),
    },
    {
        "csvs": ["blocked_by_robots_txt_inlinks.csv"],
        "labels": ["Blocked by robots.txt Inlinks"],
        "sheet": "Internal Blocked Robots.txt - I",
        "severity": "Warning",
        "priority": "Low",
        "summary": (
            "Issue Summary: We are analysing 1 report here\n"
            "1. response_codes_internal_blocked_by_robots_txt_inlinks - Analysing internal links pointing to "
            "URLs blocked by robots.txt, to identify non-preferred URLs being linked internally."
        ),
    },
    {
        "csvs": ["blocked_resource_inlinks.csv"],
        "labels": ["Blocked Resource Inlinks"],
        "sheet": "Internal Blocked Resource - Inl",
        "severity": "Warning",
        "priority": "Low",
        "summary": (
            "Issue Summary: We are analysing 1 report here\n"
            "1. response_codes_internal_blocked_resource_inlinks - Analysing internal links pointing to "
            "blocked resource URLs, to identify non-preferred URLs being linked internally."
        ),
    },
]


def safe_str(v):
    if v is None:
        return ""
    s = str(v)
    return "" if s in ("nan", "None", "NaN") else s


def safe_num(v):
    try:
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else f
    except Exception:
        return None


def _gc(df, *names):
    for n in names:
        m = next((c for c in df.columns if c.lower() == n.lower()), None)
        if m:
            return m
    return None


def load_df(report_path, filename):
    """Load a Screaming Frog export CSV.

    Matches by exact name first, then by suffix, because SF prefixes
    export filenames with the menu path, e.g.
    "client_error_(4xx)_inlinks.csv" is written to disk as
    "internal_client_error_(4xx)_inlinks.csv". Files at or under 200
    bytes are header-only exports and are treated as empty.
    """
    names = [filename] if isinstance(filename, str) else list(filename)
    candidates = []
    for name in names:
        candidates.append(os.path.join(report_path, name))
        # glob treats [] as character classes; SF names contain () only,
        # but escape defensively so future names with brackets still match
        pattern = glob.escape(name)
        candidates.extend(sorted(glob.glob(os.path.join(glob.escape(report_path), "*" + pattern))))
    seen = set()
    for p in candidates:
        if p in seen:
            continue
        seen.add(p)
        if os.path.exists(p) and os.path.getsize(p) > 200:
            try:
                return pd.read_csv(p, encoding="utf-8", low_memory=False)
            except Exception:
                pass
    return pd.DataFrame()


def build_non_functional_internal_links_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    template_path = os.path.join(settings.TEMPLATES_DIR, "Non-Functional Internal Links Analysis.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    rulebook = load_rulebook(domain)

    # Load internal_all for source URL indexability
    internal_map = {}
    total_html_crawled = 0
    df_int = load_df(report_path, "internal_all.csv")
    if not df_int.empty:
        a_col = _gc(df_int, "Address")
        idx_col = _gc(df_int, "Indexability")
        sc_col = _gc(df_int, "Status Code")
        st_col = _gc(df_int, "Status")
        ru_col = _gc(df_int, "Redirect URL")
        ct_col = _gc(df_int, "Content Type")
        # HTML-only denominator for "% Share against Total HTML URLs Crawled";
        # counting every crawled URL (images, PDFs, JS) understates the shares
        if ct_col:
            total_html_crawled = int(
                df_int[ct_col].astype(str).str.contains("html", case=False, na=False).sum()
            )
        for _, r in df_int.iterrows():
            url = safe_str(r.get(a_col, "")) if a_col else ""
            internal_map[url] = {
                "indexability": safe_str(r.get(idx_col, "")) if idx_col else "",
                "status_code": safe_str(r.get(sc_col, "")) if sc_col else "",
                "status": safe_str(r.get(st_col, "")) if st_col else "",
                "redirect_url": safe_str(r.get(ru_col, "")) if ru_col else "",
            }
    if total_html_crawled == 0:
        total_html_crawled = len(internal_map) if internal_map else 1

    # Load redirect chain data for 3xx sheet
    redirect_chain_map = {}
    df_rc = load_df(report_path, "response_codes_internal_redirect_chain.csv")
    if df_rc.empty:
        df_rc = load_df(report_path, "redirect_chains.csv")
    if not df_rc.empty:
        addr_col = _gc(df_rc, "Address")
        loop_col = next((c for c in df_rc.columns if "loop" in c.lower()), None)
        rtype_col = _gc(df_rc, "Redirect Type")
        num_redir_col = next((c for c in df_rc.columns if "number" in c.lower() or "count" in c.lower()), None)
        for _, r in df_rc.iterrows():
            url = safe_str(r.get(addr_col, "")) if addr_col else ""
            is_loop = str(r.get(loop_col, "")).strip().lower() == "true" if loop_col else False
            redirect_chain_map[url] = {
                "is_chain": not is_loop,
                "is_loop": is_loop,
                "redirect_type": safe_str(r.get(rtype_col, "")) if rtype_col else "",
                "num_redirects": safe_str(r.get(num_redir_col, "")) if num_redir_col else "",
            }

    # Redirect loops are a separate SF export; the chain CSV does not
    # reliably carry a loop column, so merge the loop report explicitly
    df_rl = load_df(report_path, "response_codes_internal_redirect_loop.csv")
    if df_rl.empty:
        df_rl = load_df(report_path, "redirect_loop.csv")
    if not df_rl.empty:
        addr_col = _gc(df_rl, "Address")
        num_redir_col = next((c for c in df_rl.columns if "number" in c.lower() or "count" in c.lower()), None)
        if addr_col:
            for _, r in df_rl.iterrows():
                url = safe_str(r.get(addr_col, ""))
                if not url:
                    continue
                entry = redirect_chain_map.get(url, {
                    "is_chain": False,
                    "redirect_type": "",
                    "num_redirects": safe_str(r.get(num_redir_col, "")) if num_redir_col else "",
                })
                entry["is_loop"] = True
                entry["is_chain"] = False
                redirect_chain_map[url] = entry

    # Load GSC and GA4
    gsc_map = {}
    df_gsc = load_df(report_path, "search_console_all.csv")
    if not df_gsc.empty:
        a = _gc(df_gsc, "Address")
        imp = next((c for c in df_gsc.columns if "impression" in c.lower()), None)
        clk = next((c for c in df_gsc.columns if "click" in c.lower()), None)
        if a:
            for _, r in df_gsc.iterrows():
                gsc_map[safe_str(r[a])] = {
                    "impressions": safe_num(r.get(imp, 0)) if imp else None,
                    "clicks": safe_num(r.get(clk, 0)) if clk else None,
                }

    ga_map = {}
    df_ga = load_df(report_path, "analytics_all.csv")
    if not df_ga.empty:
        a = _gc(df_ga, "Address")
        s = next((c for c in df_ga.columns if "session" in c.lower()), None)
        if a and s:
            for _, r in df_ga.iterrows():
                ga_map[safe_str(r[a])] = safe_num(r.get(s, 0))

    # Build workbook
    buf = io.BytesIO()
    # strings_to_urls disabled: Excel corrupts any worksheet holding more
    # than 65,530 hyperlinks, and Table 3 URL columns exceed that easily
    wb = xlsxwriter.Workbook(buf, {
        "in_memory": True,
        "nan_inf_to_errors": True,
        "strings_to_urls": False,
    })

    def f(**kw):
        return wb.add_format(kw)

    f_title = f(bold=True, font_name=FONT, font_size=14, font_color=BLACK)
    f_red_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
    f_red_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="left", valign="vcenter", text_wrap=True)
    f_dark_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                   bg_color=DARK, border=1, align="center", valign="vcenter", text_wrap=True)
    f_dark_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                   bg_color=DARK, border=1, align="left", valign="vcenter", text_wrap=True)
    f_lbl = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK)
    f_cell = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_cell_lft = f(font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_bold_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_bold_ctr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="#,##0")
    f_pct = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="0.00%")
    f_summary = f(font_name=FONT, font_size=8, font_color=BLACK,
                  text_wrap=True, valign="top", border=1)
    f_col_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                  bg_color=WHITE, border=1, align="center", valign="vcenter", text_wrap=True)
    f_col_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                  bg_color=WHITE, border=1, align="left", valign="vcenter", text_wrap=True)
    f_rgt = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="right", valign="vcenter")

    # Collect all issue data first for dashboard
    all_issue_data = {}

    # Process each issue sheet
    for cfg in ISSUE_CONFIGS:
        sheet_name = cfg["sheet"]
        labels = cfg["labels"]
        has_redirect = cfg.get("has_redirect_chain", False)

        # Load CSVs for this issue
        issue_dfs = []
        for csv_name in cfg["csvs"]:
            df = load_df(report_path, csv_name)
            if not df.empty:
                # Add error type label
                label_idx = cfg["csvs"].index(csv_name)
                lbl = labels[label_idx] if label_idx < len(labels) else labels[0]
                df["_error_type"] = lbl
                issue_dfs.append(df)

        if not issue_dfs:
            # Create empty sheet
            ws = wb.add_worksheet(sheet_name)
            ws.write(0, 0, f"No data found for {sheet_name}", f_lbl)
            all_issue_data[sheet_name] = {"rows": [], "themes": {}}
            continue

        df_all = pd.concat(issue_dfs, ignore_index=True)

        # Filter: exclude self-links and non-indexable sources
        src_col = _gc(df_all, "Source")
        dst_col = _gc(df_all, "Destination")
        lp_col = _gc(df_all, "Link Position")
        type_col = _gc(df_all, "Type")
        alt_col = _gc(df_all, "Alt Text")
        anc_col = _gc(df_all, "Anchor")
        sc_col = _gc(df_all, "Status Code")
        st_col = _gc(df_all, "Status")
        fol_col = _gc(df_all, "Follow")
        lo_col = _gc(df_all, "Link Origin")

        # Cache classify_url per unique URL; on large crawls the same
        # source and destination URLs repeat thousands of times
        _classify_cache = {}

        def classify_cached(url):
            hit = _classify_cache.get(url)
            if hit is None:
                hit = classify_url(url, rulebook)
                _classify_cache[url] = hit
            return hit

        # Build processed rows
        rows = []
        for _, row in df_all.iterrows():
            src = safe_str(row.get(src_col, "")) if src_col else ""
            dst = safe_str(row.get(dst_col, "")) if dst_col else ""
            if not src or not dst:
                continue
            if src == dst:
                continue
            src_info = internal_map.get(src, {})
            if src_info.get("indexability", "").lower() != "indexable":
                continue

            src_theme1, src_theme2, _, _ = classify_cached(src)
            dst_theme1, dst_theme2, _, _ = classify_cached(dst)
            gsc_src = gsc_map.get(src, {})
            gsc_dst = gsc_map.get(dst, {})

            row_data = {
                "error_type": safe_str(row.get("_error_type", "")),
                "source": src,
                "src_theme1": src_theme1 or "-",
                "src_theme2": src_theme2 or "-",
                "src_indexability": src_info.get("indexability", ""),
                "src_status_code": src_info.get("status_code", ""),
                "src_status": src_info.get("status", ""),
                "destination": dst,
                "dst_theme1": dst_theme1 or "-",
                "dst_theme2": dst_theme2 or "-",
                "alt_text": safe_str(row.get(alt_col, "")) if alt_col else "",
                "anchor": safe_str(row.get(anc_col, "")) if anc_col else "",
                "dst_status_code": safe_str(row.get(sc_col, "")) if sc_col else "",
                "dst_status": safe_str(row.get(st_col, "")) if st_col else "",
                "follow": safe_str(row.get(fol_col, "")) if fol_col else "",
                "link_type": safe_str(row.get(type_col, "")) if type_col else "",
                "link_position": safe_str(row.get(lp_col, "")) if lp_col else "",
                "link_origin": safe_str(row.get(lo_col, "")) if lo_col else "",
                "imp_src": gsc_src.get("impressions"),
                "clk_src": gsc_src.get("clicks"),
                "sess_src": ga_map.get(src),
                "imp_dst": gsc_dst.get("impressions"),
                "clk_dst": gsc_dst.get("clicks"),
                "sess_dst": ga_map.get(dst),
            }

            if has_redirect:
                rc = redirect_chain_map.get(dst, {})
                row_data["is_chain"] = "Yes" if rc.get("is_chain") else "No"
                row_data["is_loop"] = "Yes" if rc.get("is_loop") else "No"
                row_data["num_redirects"] = rc.get("num_redirects", "")
                row_data["temp_redirect"] = rc.get("redirect_type", "")
                row_data["final_redirect_url"] = internal_map.get(dst, {}).get("redirect_url", "")

            rows.append(row_data)

        # Sort Table 3 by Source Impressions descending
        rows.sort(key=lambda x: (x.get("imp_src") or 0), reverse=True)

        # Compute inlink counts per destination (for col Y formula base value)
        inlink_counts = defaultdict(set)
        for r in rows:
            inlink_counts[r["destination"]].add(r["source"])
        inlink_count_map = {dst: len(srcs) for dst, srcs in inlink_counts.items()}

        # Build theme summary data
        # Unique dest URLs per theme per label
        theme_unique = defaultdict(lambda: defaultdict(set))
        theme_total = defaultdict(lambda: defaultdict(int))
        theme_priority_map = {}
        for r in rows:
            theme = r["src_theme1"] if r["src_theme1"] != "-" else "Others"
            lbl = r["error_type"]
            theme_unique[theme][lbl].add(r["destination"])
            theme_total[theme][lbl] += 1
            if theme not in theme_priority_map:
                _, _, _, pri = classify_cached(r["source"])
                theme_priority_map[theme] = pri

        sorted_themes = sorted(theme_priority_map.items(),
                               key=lambda x: PRIORITY_ORDER.get(x[1], 3))

        # Table 1 totals
        t1_unique = defaultdict(set)
        t1_total = defaultdict(int)
        t1_pos_content = defaultdict(set)
        t1_pos_footer = defaultdict(set)
        t1_pos_header = defaultdict(set)
        t1_type_http = defaultdict(set)
        for r in rows:
            lbl = r["error_type"]
            t1_unique[lbl].add(r["destination"])
            t1_total[lbl] += 1
            pos = r["link_position"].lower()
            lt = r["link_type"]
            if pos == "content":
                t1_pos_content[lbl].add(r["destination"])
            elif pos == "footer":
                t1_pos_footer[lbl].add(r["destination"])
            elif pos == "header":
                t1_pos_header[lbl].add(r["destination"])
            if "http" in lt.lower() and "redirect" in lt.lower():
                t1_type_http[lbl].add(r["destination"])

        all_issue_data[sheet_name] = {
            "rows": rows,
            "themes": sorted_themes,
            "theme_unique": dict(theme_unique),
            "theme_total": dict(theme_total),
            "t1_unique": dict(t1_unique),
            "t1_total": dict(t1_total),
            "inlink_count_map": inlink_count_map,
        }

        # Build the worksheet
        ws = wb.add_worksheet(sheet_name)
        ws.set_column("A:A", 50)
        ws.set_column("B:D", 18)
        ws.set_column("E:G", 15)
        ws.set_column("H:H", 50)
        ws.set_column("I:K", 18)
        ws.set_column("L:L", 20)
        ws.set_column("M:O", 15)
        ws.set_column("P:Q", 15)
        ws.set_column("R:X", 15)
        ws.set_column("Y:Z", 30)

        # Issue Summary box rows 0-3
        ws.merge_range(0, 0, 3, 4, cfg["summary"], f_summary)
        ws.set_row(0, 50)

        # Row 4: blank
        # Row 5: Table 1 label | Table 4 label
        T1_START_COL = 0
        T4_START_COL = 6 if len(labels) == 1 else 6

        ws.write(5, T1_START_COL, "Table 1", f_lbl)
        ws.write(5, T4_START_COL, "Table 4", f_lbl)

        # Row 6: Table 1 headers
        ws.set_row(6, 42)
        ws.write(6, 0, "URL Issue Types", f_col_lft)
        for i, lbl in enumerate(labels):
            ws.write(6, i + 1, lbl, f_col_hdr)
        # Table 4 headers
        t4_cols = ["URL Issue Types", "Inlinks Link Position - Content",
                   "Inlinks Link Position - Footer", "Inlinks Link Position - Header",
                   "Inlinks Link Type - HTTP Redirect"]
        for i, h in enumerate(t4_cols):
            ws.write(6, T4_START_COL + i, h, f_dark_hdr if i > 0 else f_col_lft)

        # Row 7: Issue Priority
        ws.write(7, 0, "Issue Priority", f_bold_lft)
        for i, lbl in enumerate(labels):
            ws.write(7, i + 1, cfg["priority"], f_cell)
        ws.write(7, T4_START_COL, "Issue Priority", f_bold_lft)
        for i in range(1, 5):
            ws.write(7, T4_START_COL + i, "High", f_cell)

        # Row 8: #Affected Unique Destination URLs
        ws.write(8, 0, "#Affected Unique Destination URLs", f_bold_lft)
        for i, lbl in enumerate(labels):
            count = len(t1_unique.get(lbl, set()))
            ws.write(8, i + 1, count, f_num)
        ws.write(8, T4_START_COL, "#Affected Unique Destination URLs", f_bold_lft)
        # Table 4 unique counts by position/type
        for j, pos_set in enumerate([t1_pos_content, t1_pos_footer, t1_pos_header, t1_type_http]):
            total = sum(len(s) for s in pos_set.values())
            ws.write(8, T4_START_COL + j + 1, total, f_num)

        # Row 9: % Share
        ws.write(9, 0, "% Share against Total HTML URLs Crawled", f_bold_lft)
        for i, lbl in enumerate(labels):
            count = len(t1_unique.get(lbl, set()))
            pct = count / total_html_crawled if total_html_crawled > 0 else 0
            ws.write(9, i + 1, pct, f_pct)
        ws.write(9, T4_START_COL, "% Share against Total HTML URLs Crawled", f_bold_lft)
        for j, pos_set in enumerate([t1_pos_content, t1_pos_footer, t1_pos_header, t1_type_http]):
            total = sum(len(s) for s in pos_set.values())
            pct = total / total_html_crawled if total_html_crawled > 0 else 0
            ws.write(9, T4_START_COL + j + 1, pct, f_pct)

        # Row 11: Table 2 label
        ws.write(11, 0, "Table 2", f_lbl)

        # Row 12: Table 2 title
        ws.merge_range(12, 0, 12, len(labels) * 2, "Page Theme Wise Internal Links Analysis", f_red_lft)

        # Row 13: Table 2 headers
        ws.set_row(13, 42)
        ws.write(13, 0, "Page Theme 1", f_red_lft)
        ws.write(13, 1, "Priority Basis Page Theme 1", f_red_hdr)
        col = 2
        for lbl in labels:
            ws.write(13, col, f"{lbl} - Unique Pages", f_col_hdr)
            ws.write(13, col + 1, lbl, f_col_hdr)
            col += 2

        # Table 2 data rows
        T2_DATA_START = 14
        for row_offset, (theme, priority) in enumerate(sorted_themes):
            r = T2_DATA_START + row_offset
            ws.write(r, 0, theme, f_bold_lft)
            ws.write(r, 1, priority, f_cell)
            col = 2
            for lbl in labels:
                unique_count = len(theme_unique.get(theme, {}).get(lbl, set()))
                total_count = theme_total.get(theme, {}).get(lbl, 0)
                ws.write(r, col, unique_count, f_num)
                ws.write(r, col + 1, total_count, f_num)
                col += 2

        # Table 3 label
        T3_LABEL_ROW = T2_DATA_START + len(sorted_themes) + 1
        ws.write(T3_LABEL_ROW, 0, "Table 3", f_lbl)

        # Table 3 headers
        T3_HDR_ROW = T3_LABEL_ROW + 1
        T3_DATA_START = T3_HDR_ROW + 1

        ws.set_row(T3_HDR_ROW, 31.5)
        t3_headers = [
            "Error Type", "Source", "Page Theme 1", "Page Theme 2",
            "Source - Indexability", "Source - Status Code", "Source - Status",
            "Destination", "Page Theme 1", "Page Theme 2",
            "Alt Text", "Anchor", "Destination - Status Code", "Destination - Status",
            "Follow", "Type", "Link Position", "Link Origin",
            "Impressions - Source", "Clicks - Source", "Organic Sessions - Source",
            "Impressions - Destination", "Clicks - Destination", "Organic Sessions - Destination",
            "Number of Inlinks to the Destination Page",
        ]
        if has_redirect:
            t3_headers += [
                "Is Destination URL Impacted by redirection Chain?",
                "Is Destination URL Impacted by redirection Loop?",
                "Number of Redirects", "Temp Redirect in Chain", "Final Redirect URL"
            ]

        for i, h in enumerate(t3_headers):
            ws.write(T3_HDR_ROW, i, h, f_red_lft if i == 0 else f_red_hdr)

        # Table 3 data (capped at MAX_T3_ROWS; rows are sorted by source
        # impressions descending, so truncation drops the least important)
        t3_rows = rows[:MAX_T3_ROWS]
        truncated = len(rows) - len(t3_rows)
        for row_offset, row_data in enumerate(t3_rows):
            r = T3_DATA_START + row_offset
            ws.set_row(r, 14.5)
            ws.write(r, 0, row_data["error_type"], f_cell_lft)
            ws.write(r, 1, row_data["source"], f_cell_lft)
            ws.write(r, 2, row_data["src_theme1"], f_cell)
            ws.write(r, 3, row_data["src_theme2"], f_cell)
            ws.write(r, 4, row_data["src_indexability"], f_cell)
            ws.write(r, 5, row_data["src_status_code"], f_cell)
            ws.write(r, 6, row_data["src_status"], f_cell)
            ws.write(r, 7, row_data["destination"], f_cell_lft)
            ws.write(r, 8, row_data["dst_theme1"], f_cell)
            ws.write(r, 9, row_data["dst_theme2"], f_cell)
            ws.write(r, 10, row_data["alt_text"], f_cell_lft)
            ws.write(r, 11, row_data["anchor"], f_cell_lft)
            ws.write(r, 12, row_data["dst_status_code"], f_cell)
            ws.write(r, 13, row_data["dst_status"], f_cell)
            ws.write(r, 14, row_data["follow"], f_cell)
            ws.write(r, 15, row_data["link_type"], f_cell)
            ws.write(r, 16, row_data["link_position"], f_cell)
            ws.write(r, 17, row_data["link_origin"], f_cell)
            ws.write(r, 18, safe_num(row_data["imp_src"]), f_rgt)
            ws.write(r, 19, safe_num(row_data["clk_src"]), f_rgt)
            ws.write(r, 20, safe_num(row_data["sess_src"]), f_rgt)
            ws.write(r, 21, safe_num(row_data["imp_dst"]), f_rgt)
            ws.write(r, 22, safe_num(row_data["clk_dst"]), f_rgt)
            ws.write(r, 23, safe_num(row_data["sess_dst"]), f_rgt)
            inlink_count = inlink_count_map.get(row_data["destination"], 0)
            ws.write(r, 24, inlink_count, f_num)
            if has_redirect:
                ws.write(r, 25, row_data.get("is_chain", "No"), f_cell)
                ws.write(r, 26, row_data.get("is_loop", "No"), f_cell)
                ws.write(r, 27, row_data.get("num_redirects", ""), f_cell)
                ws.write(r, 28, row_data.get("temp_redirect", ""), f_cell)
                ws.write(r, 29, row_data.get("final_redirect_url", ""), f_cell_lft)

        if truncated > 0:
            note_row = T3_DATA_START + len(t3_rows)
            ws.write(note_row, 0,
                     f"Note: {truncated:,} additional rows omitted due to the Excel row limit; "
                     "full detail available in the source CSV exports", f_lbl)

    # BUILD DASHBOARD SHEET
    ws_dash = wb.add_worksheet("Dashboard")
    ws_dash.set_column("A:A", 25)
    for col in range(1, 11):
        ws_dash.set_column(col, col, 20)

    # Dashboard headers
    issue_labels_dash = [
        "Canonicalised Inlinks", "Nonindexable Canonical Inlinks",
        "http inlinks", "4xx Inlinks", "5xx Inlinks",
        "3xx Inlinks", "No Response Inlinks",
        "Blocked by robots.txt Inlinks", "Blocked Resource Inlinks"
    ]
    # Priorities mirror ISSUE_CONFIGS: http, 5xx and 3xx are Medium per template
    priorities_dash = ["High", "High", "Medium", "High", "Medium", "Medium", "High", "Low", "Low"]

    ws_dash.write(0, 0, "Non-functional Internal Links Summary", f_title)
    ws_dash.write(1, 0, "Add slicer for Link Type and Link Position for Table 1, Table 2 and Table 3", f_lbl)

    # Table 1 header
    ws_dash.write(3, 0, "Table 1", f_lbl)
    ws_dash.set_row(4, 42)
    ws_dash.write(4, 0, "", f_red_lft)
    for i, lbl in enumerate(issue_labels_dash):
        ws_dash.write(4, i + 1, lbl, f_red_hdr)

    ws_dash.write(5, 0, "Issue Priority", f_bold_lft)
    for i, pri in enumerate(priorities_dash):
        ws_dash.write(5, i + 1, pri, f_cell)

    # Unique Destination URLs row
    ws_dash.write(6, 0, "Unique Destination URLs", f_bold_lft)
    # Total Internal Links row
    ws_dash.write(7, 0, "Total Internal Links", f_bold_lft)

    # Populate from collected data
    dash_col_map = {
        "Canonical Issue - Inlinks": {"Canonicalised Inlinks": 1, "Nonindexable Canonical Inlinks": 2},
        "Security http URLs - Inlinks": {"Security Http URLs Inlinks": 3},
        "Internal 4xx - Inlinks": {"4xx Inlinks": 4},
        "Internal 5xx - Inlinks": {"5xx Inlinks": 5},
        "Internal 3xx - Inlinks": {"3xx Inlinks": 6},
        "Internal No Response - Inlinks": {"No Response Inlinks": 7},
        "Internal Blocked Robots.txt - I": {"Blocked by robots.txt Inlinks": 8},
        "Internal Blocked Resource - Inl": {"Blocked Resource Inlinks": 9},
    }

    for sheet_name, label_col_map in dash_col_map.items():
        data = all_issue_data.get(sheet_name, {})
        t1_unique = data.get("t1_unique", {})
        t1_total = data.get("t1_total", {})
        for lbl, col in label_col_map.items():
            unique_count = len(t1_unique.get(lbl, set()))
            total_count = t1_total.get(lbl, 0)
            ws_dash.write(6, col, unique_count, f_num)
            ws_dash.write(7, col, total_count, f_num)

    # Table 2 - Unique Destination URLs
    ws_dash.write(9, 0, "Table 2 - Unique Destination URLs", f_lbl)
    ws_dash.set_row(10, 42)
    ws_dash.write(10, 0, "Page Theme", f_red_lft)
    for i, lbl in enumerate(issue_labels_dash):
        ws_dash.write(10, i + 1, lbl, f_red_hdr)

    # Collect all themes
    all_themes = set()
    for data in all_issue_data.values():
        for theme, _ in data.get("themes", []):
            all_themes.add(theme)
    all_themes_sorted = sorted(all_themes)

    T2_DASH_START = 11
    for row_offset, theme in enumerate(all_themes_sorted):
        r = T2_DASH_START + row_offset
        ws_dash.write(r, 0, theme, f_bold_lft)
        for sheet_name, label_col_map in dash_col_map.items():
            data = all_issue_data.get(sheet_name, {})
            theme_unique = data.get("theme_unique", {})
            for lbl, col in label_col_map.items():
                count = len(theme_unique.get(theme, {}).get(lbl, set()))
                ws_dash.write(r, col, count, f_num)

    # Table 3 - Total Internal Links
    T3_DASH_START = T2_DASH_START + len(all_themes_sorted) + 2
    ws_dash.write(T3_DASH_START - 1, 0, "Table 3 - Total Internal Links", f_lbl)
    ws_dash.set_row(T3_DASH_START, 42)
    ws_dash.write(T3_DASH_START, 0, "Page Theme", f_red_lft)
    for i, lbl in enumerate(issue_labels_dash):
        ws_dash.write(T3_DASH_START, i + 1, lbl, f_red_hdr)

    for row_offset, theme in enumerate(all_themes_sorted):
        r = T3_DASH_START + 1 + row_offset
        ws_dash.write(r, 0, theme, f_bold_lft)
        for sheet_name, label_col_map in dash_col_map.items():
            data = all_issue_data.get(sheet_name, {})
            theme_total = data.get("theme_total", {})
            for lbl, col in label_col_map.items():
                count = theme_total.get(theme, {}).get(lbl, 0)
                ws_dash.write(r, col, count, f_num)

    # Move Dashboard to first position
    wb.worksheets_objs = [wb.worksheets_objs[-1]] + wb.worksheets_objs[:-1]
    wb.close()
    buf.seek(0)
    return buf.read()
