import os
import io
import math
import pandas as pd
import xlsxwriter
from collections import defaultdict
from app.config import settings
from app.services.rulebook import load_rulebook, classify_url

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2, "N/A": 3}
FONT = "Rockwell"
RED = "#FF0000"
WHITE = "#FFFFFF"
BLACK = "#000000"
DARK = "#1F3864"

# Issue definitions: (category, issue_label, severity, priority, csv_files)
# csv_files: list of CSV filenames to check for URL presence
ISSUES = [
    # Response Codes - Internal
    ("Response Codes - Internal", "3xx (Redirection)", "Warning", "High", ["response_codes_internal_redirection_(3xx).csv"]),
    ("Response Codes - Internal", "Internal Client Error (4xx)", "Issue", "High", ["response_codes_internal_client_error_(4xx).csv"]),
    ("Response Codes - Internal", "Internal Server Error (5xx)", "Issue", "High", ["response_codes_internal_server_error_(5xx).csv"]),
    ("Response Codes - Internal", "Internal No Response", "Issue", "High", ["response_codes_internal_no_response.csv"]),
    ("Response Codes - Internal", "Internal Redirect Chain", "Issue", "High", ["response_codes_internal_redirect_chain.csv"]),
    ("Response Codes - Internal", "Internal Redirect Loop", "Issue", "High", ["response_codes_internal_redirect_loop.csv"]),
    ("Response Codes - Internal", "Internal Blocked by Robots.txt", "Warning", "High", ["response_codes_internal_blocked_by_robots_txt.csv"]),
    # URL Issues
    ("URL Issues", "Uppercase", "Warning", "Low", ["url_uppercase.csv"]),
    ("URL Issues", "Underscores", "Opportunity", "Low", ["url_underscores.csv"]),
    ("URL Issues", "Parameters", "Warning", "Low", ["url_parameters.csv"]),
    ("URL Issues", "Multiple Slashes", "Issue", "Low", ["url_multiple_slashes.csv"]),
    ("URL Issues", "Repetitive Path", "Warning", "Low", ["url_repetitive_path.csv"]),
    ("URL Issues", "Contains Space", "Issue", "Medium", ["url_contains_space.csv"]),
    # Page Titles
    ("Page Titles", "Missing", "Issue", "High", ["page_titles_missing.csv"]),
    ("Page Titles", "Duplicate", "Opportunity", "Low", ["page_titles_duplicate.csv"]),
    ("Page Titles", "Over 561 Pixels", "Opportunity", "Low", ["page_titles_over_561_pixels.csv"]),
    ("Page Titles", "Below 200 Pixels", "Opportunity", "Low", ["page_titles_below_200_pixels.csv"]),
    ("Page Titles", "Same as H1", "Opportunity", "Low", ["page_titles_same_as_h1.csv"]),
    ("Page Titles", "Multiple", "Issue", "High", ["page_titles_multiple.csv"]),
    ("Page Titles", "Outside <Head>", "Issue", "High", ["page_titles_outside_head.csv"]),
    # Meta Description
    ("Meta Description", "Missing", "Opportunity", "Medium", ["meta_description_missing.csv"]),
    ("Meta Description", "Duplicate", "Opportunity", "Low", ["meta_description_duplicate.csv"]),
    ("Meta Description", "Over 985 Pixels", "Opportunity", "Low", ["meta_description_over_985_pixels.csv"]),
    ("Meta Description", "Below 400 Pixels", "Opportunity", "Low", ["meta_description_below_400_pixels.csv"]),
    ("Meta Description", "Multiple", "Issue", "Medium", ["meta_description_multiple.csv"]),
    ("Meta Description", "Outside <Head>", "Issue", "Medium", ["meta_description_outside_head.csv"]),
    # H1
    ("H1", "Missing", "Issue", "High", ["h1_missing.csv"]),
    ("H1", "Duplicate", "Opportunity", "Low", ["h1_duplicate.csv"]),
    ("H1", "Over 70 Characters", "Opportunity", "Low", ["h1_over_70_characters.csv"]),
    ("H1", "Multiple", "Warning", "Low", ["h1_multiple.csv"]),
    # Canonicals
    ("Canonicals", "Canonicalised", "Issue", "High", ["canonicals_canonicalised.csv"]),
    ("Canonicals", "Missing", "Issue", "High", ["canonicals_missing.csv"]),
    ("Canonicals", "Multiple", "Issue", "High", ["canonicals_multiple.csv"]),
    ("Canonicals", "Non-Indexable Canonical", "Issue", "High", ["canonicals_nonindexable_canonical.csv"]),
    ("Canonicals", "Multiple Conflicting", "Issue", "High", ["canonicals_multiple_conflicting.csv"]),
    ("Canonicals", "Canonical Is Relative", "Warning", "High", ["canonicals_canonical_is_relative.csv"]),
    ("Canonicals", "Unlinked", "Warning", "Medium", ["canonicals_unlinked.csv"]),
    ("Canonicals", "Invalid Attribute In Annotation", "Issue", "High", []),
    ("Canonicals", "Contains Fragment URL", "Issue", "High", []),
    ("Canonicals", "Outside <head>", "Issue", "High", ["canonicals_outside_head.csv"]),
    # Directives
    ("Directives", "Noindex", "Warning", "High", ["directives_noindex.csv"]),
    ("Directives", "Nofollow", "Warning", "Medium", ["directives_nofollow.csv"]),
    # Sitemaps
    ("Sitemaps", "URLs Not in Sitemap", "Issue", "Medium", ["sitemaps_urls_not_in_sitemap.csv"]),
    ("Sitemaps", "Orphan URLs", "Issue", "Medium", ["sitemaps_orphan_urls.csv"]),
    ("Sitemaps", "Non-Indexable URLs in Sitemap", "Issue", "Medium", ["sitemaps_nonindexable_urls_in_sitemap.csv"]),
    ("Sitemaps", "XML Sitemap with over 50k URLs", "Issue", "High", ["sitemaps_xml_sitemap_with_over_50k_urls.csv"]),
    # Security
    ("Security", "HTTP URLs", "Issue", "High", ["security_http_urls.csv"]),
    ("Security", "Mixed Content", "Issue", "High", ["security_mixed_content.csv"]),
    ("Security", "Form URL Insecure", "", "", []),
    ("Security", "Form On HTTP URL", "", "", []),
    ("Security", "Missing HSTS Header", "", "", []),
    ("Security", "Unsafe Cross Origin Links", "", "", []),
    ("Security", "Protocol-Relative Resource Links", "", "", []),
    ("Security", "Missing Content-Security-Policy Header", "", "", []),
    ("Security", "Missing X-Content-Type-Options Header", "", "", []),
    ("Security", "Missing X-Frames-Options Header", "", "", []),
    ("Security", "Missing Secure Referrer-Policy Header", "", "", []),
    ("Security", "Bad Content Type", "", "", []),
    # Page Speed/CWV
    ("Page Speed/CWV", "Largest Contentful Paint (LCP)", "Issue", "High", []),
    ("Page Speed/CWV", "Interaction to Next Paint (INP)", "Issue", "High", []),
    ("Page Speed/CWV", "Cumulative Layout Shift (CLS)", "Issue", "High", []),
    ("Page Speed/CWV", "First Contentful Paint (FCP)", "Issue", "High", []),
    ("Page Speed/CWV", "Time to First Byte (TTFB)", "Issue", "High", []),
    # Structured Tags
    ("Strucured Tags", "Missing Structured Data", "Opportunity", "High", ["structured_data_missing.csv"]),
    ("Strucured Tags", "Validation Errors", "Issue", "High", ["structured_data_validation_errors.csv"]),
    ("Strucured Tags", "Validation Warnings", "Warning", "Low", ["structured_data_validation_warnings.csv"]),
    ("Strucured Tags", "Parse Errors", "Issue", "High", ["structured_data_parse_errors.csv"]),
    ("Strucured Tags", "Rich Result Validation Errors", "Issue", "High", []),
    ("Strucured Tags", "Rich Result Validation Warnings", "Warning", "Low", []),
    # Internal Links
    ("Internal Links", "Canonical Issue - Inlinks", "Issue", "High", ["canonicalised_inlinks.csv", "nonindexable_canonical_inlinks.csv"]),
    ("Internal Links", "Security http URLs- Inlinks", "Issue", "High", ["http_urls_inlinks.csv"]),
    ("Internal Links", "Internal 4xx - Inlinks", "Issue", "High", ["client_error_(4xx)_inlinks.csv"]),
    ("Internal Links", "Internal 5xx - Inlinks", "Issue", "High", ["server_error_(5xx)_inlinks.csv"]),
    ("Internal Links", "Internal 3xx - Inlinks", "Issue", "High", ["redirection_(3xx)_inlinks.csv"]),
    ("Internal Links", "Internal No response URLs - Inlinks", "Issue", "High", ["no_response_inlinks.csv"]),
    ("Internal Links", "Internal Blocked by Robots.txt - Inlinks", "Warning", "Low", ["blocked_by_robots_txt_inlinks.csv"]),
    ("Internal Links", "Internal Blocked Resource - Inlinks", "Warning", "Low", ["blocked_resource_inlinks.csv"]),
    ("Internal Links", "Functional Internal Links Analysis", "Opportunity", "High", []),
    # Content Issues
    ("Content Issues", "Low Content Pages", "Warning", "High", ["content_low_content_pages.csv"]),
    ("Content Issues", "Soft 404 Pages", "Warning", "High", ["content_soft_404_pages.csv"]),
    ("Content Issues", "Spelling Errors", "Warning", "Medium", ["content_spelling_errors.csv"]),
    ("Content Issues", "Grammar Errors", "Warning", "Medium", ["content_grammar_errors.csv"]),
    ("Content Issues", "Redability Difficult", "Warning", "Medium", []),
    ("Content Issues", "Redability Very Difficult", "Warning", "Medium", []),
    ("Content Issues", "Lorem Ipsum Placeholder", "Opportunity", "Medium", ["content_lorem_ipsum_placeholder.csv"]),
    ("Content Issues", "Near Duplicates", "", "Medium", ["content_near_duplicates.csv"]),
    ("Content Issues", "Exact Duplicates", "", "High", ["content_exact_duplicates.csv"]),
    # Custom Search
    ("Custom Search ", "GA4 Tags Implementation", "Issue", "High", []),
    ("Custom Search ", "GTM Tags Implementation", "Issue", "High", []),
    ("Custom Search ", "OG Tags", "Issue", "Low", []),
    ("Custom Search ", "Twitter Card", "Issue", "Low", []),
    # Pagination
    ("Pagination", "Pagination URL Not in Anchor Tag", "Issue", "Low", ["pagination_pagination_url_not_in_anchor_tag.csv"]),
    ("Pagination", "Unlinked Pagination URLs", "Issue", "High", ["pagination_unlinked_pagination_urls.csv"]),
    ("Pagination", "Non-Indexable", "Issue", "High", ["pagination_nonindexable.csv"]),
    ("Pagination", "Multiple Pagination URLs", "Issue", "Low", ["pagination_multiple_pagination_urls.csv"]),
    ("Pagination", "Pagination Loop", "Issue", "Low", ["pagination_pagination_loop.csv"]),
    ("Pagination", "Sequence Error", "Issue", "High", ["pagination_sequence_error.csv"]),
    # Hreflang Tags
    ("Hreflang Tags", "Non-200 hreflang URLs", "Issue", "Medium", ["hreflang_non200_hreflang_urls.csv"]),
    ("Hreflang Tags", "Unlinked hreflang URLs", "Issue", "High", ["hreflang_unlinked_hreflang_urls.csv"]),
    ("Hreflang Tags", "Missing Return Links", "Issue", "High", ["hreflang_missing_return_links.csv"]),
    ("Hreflang Tags", "Non-Canonical Return Links", "Issue", "High", ["hreflang_noncanonical_return_links.csv"]),
    ("Hreflang Tags", "Noindex Return Links", "Issue", "High", ["hreflang_noindex_return_links.csv"]),
    ("Hreflang Tags", "Incorrect Language & Region Codes", "Issue", "High", ["hreflang_incorrect_language_region_codes.csv"]),
    ("Hreflang Tags", "Multiple Entries", "Issue", "High", ["hreflang_multiple_entries.csv"]),
    ("Hreflang Tags", "Missing Self Reference", "Issue", "Medium", ["hreflang_missing_self_reference.csv"]),
    ("Hreflang Tags", "Inconsistent Language & Region Return Links", "Issue", "High", ["hreflang_inconsistent_language_region_return_links.csv"]),
    ("Hreflang Tags", "Not Using Canonical", "Issue", "High", ["hreflang_not_using_canonical.csv"]),
    ("Hreflang Tags", "Missing X-Default", "Warning", "High", ["hreflang_missing_xdefault.csv"]),
    ("Hreflang Tags", "Missing", "Issue", "High", ["hreflang_missing.csv"]),
    ("Hreflang Tags", "Outside <head>", "Issue", "High", ["hreflang_outside_head.csv"]),
]

# Category group spans for Sheet 2 row 3
CATEGORY_GROUPS = [
    ("Response Codes - Internal", 7),
    ("URL Issues", 6),
    ("Page Titles", 7),
    ("Meta Description", 6),
    ("H1", 4),
    ("Canonicals", 10),
    ("Directives", 2),
    ("Sitemaps", 4),
    ("Security", 12),
    ("Page Speed/CWV", 5),
    ("Strucured Tags", 6),
    ("Internal Links", 9),
    ("Content Issues", 9),
    ("Custom Search ", 4),
    ("Pagination", 6),
    ("Hreflang Tags", 13),
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


def load_url_set(report_path, csv_files):
    """Load set of URLs from one or more CSVs (Address column)."""
    urls = set()
    for csv_name in csv_files:
        p = os.path.join(report_path, csv_name)
        if not os.path.exists(p) or os.path.getsize(p) < 100:
            continue
        try:
            df = pd.read_csv(p, encoding="utf-8", low_memory=False,
                             usecols=lambda c: c.lower() in ("address", "source", "destination"))
            for col in df.columns:
                urls.update(df[col].dropna().astype(str).tolist())
        except Exception:
            pass
    return urls


def build_overview_report_masterfile(crawl_id: str, domain: str, report_path: str) -> bytes:
    rulebook = load_rulebook(domain)

    # Load all URLs from internal_all.csv
    int_path = os.path.join(report_path, "internal_all.csv")
    if not os.path.exists(int_path):
        raise FileNotFoundError("internal_all.csv not found")

    df_int = pd.read_csv(int_path, encoding="utf-8", low_memory=False,
                         usecols=lambda c: c.lower() in (
                             "address", "indexability", "status code", "status",
                             "content type", "size (bytes)"
                         ))
    a_col = _gc(df_int, "Address")
    idx_col = _gc(df_int, "Indexability")
    sc_col = _gc(df_int, "Status Code")
    st_col = _gc(df_int, "Status")

    urls_list = df_int[a_col].astype(str).tolist() if a_col else []
    idx_map = dict(zip(df_int[a_col].astype(str), df_int[idx_col].astype(str))) if a_col and idx_col else {}
    sc_map = dict(zip(df_int[a_col].astype(str), df_int[sc_col].astype(str))) if a_col and sc_col else {}
    st_map = dict(zip(df_int[a_col].astype(str), df_int[st_col].astype(str))) if a_col and st_col else {}

    # Load GSC and GA4
    gsc_imp, gsc_clk, ga_sess = {}, {}, {}
    gsc_path = os.path.join(report_path, "search_console_all.csv")
    if os.path.exists(gsc_path):
        try:
            df_gsc = pd.read_csv(gsc_path, encoding="utf-8", low_memory=False)
            ac = _gc(df_gsc, "Address")
            ic = next((c for c in df_gsc.columns if "impression" in c.lower()), None)
            cc = next((c for c in df_gsc.columns if "click" in c.lower()), None)
            if ac:
                addr = df_gsc[ac].astype(str)
                if ic: gsc_imp = dict(zip(addr, df_gsc[ic]))
                if cc: gsc_clk = dict(zip(addr, df_gsc[cc]))
        except Exception:
            pass
    ga_path = os.path.join(report_path, "analytics_all.csv")
    if os.path.exists(ga_path):
        try:
            df_ga = pd.read_csv(ga_path, encoding="utf-8", low_memory=False)
            ac = _gc(df_ga, "Address")
            sc = next((c for c in df_ga.columns if "session" in c.lower()), None)
            if ac and sc:
                ga_sess = dict(zip(df_ga[ac].astype(str), df_ga[sc]))
        except Exception:
            pass

    # Load all issue URL sets
    issue_url_sets = []
    for cat, issue, sev, pri, csvs in ISSUES:
        url_set = load_url_set(report_path, csvs) if csvs else set()
        issue_url_sets.append(url_set)

    # Classify all URLs
    cls_cache = {}
    for url in urls_list:
        cls_cache[url] = classify_url(url, rulebook)

    # Get unique themes from rulebook sorted by priority then name
    theme_priority = {}
    for url in urls_list:
        t1, _, _, pri = cls_cache[url]
        theme = t1 or "Others"
        cur = theme_priority.get(theme)
        if cur is None or PRIORITY_ORDER.get(pri, 3) < PRIORITY_ORDER.get(cur, 3):
            theme_priority[theme] = pri
    sorted_themes = sorted(theme_priority.items(), key=lambda x: (PRIORITY_ORDER.get(x[1], 3), x[0]))
    theme_names = [t for t, _ in sorted_themes]

    # Build Sheet 1 data: for each issue, count affected URLs per theme
    # issue_theme_counts[issue_idx][theme] = count
    issue_theme_counts = []
    for i, (cat, issue, sev, pri, csvs) in enumerate(ISSUES):
        url_set = issue_url_sets[i]
        theme_counts = defaultdict(int)
        for url in url_set:
            if url in cls_cache:
                t1, _, _, _ = cls_cache[url]
                theme = t1 or "Others"
                theme_counts[theme] += 1
        issue_theme_counts.append(theme_counts)

    # Build workbook
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True, "strings_to_urls": False})

    def f(**kw):
        return wb.add_format(kw)

    f_title = f(bold=True, font_name=FONT, font_size=12, font_color=BLACK)
    f_red_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
    f_red_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="left", valign="vcenter", text_wrap=True)
    f_dark_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                   bg_color=DARK, border=1, align="center", valign="vcenter", text_wrap=True)
    f_cell = f(font_name=FONT, font_size=8, font_color=BLACK,
               bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_cell_lft = f(font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_bold_lft = f(bold=True, font_name=FONT, font_size=8, font_color=BLACK,
                   bg_color=WHITE, border=1, align="left", valign="vcenter")
    f_num = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="center", valign="vcenter", num_format="#,##0")
    f_yes = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color="#C6EFCE", border=1, align="center", valign="vcenter")
    f_no = f(font_name=FONT, font_size=8, font_color=BLACK,
             bg_color=WHITE, border=1, align="center", valign="vcenter")
    f_rgt = f(font_name=FONT, font_size=8, font_color=BLACK,
              bg_color=WHITE, border=1, align="right", valign="vcenter")
    f_black_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                    bg_color=BLACK, border=1, align="center", valign="vcenter", text_wrap=True)
    f_black_lft = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                    bg_color=BLACK, border=1, align="left", valign="vcenter", text_wrap=True)

    # ── Sheet 1: Dashboard ────────────────────────────────────────────────
    ws1 = wb.add_worksheet("Dashboard -Dashboard")
    ws1.set_column("A:A", 25)
    ws1.set_column("B:B", 35)
    ws1.set_column("C:C", 12)
    ws1.set_column("D:D", 15)
    for i in range(len(theme_names)):
        ws1.set_column(4 + i, 4 + i, 15)

    ws1.merge_range(0, 0, 0, 3 + len(theme_names), "RankUno SEO Audit - Overview Report", f_title)
    ws1.set_row(6, 42)

    # Header row (row 7, 0-indexed)
    ws1.write(7, 0, "Issue Category", f_black_lft)
    ws1.write(7, 1, "Issue", f_red_hdr)
    ws1.write(7, 2, "Severity", f_black_hdr)
    ws1.write(7, 3, "Priority", f_black_hdr)
    for ti, (theme, tpri) in enumerate(sorted_themes):
        ws1.write(7, 4 + ti, theme, f_black_hdr)

    # Data rows
    prev_cat = None
    for i, (cat, issue, sev, pri, csvs) in enumerate(ISSUES):
        r = 8 + i
        ws1.set_row(r, 14.5)
        cat_display = cat if cat != prev_cat else ""
        ws1.write(r, 0, cat_display, f_bold_lft)
        ws1.write(r, 1, issue, f_red_hdr)
        ws1.write(r, 2, sev, f_cell)
        ws1.write(r, 3, pri, f_cell)
        theme_counts = issue_theme_counts[i]
        for ti, (theme, _) in enumerate(sorted_themes):
            count = theme_counts.get(theme, 0)
            ws1.write(r, 4 + ti, count if count > 0 else "", f_num if count > 0 else f_cell)
        prev_cat = cat

    # ── Sheet 2: Detailed Data ────────────────────────────────────────────
    ws2 = wb.add_worksheet("Dashboard - Detailed Data")
    ws2.set_column("A:A", 60)
    ws2.set_column("B:B", 20)
    ws2.set_column("C:C", 20)
    ws2.set_column("D:D", 15)
    ws2.set_column("E:E", 15)
    ws2.set_column("F:F", 12)
    ws2.set_column("G:G", 20)
    ws2.set_column("H:J", 12)
    ws2.set_column("K:K", 15)
    for i in range(len(ISSUES)):
        ws2.set_column(11 + i, 11 + i, 10)

    # Row 3: category group headers (merged)
    col_offset = 11  # L = col 11
    for cat_name, count in CATEGORY_GROUPS:
        if count > 1:
            ws2.merge_range(2, col_offset, 2, col_offset + count - 1, cat_name, f_dark_hdr)
        else:
            ws2.write(2, col_offset, cat_name, f_dark_hdr)
        col_offset += count

    # Row 4: column headers
    ws2.set_row(3, 42)
    fixed_headers = [
        "URL", "Page Theme 1", "Page Theme 2", "Page Theme 1 - Priority",
        "Indexability", "Status Code", "Status",
        "Impressions", "Clicks", "Organic Sessions",
        "Number of Issues, URL is Impacted with"
    ]
    for ci, h in enumerate(fixed_headers):
        ws2.write(3, ci, h, f_black_hdr if ci < 11 else f_red_hdr)

    for i, (cat, issue, sev, pri, csvs) in enumerate(ISSUES):
        ws2.write(3, 11 + i, issue, f_red_hdr)

    # Data rows
    for row_idx, url in enumerate(urls_list):
        r = 4 + row_idx
        t1, t2, lang, upri = cls_cache.get(url, ("", "", "", "N/A"))
        theme1 = t1 or "-"
        theme2 = t2 or "-"
        indexability = idx_map.get(url, "")
        status_code = sc_map.get(url, "")
        status = st_map.get(url, "")
        impressions = safe_num(gsc_imp.get(url))
        clicks = safe_num(gsc_clk.get(url))
        sessions = safe_num(ga_sess.get(url))

        # Yes/No for each issue
        yn_values = []
        for i, (cat, issue, sev, pri_i, csvs) in enumerate(ISSUES):
            yn = "Yes" if url in issue_url_sets[i] else "No"
            yn_values.append(yn)

        issue_count = yn_values.count("Yes")

        ws2.write(r, 0, url, f_cell_lft)
        ws2.write(r, 1, theme1, f_cell)
        ws2.write(r, 2, theme2, f_cell)
        ws2.write(r, 3, upri, f_cell)
        ws2.write(r, 4, indexability, f_cell)
        ws2.write(r, 5, status_code, f_cell)
        ws2.write(r, 6, status, f_cell)
        ws2.write(r, 7, impressions, f_rgt)
        ws2.write(r, 8, clicks, f_rgt)
        ws2.write(r, 9, sessions, f_rgt)
        ws2.write(r, 10, issue_count, f_num)

        for i, yn in enumerate(yn_values):
            ws2.write(r, 11 + i, yn, f_yes if yn == "Yes" else f_no)

    # Dashboard first
    wb.worksheets_objs = [wb.worksheets_objs[0]] + [wb.worksheets_objs[1]]
    wb.close()
    buf.seek(0)
    return buf.read()
