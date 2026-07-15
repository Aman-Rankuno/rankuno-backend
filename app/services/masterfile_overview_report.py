import os
import io
import csv
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

# ============================================================
# Crawl Overview + Score Weights tabs: built entirely from
# scratch via xlsxwriter, no template file dependency. The row
# structure below (category/severity/priority/description per
# row) mirrors Screaming Frog's "Crawl Overview" report layout
# and was captured once from a verified reference workbook; it
# does not change per crawl, only the Count/%/Total values do.
# ============================================================

CATEGORY_WEIGHTS = [
    ("AI Search", 0.02),
    ("RESPONSE CODES", 0.13),
    ("SECURITY", 0.07),
    ("CANONICALS", 0.07),
    ("DIRECTIVES (ROBOTS, META)", 0.07),
    ("HREFLANG", 0.07),
    ("CONTENT", 0.07),
    ("JAVASCRIPT", 0.05),
    ("LINKS", 0.05),
    ("Structured Data", 0.05),
    ("PAGE TITLES", 0.04),
    ("Sitemaps", 0.03),
    ("H1", 0.03),
    ("PAGINATION", 0.03),
    ("PageSpeed", 0.03),
    ("Mobile", 0.03),
    ("Accessibility", 0.03),
    ("URL", 0.02),
    ("META DESCRIPTION", 0.02),
    ("H2", 0.02),
    ("IMAGES", 0.02),
    ("Analytics", 0.02),
    ("DIRECTIVES (robots.txt)", 0.01),
    ("Search Console", 0.01),
    ("Validation", 0.01),
    ("AMP", 0),
    ("Meta Keywords", 0),
]

SEVERITY_MULT = [
    ("Issue", 1),
    ("Warning", 0.8),
    ("Opportunity", 0.5),
]

PRIORITY_MULT = [
    ("High", 1),
    ("Medium", 0.8),
    ("Low", 0.5),
]

CRAWL_OVERVIEW_ROWS = [
    ('NA', 'Start Date', None, 'NA', 'NA', 'NA', False),
    ('NA', 'Start Time', None, 'NA', 'NA', 'NA', False),
    ('NA', 'Last Modified Date', None, 'NA', 'NA', 'NA', False),
    ('NA', 'Last Modified Time', None, 'NA', 'NA', 'NA', False),
    ('NA', 'Elapsed', None, 'NA', 'NA', 'NA', False),
    ('NA', 'Report Date', None, 'NA', 'NA', 'NA', False),
    ('NA', 'Report Time', None, 'NA', 'NA', 'NA', False),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Summary', 'Total URLs Description', 'NA', 'NA', 'NA', False),
    ('NA', 'Total URLs Encountered', 'URLs Encountered', 'NA', 'NA', 'NA', True),
    ('NA', 'Total URLs Crawled', 'URLs Encountered', 'NA', 'NA', 'NA', True),
    ('NA', 'Total Internal blocked by robots.txt', 'URLs Encountered', 'NA', 'NA', 'NA', True),
    ('NA', 'Total External blocked by robots.txt', 'URLs Encountered', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'URLs Displayed', 'URLs Displayed', 'NA', 'NA', 'NA', True),
    ('NA', 'Total Internal URLs', 'URLs Displayed', 'NA', 'NA', 'NA', True),
    ('NA', 'Total External URLs', 'URLs Displayed', 'NA', 'NA', 'NA', True),
    ('NA', 'Total Internal Indexable URLs', 'URLs Displayed', 'NA', 'NA', 'NA', True),
    ('NA', 'Total Internal Non-Indexable URLs', 'URLs Displayed', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Internal', None, 'NA', 'NA', 'NA', False),
    ('NA', 'All', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'HTML', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'JavaScript', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'CSS', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Images', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Media', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Fonts', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'XML', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'PDF', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Plugins', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Other', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Unknown', 'Internal URLs', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'External', None, 'NA', 'NA', 'NA', False),
    ('NA', 'All', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'HTML', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'JavaScript', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'CSS', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Images', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Media', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Fonts', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'XML', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'PDF', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Plugins', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Other', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Unknown', 'External URLs', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Security', None, 'NA', 'NA', 'NA', False),
    (79, 'All', 'All internal and non-HTML external URLs', 'SECURITY', 'NA', 'NA', True),
    (81, 'HTTP URLs', 'All internal and non-HTML external URLs', 'SECURITY', 'Issue', 'High', True),
    (80, 'HTTPS URLs', 'All internal and non-HTML external URLs', 'SECURITY', 'NA', 'NA', True),
    (82, 'Mixed Content', 'All internal and non-HTML external URLs', 'SECURITY', 'Issue', 'High', True),
    (83, 'Form URL Insecure', 'All internal and non-HTML external URLs', 'SECURITY', 'Issue', 'High', True),
    (84, 'Form on HTTP URL', 'All internal and non-HTML external URLs', 'SECURITY', 'Issue', 'High', True),
    (91, 'Unsafe Cross-Origin Links', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (92, 'Protocol-Relative Resource Links', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (93, 'Missing HSTS Header', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (94, 'Missing Content-Security-Policy Header', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (95, 'Missing X-Content-Type-Options Header', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (96, 'Missing X-Frame-Options Header', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (97, 'Missing Secure Referrer-Policy Header', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    (98, 'Bad Content Type', 'All internal and non-HTML external URLs', 'SECURITY', 'Warning', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (50, 'Response Codes', None, 'RESPONSE CODES', 'NA', 'NA', False),
    ('NA', 'All', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Blocked by Robots.txt', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Blocked Resource', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'No Response', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Success (2xx)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Redirection (3xx)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Redirection (JavaScript)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Redirection (Meta Refresh)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Redirection (HTTP Refresh)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Client Error (4xx)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Server Error (5xx)', 'All Internal & External Crawled URLs', 'NA', 'NA', 'NA', True),
    (51, 'Internal All', 'All Internal Crawled URLs', 'RESPONSE CODES', 'NA', 'NA', True),
    (62, 'Internal Blocked by Robots.txt', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Warning', 'High', True),
    (63, 'Internal Blocked Resource', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Warning', 'High', True),
    (53, 'Internal No Response', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Issue', 'High', True),
    (52, 'Internal Success (2xx)', 'All Internal Crawled URLs', 'RESPONSE CODES', 'NA', 'NA', True),
    ('NA', 'Internal Redirection (3xx)', 'All Internal Crawled URLs', 'NA', 'NA', 'NA', True),
    (64, 'Internal Redirection (JavaScript)', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (65, 'Internal Redirection (Meta Refresh)', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (66, 'Internal Redirection (HTTP Refresh)', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (55, 'Internal Redirect Chain', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Issue', 'High', True),
    (56, 'Internal Redirect Loop', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Issue', 'High', True),
    (57, 'Internal Client Error (4xx)', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Issue', 'High', True),
    (58, 'Internal Server Error (5xx)', 'All Internal Crawled URLs', 'RESPONSE CODES', 'Issue', 'High', True),
    (67, 'External All', 'All External Crawled URLs', 'RESPONSE CODES', 'NA', 'NA', True),
    (69, 'External Blocked by Robots.txt', 'All External Crawled URLs', 'RESPONSE CODES', 'NA', 'NA', True),
    (70, 'External Blocked Resource', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Medium', True),
    (71, 'External No Response', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (68, 'External Success (2xx)', 'All External Crawled URLs', 'RESPONSE CODES', 'NA', 'NA', True),
    (77, 'External Redirection (3xx)', 'All External Crawled URLs', 'RESPONSE CODES', 'Opportunity', 'Low', True),
    (74, 'External Redirection (JavaScript)', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (75, 'External Redirection (Meta Refresh)', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (76, 'External Redirection (HTTP Refresh)', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (72, 'External Client Error (4xx)', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    (73, 'External Server Error (5xx)', 'All External Crawled URLs', 'RESPONSE CODES', 'Warning', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (403, 'URL', None, 'URL', 'NA', 'NA', False),
    (404, 'All', 'Internal URLs', 'URL', 'NA', 'NA', True),
    (409, 'Non ASCII Characters', 'Internal URLs', 'URL', 'Warning', 'Low', True),
    (415, 'Underscores', 'Internal URLs', 'URL', 'Opportunity', 'Low', True),
    (410, 'Uppercase', 'Internal URLs', 'URL', 'Warning', 'Low', True),
    (406, 'Multiple Slashes', 'Internal URLs', 'URL', 'Issue', 'Low', True),
    (411, 'Repetitive Path', 'Internal URLs', 'URL', 'Warning', 'Low', True),
    (407, 'Contains Space', 'Internal URLs', 'URL', 'Issue', 'Low', True),
    (412, 'Internal Search', 'Internal URLs', 'URL', 'Warning', 'Low', True),
    (413, 'Parameters', 'Internal URLs', 'URL', 'Warning', 'Low', True),
    (408, 'Broken Bookmark', 'Internal URLs', 'URL', 'Issue', 'Low', True),
    (414, 'GA Tracking Parameters', 'Internal URLs', 'URL', 'Warning', 'Low', True),
    (416, 'Over 115 Characters', 'Internal URLs', 'URL', 'Opportunity', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (233, 'Page Titles', None, 'PAGE TITLES', 'NA', 'NA', False),
    (234, 'All', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'NA', 'NA', True),
    (235, 'Missing', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Issue', 'High', True),
    (238, 'Duplicate', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Opportunity', 'Low', True),
    (239, 'Over 60 Characters', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Opportunity', 'Low', True),
    (240, 'Below 30 Characters', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Opportunity', 'Low', True),
    (241, 'Over 561 Pixels', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Opportunity', 'Low', True),
    (242, 'Below 200 Pixels', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Opportunity', 'Low', True),
    (243, 'Same as H1', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Opportunity', 'Low', True),
    (236, 'Multiple', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Issue', 'High', True),
    (237, 'Outside <head>', 'Internal HTML pages with 2xx response', 'PAGE TITLES', 'Issue', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (417, 'Meta Description', None, 'META DESCRIPTION', 'NA', 'NA', False),
    (418, 'All', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'NA', 'NA', True),
    (421, 'Missing', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Opportunity', 'Low', True),
    (422, 'Duplicate', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Opportunity', 'Low', True),
    (423, 'Over 155 Characters', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Opportunity', 'Low', True),
    (424, 'Below 70 Characters', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Opportunity', 'Low', True),
    (425, 'Over 985 Pixels', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Opportunity', 'Low', True),
    (426, 'Below 400 Pixels', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Opportunity', 'Low', True),
    (419, 'Multiple', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Issue', 'Medium', True),
    (420, 'Outside <head>', 'Internal HTML pages with 2xx response', 'META DESCRIPTION', 'Issue', 'Medium', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Meta Keywords', None, 'NA', 'NA', 'NA', False),
    ('NA', 'All', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', 'Duplicate', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', 'Multiple', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (256, 'H1', None, 'H1', 'NA', 'NA', False),
    (257, 'All', 'Internal HTML pages with 2xx response', 'H1', 'NA', 'NA', True),
    (258, 'Missing', 'Internal HTML pages with 2xx response', 'H1', 'Issue', 'Medium', True),
    (262, 'Duplicate', 'Internal HTML pages with 2xx response', 'H1', 'Opportunity', 'Low', True),
    (263, 'Over 70 Characters', 'Internal HTML pages with 2xx response', 'H1', 'Opportunity', 'Low', True),
    (259, 'Multiple', 'Internal HTML pages with 2xx response', 'H1', 'Warning', 'Medium', True),
    (260, 'Alt Text in H1', 'Internal HTML pages with 2xx response', 'H1', 'Warning', 'Low', True),
    (261, 'Non-Sequential', 'Internal HTML pages with 2xx response', 'H1', 'Warning', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (427, 'H2', None, 'H2', 'NA', 'NA', False),
    (428, 'All', 'Internal HTML pages with 2xx response', 'H2', 'NA', 'NA', True),
    (429, 'Missing', 'Internal HTML pages with 2xx response', 'H2', 'Warning', 'Low', True),
    (432, 'Duplicate', 'Internal HTML pages with 2xx response', 'H2', 'Opportunity', 'Low', True),
    (433, 'Over 70 Characters', 'Internal HTML pages with 2xx response', 'H2', 'Opportunity', 'Low', True),
    (430, 'Multiple', 'Internal HTML pages with 2xx response', 'H2', 'Warning', 'Low', True),
    (431, 'Non-Sequential', 'Internal HTML pages with 2xx response', 'H2', 'Warning', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (152, 'Content', None, 'CONTENT', 'NA', 'NA', False),
    (153, 'All', 'Internal HTML pages with 2xx response', 'CONTENT', 'NA', 'NA', True),
    (156, 'Exact Duplicates', 'Internal HTML pages with 2xx response', 'CONTENT', 'Issue', 'Medium', True),
    (157, 'Near Duplicates', 'Internal HTML pages with 2xx response', 'CONTENT', 'Issue', 'Medium', True),
    (158, 'Semantically Similar', 'Internal HTML pages with 2xx response', 'CONTENT', 'Issue', 'Medium', True),
    (159, 'Low Relevance Content', 'Internal HTML pages with 2xx response', 'CONTENT', 'Warning', 'High', True),
    (160, 'Low Content Pages', 'Internal HTML pages with 2xx response', 'CONTENT', 'Warning', 'High', True),
    (161, 'Soft 404 Pages', 'Internal HTML pages with 2xx response', 'CONTENT', 'Warning', 'Medium', True),
    (162, 'Spelling Errors', 'Internal HTML pages with 2xx response', 'CONTENT', 'Warning', 'Medium', True),
    (163, 'Grammar Errors', 'Internal HTML pages with 2xx response', 'CONTENT', 'Warning', 'Medium', True),
    (164, 'Readability Difficult', 'Internal HTML pages with 2xx response', 'CONTENT', 'Warning', 'Low', True),
    (165, 'Readability Very Difficult', 'Internal HTML pages with 2xx response', 'CONTENT', 'Opportunity', 'Medium', True),
    (166, 'Lorem Ipsum Placeholder', 'Internal HTML pages with 2xx response', 'CONTENT', 'Opportunity', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (434, 'Images', None, 'IMAGES', 'NA', 'NA', False),
    (435, 'All', 'Images with 2xx response', 'IMAGES', 'NA', 'NA', True),
    (440, 'Over 100 KB', 'Images with 2xx response', 'IMAGES', 'Opportunity', 'Medium', True),
    (436, 'Missing Alt Text', 'Images with 2xx response', 'IMAGES', 'Issue', 'Low', True),
    (437, 'Missing Alt Attribute', 'Images with 2xx response', 'IMAGES', 'Issue', 'Low', True),
    (441, 'Alt Text Over 100 Characters', 'Images with 2xx response', 'IMAGES', 'Opportunity', 'Low', True),
    (438, 'Background Images', 'Images with 2xx response', 'IMAGES', 'Warning', 'Low', True),
    (442, 'Incorrectly Sized Images', 'Images with 2xx response', 'IMAGES', 'Opportunity', 'Low', True),
    (443, 'Missing Size Attributes', 'Images with 2xx response', 'IMAGES', 'Opportunity', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (100, 'Canonicals', None, 'CANONICALS', 'NA', 'NA', False),
    (101, 'All', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'NA', 'NA', True),
    ('NA', 'Contains Canonical', 'Internal HTML or PDF pages with 2xx response', 'NA', 'NA', 'NA', True),
    (102, 'Self Referencing', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'NA', 'NA', True),
    (104, 'Canonicalised', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Issue', 'High', True),
    (110, 'Missing', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Warning', 'Medium', True),
    (113, 'Multiple', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Warning', 'Low', True),
    (105, 'Multiple Conflicting', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Issue', 'High', True),
    (106, 'Non-Indexable Canonical', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Issue', 'High', True),
    (111, 'Canonical Is Relative', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Warning', 'Medium', True),
    (112, 'Unlinked', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Warning', 'Medium', True),
    (107, 'Invalid Attribute In Annotation', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Issue', 'High', True),
    (108, 'Contains Fragment URL', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Issue', 'Medium', True),
    (109, 'Outside <head>', 'Internal HTML or PDF pages with 2xx response', 'CANONICALS', 'Issue', 'Medium', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (244, 'Pagination', None, 'PAGINATION', 'NA', 'NA', False),
    (245, 'All', 'Internal HTML pages with 2xx response', 'PAGINATION', 'NA', 'NA', True),
    (246, 'Contains Pagination', 'Internal HTML pages with 2xx response', 'PAGINATION', 'NA', 'NA', True),
    (247, 'First Page', 'Internal HTML pages with 2xx response', 'PAGINATION', 'NA', 'NA', True),
    (248, 'Paginated 2+ Pages', 'Internal HTML pages with 2xx response', 'PAGINATION', 'NA', 'NA', True),
    (254, 'Pagination URL Not in Anchor Tag', 'Internal HTML pages with 2xx response', 'PAGINATION', 'Issue', 'Low', True),
    ('NA', 'Non-200 Pagination URLs', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    (249, 'Unlinked Pagination URLs', 'Internal HTML pages with 2xx response', 'PAGINATION', 'Issue', 'High', True),
    (250, 'Non-Indexable', 'Internal HTML pages with 2xx response', 'PAGINATION', 'Issue', 'High', True),
    (253, 'Multiple Pagination URLs', 'Internal HTML pages with 2xx response', 'PAGINATION', 'Issue', 'Low', True),
    (255, 'Pagination Loop', 'Internal HTML pages with 2xx response', 'PAGINATION', 'Issue', 'Low', True),
    (251, 'Sequence Error', 'Internal HTML pages with 2xx response', 'PAGINATION', 'Issue', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (114, 'Directives', None, 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', False),
    (115, 'All', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (116, 'Index', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (125, 'Noindex', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'High', True),
    (117, 'Follow', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (126, 'Nofollow', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'High', True),
    (127, 'None', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'High', True),
    (118, 'NoArchive', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (128, 'NoSnippet', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'Low', True),
    (119, 'Max-Snippet', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (120, 'Max-Image-Preview', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (121, 'Max-Video-Preview', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (129, 'NoODP', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'Low', True),
    (130, 'NoYDIR', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'Low', True),
    (124, 'NoImageIndex', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Issue', 'Low', True),
    (131, 'NoTranslate', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'Low', True),
    (132, 'Unavailable_After', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Warning', 'Medium', True),
    (122, 'Refresh', 'Internal HTML pages with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'NA', 'NA', True),
    (123, 'Outside <head>', 'Internal URLs with 2xx response', 'DIRECTIVES (ROBOTS, META)', 'Issue', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (136, 'Hreflang', None, 'HREFLANG', 'NA', 'NA', False),
    (137, 'All', 'Internal pages with 2xx response', 'HREFLANG', 'NA', 'NA', True),
    (138, 'Contains hreflang', 'Internal pages with 2xx response', 'HREFLANG', 'NA', 'NA', True),
    (139, 'Non-200 hreflang URLs', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (140, 'Unlinked hreflang URLs', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (141, 'Missing Return Links', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (142, 'Inconsistent Language & Region Return Links', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (143, 'Non-Canonical Return Links', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (144, 'Noindex Return Links', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (145, 'Incorrect Language & Region Codes', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (146, 'Multiple Entries', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (147, 'Missing Self Reference', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'Medium', True),
    (148, 'Not Using Canonical', 'Internal pages with 2xx response', 'HREFLANG', 'Issue', 'High', True),
    (149, 'Missing X-Default', 'Internal pages with 2xx response', 'HREFLANG', 'Warning', 'High', True),
    (150, 'Missing', 'Internal pages with 2xx response', 'HREFLANG', 'Warning', 'High', True),
    (151, 'Outside <head>', 'Internal pages with 2xx response', 'HREFLANG', 'Warning', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (167, 'JavaScript', None, 'JAVASCRIPT', 'NA', 'NA', False),
    ('NA', 'All', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    (172, 'Pages with Blocked Resources', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Issue', 'Medium', True),
    (173, 'Contains JavaScript Links', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Issue', 'Medium', True),
    (174, 'Contains JavaScript Content', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (175, 'Noindex Only in Original HTML', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (176, 'Nofollow Only in Original HTML', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (177, 'Canonical Only in Rendered HTML', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (178, 'Canonical Mismatch', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (179, 'Page Title Only in Rendered HTML', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (182, 'Page Title Updated by JavaScript', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Low', True),
    (168, 'Meta Description Only in Rendered HTML', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Issue', 'High', True),
    (169, 'Meta Description Updated by JavaScript', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Issue', 'High', True),
    (170, 'H1 Only in Rendered HTML', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Issue', 'High', True),
    (180, 'H1 Updated by JavaScript', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (181, 'Uses Old AJAX Crawling Scheme URLs', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Medium', True),
    (171, 'Uses Old AJAX Crawling Scheme Meta Fragment Tag', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'High', True),
    (183, 'Pages with JavaScript Errors', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Low', True),
    (184, 'Pages with JavaScript Warnings', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Low', True),
    (185, 'Pages with Chrome Issues', 'Internal HTML pages with 2xx response', 'JAVASCRIPT', 'Warning', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (187, 'Links', None, 'LINKS', 'NA', 'NA', False),
    (188, 'All', 'Internal HTML pages with 2xx response', 'LINKS', 'NA', 'NA', True),
    (192, 'Pages With High Crawl Depth', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'High', True),
    (204, 'Pages Without Internal Outlinks', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'Low', True),
    (208, 'Internal Nofollow Outlinks', 'Internal HTML pages with 2xx response', 'LINKS', 'Opportunity', 'Medium', True),
    (199, 'Internal Outlinks With No Anchor Text', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'High', True),
    (198, 'Non-Descriptive Anchor Text In Internal Outlinks', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'High', True),
    (205, 'Pages With High External Outlinks', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'Low', True),
    (206, 'Pages With High Internal Outlinks', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'Low', True),
    (207, 'Follow & Nofollow Internal Inlinks To Page', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'Low', True),
    (209, 'Internal Nofollow Inlinks Only', 'Internal HTML pages with 2xx response', 'LINKS', 'Opportunity', 'Medium', True),
    (189, 'Outlinks To Localhost', 'Internal HTML pages with 2xx response', 'LINKS', 'Issue', 'High', True),
    (193, 'Non-Indexable Page Inlinks Only', 'Internal HTML pages with 2xx response', 'LINKS', 'Warning', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'AMP', None, 'NA', 'NA', 'NA', False),
    ('NA', 'All', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Non-200 Response', 'Internal HTML URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing Non-AMP Return Link', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing Canonical to Non-AMP', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Non-Indexable Canonical', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Indexable', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Non-Indexable', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing <html amp> Tag', 'Internal HTML URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing/Invalid <!doctype html> Tag', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing <head> Tag', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing <body> Tag', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing Canonical', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing/Invalid <meta charset> Tag', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing/Invalid <meta viewport> Tag', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing/Invalid AMP Script', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Missing/Invalid AMP Boilerplate', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Contains Disallowed HTML', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', 'Other Validation Errors', 'Internal AMP URLs', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (210, 'Structured Data', None, 'Structured Data', 'NA', 'NA', False),
    (211, 'All', 'Internal HTML pages with 2xx response', 'Structured Data', 'NA', 'NA', True),
    (212, 'Contains Structured Data', 'Internal HTML pages with 2xx response', 'Structured Data', 'NA', 'NA', True),
    ('NA', 'Missing', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    (217, 'Validation Errors', 'Internal HTML pages with 2xx response', 'Structured Data', 'Issue', 'High', True),
    (222, 'Validation Warnings', 'Internal HTML pages with 2xx response', 'Structured Data', 'Warning', 'Low', True),
    (218, 'Rich Result Validation Errors', 'Internal HTML pages with 2xx response', 'Structured Data', 'Issue', 'High', True),
    (223, 'Rich Result Validation Warnings', 'Internal HTML pages with 2xx response', 'Structured Data', 'Warning', 'Low', True),
    (219, 'Parse Errors', 'Internal HTML pages with 2xx response', 'Structured Data', 'Issue', 'High', True),
    (213, 'Microdata URLs', 'Internal HTML pages with 2xx response', 'Structured Data', 'NA', 'NA', True),
    (214, 'JSON-LD URLs', 'Internal HTML pages with 2xx response', 'Structured Data', 'NA', 'NA', True),
    (215, 'RDFa URLs', 'Internal HTML pages with 2xx response', 'Structured Data', 'NA', 'NA', True),
    (216, 'Rich Result Feature Detected', 'Internal HTML pages with 2xx response', 'Structured Data', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (224, 'Sitemaps', None, 'Sitemaps', 'NA', 'NA', False),
    (225, 'All', 'Internal HTML/XML pages', 'Sitemaps', 'NA', 'NA', True),
    ('NA', 'URLs in Sitemap', 'Internal HTML/XML pages', 'NA', 'NA', 'NA', True),
    (231, 'URLs not in Sitemap', 'Internal HTML/XML pages', 'Sitemaps', 'Issue', 'Medium', True),
    (227, 'Orphan URLs', 'Internal HTML/XML pages', 'Sitemaps', 'Issue', 'High', True),
    ('NA', 'Non-Indexable URLs in Sitemap', 'Internal HTML/XML pages', 'NA', 'NA', 'NA', True),
    ('NA', 'URLs in Multiple Sitemaps', 'Internal HTML/XML pages', 'NA', 'NA', 'NA', True),
    (229, 'XML Sitemap with over 50k URLs', 'Internal HTML/XML pages', 'Sitemaps', 'Issue', 'High', True),
    (228, 'XML Sitemap over 50MB', 'Internal HTML/XML pages', 'Sitemaps', 'Issue', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (264, 'PageSpeed', None, 'PageSpeed', 'NA', 'NA', False),
    (265, 'All', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'NA', 'NA', True),
    (271, 'Document Request Latency', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (272, 'LCP Request Discovery', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (273, 'Render Blocking Requests', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (274, 'Network Dependency Tree', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (275, 'Use Efficient Cache Lifetimes', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (276, 'Layout Shift Culprits', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (289, 'Optimize DOM Size', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Low', True),
    (277, 'Improve Image Delivery', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (278, 'Forced Reflow', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (279, 'Legacy JavaScript', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (280, 'Duplicated JavaScript', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (288, 'Font Display', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Low', True),
    (281, 'Avoid Enormous Network Payloads', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (282, 'Minify CSS', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (283, 'Minify JavaScript', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (284, 'Reduce Unused CSS', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (285, 'Reduce Unused JavaScript', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (286, 'Reduce JavaScript Execution Time', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (287, 'Minimize Main-Thread Work', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Medium', True),
    (290, 'Request Errors', 'Internal HTML URLs with 2xx response', 'PageSpeed', 'Opportunity', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (291, 'Mobile', None, 'Mobile', 'NA', 'NA', False),
    (292, 'All', None, 'Mobile', 'NA', 'NA', True),
    (293, 'Viewport Not Set', None, 'Mobile', 'Issue', 'High', True),
    (297, 'Target Size', None, 'Mobile', 'Issue', 'Medium', True),
    (294, 'Content Not Sized Correctly', None, 'Mobile', 'Issue', 'High', True),
    (295, 'Illegible Font Size', None, 'Mobile', 'Issue', 'High', True),
    (296, 'Contains Unsupported Plugins', None, 'Mobile', 'Issue', 'High', True),
    (298, 'Mobile Alternate Link', None, 'Mobile', 'Warning', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (299, 'Accessibility', None, 'Accessibility', 'NA', 'NA', False),
    (300, 'All', None, 'Accessibility', 'NA', 'NA', True),
    (385, 'Accessibility Score Poor', None, 'Accessibility', 'Warning', 'Medium', True),
    (401, 'Accessibility Score Needs Improvement', None, 'Accessibility', 'Opportunity', 'Low', True),
    (301, 'Accessibility Score Good', None, 'Accessibility', 'NA', 'NA', True),
    (302, 'WCAG 2.0 A Violation', None, 'Accessibility', 'Issue', 'High', True),
    (303, 'WCAG 2.0 AA Violation', None, 'Accessibility', 'Issue', 'High', True),
    (402, 'WCAG 2.0 AAA Violation', None, 'Accessibility', 'Opportunity', 'Low', True),
    (304, 'WCAG 2.1 AA Violation', None, 'Accessibility', 'Issue', 'High', True),
    (305, 'WCAG 2.2 AA Violation', None, 'Accessibility', 'Issue', 'High', True),
    (373, 'Best Practice Violation', None, 'Accessibility', 'Warning', 'Low', True),
    (307, 'Accesskey Attribute Value Must Be Unique', None, 'Accessibility', 'Warning', 'High', True),
    (374, 'Ensure Elements Marked Presentational Are Ignored', None, 'Accessibility', 'Warning', 'Low', True),
    (308, 'Elements Must Not Have Tabindex Greater Than Zero', None, 'Accessibility', 'Warning', 'High', True),
    (309, 'Scrollable Region Requires Keyboard Access', None, 'Accessibility', 'Warning', 'High', True),
    (386, 'Skip-link Target Should Exist & Be Focusable', None, 'Accessibility', 'Warning', 'Medium', True),
    (310, 'Required ARIA Attributes Must Be Provided', None, 'Accessibility', 'Warning', 'High', True),
    (311, 'Role=text Should Have No Focusable Descendants', None, 'Accessibility', 'Warning', 'High', True),
    (312, 'ARIA Attribute Must Be Used As Specified For Role', None, 'Accessibility', 'Warning', 'High', True),
    (313, 'ARIA Attributes Require Valid Values', None, 'Accessibility', 'Warning', 'High', True),
    (314, 'ARIA Attributes Require Valid Names', None, 'Accessibility', 'Warning', 'High', True),
    (315, 'ARIA Commands Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (316, 'ARIA Dialog & Alertdialog Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (317, 'ARIA Input Fields Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (318, 'ARIA Meter Nodes Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (319, 'ARIA Progressbar Nodes Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (375, 'ARIA Role Should Be Appropriate For Element', None, 'Accessibility', 'Warning', 'Low', True),
    (320, 'ARIA Roles Must Be Contained By Required Parent', None, 'Accessibility', 'Warning', 'High', True),
    (321, 'ARIA Roles Require Valid Values', None, 'Accessibility', 'Warning', 'High', True),
    (322, 'ARIA Toggle Fields Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (323, 'ARIA Tooltip Nodes Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (324, 'ARIA Treeitem Nodes Require Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (325, 'Certain ARIA Roles Must Contain Specific Children', None, 'Accessibility', 'Warning', 'High', True),
    (376, 'Deprecated ARIA Roles Must Not Be Used', None, 'Accessibility', 'Warning', 'Low', True),
    (326, 'Aria-braille Require Non-braille Equivalent', None, 'Accessibility', 'Warning', 'High', True),
    (327, 'Aria-hidden Elements Contains Focusable Elements', None, 'Accessibility', 'Warning', 'High', True),
    (328, 'Aria-hidden=true Must Not Be Used In <body>', None, 'Accessibility', 'Warning', 'High', True),
    (329, 'Elements Must Only Use Permitted ARIA Attributes', None, 'Accessibility', 'Warning', 'High', True),
    (330, 'Elements Must Use Allowed ARIA Attributes', None, 'Accessibility', 'Warning', 'High', True),
    (331, 'IDs Used In ARIA & Labels Must Be Unique', None, 'Accessibility', 'Warning', 'High', True),
    (332, 'Page Requires Means To Bypass Repeated Blocks', None, 'Accessibility', 'Warning', 'High', True),
    (387, 'All Page Content Must Be Contained By Landmarks', None, 'Accessibility', 'Warning', 'Medium', True),
    (388, 'Page Requires One Main Landmark', None, 'Accessibility', 'Warning', 'Medium', True),
    (389, 'Page Must Not Have More Than One Banner Landmark', None, 'Accessibility', 'Warning', 'Medium', True),
    (390, 'Banner Landmark Must Not Be In Another Landmark', None, 'Accessibility', 'Warning', 'Medium', True),
    (391, 'Page Must Not Have Multiple Contentinfo Landmarks', None, 'Accessibility', 'Warning', 'Medium', True),
    (392, 'Page Requires At Most One Main Landmark', None, 'Accessibility', 'Warning', 'Medium', True),
    (393, 'Complementary Landmarks & Asides Must Be Top Level', None, 'Accessibility', 'Warning', 'Medium', True),
    (394, 'Contentinfo Landmark Must Be Top Level Landmark', None, 'Accessibility', 'Warning', 'Medium', True),
    (395, 'Main Landmark Must Not Be In Another Landmark', None, 'Accessibility', 'Warning', 'Medium', True),
    (396, 'Landmarks Require Unique Role Or Accessible Name', None, 'Accessibility', 'Warning', 'Medium', True),
    (397, 'Form Field Must Not Have Multiple Label Elements', None, 'Accessibility', 'Warning', 'Medium', True),
    (333, 'Form <input> Elements Require Labels', None, 'Accessibility', 'Warning', 'High', True),
    (334, 'Form Elements Should Have Visible Label', None, 'Accessibility', 'Warning', 'High', True),
    (335, 'Autocomplete Attribute Must Be Used Correctly', None, 'Accessibility', 'Warning', 'High', True),
    (336, 'Frames Require Title Attribute', None, 'Accessibility', 'Warning', 'High', True),
    (337, 'Frames Require Unique Title Attribute', None, 'Accessibility', 'Warning', 'High', True),
    (338, 'Frames Should Be Tested With axe-core', None, 'Accessibility', 'Warning', 'High', True),
    (339, 'Frames With Focusable Content Must Not Use tabindex=-1', None, 'Accessibility', 'Warning', 'High', True),
    (340, 'Page Must Contain <title>', None, 'Accessibility', 'Warning', 'High', True),
    (398, 'Page Must Contain <h1>', None, 'Accessibility', 'Warning', 'Medium', True),
    (399, 'Heading Levels Should Only Increase By One', None, 'Accessibility', 'Warning', 'Medium', True),
    (377, 'Headings Should Not Be Empty', None, 'Accessibility', 'Warning', 'Low', True),
    (378, 'Meta Viewport Should Allow Zoom & Scale Up to 500%', None, 'Accessibility', 'Warning', 'Low', True),
    (341, 'Meta Viewport Zoom & Scaling Disabled', None, 'Accessibility', 'Warning', 'High', True),
    (342, 'HTML Element Lang Attribute Value Must Be Valid', None, 'Accessibility', 'Warning', 'High', True),
    (343, 'HTML Element Requires Lang Attribute', None, 'Accessibility', 'Warning', 'High', True),
    (400, 'HTML Lang & XML Lang Value Should Match', None, 'Accessibility', 'Warning', 'Medium', True),
    (344, 'Lang Attribute Requires Valid Value', None, 'Accessibility', 'Warning', 'High', True),
    (379, 'Delayed Meta Refresh Must Not Be Used', None, 'Accessibility', 'Warning', 'Low', True),
    (306, 'Timed Meta Refresh Must Not Exist', None, 'Accessibility', 'Issue', 'Low', True),
    (345, 'Image Button Requires Alternate Text', None, 'Accessibility', 'Warning', 'High', True),
    (346, 'Images Require Alternate Text', None, 'Accessibility', 'Warning', 'High', True),
    (347, '<object> Elements Require Alternate Text', None, 'Accessibility', 'Warning', 'High', True),
    (348, 'Active <area> Elements Require Alternate Text', None, 'Accessibility', 'Warning', 'High', True),
    (380, 'Alt Text Should Not Be Repeated As Text', None, 'Accessibility', 'Warning', 'Low', True),
    (349, 'Elements Marked role=img Require Alternate Text', None, 'Accessibility', 'Warning', 'High', True),
    (350, 'SVG Images & Graphics Require Accessible Text', None, 'Accessibility', 'Warning', 'High', True),
    (381, 'Server-Side Image Maps Must Not Be Used', None, 'Accessibility', 'Warning', 'Low', True),
    (351, '<video> Elements Require <track> For Captions', None, 'Accessibility', 'Warning', 'High', True),
    (352, '<video> or <audio> Elements Must Not Auto-play', None, 'Accessibility', 'Warning', 'High', True),
    (353, 'Buttons Require Discernible Text', None, 'Accessibility', 'Warning', 'High', True),
    (354, 'Input Buttons Require Discernible Text', None, 'Accessibility', 'Warning', 'High', True),
    (355, 'Inline Text Spacing Must Be Adjustable', None, 'Accessibility', 'Warning', 'High', True),
    (356, 'Links Must Be Distinguishable', None, 'Accessibility', 'Warning', 'High', True),
    (357, 'Links Require Discernible Text', None, 'Accessibility', 'Warning', 'High', True),
    (382, 'Links With Same Accessible Name', None, 'Accessibility', 'Warning', 'Low', True),
    (358, 'Select Element Requires Accessible Name', None, 'Accessibility', 'Warning', 'High', True),
    (359, 'Summary Elements Require Discernible Text', None, 'Accessibility', 'Warning', 'High', True),
    (360, 'Deprecated <marquee> Element Must Not Be Used', None, 'Accessibility', 'Warning', 'High', True),
    (361, '<blink> Elements Deprecated & Must Not Be Used', None, 'Accessibility', 'Warning', 'High', True),
    (362, 'Text Requires Higher Color Contrast Ratio', None, 'Accessibility', 'Warning', 'High', True),
    (363, 'Text Requires Higher Color Contrast to Background', None, 'Accessibility', 'Warning', 'High', True),
    (364, 'Touch Targets Require Sufficient Size & Spacing', None, 'Accessibility', 'Warning', 'High', True),
    (365, 'Interactive Controls Must Not Be Nested', None, 'Accessibility', 'Warning', 'High', True),
    (366, 'List Items Must Be Contained In List Elements', None, 'Accessibility', 'Warning', 'High', True),
    (367, 'Lists Must Only Contain <li> Content Elements', None, 'Accessibility', 'Warning', 'High', True),
    (368, '<dl> Must Only Have Ordered <dt> & <dd> Groups', None, 'Accessibility', 'Warning', 'High', True),
    (369, '<dt> & <dd> Elements Must Be Contained by <dl>', None, 'Accessibility', 'Warning', 'High', True),
    (370, '<th> Element Requires Associated Data Cells', None, 'Accessibility', 'Warning', 'High', True),
    (371, 'Table Header Attr Must Refer To Cell In Same Table', None, 'Accessibility', 'Warning', 'High', True),
    (383, 'Table Headers Require Discernible Text', None, 'Accessibility', 'Warning', 'Low', True),
    (384, 'Table With Identical Summary & Caption Text', None, 'Accessibility', 'Warning', 'Low', True),
    (372, 'Scope Attribute Should Be Used Correctly On Tables', None, 'Accessibility', 'Warning', 'High', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Custom Search', None, 'NA', 'NA', 'NA', False),
    (155, 'All', 'Internal HTML pages with 2xx response', 'CONTENT', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Custom Extraction', None, 'NA', 'NA', 'NA', False),
    (154, 'All', 'Internal HTML pages with 2xx response', 'CONTENT', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Custom JavaScript', None, 'NA', 'NA', 'NA', False),
    (186, 'All', 'Internal HTML pages with 2xx response', 'CONTENT', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (444, 'Analytics', None, 'Analytics', 'NA', 'NA', False),
    (445, 'All', 'Internal URLs with associated Google Analytics data', 'Analytics', 'NA', 'NA', True),
    (446, 'Sessions Above 0', 'Internal URLs with associated Google Analytics data', 'Analytics', 'NA', 'NA', True),
    (447, 'Bounce Rate Above 70%', 'Internal URLs with associated Google Analytics data', 'Analytics', 'Warning', 'Low', True),
    (448, 'No GA Data', 'Internal URLs with associated Google Analytics data', 'Analytics', 'Warning', 'Low', True),
    (449, 'Non-Indexable with GA Data', 'Internal URLs with associated Google Analytics data', 'Analytics', 'Warning', 'Low', True),
    ('NA', 'Orphan URLs', 'Internal URLs with associated Google Analytics data', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (450, 'Search Console', None, 'Search Console', 'NA', 'NA', False),
    (451, 'All', 'Internal URLs with associated Google Search Console data', 'Search Console', 'NA', 'NA', True),
    (452, 'Clicks Above 0', 'Internal URLs with associated Google Search Console data', 'Search Console', 'NA', 'NA', True),
    (458, 'No Search Analytics Data', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Warning', 'Low', True),
    (459, 'Non-Indexable with Search Analytics Data', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Warning', 'Low', True),
    ('NA', 'Orphan URLs', 'Internal URLs with associated Google Search Console data', 'NA', 'NA', 'NA', True),
    ('NA', 'URL is Not on Google', 'Internal URLs with associated Google Search Console data', 'NA', 'NA', 'NA', True),
    (456, 'Indexable URL Not Indexed', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Warning', 'High', True),
    (454, 'URL is on Google But Has Issues', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Issue', 'Medium', True),
    (457, 'User-Declared Canonical Not Selected', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Warning', 'Medium', True),
    (453, 'Page is Not Mobile Friendly', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Issue', 'High', True),
    ('NA', 'AMP URL Invalid', 'Internal URLs with associated Google Search Console data', 'NA', 'NA', 'NA', True),
    (455, 'Rich Result Invalid', 'Internal URLs with associated Google Search Console data', 'Search Console', 'Issue', 'Medium', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    (460, 'Validation', None, 'Validation', 'NA', 'NA', False),
    (461, 'All', 'Internal URLs with 2xx response', 'Validation', 'NA', 'NA', True),
    (468, 'Invalid HTML Elements in <head>', 'Internal URLs with 2xx response', 'Validation', 'Warning', 'High', True),
    (469, '<body> Element Preceding <html>', 'Internal URLs with 2xx response', 'Validation', 'Warning', 'High', True),
    (470, '<head> Not First In <html> Element', 'Internal URLs with 2xx response', 'Validation', 'Warning', 'High', True),
    (462, 'Missing <head> Tag', 'Internal URLs with 2xx response', 'Validation', 'Issue', 'High', True),
    (463, 'Multiple <head> Tags', 'Internal URLs with 2xx response', 'Validation', 'Issue', 'High', True),
    (464, 'Missing <body> Tag', 'Internal URLs with 2xx response', 'Validation', 'Issue', 'High', True),
    (465, 'Multiple <body> Tags', 'Internal URLs with 2xx response', 'Validation', 'Issue', 'High', True),
    (466, 'HTML Document Over 15MB', 'Internal URLs with 2xx response', 'Validation', 'Issue', 'High', True),
    (467, 'Resource Over 15MB', 'Internal URLs with 2xx response', 'Validation', 'Issue', 'High', True),
    (473, 'High Carbon Rating', 'Internal URLs with 2xx response', 'Validation', 'Opportunity', 'Low', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Link Metrics', None, 'NA', 'NA', 'NA', False),
    ('NA', 'All', 'Internal HTML/PDF pages', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'AI', None, 'NA', 'NA', 'NA', False),
    ('NA', 'All', 'Internal HTML URLs with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Depth (Clicks from Start URL)', None, 'NA', 'NA', 'NA', False),
    ('NA', '0', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '1', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '2', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '3', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '4', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '5', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '6', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '7', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '8', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '9', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', '10+', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', None, None, 'NA', 'NA', 'NA', False),
    ('NA', 'Inlinks (Top 20 URLs)', None, 'NA', 'NA', 'NA', False),
    ('NA', None, 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', 'Response Time (Seconds)', 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    ('NA', None, 'Internal HTML pages with 2xx response', 'NA', 'NA', 'NA', True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, 'Internal HTML pages with 2xx response', None, None, None, True),
    (None, None, None, None, None, None, False),
    (None, None, None, None, None, None, False),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, True),
    (None, None, None, None, None, None, False),
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


def _filename_variants(name):
    """Return candidate filenames to try: the name as given, plus the
    'internal_' prefixed/unprefixed counterpart. Screaming Frog's
    --bulk-export CLI writes several Inlinks CSVs with an 'internal_'
    prefix (e.g. internal_client_error_(4xx)_inlinks.csv), while some
    manual GUI exports and other CSVs omit it (client_error_(4xx)_inlinks.csv).
    Checking both avoids silently-empty issue sets when only one variant
    exists on disk."""
    variants = [name]
    if name.startswith("internal_"):
        alt = name[len("internal_"):]
    else:
        alt = "internal_" + name
    if alt not in variants:
        variants.append(alt)
    return variants


def load_url_set(report_path, csv_files):
    """Load set of URLs from one or more CSVs (Address column).
    Tries both 'internal_' prefixed and unprefixed filename variants
    for each entry, since Screaming Frog's export mechanisms are
    inconsistent about the prefix depending on CLI export tab vs
    bulk-export vs manual GUI export."""
    urls = set()
    for csv_name in csv_files:
        for candidate in _filename_variants(csv_name):
            p = os.path.join(report_path, candidate)
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


def parse_crawl_overview_csv(csv_path):
    """Parse crawl_overview.csv into (header, data).

    header: dict with 'Site Crawled', 'Date', 'Time' from the top of the file.
    data: dict of {section_name: {row_label: (count, pct_fraction, total)}}

    The file's structure is irregular in a few ways that this function
    handles explicitly:
    - Most sections start with a bare single-cell title line (e.g. "Security")
      followed immediately by real data rows.
    - The "Summary" section instead starts with a 5-column column-header
      row ("Summary","URLs","% of Total","Total URLs","Total URLs Description")
      whose second cell is text, not a number - this is a label row, not data,
      and is discarded.
    - "URLs Displayed" is a section whose own title row IS simultaneously a
      real data row (its second cell is a real count) - this row must be
      kept as data, not discarded.
    """
    with open(csv_path, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))

    header = {}
    sections = {}
    current_section = None
    current_rows = []
    header_done = False

    for row in rows:
        if len(row) == 0 or (len(row) == 1 and row[0] == ""):
            if current_section is not None:
                sections[current_section] = current_rows
                current_section = None
                current_rows = []
            continue
        if not header_done and len(row) == 2 and row[0] in ("Site Crawled", "Date", "Time"):
            header[row[0]] = row[1]
            continue
        header_done = True
        if current_section is None:
            current_section = row[0]
            current_rows = []
            if len(row) >= 4:
                try:
                    int(row[1])
                    current_rows.append(row)  # self-titled data row (e.g. "URLs Displayed")
                except (ValueError, TypeError):
                    pass  # pure column-header row (e.g. "Summary", "Response Time (Seconds)")
            continue
        current_rows.append(row)
    if current_section is not None:
        sections[current_section] = current_rows

    data = {}
    for section, rws in sections.items():
        lookup = {}
        for r in rws:
            if len(r) >= 4:
                label, count, pct, total = r[0], r[1], r[2], r[3]
                try:
                    count_v = int(count)
                except Exception:
                    count_v = count
                try:
                    pct_v = float(pct.strip("%")) / 100.0
                except Exception:
                    pct_v = None
                try:
                    total_v = int(total)
                except Exception:
                    total_v = total
                lookup[label] = (count_v, pct_v, total_v)
        data[section] = lookup
    return header, data


def _fmt_elapsed(started, completed):
    if not started or not completed:
        return None
    delta = completed - started
    total_seconds = int(delta.total_seconds())
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return "{:02d}:{:02d}:{:02d}".format(h, m, s)


def build_score_weights_sheet(wb, f_hdr, f_cell):
    """Build the hidden Score Weights lookup sheet from scratch.
    Referenced by the Crawl Overview sheet's CatWeight/SevMult/PriMult
    VLOOKUP formulas."""
    ws = wb.add_worksheet("Score Weights")
    ws.hide()
    ws.set_column("A:A", 32)
    ws.set_column("B:B", 10)
    ws.set_column("D:D", 16)
    ws.set_column("E:E", 12)
    ws.set_column("G:G", 12)
    ws.set_column("H:H", 12)

    def write_table(start_col_idx, title_a, title_b, rows):
        ws.write(0, start_col_idx, title_a, f_hdr)
        ws.write(0, start_col_idx + 1, title_b, f_hdr)
        for i, (k, v) in enumerate(rows):
            ws.write(1 + i, start_col_idx, k, f_cell)
            ws.write(1 + i, start_col_idx + 1, v, f_cell)

    write_table(0, "Category", "Weight", CATEGORY_WEIGHTS)
    write_table(3, "Severity", "Multiplier", SEVERITY_MULT)
    write_table(6, "Priority", "Multiplier", PRIORITY_MULT)


def build_crawl_overview_sheet(wb, formats, domain, started_at, completed_at, report_path):
    """Build the Crawl Overview tab from scratch via xlsxwriter, using
    CRAWL_OVERVIEW_ROWS (the fixed label/category/severity/priority
    structure) and real values parsed from crawl_overview.csv (Screaming
    Frog's native --save-report "Crawl Overview" export).

    Every row gets the same live Excel formulas (Spread/CatWeight/
    SevMult/PriMult/Penalty) regardless of row type; title and
    administrative rows have Severity/Priority = "NA" so the formulas
    naturally resolve to 0 for them via IF() shortcuts, matching the
    original template's exhaustive formula-everywhere design.
    """
    f_title, f_hdr, f_cell, f_pct = formats

    ws = wb.add_worksheet("Crawl Overview")
    ws.set_column("A:A", 8)
    ws.set_column("B:B", 40)
    ws.set_column("C:E", 12)
    ws.set_column("F:F", 45)
    ws.set_column("G:I", 14)
    ws.set_column("J:K", 12)
    ws.set_column("L:P", 12)
    ws.set_column("Q:Q", 20)

    ws.write(0, 1, "Screaming Frog Crawl Overview Report - Issues Scoring", f_title)

    headers = ["Sr. no.", "Site Crawled", None, None, None, None, "Category",
               "Severity", "Priority", "Affected URLs", "Total URLs", "Spread",
               "CatWeight", "SevMult", "PriMult", "Penalty (Issue Score)", "Comment"]
    for ci, h in enumerate(headers):
        if h is not None:
            ws.write(1, ci, h, f_hdr)
    ws.write(1, 2, domain, f_cell)

    csv_overview_path = os.path.join(report_path, "crawl_overview.csv")
    csv_header, csv_data = ({}, {})
    if os.path.exists(csv_overview_path):
        csv_header, csv_data = parse_crawl_overview_csv(csv_overview_path)

    # Rows 3-9 (0-indexed 2-8): metadata, no CSV section equivalent
    meta_rows = {
        2: started_at.strftime("%Y-%m-%d") if started_at else None,   # Start Date
        3: started_at.strftime("%H:%M:%S") if started_at else None,  # Start Time
        4: None,                                                     # Last Modified Date
        5: None,                                                     # Last Modified Time
        6: _fmt_elapsed(started_at, completed_at),                   # Elapsed
        7: csv_header.get("Date"),                                   # Report Date
        8: csv_header.get("Time"),                                   # Report Time
    }

    section_names_lower = {s.lower(): s for s in csv_data.keys()}
    current_section = None

    for idx, (a, b, fdesc, g, h, i, is_data) in enumerate(CRAWL_OVERVIEW_ROWS):
        r = idx + 2  # CRAWL_OVERVIEW_ROWS[0] corresponds to sheet row 3 (0-indexed row 2)

        count = pct = total = None
        if r in meta_rows:
            count = meta_rows[r]
        elif b:
            if not is_data:
                current_section = section_names_lower.get(str(b).strip().lower())
            else:
                maybe_new_section = section_names_lower.get(str(b).strip().lower())
                if maybe_new_section:
                    current_section = maybe_new_section
                hit = csv_data.get(current_section, {}).get(b) if current_section else None
                if hit:
                    count, pct, total = hit

        if a is not None:
            ws.write(r, 0, a, f_cell)
        if b is not None:
            ws.write(r, 1, b, f_cell)
        if count is not None:
            ws.write(r, 2, count, f_cell)
        if pct is not None:
            ws.write(r, 3, pct, f_pct)
        if total is not None:
            ws.write(r, 4, total, f_cell)
        if fdesc is not None:
            ws.write(r, 5, fdesc, f_cell)
        # Default to "NA" (matching the sheet's own convention everywhere
        # else) rather than leaving truly blank: IF(G="NA",...) is a text
        # comparison, and a genuinely blank cell does not equal "NA", so
        # a handful of rows near the end of the sheet (a Screaming Frog
        # export quirk, not present as literal "NA" text like every other
        # unscored row) would otherwise fall through to VLOOKUP on an
        # empty value and raise #N/A.
        ws.write(r, 6, g if g is not None else "NA", f_cell)
        ws.write(r, 7, h if h is not None else "NA", f_cell)
        ws.write(r, 8, i if i is not None else "NA", f_cell)

        rn = r + 1  # 1-indexed row number for formula strings
        ws.write_formula(r, 9, '=IF(H{0}="NA","-",C{0})'.format(rn), f_cell)
        ws.write_formula(r, 10, '=IF(H{0}="NA","-",E{0})'.format(rn), f_cell)
        ws.write_formula(r, 11, '=IF(AND(ISNUMBER(J{0}),J{0}>0),J{0}/E{0},0)'.format(rn), f_cell)
        ws.write_formula(r, 12, '=IF(G{0}="NA",0,VLOOKUP(G{0},\'Score Weights\'!$A$2:$B$28,2,FALSE))'.format(rn), f_cell)
        ws.write_formula(r, 13, '=IF(H{0}="NA",0,VLOOKUP(H{0},\'Score Weights\'!$D$2:$E$4,2,FALSE))'.format(rn), f_cell)
        ws.write_formula(r, 14, '=IF(I{0}="NA",0,VLOOKUP(I{0},\'Score Weights\'!$G$2:$H$4,2,FALSE))'.format(rn), f_cell)
        ws.write_formula(r, 15, '=L{0}*M{0}*N{0}*O{0}*100'.format(rn), f_cell)


def build_overview_report_masterfile(crawl_id: str, domain: str, report_path: str,
                                      started_at=None, completed_at=None) -> bytes:
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

    # ---- Build entire workbook via xlsxwriter (no template file) ----
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True, "nan_inf_to_errors": True, "strings_to_urls": False})

    def f(**kw):
        return wb.add_format(kw)

    f_title = f(bold=True, font_name=FONT, font_size=12, font_color=BLACK)
    f_red_hdr = f(bold=True, font_name=FONT, font_size=8, font_color=WHITE,
                  bg_color=RED, border=1, align="center", valign="vcenter", text_wrap=True)
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
    f_caption = f(italic=True, font_name=FONT, font_size=9, font_color="595959")

    # -- Sheet 1: Dashboard --
    ws1 = wb.add_worksheet("Dashboard -Dashboard")
    ws1.set_column("A:A", 25)
    ws1.set_column("B:B", 35)
    ws1.set_column("C:C", 12)
    ws1.set_column("D:D", 15)
    for i in range(len(theme_names)):
        ws1.set_column(4 + i, 4 + i, 15)

    ws1.write(0, 0, "Summary: Consolidated Dashboard of all the issues identified.", f_caption)
    ws1.set_row(6, 42)

    ws1.write(7, 0, "Issue Category", f_black_lft)
    ws1.write(7, 1, "Issue", f_red_hdr)
    ws1.write(7, 2, "Severity", f_black_hdr)
    ws1.write(7, 3, "Priority", f_black_hdr)
    for ti, (theme, tpri) in enumerate(sorted_themes):
        ws1.write(7, 4 + ti, theme, f_black_hdr)

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

    # -- Sheet 2: Detailed Data --
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

    ws2.write(0, 0, "Summary: Consolidated Dashboard of all the issues identified per URL", f_caption)

    col_offset = 11
    for cat_name, count in CATEGORY_GROUPS:
        if count > 1:
            ws2.merge_range(2, col_offset, 2, col_offset + count - 1, cat_name, f_dark_hdr)
        else:
            ws2.write(2, col_offset, cat_name, f_dark_hdr)
        col_offset += count

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

    # -- Sheet 3: Crawl Overview + hidden Score Weights lookup --
    co_f_title = f(bold=True, font_name=FONT, font_size=12, font_color=BLACK)
    co_f_hdr = f(bold=True, font_name=FONT, font_size=9, font_color=BLACK,
                 bg_color="#D9D9D9", border=1, align="center", valign="vcenter")
    co_f_cell = f(font_name=FONT, font_size=9, font_color=BLACK, align="left", valign="vcenter")
    co_f_pct = f(font_name=FONT, font_size=9, font_color=BLACK, align="left", valign="vcenter", num_format="0.00%")

    build_score_weights_sheet(wb, co_f_hdr, co_f_cell)
    build_crawl_overview_sheet(wb, (co_f_title, co_f_hdr, co_f_cell, co_f_pct), domain, started_at, completed_at, report_path)

    wb.worksheets_objs = [
        wb.worksheets_objs[0],  # Dashboard -Dashboard
        wb.worksheets_objs[1],  # Dashboard - Detailed Data
        wb.worksheets_objs[3],  # Crawl Overview
        wb.worksheets_objs[2],  # Score Weights (hidden)
    ]

    wb.close()
    buf.seek(0)
    return buf.read()