import subprocess
import os
import zipfile
from datetime import datetime, timezone
from sqlalchemy.orm import Session
from app.config import settings
from app.database import SessionLocal
from app.models.crawl import Crawl
from app.tasks.crawl_runner import celery_app, BULK_EXPORTS, count_crawled_pages


def get_db() -> Session:
    return SessionLocal()


# Same tab list as crawl_runner.run_crawl, kept in one place would be ideal
# but crawl_runner's is defined inline in that function; duplicated here
# deliberately to avoid a risky refactor of the working crawl path.
EXPORT_TABS = (
    "Internal:All,Response Codes:All,Response Codes:Internal Redirection (3xx),"
    "Response Codes:Internal Redirect Chain,Response Codes:Internal Redirect Loop,"
    "Response Codes:Internal Client Error (4xx),Response Codes:Internal Server Error (5xx),"
    "Response Codes:Internal No Response,Response Codes:Internal Blocked by Robots.txt,"
    "URL:All,URL:Uppercase,URL:Underscores,URL:Parameters,URL:Multiple Slashes,"
    "URL:Repetitive Path,URL:Contains Space,Page Titles:All,Page Titles:Missing,"
    "Page Titles:Duplicate,Page Titles:Over X Pixels,Page Titles:Below X Pixels,"
    "Page Titles:Same as H1,Page Titles:Multiple,Page Titles:Outside <head>,"
    "Meta Description:All,Meta Description:Missing,Meta Description:Duplicate,"
    "Meta Description:Over X Pixels,Meta Description:Below X Pixels,Meta Description:Multiple,"
    "Meta Description:Outside <head>,H1:All,H1:Missing,H1:Duplicate,H1:Over X Characters,"
    "H1:Multiple,Canonicals:All,Canonicals:Canonicalised,Canonicals:Missing,"
    "Canonicals:Multiple,Canonicals:Non-Indexable Canonical,Canonicals:Multiple Conflicting,"
    "Canonicals:Canonical Is Relative,Canonicals:Unlinked,Canonicals:Outside <head>,"
    "Directives:All,Directives:Noindex,Directives:Nofollow,Sitemaps:All,"
    "Sitemaps:URLs not in Sitemap,Sitemaps:Orphan URLs,Sitemaps:Non-Indexable URLs in Sitemap,"
    "Sitemaps:XML Sitemap with over 50k URLs,Security:All,Security:HTTP URLs,"
    "Security:Mixed Content,Security:Bad Content Type,Security:Missing Secure Referrer-Policy Header,"
    "Security:Missing Content-Security-Policy Header,Security:Missing HSTS Header,"
    "Security:Missing X-Content-Type-Options Header,Security:Missing X-Frame-Options Header,"
    "Security:Form on HTTP URL,Hreflang:All,Hreflang:Non-200 hreflang URLs,"
    "Hreflang:Unlinked hreflang URLs,Hreflang:Missing Return Links,"
    "Hreflang:Inconsistent Language & Region Return Links,Hreflang:Non-Canonical Return Links,"
    "Hreflang:Noindex Return Links,Hreflang:Incorrect Language & Region Codes,"
    "Hreflang:Multiple Entries,Hreflang:Missing Self Reference,Hreflang:Not Using Canonical,"
    "Hreflang:Missing X-Default,Hreflang:Missing,Hreflang:Outside <head>,Structured Data:All,"
    "Structured Data:Missing,Structured Data:Validation Errors,Structured Data:Validation Warnings,"
    "Structured Data:Parse Errors,Pagination:All,Pagination:Pagination URL Not in Anchor Tag,"
    "Pagination:Unlinked Pagination URLs,Pagination:Non-Indexable,Pagination:Multiple Pagination URLs,"
    "Pagination:Pagination Loop,Pagination:Sequence Error,Content:All,Content:Exact Duplicates,"
    "Content:Near Duplicates,Content:Low Content Pages,Content:Soft 404 Pages,"
    "Content:Spelling Errors,Content:Grammar Errors,Content:Lorem Ipsum Placeholder,"
    "Custom Extraction:All,Search Console:All,Analytics:All"
)


@celery_app.task(bind=True, time_limit=None, soft_time_limit=None)
def run_import_job(self, crawl_id: str, dbseospider_path: str):
    db = get_db()
    try:
        crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
        if not crawl:
            return

        crawl.status = "running"
        db.commit()

        os.makedirs(settings.CRAWL_OUTPUT_DIR, exist_ok=True)
        output_dir = os.path.join(settings.CRAWL_OUTPUT_DIR, crawl_id)
        os.makedirs(output_dir, exist_ok=True)  # --load-crawl needs this to pre-exist

        cmd = [
            settings.SCREAMING_FROG_CLI,
            "--headless",
            "--load-crawl", dbseospider_path,
            "--output-folder", output_dir,
            "--overwrite",
            "--export-tabs", EXPORT_TABS,
            "--bulk-export", BULK_EXPORTS,
            "--save-report", "Crawl Overview",
        ]

        log_file = open(os.path.join(output_dir, "crawl.log"), "w")
        process = subprocess.Popen(
            cmd, stdout=log_file, stderr=log_file, text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        process.wait()
        exit_code = process.returncode
        log_file.close()

        if exit_code == 0:
            crawl.status = "completed"
            crawl.report_path = output_dir
            crawl.pages_crawled = count_crawled_pages(output_dir)
            crawl.completed_at = datetime.now(timezone.utc)
            try:
                zip_path = os.path.join(output_dir, "raw_files.zip")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname in os.listdir(output_dir):
                        if fname == "raw_files.zip":
                            continue
                        fpath = os.path.join(output_dir, fname)
                        if os.path.isfile(fpath):
                            zf.write(fpath, arcname=fname)
            except Exception:
                pass
        else:
            crawl.status = "failed"
            try:
                with open(os.path.join(output_dir, "crawl.log"), "r") as f:
                    error_detail = f.read()[-500:]
            except Exception:
                error_detail = f"Screaming Frog exited with code {exit_code}"
            crawl.error_message = error_detail
            crawl.completed_at = datetime.now(timezone.utc)

        db.commit()

        # Clean up the uploaded source file once exported successfully, to
        # keep disk usage bounded; keep it around on failure for diagnosis
        if exit_code == 0 and os.path.exists(dbseospider_path):
            try:
                os.remove(dbseospider_path)
            except OSError:
                pass

    except Exception as e:
        try:
            crawl.status = "failed"
            crawl.error_message = str(e)[:500]
            crawl.completed_at = datetime.now(timezone.utc)
            db.commit()
        except Exception:
            db.rollback()
    finally:
        db.close()
