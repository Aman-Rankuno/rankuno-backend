import subprocess
import os
import zipfile
import time
from datetime import datetime, timezone
from celery import Celery
from sqlalchemy.orm import Session
from app.config import settings
from app.database import SessionLocal
from app.models.crawl import Crawl

celery_app = Celery(
    "rankuno",
    broker=settings.REDIS_URL,
    backend=settings.REDIS_URL,
)

celery_app.conf.task_time_limit = None
celery_app.conf.task_soft_time_limit = None


def get_db() -> Session:
    return SessionLocal()


BULK_EXPORTS = ",".join([
    "Response Codes:Internal:Internal Success (2xx) Inlinks",
    "Response Codes:Internal:Internal Client Error (4xx) Inlinks",
    "Response Codes:Internal:Internal Server Error (5xx) Inlinks",
    "Response Codes:Internal:Internal Redirection (3xx) Inlinks",
    "Response Codes:Internal:Internal No Response Inlinks",
    "Response Codes:Internal:Internal Blocked by Robots.txt Inlinks",
    "Response Codes:Internal:Internal Blocked Resource Inlinks",
    "Canonicals:Canonicalised Inlinks",
    "Canonicals:Non-Indexable Canonical Inlinks",
    "Security:HTTP URLs Inlinks",
])


@celery_app.task(bind=True, time_limit=None, soft_time_limit=None)
def run_crawl(self, crawl_id: str, domain: str, crawl_type: str, urls: str = None, config_file: str = None, gsc_account: str = None, gsc_property: str = None, ga_account: str = None, ga4_account: str = None, ga4_property: str = None, ga4_stream: str = None, include_patterns: str = None, exclude_patterns: str = None):
    db = get_db()
    try:
        crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
        if not crawl:
            return

        crawl.status = "running"
        db.commit()

        os.makedirs(settings.CRAWL_OUTPUT_DIR, exist_ok=True)
        output_dir = os.path.join(settings.CRAWL_OUTPUT_DIR, crawl_id)
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(output_dir):
            raise Exception(f"Failed to create output directory: {output_dir}")

        time.sleep(1)

        cmd = [
            settings.SCREAMING_FROG_CLI,
            "--headless",
            "--output-folder", output_dir,
            "--overwrite",
            "--export-tabs", "Internal:All,Response Codes:All,Response Codes:Internal Redirection (3xx),Response Codes:Internal Redirect Chain,Response Codes:Internal Redirect Loop,Response Codes:Internal Client Error (4xx),Response Codes:Internal Server Error (5xx),Response Codes:Internal No Response,Response Codes:Internal Blocked by Robots.txt,URL:All,URL:Uppercase,URL:Underscores,URL:Parameters,URL:Multiple Slashes,URL:Repetitive Path,URL:Contains Space,Page Titles:All,Page Titles:Missing,Page Titles:Duplicate,Page Titles:Over X Pixels,Page Titles:Below X Pixels,Page Titles:Same as H1,Page Titles:Multiple,Page Titles:Outside <head>,Meta Description:All,Meta Description:Missing,Meta Description:Duplicate,Meta Description:Over X Pixels,Meta Description:Below X Pixels,Meta Description:Multiple,Meta Description:Outside <head>,H1:All,H1:Missing,H1:Duplicate,H1:Over X Characters,H1:Multiple,Canonicals:All,Canonicals:Canonicalised,Canonicals:Missing,Canonicals:Multiple,Canonicals:Non-Indexable Canonical,Canonicals:Multiple Conflicting,Canonicals:Canonical Is Relative,Canonicals:Unlinked,Canonicals:Outside <head>,Directives:All,Directives:Noindex,Directives:Nofollow,Sitemaps:All,Sitemaps:URLs not in Sitemap,Sitemaps:Orphan URLs,Sitemaps:Non-Indexable URLs in Sitemap,Sitemaps:XML Sitemap with over 50k URLs,Security:All,Security:HTTP URLs,Security:Mixed Content,Hreflang:All,Hreflang:Non-200 hreflang URLs,Hreflang:Unlinked hreflang URLs,Hreflang:Missing Return Links,Hreflang:Inconsistent Language & Region Return Links,Hreflang:Non-Canonical Return Links,Hreflang:Noindex Return Links,Hreflang:Incorrect Language & Region Codes,Hreflang:Multiple Entries,Hreflang:Missing Self Reference,Hreflang:Not Using Canonical,Hreflang:Missing X-Default,Hreflang:Missing,Hreflang:Outside <head>,Structured Data:All,Structured Data:Missing,Structured Data:Validation Errors,Structured Data:Validation Warnings,Structured Data:Parse Errors,Pagination:All,Pagination:Pagination URL Not in Anchor Tag,Pagination:Unlinked Pagination URLs,Pagination:Non-Indexable,Pagination:Multiple Pagination URLs,Pagination:Pagination Loop,Pagination:Sequence Error,Content:All,Content:Exact Duplicates,Content:Near Duplicates,Content:Low Content Pages,Content:Soft 404 Pages,Content:Spelling Errors,Content:Grammar Errors,Content:Lorem Ipsum Placeholder,Custom Extraction:All,Search Console:All,Analytics:All",
            "--bulk-export", BULK_EXPORTS,
            "--save-report", "Crawl Overview",
            "--save-crawl",
        ]

        if config_file:
            config_path = os.path.join(settings.CRAWL_CONFIGS_DIR, config_file)
            if os.path.exists(config_path):
                cmd += ["--config", config_path]
        if gsc_account and gsc_property:
            cmd += ["--use-google-search-console", gsc_account, gsc_property]


        if crawl_type in ("full-site", "full-audit", "advanced-audit", "js-crawl", "orphan-pages", "sitemap-generator"):
            crawl_url = domain if domain.startswith("http") else f"https://{domain}"
            cmd += ["--crawl", crawl_url]
        else:
            url_file = os.path.join(output_dir, "urls.txt")
            with open(url_file, "w") as f:
                f.write(urls or "")
            cmd += ["--crawl-list", url_file]

        log_file = open(os.path.join(output_dir, "crawl.log"), "w")

        process = subprocess.Popen(
            cmd,
            stdout=log_file,
            stderr=log_file,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

        while process.poll() is None:
            time.sleep(30)
            try:
                db.refresh(crawl)
            except Exception:
                db.rollback()

        exit_code = process.wait()
        log_file.close()

        if exit_code == 0:
            crawl.status = "completed"
            crawl.report_path = output_dir
            crawl.pages_crawled = count_crawled_pages(output_dir)
            crawl.completed_at = datetime.now(timezone.utc)
            # Pre-build raw files zip for instant download
            try:
                zip_path = os.path.join(output_dir, "raw_files.zip")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname in os.listdir(output_dir):
                        if fname == "raw_files.zip":
                            continue
                        fpath = os.path.join(output_dir, fname)
                        if os.path.isfile(fpath):
                            zf.write(fpath, arcname=fname)
            except Exception as zip_err:
                pass  # Non-fatal: zip failed but crawl succeeded
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


def count_crawled_pages(output_dir: str) -> int:
    try:
        for filename in os.listdir(output_dir):
            if filename.endswith(".csv") and "internal" in filename.lower():
                filepath = os.path.join(output_dir, filename)
                with open(filepath, "r", encoding="utf-8") as f:
                    return max(0, sum(1 for _ in f) - 1)
    except Exception:
        pass
    return 0
