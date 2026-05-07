import subprocess
import os
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


def get_db() -> Session:
    return SessionLocal()


@celery_app.task(bind=True)
def run_crawl(self, crawl_id: str, domain: str, crawl_type: str, urls: str = None):
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

        os.makedirs(output_dir, exist_ok=True)

        cmd = [
            settings.SCREAMING_FROG_CLI,
            "--headless",
            "--output-folder", output_dir,
            "--overwrite",
            "--export-tabs", "Internal:All",
            "--save-crawl",
        ]

        if crawl_type in ("full-site", "full-audit", "advanced-audit", "js-crawl", "orphan-pages", "sitemap-generator"):
            crawl_url = domain if domain.startswith("http") else f"https://{domain}"
            cmd += ["--crawl", crawl_url]
        else:
            url_file = os.path.join(output_dir, "urls.txt")
            with open(url_file, "w") as f:
                f.write(urls or "")
            cmd += ["--crawl-list", url_file]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=3600,
        )

        if result.returncode == 0:
            crawl.status = "completed"
            crawl.report_path = output_dir
            crawl.pages_crawled = count_crawled_pages(output_dir)
            crawl.completed_at = datetime.now(timezone.utc)
        else:
            crawl.status = "failed"
            error_detail = ""
            if result.stderr:
                error_detail = result.stderr[:500]
            elif result.stdout:
                error_detail = result.stdout[:500]
            else:
                error_detail = f"Screaming Frog exited with code {result.returncode}"
            crawl.error_message = error_detail
            crawl.completed_at = datetime.now(timezone.utc)

        db.commit()

    except subprocess.TimeoutExpired:
        crawl.status = "failed"
        crawl.error_message = "Crawl timed out after 60 minutes"
        crawl.completed_at = datetime.now(timezone.utc)
        db.commit()
    except Exception as e:
        crawl.status = "failed"
        crawl.error_message = str(e)[:500]
        crawl.completed_at = datetime.now(timezone.utc)
        db.commit()
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