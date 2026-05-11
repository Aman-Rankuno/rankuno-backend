import subprocess
import os
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

# No task time limit, crawls can run as long as needed
celery_app.conf.task_time_limit = None
celery_app.conf.task_soft_time_limit = None


def get_db() -> Session:
    return SessionLocal()


@celery_app.task(bind=True, time_limit=None, soft_time_limit=None)
def run_crawl(self, crawl_id: str, domain: str, crawl_type: str, urls: str = None, config_file: str = None):
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
        
        # Verify folder exists before launching Screaming Frog
        if not os.path.exists(output_dir):
            raise Exception(f"Failed to create output directory: {output_dir}")
        
        # Small delay to ensure filesystem sync
        time.sleep(1)

        cmd = [
            settings.SCREAMING_FROG_CLI,
            "--headless",
            "--output-folder", output_dir,
            "--overwrite",
            "--export-tabs", "Internal:All",
            "--save-crawl",
        ]
        
        if config_file:
            config_path = os.path.join(settings.CRAWL_CONFIGS_DIR, config_file)
            if os.path.exists(config_path):
                cmd += ["--config", config_path]

        if crawl_type in ("full-site", "full-audit", "advanced-audit", "js-crawl", "orphan-pages", "sitemap-generator"):
            crawl_url = domain if domain.startswith("http") else f"https://{domain}"
            cmd += ["--crawl", crawl_url]
        else:
            url_file = os.path.join(output_dir, "urls.txt")
            with open(url_file, "w") as f:
                f.write(urls or "")
            cmd += ["--crawl-list", url_file]

        # Use Popen so the process runs freely with no timeout
        log_file = open(os.path.join(output_dir, "crawl.log"), "w")
        
        process = subprocess.Popen(
            cmd,
            stdout=log_file,
            stderr=log_file,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

        # Poll every 30 seconds, update DB so frontend knows it's still running
        while process.poll() is None:
            time.sleep(30)
            # Keep the DB session alive and status as running
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
        else:
            crawl.status = "failed"
            try:
                with open(os.path.join(output_dir, "crawl.log"), "r") as f:
                    error_detail = f.read()[-500:]
            except Exception:
                error_detail = f"Screaming Frog exited with code {exit_code}"
            crawl.error_message = error_detail
            crawl.completed_at = datetime.now(timezone.utc)

        if process.returncode == 0:
            crawl.status = "completed"
            crawl.report_path = output_dir
            crawl.pages_crawled = count_crawled_pages(output_dir)
            crawl.completed_at = datetime.now(timezone.utc)
        else:
            crawl.status = "failed"
            error_detail = ""
            if stderr:
                error_detail = stderr[:500]
            elif stdout:
                error_detail = stdout[:500]
            else:
                error_detail = f"Screaming Frog exited with code {process.returncode}"
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