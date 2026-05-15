from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional
from app.database import get_db
from app.models.crawl import Crawl
from app.tasks.crawl_runner import run_crawl
import zipfile
import io
import os

router = APIRouter()

class CrawlCreate(BaseModel):
    domain: str
    crawl_type: str
    urls: Optional[str] = None
    config_file: Optional[str] = None
    gsc_account: Optional[str] = None
    gsc_property: Optional[str] = None
    ga_account: Optional[str] = None
    ga4_account: Optional[str] = None
    ga4_property: Optional[str] = None
    ga4_stream: Optional[str] = None
    include_patterns: Optional[str] = None
    exclude_patterns: Optional[str] = None

class CrawlResponse(BaseModel):
    id: str
    domain: str
    crawl_type: str
    status: str
    pages_crawled: int
    report_path: Optional[str]
    error_message: Optional[str]
    created_at: str
    completed_at: Optional[str]

    class Config:
        from_attributes = True

@router.post("/")
def create_crawl(payload: CrawlCreate, db: Session = Depends(get_db)):
    crawl = Crawl(
        domain=payload.domain,
        crawl_type=payload.crawl_type,
        status="queued",
    )
    db.add(crawl)
    db.commit()
    db.refresh(crawl)
    run_crawl.delay(
        crawl.id,
        payload.domain,
        payload.crawl_type,
        payload.urls,
        payload.config_file,
        payload.gsc_account,
        payload.gsc_property,
        payload.ga_account,
        payload.ga4_account,
        payload.ga4_property,
        payload.ga4_stream,
    )
    return {"id": crawl.id, "status": "queued", "message": "Crawl queued successfully"}

@router.get("/")
def list_crawls(db: Session = Depends(get_db)):
    crawls = db.query(Crawl).order_by(Crawl.created_at.desc()).all()
    return crawls

@router.get("/{crawl_id}")
def get_crawl(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    return crawl

@router.get("/{crawl_id}/download/zip")
def download_zip(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename in os.listdir(crawl.report_path):
            filepath = os.path.join(crawl.report_path, filename)
            if os.path.isfile(filepath):
                zf.write(filepath, arcname=filename)
    zip_buffer.seek(0)
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    zip_filename = f"{domain_safe}_crawl.zip"
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={zip_filename}"},
    )