from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional
from app.database import get_db
from app.models.crawl import Crawl
from app.tasks.crawl_runner import run_crawl

router = APIRouter()


class CrawlCreate(BaseModel):
    domain: str
    crawl_type: str
    urls: Optional[str] = None
    config_file: Optional[str] = None


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

    run_crawl.delay(crawl.id, payload.domain, payload.crawl_type, payload.urls, payload.config_file)

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