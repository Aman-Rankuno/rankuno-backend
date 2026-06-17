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
from app.services.masterfile_response_codes import build_response_codes_masterfile
from app.services.masterfile_url_issues import build_url_issues_masterfile
from app.services.masterfile_page_titles import build_page_titles_masterfile
from app.services.masterfile_meta_description import build_meta_description_masterfile
from app.services.masterfile_h1 import generate as generate_h1_masterfile
from app.services.masterfile_directives import generate as generate_directives_masterfile
from app.services.masterfile_sitemaps import generate as generate_sitemaps_masterfile
from app.services.masterfile_content_issues import generate as generate_content_issues_masterfile
from app.services.masterfile_duplicate_content import generate as generate_duplicate_content_masterfile
from app.services.masterfile_custom_search_ga4_gtm import generate as generate_custom_search_ga4_gtm_masterfile
from app.services.masterfile_custom_search_og_twitter import generate as generate_custom_search_og_twitter_masterfile
from app.services.masterfile_pagination import build_pagination_masterfile
from app.services.masterfile_functional_internal_links import build_functional_internal_links_masterfile
from app.services.masterfile_non_functional_internal_links import build_non_functional_internal_links_masterfile

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

@router.get("/{crawl_id}/download/masterfile/response-codes-internal")
def download_masterfile_response_codes(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_response_codes_masterfile(
            crawl.id,
            crawl.domain,
            crawl.report_path,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_response_codes_internal.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/url-issues")
def download_masterfile_url_issues(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_url_issues_masterfile(
            crawl.id,
            crawl.domain,
            crawl.report_path,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_url_issues.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/page-titles")
def download_masterfile_page_titles(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_page_titles_masterfile(
            crawl.id,
            crawl.domain,
            crawl.report_path,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_page_titles.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/meta-description")
def download_masterfile_meta_description(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_meta_description_masterfile(crawl.id, crawl.domain, crawl.report_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_meta_description.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/h1")
def download_masterfile_h1(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_h1_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_h1.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/directives")
def download_masterfile_directives(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_directives_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_directives.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/sitemaps")
def download_masterfile_sitemaps(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_sitemaps_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_sitemaps.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/content-issues")
def download_masterfile_content_issues(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_content_issues_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_content_issues.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/duplicate-content")
def download_masterfile_duplicate_content(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_duplicate_content_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_duplicate_content.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/custom-search-ga4-gtm")
def download_masterfile_custom_search_ga4_gtm(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_custom_search_ga4_gtm_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_custom_search_ga4_gtm.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/custom-search-og-twitter")
def download_masterfile_custom_search_og_twitter(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = generate_custom_search_og_twitter_masterfile(crawl.report_path, crawl.domain)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_custom_search_og_twitter.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/pagination")
def download_masterfile_pagination(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_pagination_masterfile(crawl.id, crawl.domain, crawl.report_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_pagination.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/internal-links-non-functional")
def download_masterfile_non_functional_internal_links(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_non_functional_internal_links_masterfile(crawl.id, crawl.domain, crawl.report_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_non_functional_internal_links.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@router.get("/{crawl_id}/download/masterfile/internal-links-non-functional")
def download_masterfile_non_functional_internal_links(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")
    try:
        excel_bytes = build_non_functional_internal_links_masterfile(crawl.id, crawl.domain, crawl.report_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate masterfile: {str(e)}")
    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")
    filename = f"{domain_safe}_non_functional_internal_links.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )

@router.get("/{crawl_id}/download/masterfile/all")
def download_all_masterfiles(crawl_id: str, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")
    if not crawl.report_path or not os.path.exists(crawl.report_path):
        raise HTTPException(status_code=404, detail="Crawl output folder not found")

    domain_safe = crawl.domain.replace("https://", "").replace("http://", "").replace("/", "_").rstrip("_")

    masterfiles = [
        ("response_codes_internal", build_response_codes_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
        ("url_issues", build_url_issues_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
        ("page_titles", build_page_titles_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
        ("meta_description", build_meta_description_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
        ("h1", generate_h1_masterfile, [crawl.report_path, crawl.domain]),
        ("directives", generate_directives_masterfile, [crawl.report_path, crawl.domain]),
        ("sitemaps", generate_sitemaps_masterfile, [crawl.report_path, crawl.domain]),
        ("content_issues", generate_content_issues_masterfile, [crawl.report_path, crawl.domain]),
        ("duplicate_content", generate_duplicate_content_masterfile, [crawl.report_path, crawl.domain]),
        ("custom_search_ga4_gtm", generate_custom_search_ga4_gtm_masterfile, [crawl.report_path, crawl.domain]),
        ("custom_search_og_twitter", generate_custom_search_og_twitter_masterfile, [crawl.report_path, crawl.domain]),
        ("pagination", build_pagination_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
        ("functional_internal_links", build_functional_internal_links_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
        ("non_functional_internal_links", build_non_functional_internal_links_masterfile, [crawl.id, crawl.domain, crawl.report_path]),
    ]

    zip_buffer = io.BytesIO()
    errors = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, fn, args in masterfiles:
            try:
                excel_bytes = fn(*args)
                zf.writestr(f"{domain_safe}_{name}.xlsx", excel_bytes)
            except Exception as e:
                errors.append(f"{name}: {str(e)}")

    if errors:
        with zipfile.ZipFile(zip_buffer, "a") as zf:
            zf.writestr("errors.txt", "\n".join(errors))

    zip_buffer.seek(0)
    filename = f"{domain_safe}_all_masterfiles.zip"
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
