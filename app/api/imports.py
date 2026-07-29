import os
import uuid
import shutil
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form
from sqlalchemy.orm import Session
from app.database import get_db
from app.models.crawl import Crawl
from app.config import settings
from app.api.import_auth import require_import_token
from app.tasks.import_runner import run_import_job

router = APIRouter()

ALLOWED_EXTENSIONS = (".dbseospider", ".seospider")
MAX_UPLOAD_BYTES = 5 * 1024 * 1024 * 1024  # 5 GB; large crawl files are expected
JAVA_MAGIC = b"\xac\xed"  # same Java-serialized signature as .seospiderconfig


@router.post("/", dependencies=[Depends(require_import_token)])
async def create_import(
    domain: str = Form(...),
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    """Upload a saved Screaming Frog crawl file (.dbseospider) and queue it
    for export. Gated by require_import_token: an unauthenticated request is
    rejected before any file processing happens."""
    original = file.filename or ""
    if not original.lower().endswith(ALLOWED_EXTENSIONS):
        raise HTTPException(
            status_code=400,
            detail=f"File must be a Screaming Frog crawl file ({', '.join(ALLOWED_EXTENSIONS)})",
        )
    if not domain.strip():
        raise HTTPException(status_code=400, detail="Domain is required")

    os.makedirs(settings.CRAWL_IMPORTS_DIR, exist_ok=True)
    import_id = str(uuid.uuid4())
    ext = os.path.splitext(original)[1]
    dest_path = os.path.join(settings.CRAWL_IMPORTS_DIR, f"{import_id}{ext}")

    # Stream to disk in chunks; never hold the whole file in memory, since
    # these can be multiple GB for large crawls
    total_bytes = 0
    magic_checked = False
    try:
        with open(dest_path, "wb") as out:
            while True:
                chunk = await file.read(1024 * 1024)  # 1 MB chunks
                if not chunk:
                    break
                if not magic_checked:
                    if not chunk.startswith(JAVA_MAGIC):
                        raise HTTPException(
                            status_code=400,
                            detail="File does not look like a valid Screaming Frog crawl file.",
                        )
                    magic_checked = True
                total_bytes += len(chunk)
                if total_bytes > MAX_UPLOAD_BYTES:
                    raise HTTPException(status_code=400, detail="File exceeds the 5 GB limit.")
                out.write(chunk)
    except HTTPException:
        if os.path.exists(dest_path):
            os.remove(dest_path)
        raise
    except Exception as e:
        if os.path.exists(dest_path):
            os.remove(dest_path)
        raise HTTPException(status_code=500, detail=f"Upload failed: {e}")

    if total_bytes == 0:
        os.remove(dest_path)
        raise HTTPException(status_code=400, detail="Uploaded file is empty")

    crawl = Crawl(
        domain=domain.strip(),
        crawl_type="imported",
        status="queued",
    )
    db.add(crawl)
    db.commit()
    db.refresh(crawl)

    run_import_job.delay(crawl.id, dest_path)
    return {
        "id": crawl.id,
        "status": "queued",
        "message": f"Crawl file uploaded ({round(total_bytes / (1024*1024))} MB) and queued for export",
    }
