import os
import io
import json
import uuid
import threading
from datetime import datetime, timezone
from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from app.config import settings

router = APIRouter()

# Custom configs live in a subfolder of the existing configs directory so the
# crawl task resolves them with the same os.path.join(CRAWL_CONFIGS_DIR, ...)
# mechanism the 6 presets already use. Metadata is a JSON manifest, no DB change.
CUSTOM_SUBDIR = "custom"
MANIFEST_NAME = "manifest.json"
MAX_CONFIG_BYTES = 10 * 1024 * 1024  # 10 MB
JAVA_MAGIC = b"\xac\xed"  # Java serialization header; every real .seospiderconfig starts with this

_manifest_lock = threading.Lock()


def _custom_dir() -> str:
    d = os.path.join(settings.CRAWL_CONFIGS_DIR, CUSTOM_SUBDIR)
    os.makedirs(d, exist_ok=True)
    return d


def _manifest_path() -> str:
    return os.path.join(_custom_dir(), MANIFEST_NAME)


def _read_manifest() -> list:
    p = _manifest_path()
    if not os.path.exists(p):
        return []
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _write_manifest(entries: list) -> None:
    tmp = _manifest_path() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)
    os.replace(tmp, _manifest_path())


@router.get("/")
def list_configs():
    """List the built-in presets plus every uploaded custom config."""
    presets = []
    try:
        for fname in sorted(os.listdir(settings.CRAWL_CONFIGS_DIR)):
            full = os.path.join(settings.CRAWL_CONFIGS_DIR, fname)
            if os.path.isfile(full) and fname.lower().endswith(".seospiderconfig"):
                presets.append({
                    "id": fname,
                    "name": os.path.splitext(fname)[0],
                    "kind": "preset",
                    "config_file": fname,
                })
    except FileNotFoundError:
        pass

    with _manifest_lock:
        entries = _read_manifest()
    custom = []
    for e in entries:
        # config_file is the path relative to CRAWL_CONFIGS_DIR, e.g. custom/<uuid>.seospiderconfig
        full = os.path.join(settings.CRAWL_CONFIGS_DIR, e.get("config_file", ""))
        custom.append({
            "id": e.get("id"),
            "name": e.get("name"),
            "kind": "custom",
            "config_file": e.get("config_file"),
            "original_filename": e.get("original_filename"),
            "uploaded_at": e.get("uploaded_at"),
            "size_bytes": os.path.getsize(full) if os.path.exists(full) else None,
            "missing_on_disk": not os.path.exists(full),
        })

    return {"presets": presets, "custom": custom}


@router.post("/upload")
async def upload_config(file: UploadFile = File(...), name: str = Form("")):
    """Upload a .seospiderconfig built in the Screaming Frog GUI.

    Validation:
    - extension must be .seospiderconfig
    - first two bytes must be the Java serialization magic header (AC ED);
      SF configs are Java-serialized binaries, anything else is not a real config
    - size capped at 10 MB
    The file is stored under a UUID; the original filename is kept as metadata only.
    """
    original = file.filename or ""
    if not original.lower().endswith(".seospiderconfig"):
        raise HTTPException(status_code=400, detail="File must be a .seospiderconfig exported from Screaming Frog")

    content = await file.read()
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="Uploaded file is empty")
    if len(content) > MAX_CONFIG_BYTES:
        raise HTTPException(status_code=400, detail="Config file exceeds the 10 MB limit")
    if content[:2] != JAVA_MAGIC:
        raise HTTPException(
            status_code=400,
            detail="Not a valid Screaming Frog config file. Open Screaming Frog, set up the "
                   "configuration, then File > Configuration > Save As to produce a .seospiderconfig.",
        )

    config_id = str(uuid.uuid4())
    stored_name = f"{config_id}.seospiderconfig"
    stored_path = os.path.join(_custom_dir(), stored_name)
    with open(stored_path, "wb") as f:
        f.write(content)

    display_name = (name or "").strip() or os.path.splitext(original)[0]

    entry = {
        "id": config_id,
        "name": display_name[:200],
        "original_filename": original[:255],
        "config_file": f"{CUSTOM_SUBDIR}/{stored_name}",
        "uploaded_at": datetime.now(timezone.utc).isoformat(),
    }

    with _manifest_lock:
        entries = _read_manifest()
        entries.append(entry)
        _write_manifest(entries)

    return {
        "id": config_id,
        "name": entry["name"],
        "config_file": entry["config_file"],
        "message": "Config uploaded. Note: it must have been saved from Screaming Frog 19.x "
                   "to be compatible with the crawl engine.",
    }


@router.delete("/{config_id}")
def delete_config(config_id: str):
    """Delete an uploaded custom config. Presets cannot be deleted through the API."""
    with _manifest_lock:
        entries = _read_manifest()
        match = next((e for e in entries if e.get("id") == config_id), None)
        if not match:
            raise HTTPException(status_code=404, detail="Custom config not found")
        entries = [e for e in entries if e.get("id") != config_id]
        _write_manifest(entries)

    full = os.path.join(settings.CRAWL_CONFIGS_DIR, match.get("config_file", ""))
    # Path safety: only ever delete inside the custom subfolder
    custom_root = os.path.abspath(_custom_dir())
    if os.path.abspath(full).startswith(custom_root) and os.path.exists(full):
        try:
            os.remove(full)
        except OSError:
            pass

    return {"id": config_id, "deleted": True}
