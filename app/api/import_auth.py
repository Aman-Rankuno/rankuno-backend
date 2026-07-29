"""
Single shared-credential gate for the crawl import feature.

Not a full user system: one username/password pair configured in the
backend's .env unlocks a signed, time-limited token. The token is stateless
(HMAC-signed with SECRET_KEY, like a lightweight JWT) so it survives uvicorn
restarts with no server-side session storage needed.

The actual security boundary is server-side: every import endpoint depends
on require_import_token, so a request without a valid token is rejected
regardless of what the frontend does or doesn't show.
"""
import hashlib
import hmac
import base64
import json
import time
from fastapi import APIRouter, HTTPException, Header
from pydantic import BaseModel
from app.config import settings

router = APIRouter()

TOKEN_TTL_SECONDS = 8 * 60 * 60  # 8 hours


def _sign(payload: str) -> str:
    return hmac.new(
        settings.SECRET_KEY.encode(), payload.encode(), hashlib.sha256
    ).hexdigest()


def _make_token() -> str:
    body = json.dumps({"exp": int(time.time()) + TOKEN_TTL_SECONDS})
    body_b64 = base64.urlsafe_b64encode(body.encode()).decode()
    sig = _sign(body_b64)
    return f"{body_b64}.{sig}"


def _verify_token(token: str) -> bool:
    try:
        body_b64, sig = token.split(".", 1)
    except ValueError:
        return False
    if not hmac.compare_digest(_sign(body_b64), sig):
        return False
    try:
        body = json.loads(base64.urlsafe_b64decode(body_b64.encode()))
    except Exception:
        return False
    return int(time.time()) < body.get("exp", 0)


def require_import_token(authorization: str = Header(default="")) -> None:
    """FastAPI dependency: raises 401 unless a valid Bearer token is present.
    This is the real access-control boundary; the frontend login form only
    exists to obtain this token, it does not enforce anything by itself."""
    prefix = "Bearer "
    if not authorization.startswith(prefix):
        raise HTTPException(status_code=401, detail="Login required to access crawl imports")
    token = authorization[len(prefix):]
    if not _verify_token(token):
        raise HTTPException(status_code=401, detail="Session expired or invalid, please log in again")


class LoginRequest(BaseModel):
    username: str
    password: str


@router.post("/login")
def login(payload: LoginRequest):
    if not settings.IMPORT_UPLOAD_USERNAME or not settings.IMPORT_UPLOAD_PASSWORD:
        raise HTTPException(
            status_code=500,
            detail="Import upload credentials are not configured on the server.",
        )
    # Constant-time comparison to avoid leaking match-length via timing
    user_ok = hmac.compare_digest(payload.username, settings.IMPORT_UPLOAD_USERNAME)
    pass_ok = hmac.compare_digest(payload.password, settings.IMPORT_UPLOAD_PASSWORD)
    if not (user_ok and pass_ok):
        raise HTTPException(status_code=401, detail="Incorrect username or password")
    return {"token": _make_token(), "expires_in": TOKEN_TTL_SECONDS}
