from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
from sqlalchemy.orm import Session
from app.database import get_db
from app.models.crawl import Crawl
from app.services.ai_chat.providers import get_provider, ProviderNotConfigured

router = APIRouter()


class ChatMessage(BaseModel):
    role: str  # "user" or "assistant"
    content: str


class ChatRequest(BaseModel):
    crawl_id: str
    message: str
    provider: str = "openai"  # "openai" or "claude"
    history: Optional[List[ChatMessage]] = None


class ChatResponse(BaseModel):
    reply: str
    provider: str


SYSTEM_PROMPT_TEMPLATE = """You are an SEO audit assistant for RankUno, helping analyze a specific website crawl.

You have tools available to look up real data about this crawl: issue counts, top issues, and affected URLs.
Always use these tools to get real numbers - never guess or estimate a count yourself.

Current crawl context:
- Domain: {domain}
- Crawl type: {crawl_type}
- Status: {status}
- Pages crawled: {pages_crawled}

Answer questions clearly and concisely. When asked for a summary or top issues, use list_top_issues.
When asked about a specific issue type, use get_issue_count or list_affected_urls.
If a question is outside what your tools can answer, say so plainly rather than guessing."""


@router.post("/chat", response_model=ChatResponse)
def chat(payload: ChatRequest, db: Session = Depends(get_db)):
    crawl = db.query(Crawl).filter(Crawl.id == payload.crawl_id).first()
    if not crawl:
        raise HTTPException(status_code=404, detail="Crawl not found")

    try:
        provider = get_provider(payload.provider)
    except ProviderNotConfigured as e:
        raise HTTPException(status_code=400, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    system_prompt = SYSTEM_PROMPT_TEMPLATE.format(
        domain=crawl.domain,
        crawl_type=crawl.crawl_type,
        status=crawl.status,
        pages_crawled=crawl.pages_crawled,
    )

    history = [{"role": m.role, "content": m.content} for m in (payload.history or [])]
    history.append({"role": "user", "content": payload.message})

    try:
        reply = provider.chat(system_prompt, history, crawl)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Chat failed: {str(e)}")

    return ChatResponse(reply=reply, provider=payload.provider)


@router.get("/chat/providers")
def available_providers():
    """Tells the frontend which providers are actually usable right now
    (i.e. have an API key configured), so the UI can disable/hide options
    that aren't ready rather than letting the user pick one that 400s."""
    from app.config import settings
    return {
        "openai": bool(settings.OPENAI_API_KEY),
        "claude": bool(settings.ANTHROPIC_API_KEY),
    }
