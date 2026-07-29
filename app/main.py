from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.database import Base, engine
from app.api import crawls, chat
Base.metadata.create_all(bind=engine)
app = FastAPI(title="RankUno Crawl Toolkit API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(crawls.router, prefix="/api/crawls", tags=["crawls"])
app.include_router(chat.router, prefix="/api", tags=["chat"])
from app.api import configs
app.include_router(configs.router, prefix="/api/configs", tags=["configs"])
from app.api import import_auth, imports
app.include_router(import_auth.router, prefix="/api/imports/auth", tags=["imports"])
app.include_router(imports.router, prefix="/api/imports", tags=["imports"])
@app.get("/")
def health_check():
    return {"status": "ok", "service": "RankUno Crawl Toolkit API"}