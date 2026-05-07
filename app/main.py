from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.database import Base, engine
from app.api import crawls

Base.metadata.create_all(bind=engine)

app = FastAPI(title="RankUno Crawl Toolkit API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://192.168.1.106:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(crawls.router, prefix="/api/crawls", tags=["crawls"])


@app.get("/")
def health_check():
    return {"status": "ok", "service": "RankUno Crawl Toolkit API"}