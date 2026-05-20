from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    DATABASE_URL: str
    REDIS_URL: str
    SCREAMING_FROG_CLI: str
    CRAWL_OUTPUT_DIR: str
    CRAWL_CONFIGS_DIR: str
    RULEBOOKS_DIR: str = r"D:\projects\rankuno-rulebooks"
    SECRET_KEY: str

    class Config:
        env_file = ".env"

settings = Settings()