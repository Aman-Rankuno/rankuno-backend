from pydantic_settings import BaseSettings
class Settings(BaseSettings):
    DATABASE_URL: str
    REDIS_URL: str
    SCREAMING_FROG_CLI: str
    CRAWL_OUTPUT_DIR: str
    CRAWL_CONFIGS_DIR: str
    RULEBOOKS_DIR: str = r"D:\projects\rankuno-rulebooks"
    TEMPLATES_DIR: str = r"D:\projects\rankuno-templates"
    SECRET_KEY: str
    ANTHROPIC_API_KEY: str = ""
    OPENAI_API_KEY: str = ""
    IMPORT_UPLOAD_USERNAME: str = ""
    IMPORT_UPLOAD_PASSWORD: str = ""
    CRAWL_IMPORTS_DIR: str = r"D:\projects\rankuno-imports"
    class Config:
        env_file = ".env"
settings = Settings()