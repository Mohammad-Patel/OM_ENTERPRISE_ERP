
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


class Config:
    SECRET_KEY = os.getenv("SECRET_KEY")

    if not SECRET_KEY:
        if os.getenv("ERP_ENV", "development").lower() == "production":
            raise RuntimeError(
                "SECRET_KEY environment variable must be set in production."
            )

        # Local development only.
        SECRET_KEY = "om-enterprise-local-development-key"

    WTF_CSRF_ENABLED = True

    BASE_DIR = BASE_DIR
    DATA_DIR = BASE_DIR / "data"
    DOCUMENTS_DIR = BASE_DIR / "documents"
    BACKUP_DIR = BASE_DIR / "backups"
    LOG_DIR = BASE_DIR / "logs"

    WORKBOOK_PATH = DATA_DIR / "om_enterprise.xlsx"

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DOCUMENTS_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
