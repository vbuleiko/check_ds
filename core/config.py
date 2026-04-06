"""Конфигурация приложения."""
import os
from pathlib import Path
from dataclasses import dataclass

# Загружаем .env файл
def load_dotenv():
    env_path = Path(__file__).parent.parent / ".env"
    if env_path.exists():
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, value = line.partition("=")
                    os.environ.setdefault(key.strip(), value.strip())

load_dotenv()


@dataclass
class Settings:
    """Настройки приложения."""

    # База данных приложения (SQLite для начала)
    database_url: str = "sqlite:///./check_ds.db"

    # Внешняя БД PostgreSQL (справочники)
    external_db_url: str | None = None

    # Пути
    base_dir: Path = Path(__file__).parent.parent
    upload_dir: Path = None
    old_data_dir: Path = None

    # API
    api_prefix: str = "/api"

    # Путь к 7z (для распаковки RAR архивов)
    sevenzip_path: str | None = None

    def __post_init__(self):
        if self.upload_dir is None:
            self.upload_dir = self.base_dir / "uploads"
        if self.old_data_dir is None:
            self.old_data_dir = self.base_dir / "old"

        # Загружаем из переменных окружения
        self.database_url = os.environ.get("DATABASE_URL", self.database_url)
        self.external_db_url = os.environ.get("EXTERNAL_DB_URL", self.external_db_url)
        self.sevenzip_path = os.environ.get("SEVENZIP_PATH", self.sevenzip_path)

        # Автоопределение 7z если не задан явно
        if not self.sevenzip_path:
            self.sevenzip_path = self._find_7z()

    @staticmethod
    def _find_7z() -> str | None:
        """Автоопределение пути к 7z."""
        import shutil
        candidates = [
            "C:/Program Files/7-Zip/7z.exe",
            "C:/Program Files (x86)/7-Zip/7z.exe",
            "7z",
        ]
        for path in candidates:
            if Path(path).exists() or shutil.which(path):
                return path
        return None


settings = Settings()
