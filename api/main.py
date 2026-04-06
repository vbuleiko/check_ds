"""FastAPI приложение для проверки ДС."""
import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Depends, Request

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
from sqlalchemy.orm import Session

from core.config import settings
from db.database import init_db, get_db
from db.models import Contract, Agreement, AgreementStatus

# Пути
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Инициализация при старте."""
    init_db()
    settings.upload_dir.mkdir(exist_ok=True)
    yield


# Инициализация
app = FastAPI(
    title="Проверка ДС",
    description="Веб-сервис проверки дополнительных соглашений к госконтрактам",
    version="0.1.0",
    lifespan=lifespan,
)

# Статика и шаблоны
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# Подключаем API роуты
from api.routes import upload, agreements, calendar, calculations, summary, table_checks, acts, route_params_editor

app.include_router(upload.router, prefix="/api", tags=["Upload"])
app.include_router(agreements.router, prefix="/api/agreements", tags=["Agreements"])
app.include_router(calendar.router, prefix="/api/calendar", tags=["Calendar"])
app.include_router(calculations.router, prefix="/api/calculations", tags=["Calculations"])
app.include_router(summary.router, prefix="/api/summary", tags=["Summary"])
app.include_router(table_checks.router, prefix="/api/table-checks", tags=["TableChecks"])
app.include_router(acts.router, prefix="/api/acts", tags=["Acts"])
app.include_router(route_params_editor.router, prefix="/api/params-editor", tags=["ParamsEditor"])


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Главная страница."""
    return templates.TemplateResponse(request=request, name="index.html")


@app.get("/api/contracts")
async def list_contracts(db: Session = Depends(get_db)):
    """Список контрактов."""
    contracts = db.query(Contract).all()
    return [
        {
            "id": c.id,
            "number": c.number,
            "full_number": c.full_number,
            "date_from": str(c.date_from) if c.date_from else None,
            "date_to": str(c.date_to) if c.date_to else None,
        }
        for c in contracts
    ]


@app.get("/api/stats")
async def get_stats(db: Session = Depends(get_db)):
    """Статистика системы."""
    contracts_count = db.query(Contract).count()
    agreements_count = db.query(Agreement).count()
    agreements_checked = db.query(Agreement).filter(
        Agreement.status != AgreementStatus.DRAFT
    ).count()

    return {
        "contracts": contracts_count,
        "agreements": agreements_count,
        "agreements_checked": agreements_checked,
    }
