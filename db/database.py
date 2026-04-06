"""Настройка подключения к базе данных."""
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, Session
from contextlib import contextmanager

from core.config import settings
from db.models import Base


# Движок SQLite для локальной БД
engine = create_engine(
    settings.database_url,
    connect_args={"check_same_thread": False} if "sqlite" in settings.database_url else {},
    echo=False,
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


def init_db():
    """Инициализация базы данных (создание таблиц)."""
    Base.metadata.create_all(bind=engine)
    _seed_route_season_configs()
    _seed_route_season_periods()


def _seed_route_season_configs():
    """Заполняет route_season_configs из констант, если таблица пуста."""
    from db.models import RouteSeasonConfig
    from core.constants import ROUTE_SEASON_PERIODS

    with get_db_session() as db:
        if db.query(RouteSeasonConfig).count() == 0:
            for route, seasons in ROUTE_SEASON_PERIODS.items():
                ws, wd, we, wed = seasons["winter"]
                ss, sd, se, sed = seasons["summer"]
                db.add(RouteSeasonConfig(
                    route=route,
                    winter_start_month=ws, winter_start_day=wd,
                    winter_end_month=we, winter_end_day=wed,
                    summer_start_month=ss, summer_start_day=sd,
                    summer_end_month=se, summer_end_day=sed,
                ))


def _seed_route_season_periods():
    """Заполняет route_season_periods конкретными датами по годам, если таблица пуста."""
    from db.models import RouteSeasonPeriod, SeasonType
    from core.constants import ROUTE_SEASON_PERIODS
    from datetime import date as _date

    with get_db_session() as db:
        if db.query(RouteSeasonPeriod).count() == 0:
            for route, seasons in ROUTE_SEASON_PERIODS.items():
                ws_m, ws_d, we_m, we_d = seasons["winter"]
                ss_m, ss_d, se_m, se_d = seasons["summer"]
                winter_cross = ws_m > we_m
                summer_cross = ss_m > se_m
                for year in range(2022, 2029):
                    db.add(RouteSeasonPeriod(
                        route=route,
                        season=SeasonType.WINTER,
                        date_from=_date(year, ws_m, ws_d),
                        date_to=_date(year + 1 if winter_cross else year, we_m, we_d),
                    ))
                    db.add(RouteSeasonPeriod(
                        route=route,
                        season=SeasonType.SUMMER,
                        date_from=_date(year, ss_m, ss_d),
                        date_to=_date(year + 1 if summer_cross else year, se_m, se_d),
                    ))


def get_db():
    """Генератор сессии БД для FastAPI Depends."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


@contextmanager
def get_db_session() -> Session:
    """Контекстный менеджер для работы с БД."""
    db = SessionLocal()
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()
