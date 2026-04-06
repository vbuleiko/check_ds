"""Модели параметров маршрутов и рейсов."""
from datetime import date, datetime, timezone
from typing import Optional
from sqlalchemy import Date, DateTime, Enum, Float, ForeignKey, Integer, String, JSON, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from db.models.base import Base, SeasonType


class RouteParams(Base):
    """Параметры маршрута на определённый период."""
    __tablename__ = "route_params"

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contracts.id"))
    route: Mapped[str] = mapped_column(String(20))

    # Период действия
    date_from: Mapped[date] = mapped_column(Date)
    date_to: Mapped[Optional[date]] = mapped_column(Date)

    # Сезон
    season: Mapped[SeasonType] = mapped_column(
        Enum(SeasonType), default=SeasonType.ALL_YEAR
    )

    # Протяжённость
    length_total: Mapped[Optional[float]] = mapped_column(Float)
    length_forward: Mapped[Optional[float]] = mapped_column(Float)
    length_reverse: Mapped[Optional[float]] = mapped_column(Float)

    # Источник
    source_agreement_id: Mapped[Optional[int]] = mapped_column(ForeignKey("agreements.id"))
    source_appendix: Mapped[Optional[str]] = mapped_column(String(50))

    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))

    # Связи
    contract: Mapped["Contract"] = relationship(back_populates="route_params")
    source_agreement: Mapped[Optional["Agreement"]] = relationship(back_populates="route_params")
    trips: Mapped[list["RouteTrips"]] = relationship(back_populates="route_params", cascade="all, delete-orphan")


class RouteTrips(Base):
    """Количество рейсов по типам дней."""
    __tablename__ = "route_trips"

    id: Mapped[int] = mapped_column(primary_key=True)
    route_params_id: Mapped[int] = mapped_column(ForeignKey("route_params.id"))

    day_type_name: Mapped[str] = mapped_column(String(100))
    weekdays: Mapped[list[int]] = mapped_column(JSON)
    forward_number: Mapped[int] = mapped_column(Integer, default=0)
    reverse_number: Mapped[int] = mapped_column(Integer, default=0)

    # Связи
    route_params: Mapped["RouteParams"] = relationship(back_populates="trips")


class RouteEndDate(Base):
    """Дата окончания работы маршрута."""
    __tablename__ = "route_end_dates"

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contracts.id"))
    route: Mapped[str] = mapped_column(String(20))
    end_date: Mapped[date] = mapped_column(Date)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))

    # Связи
    contract: Mapped["Contract"] = relationship()

    __table_args__ = (
        UniqueConstraint("contract_id", "route", name="uq_route_end_date"),
    )


class RouteSeasonConfig(Base):
    """Конфигурация сезонных периодов для маршрута."""
    __tablename__ = "route_season_configs"

    id: Mapped[int] = mapped_column(primary_key=True)
    route: Mapped[str] = mapped_column(String(20), unique=True)

    winter_start_month: Mapped[int] = mapped_column(Integer)
    winter_start_day: Mapped[int] = mapped_column(Integer)
    winter_end_month: Mapped[int] = mapped_column(Integer)
    winter_end_day: Mapped[int] = mapped_column(Integer)

    summer_start_month: Mapped[int] = mapped_column(Integer)
    summer_start_day: Mapped[int] = mapped_column(Integer)
    summer_end_month: Mapped[int] = mapped_column(Integer)
    summer_end_day: Mapped[int] = mapped_column(Integer)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc)
    )


class RouteSeasonPeriod(Base):
    """Конкретный сезонный период маршрута с реальными датами."""
    __tablename__ = "route_season_periods"

    id: Mapped[int] = mapped_column(primary_key=True)
    route: Mapped[str] = mapped_column(String(20))
    season: Mapped[SeasonType] = mapped_column(Enum(SeasonType))
    date_from: Mapped[date] = mapped_column(Date)
    date_to: Mapped[date] = mapped_column(Date)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc)
    )

    __table_args__ = (
        UniqueConstraint("route", "season", "date_from", name="uq_route_season_period"),
    )


# Отложенные импорты для типизации
from db.models.contracts import Contract, Agreement  # noqa: E402, F401
