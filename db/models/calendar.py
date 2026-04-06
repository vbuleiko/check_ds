"""Модели календаря."""
from datetime import date, datetime, timezone
from typing import Optional
from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, String, Text, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from db.models.base import Base


class CalendarBase(Base):
    """Базовый календарь (общие праздники)."""
    __tablename__ = "calendar_base"

    date: Mapped[date] = mapped_column(Date, primary_key=True)
    weekday: Mapped[int] = mapped_column(Integer)
    is_holiday: Mapped[bool] = mapped_column(Boolean, default=False)
    treat_as: Mapped[Optional[int]] = mapped_column(Integer)
    note: Mapped[Optional[str]] = mapped_column(String(200))


class CalendarRouteOverride(Base):
    """Переопределение календаря для конкретного маршрута."""
    __tablename__ = "calendar_route_override"

    id: Mapped[int] = mapped_column(primary_key=True)
    date: Mapped[date] = mapped_column(Date)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contracts.id"))
    route: Mapped[str] = mapped_column(String(20))

    treat_as: Mapped[int] = mapped_column(Integer)

    # Источник
    source_agreement_id: Mapped[Optional[int]] = mapped_column(ForeignKey("agreements.id"))
    source_text: Mapped[Optional[str]] = mapped_column(Text)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))

    # Связи
    contract: Mapped["Contract"] = relationship(back_populates="calendar_overrides")
    source_agreement: Mapped[Optional["Agreement"]] = relationship(back_populates="calendar_overrides")

    __table_args__ = (
        UniqueConstraint("date", "contract_id", "route", name="uq_calendar_override"),
    )


# Отложенные импорты для типизации
from db.models.contracts import Contract, Agreement  # noqa: E402, F401
