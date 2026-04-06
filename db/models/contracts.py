"""Модели госконтрактов и дополнительных соглашений."""
from datetime import date, datetime, timezone
from typing import Optional
from sqlalchemy import Date, DateTime, Enum, Float, ForeignKey, Integer, String, Text, JSON, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from db.models.base import Base, AgreementStatus


class Contract(Base):
    """Госконтракт."""
    __tablename__ = "contracts"

    id: Mapped[int] = mapped_column(primary_key=True)
    number: Mapped[str] = mapped_column(String(10), unique=True)
    full_number: Mapped[Optional[str]] = mapped_column(String(50))
    carrier: Mapped[Optional[str]] = mapped_column(String(200))
    date_from: Mapped[Optional[date]] = mapped_column(Date)
    date_to: Mapped[Optional[date]] = mapped_column(Date)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))

    # Связи
    agreements: Mapped[list["Agreement"]] = relationship(back_populates="contract")
    route_params: Mapped[list["RouteParams"]] = relationship(back_populates="contract")
    calendar_overrides: Mapped[list["CalendarRouteOverride"]] = relationship(back_populates="contract")


class Agreement(Base):
    """Дополнительное соглашение."""
    __tablename__ = "agreements"

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contracts.id"))
    number: Mapped[str] = mapped_column(String(10))
    date_signed: Mapped[Optional[date]] = mapped_column(Date)
    status: Mapped[AgreementStatus] = mapped_column(
        Enum(AgreementStatus), default=AgreementStatus.DRAFT
    )

    # Данные
    json_data: Mapped[Optional[dict]] = mapped_column(JSON)
    check_errors: Mapped[Optional[list]] = mapped_column(JSON)
    check_warnings: Mapped[Optional[list]] = mapped_column(JSON)

    # Расчёты
    total_km_before: Mapped[Optional[float]] = mapped_column(Float)
    total_km_after: Mapped[Optional[float]] = mapped_column(Float)
    total_price_before: Mapped[Optional[float]] = mapped_column(Float)
    total_price_after: Mapped[Optional[float]] = mapped_column(Float)
    calculation_details: Mapped[Optional[dict]] = mapped_column(JSON)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc)
    )

    # Связи
    contract: Mapped["Contract"] = relationship(back_populates="agreements")
    route_params: Mapped[list["RouteParams"]] = relationship(back_populates="source_agreement")
    calendar_overrides: Mapped[list["CalendarRouteOverride"]] = relationship(back_populates="source_agreement")

    __table_args__ = (
        UniqueConstraint("contract_id", "number", name="uq_agreement_contract_number"),
    )


class AgreementReference(Base):
    """Эталонные данные из ДС, которого нет в системе."""
    __tablename__ = "agreement_references"

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contracts.id"))
    reference_ds_number: Mapped[str] = mapped_column(String(10))
    initial_km: Mapped[Optional[float]] = mapped_column(Float)
    probeg_etapy: Mapped[Optional[float]] = mapped_column(Float)
    sum_price: Mapped[Optional[float]] = mapped_column(Float)
    note: Mapped[Optional[str]] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))

    # Связи
    contract: Mapped["Contract"] = relationship()

    __table_args__ = (
        UniqueConstraint("contract_id", "reference_ds_number", name="uq_agreement_reference"),
    )


# Отложенные импорты для типизации (avoid circular imports)
from db.models.routes import RouteParams  # noqa: E402, F401
from db.models.calendar import CalendarRouteOverride  # noqa: E402, F401
