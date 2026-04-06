"""Модели рассчитанных этапов."""
from datetime import date, datetime, timezone
from typing import Optional
from sqlalchemy import Date, DateTime, Enum, Float, ForeignKey, Integer, String, JSON, UniqueConstraint
from sqlalchemy.orm import Mapped, mapped_column, relationship

from db.models.base import Base, StageStatus


class CalculatedStage(Base):
    """Рассчитанный этап контракта."""
    __tablename__ = "calculated_stages"

    id: Mapped[int] = mapped_column(primary_key=True)
    contract_id: Mapped[int] = mapped_column(ForeignKey("contracts.id"))

    stage: Mapped[int] = mapped_column(Integer)
    year: Mapped[int] = mapped_column(Integer)
    month: Mapped[int] = mapped_column(Integer)
    period_name: Mapped[str] = mapped_column(String(20))
    date_from: Mapped[date] = mapped_column(Date)
    date_to: Mapped[date] = mapped_column(Date)

    total_km: Mapped[float] = mapped_column(Float, default=0.0)
    total_price: Mapped[Optional[float]] = mapped_column(Float)

    status: Mapped[StageStatus] = mapped_column(
        Enum(StageStatus), default=StageStatus.SAVED
    )

    routes_data: Mapped[Optional[dict]] = mapped_column(JSON)

    calculated_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.now(timezone.utc))
    source_agreement_id: Mapped[Optional[int]] = mapped_column(ForeignKey("agreements.id"))

    # Связи
    contract: Mapped["Contract"] = relationship()

    __table_args__ = (
        UniqueConstraint("contract_id", "stage", name="uq_calculated_stage"),
    )


# Отложенные импорты
from db.models.contracts import Contract  # noqa: E402, F401
