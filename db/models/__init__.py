"""
Модели базы данных.

Все модели доступны через `from db.models import ...` для обратной совместимости.
Модели организованы по доменам:
- base.py: Base, перечисления (AgreementStatus, StageStatus, SeasonType)
- contracts.py: Contract, Agreement, AgreementReference
- routes.py: RouteParams, RouteTrips, RouteEndDate, RouteSeasonConfig, RouteSeasonPeriod
- calendar.py: CalendarBase, CalendarRouteOverride
- stages.py: CalculatedStage
"""

from db.models.base import Base, AgreementStatus, StageStatus, SeasonType
from db.models.contracts import Contract, Agreement, AgreementReference
from db.models.routes import (
    RouteParams, RouteTrips, RouteEndDate,
    RouteSeasonConfig, RouteSeasonPeriod,
)
from db.models.calendar import CalendarBase, CalendarRouteOverride
from db.models.stages import CalculatedStage

__all__ = [
    "Base",
    "AgreementStatus",
    "StageStatus",
    "SeasonType",
    "Contract",
    "Agreement",
    "AgreementReference",
    "RouteParams",
    "RouteTrips",
    "RouteEndDate",
    "RouteSeasonConfig",
    "RouteSeasonPeriod",
    "CalendarBase",
    "CalendarRouteOverride",
    "CalculatedStage",
]
