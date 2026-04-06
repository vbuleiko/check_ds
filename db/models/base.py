"""Базовые классы и перечисления для моделей."""
import enum
from sqlalchemy.orm import DeclarativeBase


class Base(DeclarativeBase):
    """Базовый класс для моделей."""
    pass


class AgreementStatus(str, enum.Enum):
    """Статус ДС."""
    DRAFT = "draft"           # Загружен, не проверен
    CHECKED = "checked"       # Проверен, есть ошибки или нет
    APPLIED = "applied"       # Применён (данные добавлены в route_params)


class StageStatus(str, enum.Enum):
    """Статус этапа контракта."""
    SAVED = "saved"           # Сохранён (будущие этапы)
    NOT_CLOSED = "not_closed" # Не закрыт (прошедшие, можно редактировать)
    CLOSED = "closed"         # Закрыт (не пересчитывается)


class SeasonType(str, enum.Enum):
    """Тип сезона."""
    ALL_YEAR = "all_year"     # Круглогодично
    WINTER = "winter"         # Зимний период
    SUMMER = "summer"         # Летний период
