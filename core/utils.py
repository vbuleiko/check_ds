"""Общие утилиты проекта."""
from datetime import date


def parse_date(date_str: str | None) -> date | None:
    """Парсит дату из строки формата DD.MM.YYYY или YYYY-MM-DD."""
    if not date_str:
        return None
    try:
        if "-" in date_str:
            return date.fromisoformat(date_str)
        parts = date_str.split(".")
        if len(parts) == 3:
            return date(int(parts[2]), int(parts[1]), int(parts[0]))
    except (ValueError, IndexError, TypeError):
        pass
    return None
