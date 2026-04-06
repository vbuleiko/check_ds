"""
Миграция: добавление поля status в таблицу calculated_stages.

Запуск: python scripts/migrate_add_stage_status.py
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from datetime import date
from sqlalchemy import text
from db.database import engine, SessionLocal
from db.models import CalculatedStage, StageStatus


def migrate():
    """Добавляет поле status в calculated_stages и устанавливает значения."""

    with engine.connect() as conn:
        # Проверяем, существует ли поле status
        result = conn.execute(text("PRAGMA table_info(calculated_stages)"))
        columns = [row[1] for row in result.fetchall()]

        if 'status' not in columns:
            print("Добавляю поле status в таблицу calculated_stages...")
            conn.execute(text(
                "ALTER TABLE calculated_stages ADD COLUMN status VARCHAR(20) DEFAULT 'saved'"
            ))
            conn.commit()
            print("Поле status добавлено.")
        else:
            print("Поле status уже существует.")

    # Обновляем статусы для существующих записей
    session = SessionLocal()
    try:
        today = date.today()

        # Прошедшие этапы -> NOT_CLOSED
        past_count = session.query(CalculatedStage).filter(
            CalculatedStage.date_to < today,
            CalculatedStage.status != StageStatus.CLOSED,
        ).update({CalculatedStage.status: StageStatus.NOT_CLOSED})

        # Будущие этапы -> SAVED
        future_count = session.query(CalculatedStage).filter(
            CalculatedStage.date_to >= today,
            CalculatedStage.status != StageStatus.CLOSED,
        ).update({CalculatedStage.status: StageStatus.SAVED})

        session.commit()

        print(f"Обновлено статусов: {past_count} прошедших (not_closed), {future_count} будущих (saved)")

    finally:
        session.close()

    print("Миграция завершена.")


if __name__ == "__main__":
    migrate()
