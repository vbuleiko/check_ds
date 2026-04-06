#!/usr/bin/env python3
"""
Скрипт импорта дат окончания работы маршрутов из Excel.

Импортирует данные из Книга1.xlsx → route_end_dates
"""
import sys
from pathlib import Path
from datetime import datetime, date

import pandas as pd

# Добавляем корень проекта в path
sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import init_db, get_db_session
from db.models import Contract, RouteEndDate


def normalize_route(route: str) -> str:
    """Нормализует название маршрута."""
    route = str(route).strip()
    # Убираем точку в конце
    if route.endswith('.'):
        route = route[:-1]
    return route


def import_route_end_dates(xlsx_path: str, contract_number: str = None):
    """
    Импортирует даты окончания маршрутов из Excel.

    Args:
        xlsx_path: Путь к Excel файлу
        contract_number: Номер контракта (если None - применяется ко всем)
    """
    print(f"Импорт дат окончания маршрутов из {xlsx_path}...")

    # Читаем Excel
    df = pd.read_excel(xlsx_path)

    # Определяем колонки (первая - маршрут, вторая - дата)
    columns = df.columns.tolist()
    route_col = columns[0]
    date_col = columns[1]

    print(f"  Колонка маршрутов: {route_col}")
    print(f"  Колонка дат: {date_col}")

    init_db()

    with get_db_session() as session:
        # Получаем контракты
        if contract_number:
            contracts = session.query(Contract).filter(
                Contract.number == contract_number
            ).all()
        else:
            contracts = session.query(Contract).all()

        if not contracts:
            print("Контракты не найдены!")
            return

        count = 0
        for _, row in df.iterrows():
            route = normalize_route(row[route_col])
            end_date_value = row[date_col]

            if pd.isna(route) or pd.isna(end_date_value):
                continue

            # Парсим дату
            if isinstance(end_date_value, datetime):
                end_date = end_date_value.date()
            elif isinstance(end_date_value, date):
                end_date = end_date_value
            else:
                try:
                    end_date = datetime.strptime(str(end_date_value), "%Y-%m-%d").date()
                except ValueError:
                    print(f"  Не удалось распарсить дату: {end_date_value}")
                    continue

            # Добавляем для каждого контракта
            for contract in contracts:
                # Проверяем существование
                existing = session.query(RouteEndDate).filter(
                    RouteEndDate.contract_id == contract.id,
                    RouteEndDate.route == route,
                ).first()

                if existing:
                    # Обновляем дату
                    existing.end_date = end_date
                    print(f"  Обновлён маршрут {route} ({contract.number}): {end_date}")
                else:
                    # Создаём новую запись
                    route_end = RouteEndDate(
                        contract_id=contract.id,
                        route=route,
                        end_date=end_date,
                    )
                    session.add(route_end)
                    print(f"  Добавлен маршрут {route} ({contract.number}): {end_date}")

                count += 1

        session.commit()
        print(f"\nИмпортировано/обновлено {count} записей")


def main():
    """Основная функция."""
    import argparse

    parser = argparse.ArgumentParser(description="Импорт дат окончания маршрутов")
    parser.add_argument("xlsx_path", help="Путь к Excel файлу")
    parser.add_argument("--contract", "-c", help="Номер контракта (если не указан - все)")

    args = parser.parse_args()

    import_route_end_dates(args.xlsx_path, args.contract)


if __name__ == "__main__":
    main()
