#!/usr/bin/env python3
"""
Скрипт импорта данных из xlsx файлов в БД.

Импортирует:
- 219.xlsx, 220.xlsx, 222.xlsx, 252.xlsx → route_params + route_trips
- holidays.xlsx → calendar_route_override
"""
import sys
from pathlib import Path
from datetime import datetime, date

import pandas as pd

# Добавляем корень проекта в path
sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import init_db, get_db_session
from db.models import (
    Contract, RouteParams, RouteTrips, CalendarRouteOverride, SeasonType
)
from core.constants import get_weekdays_for_day_type, parse_point_to_weekday


# Путь к старым данным
OLD_DATA_DIR = Path(__file__).parent.parent / "old"

# Контракты
CONTRACTS = ["219", "220", "222", "252"]


def parse_date(value) -> date | None:
    """Парсит дату из разных форматов."""
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        for fmt in ["%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"]:
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def parse_date_range(value: str) -> tuple[date | None, date | None]:
    """Парсит диапазон дат вида '01.01.2024-31.01.2024'."""
    if pd.isna(value) or not value:
        return None, None

    value = str(value).strip()
    if "-" in value and value.count("-") >= 2:
        # Может быть диапазон или одиночная дата
        parts = value.split("-")
        if len(parts) == 2:
            # Диапазон: 01.01.2024-31.01.2024
            return parse_date(parts[0]), parse_date(parts[1])
        elif len(parts) == 3 and len(parts[0]) == 2:
            # Одиночная дата: 01.01.2024 (разделитель - точка заменена)
            return parse_date(value), parse_date(value)

    # Одиночная дата
    d = parse_date(value)
    return d, d


def import_contracts(session):
    """Создаёт записи контрактов."""
    print("Импорт контрактов...")

    for num in CONTRACTS:
        existing = session.query(Contract).filter(Contract.number == num).first()
        if not existing:
            contract = Contract(
                number=num,
                date_from=date(2022, 4, 1),
                date_to=date(2028, 7, 14),
            )
            session.add(contract)
            print(f"  Создан контракт {num}")
        else:
            print(f"  Контракт {num} уже существует")

    session.flush()


def import_route_params(session, contract_number: str):
    """
    Импортирует параметры маршрутов из xlsx файла.

    Структура xlsx:
    - Маршрут, Дата, Всего, от НП, от КП, Источник, Номер ГК, Тип даты
    - Прямо - 1..7, Обратно - 1..7 (рейсы по дням недели)
    """
    xlsx_path = OLD_DATA_DIR / f"{contract_number}.xlsx"
    if not xlsx_path.exists():
        print(f"  Файл {xlsx_path} не найден, пропуск")
        return

    print(f"  Импорт из {xlsx_path.name}...")

    # Читаем Excel
    df = pd.read_excel(xlsx_path, sheet_name=0)

    # Получаем контракт
    contract = session.query(Contract).filter(Contract.number == contract_number).first()
    if not contract:
        print(f"    Контракт {contract_number} не найден!")
        return

    # Группируем рейсы по дням недели в типы
    # Столбцы: Прямо - 1, Обратно - 1, ... Прямо - 7, Обратно - 7
    trip_columns = {}
    for col in df.columns:
        col_str = str(col)
        if col_str.startswith("Прямо - ") or col_str.startswith("Обратно - "):
            direction = "forward" if "Прямо" in col_str else "reverse"
            weekday = int(col_str.split(" - ")[1])
            trip_columns[(direction, weekday)] = col

    count = 0
    for _, row in df.iterrows():
        route = str(row["Маршрут"]).strip()
        date_value = row["Дата"]
        date_from = parse_date(date_value)

        if not date_from or not route:
            continue

        # Пропускаем строки "Завершение" (конец контракта)
        date_type = str(row.get("Тип даты", "")).strip().lower()
        if date_type == "завершение":
            continue

        # Протяжённость
        length_total = row.get("Всего")
        length_forward = row.get("от НП")
        length_reverse = row.get("от КП")

        if pd.isna(length_forward):
            length_forward = None
        if pd.isna(length_reverse):
            length_reverse = None
        if pd.isna(length_total):
            length_total = None

        # Источник
        source = str(row.get("Источник", "")).strip()

        # Создаём RouteParams
        route_params = RouteParams(
            contract_id=contract.id,
            route=route,
            date_from=date_from,
            date_to=None,  # Действует до следующего изменения
            season=SeasonType.ALL_YEAR,
            length_total=float(length_total) if length_total else None,
            length_forward=float(length_forward) if length_forward else None,
            length_reverse=float(length_reverse) if length_reverse else None,
            source_appendix=source if source else None,
        )
        session.add(route_params)
        session.flush()

        # Собираем рейсы по типам дней
        # Анализируем, какие дни имеют одинаковые значения
        trips_by_weekday = {}
        for weekday in range(1, 8):
            fwd_col = trip_columns.get(("forward", weekday))
            rev_col = trip_columns.get(("reverse", weekday))

            fwd = row.get(fwd_col, 0) if fwd_col else 0
            rev = row.get(rev_col, 0) if rev_col else 0

            if pd.isna(fwd):
                fwd = 0
            if pd.isna(rev):
                rev = 0

            trips_by_weekday[weekday] = (int(fwd), int(rev))

        # Группируем дни с одинаковыми значениями
        grouped = {}
        for weekday, (fwd, rev) in trips_by_weekday.items():
            key = (fwd, rev)
            if key not in grouped:
                grouped[key] = []
            grouped[key].append(weekday)

        # Создаём RouteTrips для каждой группы
        for (fwd, rev), weekdays in grouped.items():
            if fwd == 0 and rev == 0:
                continue  # Пропускаем пустые

            # Определяем название типа дня
            weekdays_sorted = sorted(weekdays)
            if weekdays_sorted == [1, 2, 3, 4, 5]:
                day_type_name = "Рабочие дни"
            elif weekdays_sorted == [1, 2, 3, 4]:
                day_type_name = "Рабочие дни кроме пятницы"
            elif weekdays_sorted == [5]:
                day_type_name = "Пятница"
            elif weekdays_sorted == [6]:
                day_type_name = "Субботние дни"
            elif weekdays_sorted == [7]:
                day_type_name = "Воскресные и праздничные дни"
            elif weekdays_sorted == [6, 7]:
                day_type_name = "Выходные и праздничные дни"
            elif weekdays_sorted == [1, 2, 3, 4, 5, 6]:
                day_type_name = "Рабочие и субботние дни"
            elif weekdays_sorted == [1, 2, 3, 4, 5, 6, 7]:
                day_type_name = "Все дни"
            else:
                day_type_name = f"Дни {weekdays_sorted}"

            route_trip = RouteTrips(
                route_params_id=route_params.id,
                day_type_name=day_type_name,
                weekdays=weekdays_sorted,
                forward_number=fwd,
                reverse_number=rev,
            )
            session.add(route_trip)

        count += 1

    print(f"    Импортировано {count} записей параметров маршрутов")


def import_holidays(session):
    """
    Импортирует holidays.xlsx → calendar_route_override.

    Структура:
    - Дата (одиночная или диапазон)
    - Маршруты (через запятую)
    - День недели (как считать)
    - ДС (номер)
    """
    xlsx_path = OLD_DATA_DIR / "holidays.xlsx"
    if not xlsx_path.exists():
        print(f"Файл {xlsx_path} не найден, пропуск")
        return

    print(f"Импорт из {xlsx_path.name}...")

    xl = pd.ExcelFile(xlsx_path)

    for sheet_name in xl.sheet_names:
        if sheet_name == "Sheet1":
            continue

        contract = session.query(Contract).filter(Contract.number == sheet_name).first()
        if not contract:
            print(f"  Контракт {sheet_name} не найден, пропуск листа")
            continue

        print(f"  Лист {sheet_name}...")

        df = pd.read_excel(xl, sheet_name=sheet_name)
        count = 0

        for _, row in df.iterrows():
            date_str = str(row.get("Дата", ""))
            routes_str = str(row.get("Маршруты", ""))
            weekday = row.get("День недели")
            ds_num = row.get("ДС")

            if pd.isna(weekday) or not routes_str:
                continue

            weekday = int(weekday)

            # Парсим даты (одиночная или диапазон)
            date_from, date_to = parse_date_range(date_str)
            if not date_from:
                continue

            # Парсим маршруты
            routes = [r.strip() for r in routes_str.split(",") if r.strip()]

            # Генерируем записи для каждой даты и маршрута
            current_date = date_from
            while current_date <= (date_to or date_from):
                for route in routes:
                    # Проверяем, не существует ли уже
                    existing = session.query(CalendarRouteOverride).filter(
                        CalendarRouteOverride.date == current_date,
                        CalendarRouteOverride.contract_id == contract.id,
                        CalendarRouteOverride.route == route,
                    ).first()

                    if not existing:
                        override = CalendarRouteOverride(
                            date=current_date,
                            contract_id=contract.id,
                            route=route,
                            treat_as=weekday,
                            source_text=f"ДС №{ds_num}" if not pd.isna(ds_num) else None,
                        )
                        session.add(override)
                        count += 1

                # Следующий день
                from datetime import timedelta
                current_date = current_date + timedelta(days=1)

        print(f"    Импортировано {count} записей календаря")


def main():
    """Основная функция импорта."""
    print("=" * 60)
    print("Импорт данных из xlsx в БД")
    print("=" * 60)

    # Инициализация БД
    print("\nИнициализация БД...")
    init_db()

    with get_db_session() as session:
        # Импорт контрактов
        import_contracts(session)

        # Импорт параметров маршрутов
        print("\nИмпорт параметров маршрутов...")
        for contract_num in CONTRACTS:
            import_route_params(session, contract_num)

        # Импорт holidays
        print("\nИмпорт календаря (holidays)...")
        import_holidays(session)

        session.commit()

    print("\n" + "=" * 60)
    print("Импорт завершён!")
    print("=" * 60)


if __name__ == "__main__":
    main()
