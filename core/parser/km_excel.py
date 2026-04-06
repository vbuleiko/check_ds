"""
Парсер Excel-файлов с километрами по маршрутам.

Парсит файлы "Приложение №12/13" для ГК219, ГК220, ГК222.
Читает лист "Приложение 3 (поквартально)", который содержит
поквартальные данные в начале и помесячные данные далее.
"""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import re

import pandas as pd


# Маппинг месяцев
MONTH_NAMES = {
    "январь": 1, "февраль": 2, "март": 3, "апрель": 4,
    "май": 5, "июнь": 6, "июль": 7, "август": 8,
    "сентябрь": 9, "октябрь": 10, "ноябрь": 11, "декабрь": 12,
}


@dataclass
class MonthlyKm:
    """Данные км за один период (месяц или квартал)."""
    year: int
    month: int  # начальный месяц (1-12)
    period_start: int  # день начала (1)
    period_end: int  # день конца (30, 31, 28, 29)
    routes: dict[str, float] = field(default_factory=dict)  # {route: km}
    total: Optional[float] = None
    month_end: Optional[int] = None  # конечный месяц для квартальных периодов


@dataclass
class KmData:
    """Данные км из Excel-файла."""
    contract: Optional[str] = None  # "219", "220", "222"
    ds_number: Optional[str] = None
    appendix_number: Optional[str] = None  # "12" или "13"
    monthly: list[MonthlyKm] = field(default_factory=list)
    grand_total: Optional[float] = None  # Итого км по всем маршрутам и периодам

    def to_dict(self) -> dict:
        """Преобразует в словарь для JSON."""
        return {
            "contract": self.contract,
            "ds_number": self.ds_number,
            "appendix_number": self.appendix_number,
            "grand_total": self.grand_total,
            "monthly": [
                {
                    "date_from": f"{m.year}-{m.month:02d}-{m.period_start:02d}",
                    "date_to": (
                        f"{m.year}-{m.month_end:02d}-{m.period_end:02d}"
                        if m.month_end is not None
                        else f"{m.year}-{m.month:02d}-{m.period_end:02d}"
                    ),
                    "year": m.year,
                    "month": m.month,
                    "month_end": m.month_end,
                    "period_start": m.period_start,
                    "period_end": m.period_end,
                    "routes": m.routes,
                    "total": m.total,
                }
                for m in self.monthly
            ],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "KmData":
        """Создаёт из словаря."""
        if not data:
            return cls()
        km_data = cls(
            contract=data.get("contract"),
            ds_number=data.get("ds_number"),
            appendix_number=data.get("appendix_number"),
            grand_total=data.get("grand_total"),
        )
        for m in data.get("monthly", []):
            km_data.monthly.append(MonthlyKm(
                year=m["year"],
                month=m["month"],
                period_start=m["period_start"],
                period_end=m["period_end"],
                routes=m.get("routes", {}),
                total=m.get("total"),
                month_end=m.get("month_end"),
            ))
        return km_data


def detect_contract_and_ds(df: pd.DataFrame) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Определяет номер приложения, контракта и ДС из заголовка.

    Ищет строки вида:
    - "Приложение №12 к дополнительному соглашению №53"
    - "к Государственному контракту №01722000025210002220001"

    Returns:
        (contract, ds_number, appendix_number)
    """
    contract = None
    ds_number = None
    appendix_number = None

    # Ищем в первых 10 строках
    for i in range(min(10, len(df))):
        for j in range(len(df.columns)):
            val = df.iloc[i, j]
            if pd.isna(val):
                continue
            val = str(val)

            # Ищем номер приложения
            app_match = re.search(r'приложение\s*№\s*(\d+)', val, re.IGNORECASE)
            if app_match and appendix_number is None:
                appendix_number = app_match.group(1)

            # Ищем номер ДС
            ds_match = re.search(r'соглашению\s*№\s*(\d+)', val, re.IGNORECASE)
            if ds_match:
                ds_number = ds_match.group(1)

            # Ищем номер контракта
            contract_match = re.search(r'контракту?\s*№\s*\d*(\d{3})0001', val, re.IGNORECASE)
            if contract_match:
                contract = contract_match.group(1)

    return contract, ds_number, appendix_number


def is_column_number(val_str: str) -> bool:
    """
    Проверяет, является ли значение номером колонки (1, 2, 3, ..., 76).

    Номера колонок - это последовательные целые числа в первой строке.
    """
    try:
        num = int(val_str)
        return 1 <= num <= 80
    except (ValueError, TypeError):
        return False


def is_km_value(val_str: str) -> bool:
    """
    Проверяет, является ли значение километражом (число с точкой > 100).
    """
    if '.' not in val_str:
        return False
    try:
        num = float(val_str)
        return num > 100  # км обычно > 100
    except (ValueError, TypeError):
        return False


def is_service_word(val_str: str) -> bool:
    """Проверяет, является ли значение служебным словом."""
    skip_words = ["итого", "всего", "большой", "средний", "малый", "маршрут", "км", "nan", "большой-1"]
    return val_str.lower().strip() in skip_words


def find_routes_row(df: pd.DataFrame) -> tuple[int, dict[int, str]]:
    """
    Находит строку с номерами маршрутов.

    Логика: ищем строку, где много значений похожи на номера маршрутов
    (содержат цифры, не являются км с точкой, не служебные слова).

    Returns:
        (row_index, {col_index: route_name})
    """
    best_row = -1
    best_routes = {}
    best_unique_count = 0

    # Сначала найдём строку с номерами колонок (1, 2, 3, ...)
    column_numbers_row = -1
    for i in range(min(5, len(df))):
        col_nums = 0
        for j in range(4, min(50, len(df.columns))):
            val = df.iloc[i, j]
            if pd.isna(val):
                continue
            if is_column_number(str(val).strip()):
                col_nums += 1
        if col_nums > 20:
            column_numbers_row = i
            break

    # Теперь ищем строку с маршрутами (после строки с номерами колонок)
    start_row = column_numbers_row + 1 if column_numbers_row >= 0 else 0

    for i in range(start_row, min(15, len(df))):
        routes = {}
        for j in range(4, len(df.columns)):
            val = df.iloc[i, j]
            if pd.isna(val):
                continue
            val_str = str(val).strip()

            # Пропускаем служебные слова
            if is_service_word(val_str):
                continue

            # Пропускаем км (числа с точкой > 100)
            if is_km_value(val_str):
                continue

            # Пропускаем пустые и слишком длинные
            if len(val_str) > 15 or len(val_str) == 0:
                continue

            # Должны быть цифры
            if not any(c.isdigit() for c in val_str):
                continue

            # Это маршрут
            routes[j] = val_str

        # Если нашли много маршрутов (> 15), это строка маршрутов.
        # При равном кол-ве предпочитаем строку с более разнообразными значениями
        # (чтобы выбрать реальные номера маршрутов, а не коды типов вроде 49/80).
        if len(routes) > 15:
            unique_count = len(set(routes.values()))
            is_better = (
                len(routes) > len(best_routes) or
                (len(routes) == len(best_routes) and unique_count > best_unique_count)
            )
            if is_better:
                best_row = i
                best_routes = routes
                best_unique_count = unique_count

    return best_row, best_routes


def find_data_start_row(df: pd.DataFrame, routes_row: int) -> int:
    """Находит первую строку с данными (после заголовков)."""
    # Данные начинаются после строки с маршрутами
    # Ищем строку, где в первых колонках есть числа (№, Год)
    for i in range(routes_row + 1, min(routes_row + 5, len(df))):
        val0 = df.iloc[i, 0]  # №
        val1 = df.iloc[i, 1]  # Год
        if not pd.isna(val0) and not pd.isna(val1):
            try:
                int(val0)  # № должен быть числом
                year = int(val1)
                if 2020 <= year <= 2030:
                    return i
            except (ValueError, TypeError):
                continue
    return routes_row + 1


def parse_quarterly_period(period_str: str) -> tuple[int, int, int, int]:
    """
    Парсит квартальный период из строки "01.04-30.06".

    Returns:
        (start_day, start_month, end_day, end_month)
    """
    period_str = str(period_str).strip()
    match = re.match(r'^(\d+)\.(\d+)-(\d+)\.(\d+)$', period_str)
    if match:
        return (
            int(match.group(1)),  # start_day
            int(match.group(2)),  # start_month
            int(match.group(3)),  # end_day
            int(match.group(4)),  # end_month
        )
    return 0, 0, 0, 0


def parse_period(period_str: str) -> tuple[int, int]:
    """
    Парсит период из строки.

    Примеры: "1-30", "1-31", "01.04-30.06"
    """
    if not period_str:
        return 1, 31

    period_str = str(period_str).strip()

    # Формат "1-30" или "1-31"
    match = re.match(r'^(\d+)-(\d+)$', period_str)
    if match:
        return int(match.group(1)), int(match.group(2))

    # Формат "01.04-30.06"
    match = re.match(r'^(\d+)\.(\d+)-(\d+)\.(\d+)$', period_str)
    if match:
        return int(match.group(1)), int(match.group(3))

    return 1, 31


def parse_month_name(month_str: str) -> Optional[int]:
    """Парсит название месяца в номер."""
    if not month_str:
        return None
    month_lower = str(month_str).lower().strip()
    return MONTH_NAMES.get(month_lower)


def _find_km_sheet(xl: pd.ExcelFile) -> Optional[str]:
    """
    Ищет лист с данными км.

    Приоритет: "поквартально" > "помесячно"
    """
    for sheet_name in xl.sheet_names:
        if "поквартально" in sheet_name.lower() or "квартально" in sheet_name.lower():
            return sheet_name

    for sheet_name in xl.sheet_names:
        if "помесячно" in sheet_name.lower():
            return sheet_name

    return None


def parse_km_excel(file_path: Path) -> KmData:
    """
    Парсит Excel-файл с километрами.

    Читает лист "Приложение 3 (поквартально)" (приоритет) или "помесячно".
    Лист может содержать поквартальные периоды ("01.04-30.06") и/или
    помесячные ("Март", "1-31"). Оба формата поддерживаются.

    Динамически определяет структуру файла через find_routes_row().
    """
    result = KmData()

    try:
        xl = pd.ExcelFile(file_path)
    except Exception as e:
        raise ValueError(f"Не удалось открыть Excel файл: {e}")

    try:
        target_sheet = _find_km_sheet(xl)
        if not target_sheet:
            raise ValueError("Не найден лист 'Приложение 3 (поквартально)' или '(помесячно)'")

        df = pd.read_excel(xl, sheet_name=target_sheet, header=None)

        # Определяем контракт, ДС и номер приложения
        result.contract, result.ds_number, result.appendix_number = detect_contract_and_ds(df)

        # Динамически определяем структуру файла
        routes_row, routes_map = find_routes_row(df)
        if routes_row < 0 or not routes_map:
            raise ValueError("Не удалось найти строку с номерами маршрутов")

        data_start = find_data_start_row(df, routes_row)

        # Определяем начальный столбец маршрутов (минимальный индекс в routes_map)
        routes_start_col = min(routes_map.keys())

        # Столбцы: год всегда в колонке 1, период/месяц в колонке 2
        year_col = 1
        period_or_month_col = 2
        # В помесячном листе может быть отдельная колонка с периодом (1-31)
        # В поквартальном — период в колонке 2 ("01.04-30.06" или "Март")
        extra_period_col = 3  # Для помесячных данных: "1-31" в колонке 3

        # Находим колонку "Итого" — ищем в нескольких строках вокруг routes_row
        itogo_col = None
        itogo_search_rows = list(dict.fromkeys([
            routes_row,
            routes_row - 1,
            routes_row - 2,
            routes_row + 1,
        ]))
        for search_row in itogo_search_rows:
            if search_row < 0 or search_row >= len(df):
                continue
            for j in range(len(df.columns) - 1, routes_start_col, -1):
                val = df.iloc[search_row, j]
                if pd.notna(val) and str(val).lower().strip() in ["итого", "всего км", "итого км"]:
                    itogo_col = j
                    break
            if itogo_col is not None:
                break

        # Парсим данные
        for i in range(data_start, len(df)):
            row = df.iloc[i]

            # Проверяем, есть ли данные в строке
            year_val = row.iloc[year_col] if year_col < len(row) else None
            if pd.isna(year_val):
                continue

            try:
                year = int(year_val)
                if year < 2020 or year > 2030:
                    continue
            except (ValueError, TypeError):
                continue

            # Определяем тип периода из колонки 2
            period_val = row.iloc[period_or_month_col] if period_or_month_col < len(row) else None
            if pd.isna(period_val):
                continue

            period_str = str(period_val).strip()

            # ГК220: col2 содержит номер квартала (1-4), а период — в col3.
            # Пример: col2="2", col3="01.04-30.06"
            if re.match(r'^[1-4]$', period_str):
                col3_val = row.iloc[extra_period_col] if extra_period_col < len(row) else None
                if pd.isna(col3_val):
                    continue
                period_str = str(col3_val).strip()

            # Пробуем как квартальный/помесячный период в формате "дд.мм-дд.мм"
            # Этот формат используется и для кварталов ("01.04-30.06")
            # и для месяцев в ГК220 ("01.10-31.10")
            sd, sm, ed, em = parse_quarterly_period(period_str)
            if sm > 0 and em > 0:
                month = sm
                month_end = em if em != sm else None
                period_start = sd
                period_end = ed
            else:
                # Пробуем как название месяца (ГК219/222 помесячные строки)
                month = parse_month_name(period_str)
                if month is None:
                    continue
                month_end = None
                # Период из следующей колонки (3)
                extra_val = row.iloc[extra_period_col] if extra_period_col < len(row) else None
                if pd.notna(extra_val):
                    extra_str = str(extra_val).strip()
                    # Проверяем формат "дд.мм-дд.мм" (ГК220 помесячные строки)
                    sd2, _sm2, ed2, _em2 = parse_quarterly_period(extra_str)
                    if sd2 > 0:
                        period_start, period_end = sd2, ed2
                    else:
                        period_start, period_end = parse_period(extra_str)
                else:
                    period_start, period_end = 1, 31

            # Км по маршрутам
            routes_km = {}
            for col, route in routes_map.items():
                if col < len(row):
                    km_val = row.iloc[col]
                    if pd.notna(km_val):
                        try:
                            km = float(km_val)
                            if km > 0:
                                if route in routes_km:
                                    routes_km[route] = round(routes_km[route] + km, 2)
                                else:
                                    routes_km[route] = round(km, 2)
                        except (ValueError, TypeError):
                            pass

            # Итого
            total = None
            if itogo_col and itogo_col < len(row):
                total_val = row.iloc[itogo_col]
                if pd.notna(total_val):
                    try:
                        total = float(total_val)
                    except (ValueError, TypeError):
                        pass
            # Fallback: считаем как сумму маршрутов
            if total is None and routes_km:
                total = round(sum(routes_km.values()), 2)

            if routes_km:
                result.monthly.append(MonthlyKm(
                    year=year,
                    month=month,
                    period_start=period_start,
                    period_end=period_end,
                    routes=routes_km,
                    total=total,
                    month_end=month_end,
                ))

        # Парсим строку "Итого" — ищем снизу вверх
        for i in range(len(df) - 1, data_start, -1):
            val0 = df.iloc[i, 0]
            if pd.notna(val0) and "итого" in str(val0).lower().strip():
                # Суммируем все колонки маршрутов в строке Итого
                grand_total = 0.0
                for col in routes_map.keys():
                    if col < len(df.columns):
                        v = df.iloc[i, col]
                        if pd.notna(v):
                            try:
                                grand_total += float(v)
                            except (ValueError, TypeError):
                                pass
                result.grand_total = round(grand_total, 2)
                break

        # Fallback: если строки "Итого" нет — суммируем period.total
        if result.grand_total is None and result.monthly:
            result.grand_total = round(
                sum(m.total for m in result.monthly if m.total is not None), 2
            )

        return result
    finally:
        xl.close()


def is_inside_appendix_folder(file_path: Path, root_dir: Path) -> bool:
    """
    Проверяет, находится ли файл внутри папки конкретного приложения
    (Приложение №1, Приложение №2 и т.д. — с номером).
    Папка «Приложения» (без номера) не считается исключением.
    """
    try:
        rel_path = file_path.relative_to(root_dir)
    except ValueError:
        return False

    for part in rel_path.parts[:-1]:
        part_lower = part.lower()
        if "приложение" in part_lower:
            match = re.search(r'№?\s*(\d+)', part)
            if match:
                return True
    return False


def has_km_sheet(file_path: Path) -> bool:
    """Проверяет, есть ли в файле лист с данными км."""
    xl = None
    try:
        xl = pd.ExcelFile(file_path)
        for sheet_name in xl.sheet_names:
            lower = sheet_name.lower()
            if "поквартально" in lower or "квартально" in lower or "помесячно" in lower:
                return True
    except Exception:
        pass
    finally:
        if xl is not None:
            xl.close()
    return False


def find_km_excel_in_directory(directory: Path) -> Optional[Path]:
    """
    Ищет файл с км в директории.

    Файл км:
    - Имеет "Приложение" в названии
    - Содержит лист "Приложение 3 (поквартально)" или "(помесячно)"
    - НЕ находится внутри папки отдельного приложения (Приложение №1, №2, ...)
    """
    candidates = []

    for path in directory.rglob("*.xlsx"):
        # Пропускаем временные файлы
        if path.name.startswith("~$"):
            continue

        # Пропускаем файлы внутри папок конкретных приложений (Приложение №X)
        if is_inside_appendix_folder(path, directory):
            continue

        candidates.append(path)

    # Проверяем кандидатов на наличие нужного листа
    for path in candidates:
        if has_km_sheet(path):
            return path

    return None
