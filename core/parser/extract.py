"""
Основной модуль извлечения данных из архива ДС.

Извлекает:
- general: общая информация (суммы, пробеги)
- change_with_money: изменения с изменением стоимости
- change_without_money: изменения без изменения стоимости
- appendices: приложения с параметрами маршрутов
"""
import re
from pathlib import Path
from dataclasses import dataclass, field

from .archive import extract_archive, find_files_by_extension
from .docx_parser import (
    get_docx_data, extract_number, extract_date, extract_date_range,
    find_table_by_header
)


@dataclass
class GeneralInfo:
    """Общая информация из ДС."""
    ds_number: str | None = None
    contract_number: str | None = None
    sum_text: float | None = None
    sum_finansirovanie_text: float | None = None
    sum_etapy: float | None = None
    sum_etapy_avans: float | None = None
    sum_finansirovanie_table: float | None = None
    probeg_sravnenie: float | None = None
    probeg_etapy: float | None = None
    probeg_etapy_avans: float | None = None


@dataclass
class RouteChange:
    """Изменение маршрута."""
    appendix: str | None = None
    route: str | None = None
    date_from: str | None = None
    date_to: str | None = None
    date_on: str | None = None
    point: str | None = None  # Для изменений без приложения


@dataclass
class AppendixData:
    """Данные приложения."""
    appendix_num: str | None = None
    ds_num: str | None = None
    route: str | None = None
    date_from: str | None = None
    date_to: str | None = None
    date_on: str | None = None
    length_forward: float | None = None
    length_reverse: float | None = None
    length_sum: float | None = None
    num_of_types: int = 0
    # Динамические поля для type_N_*
    types: dict = field(default_factory=dict)
    # Сезонные периоды
    period_winter: dict | None = None
    period_summer: dict | None = None


@dataclass
class ExtractResult:
    """Результат извлечения данных из ДС."""
    general: GeneralInfo = field(default_factory=GeneralInfo)
    change_with_money: list[RouteChange] = field(default_factory=list)
    change_with_money_no_appendix: list[RouteChange] = field(default_factory=list)
    change_without_money: list[RouteChange] = field(default_factory=list)
    change_without_money_no_appendix: list[RouteChange] = field(default_factory=list)
    appendices: dict[str, dict] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    # Сырые таблицы для JSON (не используются в проверке)
    table_raschet_izm_objema: list[list[str]] | None = None
    table_etapy_sroki: list[list[str]] | None = None
    table_finansirovanie_po_godam: list[list[str]] | None = None
    table_etapy_avans: list[list[str]] | None = None
    table_objemy_rabot: list[list[str]] | None = None

    def to_dict(self) -> dict:
        """Преобразует результат в словарь."""
        return {
            "general": {
                "ds_number": self.general.ds_number,
                "contract_number": self.general.contract_number,
                "sum_text": self.general.sum_text,
                "sum_finansirovanie_text": self.general.sum_finansirovanie_text,
                "sum_etapy": self.general.sum_etapy,
                "sum_etapy_avans": self.general.sum_etapy_avans,
                "sum_finansirovanie_table": self.general.sum_finansirovanie_table,
                "probeg_sravnenie": self.general.probeg_sravnenie,
                "probeg_etapy": self.general.probeg_etapy,
                "probeg_etapy_avans": self.general.probeg_etapy_avans,
            },
            "change_with_money": [
                {k: v for k, v in c.__dict__.items() if k != "point"}
                for c in self.change_with_money
            ],
            "change_with_money_no_appendix": [
                {k: v for k, v in c.__dict__.items()}
                for c in self.change_with_money_no_appendix
            ],
            "change_without_money": [
                {k: v for k, v in c.__dict__.items() if k != "point"}
                for c in self.change_without_money
            ],
            "change_without_money_no_appendix": [
                {k: v for k, v in c.__dict__.items()}
                for c in self.change_without_money_no_appendix
            ],
            "appendices": self.appendices,
            "table_raschet_izm_objema": self.table_raschet_izm_objema,
            "table_etapy_sroki": self.table_etapy_sroki,
            "table_finansirovanie_po_godam": self.table_finansirovanie_po_godam,
            "table_etapy_avans": self.table_etapy_avans,
            "table_objemy_rabot": self.table_objemy_rabot,
        }


def parse_main_document(file_path: Path) -> tuple[GeneralInfo, list[RouteChange], list[RouteChange], list[RouteChange], list[RouteChange]]:
    """
    Парсит основной документ ДС.

    Returns:
        (general_info, changes_with_money, changes_with_money_no_appendix, changes_without_money, changes_without_money_no_appendix)
    """
    paragraphs, tables = get_docx_data(file_path)
    general = GeneralInfo()
    changes_with_money = []
    changes_with_money_no_appendix = []
    changes_without_money = []
    changes_without_money_no_appendix = []

    full_text = "\n".join(paragraphs)

    # Извлекаем номер ДС
    match = re.search(r"дополнительн\w+\s+соглашени\w+\s*№?\s*(\d+)", full_text.lower())
    if match:
        general.ds_number = match.group(1)

    # Извлекаем номер контракта
    match = re.search(r"(\d{20,})", full_text)
    if match:
        general.contract_number = match.group(1)

    # Извлекаем суммы и пробеги из текста
    for para in paragraphs:
        para_lower = para.lower()

        # Сумма контракта
        if "цена контракта" in para_lower or "стоимость" in para_lower:
            num = extract_number(para)
            if num and num > 1000000:
                if general.sum_text is None:
                    general.sum_text = num

        # Пробег
        if "километр" in para_lower or "пробег" in para_lower:
            num = extract_number(para)
            if num and num > 10000:
                if general.probeg_sravnenie is None:
                    general.probeg_sravnenie = num

    # Извлекаем значения из таблиц (суммы и пробеги)
    _extract_values_from_tables(tables, general)

    # Извлекаем изменения из параграфов (основной метод)
    changes_with_money, changes_with_money_no_appendix, changes_without_money, changes_without_money_no_appendix = _extract_changes_from_paragraphs(
        paragraphs, general.contract_number
    )

    return general, changes_with_money, changes_with_money_no_appendix, changes_without_money, changes_without_money_no_appendix


def _parse_changes_table(table: list[list[str]]) -> list[RouteChange]:
    """Парсит таблицу изменений маршрутов."""
    changes = []

    # Пропускаем заголовок
    for row in table[1:]:
        if len(row) < 3:
            continue

        # Ищем номер приложения, маршрут и даты
        appendix = None
        route = None
        date_from = None
        date_to = None
        date_on = None
        point = None

        for cell in row:
            cell_text = cell.strip()

            # Номер приложения
            match = re.search(r"приложени\w*\s*№?\s*(\d+)", cell_text.lower())
            if match:
                appendix = match.group(1)

            # Номер маршрута
            match = re.search(r"маршрут\w*\s*№?\s*([А-Яа-яA-Za-z0-9\.]+)", cell_text)
            if match:
                route = match.group(1).upper()

            # Даты
            df, dt, do = extract_date_range(cell_text)
            if df:
                date_from = df
            if dt:
                date_to = dt
            if do:
                date_on = do

            # Текст "по графику движения..."
            if "по графику движения" in cell_text.lower():
                point = cell_text

        if route:
            change = RouteChange(
                appendix=appendix,
                route=route,
                date_from=date_from,
                date_to=date_to,
                date_on=date_on,
                point=point,
            )
            changes.append(change)

    return changes


def _extract_sums_from_table(table: list[list[str]], general: GeneralInfo):
    """Извлекает суммы из таблицы этапов/финансирования."""
    for row in table:
        row_text = " ".join(row).lower()

        # Ищем строку "Итого"
        if "итого" in row_text:
            for cell in row:
                num = extract_number(cell)
                if num:
                    if num > 1000000000:  # Сумма
                        if general.sum_etapy is None:
                            general.sum_etapy = num
                    elif num > 1000000:  # Пробег
                        if general.probeg_etapy is None:
                            general.probeg_etapy = num


def _extract_values_from_tables(tables: list[list[list[str]]], general: GeneralInfo):
    """
    Извлекает значения probeg и sum из разных таблиц документа.
    Аналогично функции extract_values_from_tables из старого скрипта.
    """
    etapy_table_count = 0

    for table in tables:
        if len(table) < 2:
            continue

        header_row = table[0]

        # 1. Таблица "Расчет изменения объема" -> probeg_sravnenie
        if general.probeg_sravnenie is None:
            for col_idx, cell in enumerate(header_row):
                cell_lower = cell.lower()
                if "предлагаемые изменения" in cell_lower and "км" in cell_lower:
                    # Значение в строке 1 (первая строка данных после заголовка)
                    if len(table) > 1 and col_idx < len(table[1]):
                        general.probeg_sravnenie = extract_number(table[1][col_idx])
                    break

        # 2. Таблица "Финансирование по годам" -> sum_finansirovanie_table
        if general.sum_finansirovanie_table is None:
            for col_idx, cell in enumerate(header_row):
                cell_lower = cell.lower()
                if "финансирование по годам" in cell_lower:
                    # Ищем строку "Итого"
                    for row in table:
                        if row and "итого" in row[0].lower():
                            if col_idx < len(row):
                                general.sum_finansirovanie_table = extract_number(row[col_idx])
                            break
                    break

        # 3-4. Таблицы "Этапы исполнения Контракта" -> probeg_etapy/sum_etapy и probeg_etapy_avans/sum_etapy_avans
        # Определяем по наличию столбца "Максимальная транспортная работа" и строки ИТОГО
        # Пропускаем таблицу "Расчет изменения объема" (она имеет столбец "Предлагаемые изменения")
        is_sravnenie_table = any("предлагаемые изменения" in cell.lower() for cell in header_row)
        if is_sravnenie_table:
            continue

        probeg_col_idx = None
        sum_col_idx = None
        sum_avans_col_idx = None

        for col_idx, cell in enumerate(header_row):
            cell_lower = cell.lower()
            if "максимальная транспортная работа" in cell_lower:
                probeg_col_idx = col_idx
            if "стоимость транспортной работы" in cell_lower and "руб" in cell_lower:
                if "с учетом выплаченных авансов" in cell_lower or "с учётом выплаченных авансов" in cell_lower:
                    sum_avans_col_idx = col_idx
                elif sum_col_idx is None:
                    sum_col_idx = col_idx

        # Таблица этапов должна содержать И пробег И стоимость
        if probeg_col_idx is not None and sum_col_idx is not None:
            # Ищем строку "Итого"
            for row in table:
                if row and "итого" in row[0].lower():
                    etapy_table_count += 1
                    if etapy_table_count == 1:
                        # Первая таблица -> probeg_etapy, sum_etapy
                        if probeg_col_idx < len(row):
                            general.probeg_etapy = extract_number(row[probeg_col_idx])
                        if sum_col_idx < len(row):
                            general.sum_etapy = extract_number(row[sum_col_idx])
                    elif etapy_table_count == 2:
                        # Вторая таблица -> probeg_etapy_avans, sum_etapy_avans
                        if probeg_col_idx < len(row):
                            general.probeg_etapy_avans = extract_number(row[probeg_col_idx])
                        if sum_avans_col_idx is not None and sum_avans_col_idx < len(row):
                            general.sum_etapy_avans = extract_number(row[sum_avans_col_idx])
                    break


def _extract_changes_from_paragraphs(paragraphs: list[str], contract_number: str | None = None) -> tuple:
    """
    Извлекает данные об изменениях из пунктов документа (аналогично старому скрипту).
    """
    change_with_money = []
    change_with_money_no_appendix = []
    change_without_money = []
    change_without_money_no_appendix = []

    section1_start = None
    section1_end = None
    section2_start = None
    section2_end = None
    route_120_section_start = None
    route_120_section_end = None

    # Специальная обработка для контракта 2520001 (маршрут 120)
    is_route_120_contract = contract_number and contract_number.endswith("2520001")

    for i, text in enumerate(paragraphs):
        if "Стороны пришли к соглашению выполнять работы:" in text and section1_start is None:
            section1_start = i
        elif "Стороны пришли к соглашению выполнять работы без изменения" in text and section2_start is None:
            section2_start = i
        # Специальный случай для маршрута 120
        elif is_route_120_contract and "Стороны пришли к соглашению выполнять работы по маршруту № 120" in text:
            route_120_section_start = i

    # Определяем конец секции 1
    if section1_start is not None:
        if section2_start is not None:
            section1_end = section2_start
        else:
            for i in range(section1_start + 1, len(paragraphs)):
                if "Стороны пришли к соглашению" in paragraphs[i]:
                    section1_end = i
                    break
            if section1_end is None:
                section1_end = len(paragraphs)

    # Определяем конец секции 2
    if section2_start is not None:
        for i in range(section2_start + 1, len(paragraphs)):
            if paragraphs[i].startswith("Стороны"):
                section2_end = i
                break
        if section2_end is None:
            section2_end = len(paragraphs)

    # Определяем конец секции маршрута 120
    if route_120_section_start is not None:
        for i in range(route_120_section_start + 1, len(paragraphs)):
            if re.match(r"^\d+\.\s+[А-ЯЁA-Z]", paragraphs[i]):
                route_120_section_end = i
                break
        if route_120_section_end is None:
            route_120_section_end = len(paragraphs)

    # Извлекаем данные из секции 1 (с изменением цены)
    if section1_start is not None and section1_end is not None:
        for i in range(section1_start + 1, section1_end):
            route_infos = _extract_route_info(paragraphs[i])
            for info in route_infos:
                if info.get("appendix"):
                    change_with_money.append(_dict_to_change(info))
                else:
                    change_with_money_no_appendix.append(_dict_to_change(info))

    # Извлекаем данные из секции 2 (без изменения цены)
    if section2_start is not None:
        end = section2_end if section2_end else len(paragraphs)
        for i in range(section2_start + 1, end):
            route_infos = _extract_route_info(paragraphs[i])
            for info in route_infos:
                if info.get("appendix"):
                    change_without_money.append(_dict_to_change(info))
                else:
                    change_without_money_no_appendix.append(_dict_to_change(info))

    # Извлекаем данные из секции маршрута 120 (специальный случай)
    if route_120_section_start is not None and route_120_section_end is not None:
        for i in range(route_120_section_start + 1, route_120_section_end):
            text = paragraphs[i]
            if not text.strip():
                continue
            if "по графику движения" in text.lower() and re.search(r"\d{2}\.\d{2}\.\d{4}", text):
                infos = _extract_route_120_subitem(text)
                for info in infos:
                    change_with_money_no_appendix.append(_dict_to_change(info))

    return change_with_money, change_with_money_no_appendix, change_without_money, change_without_money_no_appendix


def _extract_route_info(text: str) -> list[dict]:
    """Извлекает информацию о маршрутах из текста подпункта."""
    results = []
    routes = []

    if "маршрут" not in text.lower():
        return results

    # Список маршрутов может заканчиваться на " - по ", " по ", " согласно" или конец строки
    # Поддерживаем оба варианта: "по маршруту №" и "маршруту №" (без "по")
    route_match = re.search(r"(?:по\s+)?маршрут\w*\s*№\s*([\d\s,А-ЯЁа-яёA-Za-z]{1,100}?)(?:\s*-\s+по\s|\s+по\s|\s+согласно|\s*$)", text, re.IGNORECASE)
    if route_match:
        routes_str = route_match.group(1)
        route_parts = re.findall(r"(\d+[А-ЯЁа-яёA-Za-z]*)", routes_str)
        routes.extend(route_parts)

    appendix_match = re.search(r"Приложению\s*№\s*(\d+)", text)
    appendix_num = appendix_match.group(1) if appendix_match else None

    point = None
    if not appendix_num and routes:
        last_route = routes[-1]
        last_route_pos = text.rfind(last_route)
        if last_route_pos != -1:
            after_routes = text[last_route_pos + len(last_route):]
            point_match = re.search(r"[,\s]*(по\s+.+?)\.?\s*$", after_routes, re.IGNORECASE)
            if point_match:
                point = point_match.group(1).strip().rstrip('.')

    dates = extract_date_range(text)

    for route in routes:
        result = {
            "appendix": appendix_num,
            "route": route.upper(),
            "date_from": dates[0] if dates and len(dates) > 0 else None,
            "date_to": dates[1] if dates and len(dates) > 1 else None,
            "date_on": dates[2] if dates and len(dates) > 2 else None,
        }
        if point is not None:
            result["point"] = point
        results.append(result)

    return results


def _extract_route_120_subitem(text: str) -> list[dict]:
    """Извлекает информацию из подпункта для маршрута 120."""
    results = []

    text = re.sub(r"^\d{1,2}\.\d{1,2}\.(?!\d{4})\s*", "", text)

    parts = [p.strip() for p in text.split(";") if p.strip()]

    for part in parts:
        result = {
            "appendix": None,
            "route": "120",
            "date_from": None,
            "date_to": None,
            "date_on": None,
        }

        point_match = re.search(r"\s+-\s+(.+?)\.?\s*$", part)
        if point_match:
            result["point"] = point_match.group(1).strip().rstrip('.')

        range_match = re.search(r"с\s+(\d{2}\.\d{2}\.\d{4})\s+года\s+по\s+(\d{2}\.\d{2}\.\d{4})\s+года", part)
        if range_match:
            df, dt = range_match.groups()
            result["date_from"] = df[-4:] + "-" + df[3:5] + "-" + df[:2]
            result["date_to"] = dt[-4:] + "-" + dt[3:5] + "-" + dt[:2]
            results.append(result)
            continue

        on_match = re.match(r"^(\d{2}\.\d{2}\.\d{4})\s+года\s+-", part)
        if on_match:
            d = on_match.group(1)
            result["date_on"] = d[-4:] + "-" + d[3:5] + "-" + d[:2]
            results.append(result)
            continue

        if result.get("point"):
            results.append(result)

    return results


def _dict_to_change(d: dict) -> RouteChange:
    """Преобразует словарь в RouteChange."""
    return RouteChange(
        appendix=d.get("appendix"),
        route=d.get("route"),
        date_from=d.get("date_from"),
        date_to=d.get("date_to"),
        date_on=d.get("date_on"),
        point=d.get("point"),
    )


def parse_appendix(file_path: Path) -> AppendixData | None:
    """Парсит файл приложения с параметрами маршрута."""
    paragraphs, tables = get_docx_data(file_path)

    if not paragraphs:
        return None

    appendix = AppendixData()

    # Извлекаем номер приложения и номер ДС
    for para in paragraphs[:10]:
        match = re.search(r"приложени\w*\s*№?\s*(\d+)", para.lower())
        if match:
            appendix.appendix_num = match.group(1)
        ds_match = re.search(r"соглашени\w*\s*№\s*(\d+)", para, re.IGNORECASE)
        if ds_match:
            appendix.ds_num = ds_match.group(1)
        if appendix.appendix_num and appendix.ds_num:
            break

    # Извлекаем маршрут
    for para in paragraphs[:20]:
        match = re.search(r"маршрут\w*\s*№?\s*([А-Яа-яA-Za-z0-9\.]+)", para)
        if match:
            appendix.route = match.group(1).upper()
            break

    # Извлекаем даты
    for para in paragraphs[:20]:
        df, dt, do = extract_date_range(para)
        if df or do:
            appendix.date_from = df
            appendix.date_to = dt
            appendix.date_on = do
            break

    # Извлекаем протяжённость
    full_text = "\n".join(paragraphs)

    # Парсим общую протяжённость (всего)
    sum_match = re.search(
        r"протяж[её]нност[ьъ][^.]*?всего[:\s]+(\d+)[,.](\d+)\s*км",
        full_text,
        re.IGNORECASE
    )
    if sum_match:
        appendix.length_sum = float(f"{sum_match.group(1)}.{sum_match.group(2)}")

    # Сначала пробуем найти через формат "в прямом направлении: X,X км"
    np_match = re.search(
        r"в\s+прямом\s+направлении[:\s]+(\d+)[,.](\d+)\s*км",
        full_text,
        re.IGNORECASE
    )
    if np_match:
        appendix.length_forward = float(f"{np_match.group(1)}.{np_match.group(2)}")

    op_match = re.search(
        r"в\s+обратном\s+направлении[:\s]+(\d+)[,.](\d+)\s*км",
        full_text,
        re.IGNORECASE
    )
    if op_match:
        appendix.length_reverse = float(f"{op_match.group(1)}.{op_match.group(2)}")

    # Если не нашли, пробуем через ключевое слово "протяжённость"
    if appendix.length_forward is None or appendix.length_reverse is None:
        for para in paragraphs:
            para_lower = para.lower()
            if ("протяжённость" in para_lower or "протяженность" in para_lower):
                if appendix.length_forward is None and ("прямом" in para_lower or "от нп" in para_lower):
                    num = extract_number(para)
                    if num:
                        appendix.length_forward = num
                elif appendix.length_reverse is None and ("обратном" in para_lower or "от кп" in para_lower):
                    num = extract_number(para)
                    if num:
                        appendix.length_reverse = num

    # Если length_sum не нашли через регулярку, ищем по параграфам
    if appendix.length_sum is None:
        for para in paragraphs:
            para_lower = para.lower()
            if ("протяжённость" in para_lower or "протяженность" in para_lower) and "всего" in para_lower:
                num = extract_number(para)
                if num:
                    appendix.length_sum = num
                    break

    # Парсим таблицу рейсов
    trips_table = find_table_by_header(tables, ["количество", "рейс"])
    if trips_table:
        _parse_trips_table(trips_table, appendix)

    return appendix


def _parse_trips_table(table: list[list[str]], appendix: AppendixData):
    """
    Парсит таблицу количества рейсов.

    Ожидаемая структура таблицы:
    | Направление | ... | Тип дня 1 |           | Тип дня 2 |           |
    |             |     | Кол-во    | Пробег    | Кол-во    | Пробег    |
    | Прямое      |     | 12        | 289.8     | 10        | 241.5     |
    | Обратное    |     | 12        | 297.96    | 10        | 247.96    |
    | ИТОГО       |     | 24        | 587.76    | 20        | 489.46    |
    """
    if len(table) < 3:
        return

    # Ищем заголовки столбцов (типы дней)
    header_row = None
    for i, row in enumerate(table[:5]):
        row_text = " ".join(row).lower()
        if "рабочие" in row_text or "выходные" in row_text or "пятниц" in row_text:
            header_row = i
            break

    if header_row is None:
        return

    # Определяем структуру столбцов
    # Ищем пары столбцов: (индекс кол-ва, индекс пробега) для каждого типа дня
    headers = table[header_row]

    # Проверяем есть ли подзаголовки (кол-во/пробег) в следующей строке
    subheader_row = None
    if header_row + 1 < len(table):
        next_row_text = " ".join(table[header_row + 1]).lower()
        if "кол" in next_row_text or "пробег" in next_row_text or "рейс" in next_row_text:
            subheader_row = header_row + 1

    # Собираем типы дней и их столбцы
    type_columns = []  # [(type_name, number_col_idx, probeg_col_idx), ...]

    type_num = 0
    i = 0
    while i < len(headers):
        cell_lower = headers[i].lower().strip()
        if any(kw in cell_lower for kw in ["рабочие", "выходные", "пятниц", "суббот", "воскресн", "праздничн"]):
            type_num += 1
            type_name = headers[i].strip()

            # Определяем индексы столбцов для кол-ва и пробега
            # Вариант 1: два соседних столбца под одним заголовком
            # Вариант 2: заголовок повторяется дважды
            number_col = i
            probeg_col = i + 1 if i + 1 < len(headers) else None

            # Проверяем следующий столбец - если там тот же тип дня, пропускаем
            if probeg_col and headers[probeg_col].lower().strip() == cell_lower:
                # Заголовок повторяется - первый для кол-ва, второй для пробега
                i += 1
            elif probeg_col:
                # Проверяем подзаголовки для определения какой столбец что содержит
                if subheader_row:
                    sub_headers = table[subheader_row]
                    if number_col < len(sub_headers) and probeg_col < len(sub_headers):
                        sub1 = sub_headers[number_col].lower()
                        sub2 = sub_headers[probeg_col].lower()
                        # Определяем по подзаголовкам
                        # Если в первом явно указано "кол" или "рейс", а во втором "пробег" или "км"
                        if "кол" in sub1 or "рейс" in sub1:
                            if "пробег" in sub2 or "км" in sub2:
                                # Явно определено: первый - кол-во, второй - пробег
                                pass
                            else:
                                # Неоднозначно - пробуем определить по данным
                                # Оставляем probeg_col как есть (i+1)
                                pass
                        else:
                            # Не можем определить по подзаголовкам
                            # Оставляем probeg_col как есть (i+1)
                            pass

            type_columns.append((type_name, type_num, number_col, probeg_col))
            appendix.types[f"type_{type_num}_name"] = type_name

        i += 1

    appendix.num_of_types = len(type_columns)

    # Определяем начало данных (используется и для анализа, и для извлечения)
    data_start = (subheader_row + 1) if subheader_row else (header_row + 1)

    # Если есть неоднозначность в столбцах, пытаемся определить по данным
    # Собираем индексы всех столбцов, которые могут быть кол-во или пробег
    ambiguous_cols = set()
    for type_name, type_num, number_col, probeg_col in type_columns:
        ambiguous_cols.add(number_col)
        if probeg_col:
            ambiguous_cols.add(probeg_col)

    # Анализируем данные для определения типа столбцов
    if len(ambiguous_cols) > 0 and len(type_columns) > 0:
        # Находим строки с данными для анализа
        data_rows = []
        for row in table[data_start:]:
            if len(row) >= 2:
                direction = row[0].lower().strip()
                if "прямо" in direction or "от нп" in direction or \
                   "обратно" in direction or "от кп" in direction:
                    data_rows.append(row)
            if len(data_rows) >= 2:  # Хватит двух строк для анализа
                break

        # Анализируем каждый столбец
        col_values = {}  # col_idx -> [значения]
        for col_idx in ambiguous_cols:
            col_values[col_idx] = []
            for row in data_rows:
                if col_idx < len(row):
                    val = extract_number(row[col_idx])
                    if val is not None:
                        col_values[col_idx].append(val)

        # Для каждого типа дня определяем, какой столбец что содержит
        for i, (type_name, type_num, number_col, probeg_col) in enumerate(type_columns):
            if number_col is None or probeg_col is None:
                continue

            # Если подзаголовки явно не указали, пробуем определить по данным
            if subheader_row:
                sub_headers = table[subheader_row]
                if number_col < len(sub_headers) and probeg_col < len(sub_headers):
                    sub1 = sub_headers[number_col].lower()
                    sub2 = sub_headers[probeg_col].lower()
                    
                    # Если явно указано, какой столбец что содержит - пропускаем
                    if ("кол" in sub1 or "рейс" in sub1) and \
                       ("пробег" in sub2 or "км" in sub2):
                        continue

            # Анализируем данные
            vals1 = col_values.get(number_col, [])
            vals2 = col_values.get(probeg_col, [])

            if len(vals1) > 0 and len(vals2) > 0:
                # Статистика по столбцу 1
                avg1 = sum(vals1) / len(vals1)
                has_float1 = any(v != int(v) for v in vals1)
                avg_int1 = sum(v for v in vals1 if v == int(v)) / [v for v in vals1 if v == int(v)].count(v) if any(v == int(v) for v in vals1) else avg1
                
                # Статистика по столбцу 2
                avg2 = sum(vals2) / len(vals2)
                has_float2 = any(v != int(v) for v in vals2)
                avg_int2 = sum(v for v in vals2 if v == int(v)) / [v for v in vals2 if v == int(v)].count(v) if any(v == int(v) for v in vals2) else avg2

                # Логика: столбец с пробегом обычно имеет большие значения и/или дробные числа
                # Столбец с количеством рейсов обычно имеет небольшие целые числа
                
                # Если во втором столбце есть дробные числа и их среднее > 100, то это пробег
                if has_float2 and avg2 > 100:
                    # Считаем, что второй столбец - пробег
                    pass
                elif has_float1 and avg1 > 100:
                    # Считаем, что первый столбец - пробег, меняем местами
                    type_columns[i] = (type_name, type_num, probeg_col, number_col)
                elif avg2 > avg1 * 2:
                    # Если среднее во втором столбце значительно больше
                    pass
                elif avg1 > avg2 * 2:
                    # Если среднее в первом столбце значительно больше, меняем местами
                    type_columns[i] = (type_name, type_num, probeg_col, number_col)

    # Обновляем appendix.types с новыми индексами
    for type_name, type_num, number_col, probeg_col in type_columns:
        appendix.types[f"type_{type_num}_name"] = type_name

    # Ищем строки с данными (Прямое, Обратное, ИТОГО)
    for row in table[data_start:]:
        if len(row) < 2:
            continue

        direction = row[0].lower().strip()

        if "прямо" in direction or "от нп" in direction:
            suffix = "forward"
        elif "обратно" in direction or "от кп" in direction:
            suffix = "reverse"
        elif "итого" in direction or "всего" in direction:
            suffix = "sum"
        else:
            continue

        for type_name, type_num, number_col, probeg_col in type_columns:
            # Извлекаем количество рейсов
            if number_col < len(row):
                val = extract_number(row[number_col])
                if val is not None:
                    appendix.types[f"type_{type_num}_{suffix}_number"] = val

            # Извлекаем пробег
            if probeg_col and probeg_col < len(row):
                val = extract_number(row[probeg_col])
                if val is not None:
                    appendix.types[f"type_{type_num}_{suffix}_probeg"] = val


def _extract_raw_tables(tables: list[list[list[str]]]) -> dict:
    """
    Идентифицирует и извлекает сырые таблицы из документа ДС для сохранения в JSON.

    Возвращает словарь с ключами:
    - table_raschet_izm_objema  — «Расчет изменения объема»
    - table_etapy_sroki         — «Этапы исполнения Контракта (в части сроков выполнения работ)»
    - table_finansirovanie_po_godam — «Год / Финансирование по годам, руб.»
    - table_etapy_avans         — «Этапы исполнения Контракта (с учетом порядка погашения авансов)»
    - table_objemy_rabot        — «Объемы работ» (только для ГК252)
    """
    result = {
        "table_raschet_izm_objema": None,
        "table_etapy_sroki": None,
        "table_finansirovanie_po_godam": None,
        "table_etapy_avans": None,
        "table_objemy_rabot": None,
    }

    for table in tables:
        if not table or len(table) < 2:
            continue

        # Объединяем текст первых 3 строк для идентификации заголовка
        header_text = " ".join(
            " ".join(cell.lower() for cell in row)
            for row in table[:3]
        )

        # «Расчет изменения объема»: содержит «предлагаемые изменения» в заголовке
        if result["table_raschet_izm_objema"] is None and "предлагаемые изменения" in header_text:
            result["table_raschet_izm_objema"] = table
            continue

        # «Финансирование по годам»: содержит «финансирование по годам» в заголовке
        if result["table_finansirovanie_po_godam"] is None and "финансирование по годам" in header_text:
            result["table_finansirovanie_po_godam"] = table
            continue

        # Таблицы этапов — содержат «максимальная транспортная работа»
        if "максимальная транспортная работа" in header_text:
            has_price = "стоимость" in header_text
            has_avans = "авансов" in header_text
            has_raschet_period = "расчетный период" in header_text

            # «Этапы с учетом авансов»: есть столбец «с учетом авансов»
            if result["table_etapy_avans"] is None and has_avans:
                result["table_etapy_avans"] = table
                continue

            # «Объемы работ» (ГК252): нет столбца цены или есть «расчетный период»
            if result["table_objemy_rabot"] is None and (not has_price or has_raschet_period):
                result["table_objemy_rabot"] = table
                continue

            # «Этапы в части сроков»: есть столбец цены, нет авансов
            if result["table_etapy_sroki"] is None and has_price:
                result["table_etapy_sroki"] = table
                continue

    return result


def extract_from_archive(archive_path: str | Path) -> ExtractResult:
    """
    Извлекает все данные из архива ДС.

    Args:
        archive_path: Путь к архиву (ZIP или RAR)

    Returns:
        ExtractResult с извлечёнными данными
    """
    result = ExtractResult()

    with extract_archive(archive_path) as tmpdir:
        # Находим все документы
        docx_files = find_files_by_extension(tmpdir, [".docx", ".doc"])

        if not docx_files:
            result.errors.append("В архиве не найдены документы DOCX/DOC")
            return result

        # Определяем основной документ (обычно самый короткий путь или содержит "ДС")
        main_doc = None
        appendix_docs = []

        for doc in docx_files:
            name_lower = doc.name.lower()
            if "приложение" in name_lower or "прил" in name_lower:
                appendix_docs.append(doc)
            elif "дс" in name_lower or "соглашение" in name_lower:
                main_doc = doc
            else:
                appendix_docs.append(doc)

        if not main_doc and docx_files:
            main_doc = docx_files[0]

        # Парсим основной документ
        if main_doc:
            try:
                general, changes_money, changes_money_no_appendix, changes_no_money, changes_no_money_no_appendix = parse_main_document(main_doc)
                result.general = general

                result.change_with_money.extend(changes_money)
                result.change_with_money_no_appendix.extend(changes_money_no_appendix)
                result.change_without_money.extend(changes_no_money)
                result.change_without_money_no_appendix.extend(changes_no_money_no_appendix)

            except Exception as e:
                result.errors.append(f"Ошибка парсинга основного документа: {e}")

            # Извлекаем сырые таблицы для JSON
            try:
                _, doc_tables = get_docx_data(main_doc)
                raw = _extract_raw_tables(doc_tables)
                result.table_raschet_izm_objema = raw["table_raschet_izm_objema"]
                result.table_etapy_sroki = raw["table_etapy_sroki"]
                result.table_finansirovanie_po_godam = raw["table_finansirovanie_po_godam"]
                result.table_etapy_avans = raw["table_etapy_avans"]
                result.table_objemy_rabot = raw["table_objemy_rabot"]
            except Exception as e:
                result.errors.append(f"Ошибка извлечения сырых таблиц: {e}")

        # Парсим приложения
        for doc in appendix_docs:
            try:
                appendix = parse_appendix(doc)
                if appendix and appendix.appendix_num:
                    # Преобразуем AppendixData в словарь
                    appendix_dict = {
                        "appendix_num": appendix.appendix_num,
                        "route": appendix.route,
                        "date_from": appendix.date_from,
                        "date_to": appendix.date_to,
                        "date_on": appendix.date_on,
                        "length_forward": appendix.length_forward,
                        "length_reverse": appendix.length_reverse,
                        "length_sum": appendix.length_sum,
                        "num_of_types": appendix.num_of_types,
                    }
                    appendix_dict.update(appendix.types)

                    if appendix.period_winter:
                        appendix_dict["period_winter"] = appendix.period_winter
                    if appendix.period_summer:
                        appendix_dict["period_summer"] = appendix.period_summer

                    result.appendices[appendix.appendix_num] = appendix_dict

            except Exception as e:
                result.errors.append(f"Ошибка парсинга {doc.name}: {e}")

    return result
