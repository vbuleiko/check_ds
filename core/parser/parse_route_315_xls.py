#!/usr/bin/env python3
"""
Скрипт для парсинга приложения к контракту для маршрута 315.
Особенность: данные в XLS/XLSX формате, маршрут содержит два подмаршрута (315 и 315.).
"""

import json
import re
import sys
from pathlib import Path

try:
    import xlrd
except ImportError:
    print("Требуется установить xlrd: pip install xlrd")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("Требуется установить openpyxl: pip install openpyxl")
    sys.exit(1)


def parse_number(val) -> float | int | None:
    """Парсит число из значения ячейки."""
    if isinstance(val, (int, float)):
        if val == int(val):
            return int(val)
        return round(val, 2)
    if isinstance(val, str):
        val = val.replace(" ", "").replace("\u00a0", "").replace(",", ".")
        try:
            num = float(val)
            if num == int(num):
                return int(num)
            return round(num, 2)
        except ValueError:
            return None
    return None


def read_xls_file(file_path: str) -> list[list]:
    """Читает XLS файл и возвращает данные в виде списка строк."""
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)

    data = []
    for row_idx in range(sheet.nrows):
        row = []
        for col_idx in range(sheet.ncols):
            val = sheet.cell_value(row_idx, col_idx)
            row.append(val)
        data.append(row)

    return data


def read_xlsx_file(file_path: str) -> list[list]:
    """Читает XLSX файл и возвращает данные в виде списка строк."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(list(row if row else []))

    return data


def read_excel_file(file_path: str) -> list[list]:
    """Читает Excel файл (XLS или XLSX)."""
    path = Path(file_path)
    if path.suffix.lower() == ".xls":
        return read_xls_file(file_path)
    elif path.suffix.lower() == ".xlsx":
        return read_xlsx_file(file_path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {path.suffix}")


def extract_header_info(data: list[list]) -> dict:
    """Извлекает информацию из заголовка: номер приложения, маршрут, даты."""
    result = {
        "appendix_num": None,
        "route": None,
        "date_from": None,
        "date_to": None,
        "date_on": None,
    }

    for row_idx, row in enumerate(data[:20]):
        for col_idx, cell in enumerate(row):
            if not cell:
                continue
            cell_str = str(cell)

            # Ищем номер приложения
            if result["appendix_num"] is None:
                match = re.search(r"[Пп]риложение\s*№?\s*(\d+)", cell_str)
                if match:
                    result["appendix_num"] = match.group(1)

            # Ищем номер маршрута
            if result["route"] is None:
                # Маршрут может быть числом в соседней ячейке
                if "маршрут" in cell_str.lower() or "параметр" in cell_str.lower():
                    # Проверяем следующие ячейки
                    for next_col in range(col_idx + 1, min(col_idx + 5, len(row))):
                        if next_col < len(row) and row[next_col]:
                            val = row[next_col]
                            if isinstance(val, (int, float)):
                                # 315.0 -> "315", 315.1 -> "315."
                                if val == int(val):
                                    result["route"] = str(int(val))
                                else:
                                    # Дробное число - вероятно 315. (точка в конце)
                                    result["route"] = str(int(val))
                                break

            # Ищем даты
            if result["date_from"] is None:
                # Формат: "с DD.MM.YYYY года по DD.MM.YYYY года"
                match = re.search(
                    r"с\s+(\d{2})\.(\d{2})\.(\d{4})\s+года?\s+по\s+(\d{2})\.(\d{2})\.(\d{4})",
                    cell_str
                )
                if match:
                    d1, m1, y1, d2, m2, y2 = match.groups()
                    result["date_from"] = f"{y1}-{m1}-{d1}"
                    result["date_to"] = f"{y2}-{m2}-{d2}"

    return result


def find_section_row(data: list[list], section_marker: str) -> int | None:
    """Находит строку с указанным маркером секции (например '4.' или '8.')."""
    for row_idx, row in enumerate(data):
        if row and row[0]:
            cell = str(row[0]).strip()
            if cell == section_marker:
                return row_idx
    return None


def extract_length_data(data: list[list]) -> dict:
    """
    Извлекает данные о протяженности для обоих подмаршрутов.

    п.4 - для маршрута 315
    п.4.1 - для маршрута 315.
    """
    result = {
        "route_315": {
            "length_total": None,
            "length_forward": None,
            "length_reverse": None,
        },
        "route_315_dot": {
            "length_total": None,
            "length_forward": None,
            "length_reverse": None,
        }
    }

    # Ищем секцию 4.
    row_4 = find_section_row(data, "4.")
    if row_4 is not None:
        # Протяженность всего - в той же строке
        for col_idx, cell in enumerate(data[row_4]):
            val = parse_number(cell)
            if val is not None and val > 1:
                result["route_315"]["length_total"] = val
                break

        # Следующие строки содержат прямое и обратное направление
        for i in range(row_4 + 1, min(row_4 + 5, len(data))):
            row = data[i]
            row_text = " ".join(str(c) for c in row if c).lower()

            if "прямом" in row_text:
                for cell in row:
                    val = parse_number(cell)
                    if val is not None and val > 1:
                        result["route_315"]["length_forward"] = val
                        break
            elif "обратном" in row_text:
                for cell in row:
                    val = parse_number(cell)
                    if val is not None and val > 1:
                        result["route_315"]["length_reverse"] = val
                        break

    # Ищем секцию 4.1
    row_41 = find_section_row(data, "4.1")
    if row_41 is not None:
        # Протяженность всего
        for col_idx, cell in enumerate(data[row_41]):
            val = parse_number(cell)
            if val is not None and val > 1:
                result["route_315_dot"]["length_total"] = val
                break

        # Следующие строки
        for i in range(row_41 + 1, min(row_41 + 5, len(data))):
            row = data[i]
            row_text = " ".join(str(c) for c in row if c).lower()

            if "прямом" in row_text:
                for cell in row:
                    val = parse_number(cell)
                    if val is not None and val > 1:
                        result["route_315_dot"]["length_forward"] = val
                        break
            elif "обратном" in row_text:
                for cell in row:
                    val = parse_number(cell)
                    if val is not None and val > 1:
                        result["route_315_dot"]["length_reverse"] = val
                        break

    return result


def extract_trips_data(data: list[list]) -> dict:
    """
    Извлекает данные о рейсах для обоих подмаршрутов.

    п.8 - для маршрута 315
    п.8.1 - для маршрута 315.

    Структура таблицы:
    Направление | Рабочие дни (кол-во, пробег) | Субботние дни (кол-во, пробег) | Воскресные и праздничные (кол-во, пробег)
    """
    result = {
        "route_315": {
            "num_of_types": 3,
            "type_1_name": "Рабочие дни",
            "type_2_name": "Субботние дни",
            "type_3_name": "Воскресные и праздничные дни",
        },
        "route_315_dot": {
            "num_of_types": 3,
            "type_1_name": "Рабочие дни",
            "type_2_name": "Субботние дни",
            "type_3_name": "Воскресные и праздничные дни",
        }
    }

    def is_data_row(row: list) -> bool:
        """Проверяет, является ли строка строкой с данными (числа в столбцах 3-4)."""
        if len(row) <= 4:
            return False
        val3 = parse_number(row[3]) if row[3] else None
        val4 = parse_number(row[4]) if row[4] else None
        return val3 is not None and val4 is not None

    def parse_trips_section(start_row: int, target_dict: dict):
        """Парсит секцию с рейсами начиная с указанной строки."""
        # Ищем строки с числовыми данными в столбцах 3-4
        # Первая такая строка - "Прямое", вторая - "Обратное", третья - "ИТОГО"
        data_rows = []
        for i in range(start_row + 1, min(start_row + 15, len(data))):
            row = data[i]
            if is_data_row(row):
                data_rows.append(row)
                if len(data_rows) >= 3:
                    break

        # Первая строка с данными - Прямое направление
        if len(data_rows) >= 1:
            row = data_rows[0]
            target_dict["type_1_forward_number"] = parse_number(row[3]) if len(row) > 3 else None
            target_dict["type_1_forward_probeg"] = parse_number(row[4]) if len(row) > 4 else None
            target_dict["type_2_forward_number"] = parse_number(row[5]) if len(row) > 5 else None
            target_dict["type_2_forward_probeg"] = parse_number(row[6]) if len(row) > 6 else None
            target_dict["type_3_forward_number"] = parse_number(row[7]) if len(row) > 7 else None
            target_dict["type_3_forward_probeg"] = parse_number(row[8]) if len(row) > 8 else None

        # Вторая строка с данными - Обратное направление
        if len(data_rows) >= 2:
            row = data_rows[1]
            target_dict["type_1_reverse_number"] = parse_number(row[3]) if len(row) > 3 else None
            target_dict["type_1_reverse_probeg"] = parse_number(row[4]) if len(row) > 4 else None
            target_dict["type_2_reverse_number"] = parse_number(row[5]) if len(row) > 5 else None
            target_dict["type_2_reverse_probeg"] = parse_number(row[6]) if len(row) > 6 else None
            target_dict["type_3_reverse_number"] = parse_number(row[7]) if len(row) > 7 else None
            target_dict["type_3_reverse_probeg"] = parse_number(row[8]) if len(row) > 8 else None

        # Третья строка с данными - ИТОГО
        if len(data_rows) >= 3:
            row = data_rows[2]
            target_dict["type_1_sum_number"] = parse_number(row[3]) if len(row) > 3 else None
            target_dict["type_1_sum_probeg"] = parse_number(row[4]) if len(row) > 4 else None
            target_dict["type_2_sum_number"] = parse_number(row[5]) if len(row) > 5 else None
            target_dict["type_2_sum_probeg"] = parse_number(row[6]) if len(row) > 6 else None
            target_dict["type_3_sum_number"] = parse_number(row[7]) if len(row) > 7 else None
            target_dict["type_3_sum_probeg"] = parse_number(row[8]) if len(row) > 8 else None

    # Секция 8. для маршрута 315
    row_8 = find_section_row(data, "8.")
    if row_8 is not None:
        parse_trips_section(row_8, result["route_315"])

    # Секция 8.1. для маршрута 315.
    row_81 = find_section_row(data, "8.1.")
    if row_81 is not None:
        parse_trips_section(row_81, result["route_315_dot"])

    return result


def calc_total_trips(trips_dict: dict) -> int:
    """Считает общую сумму рейсов."""
    total = 0
    for key in ["type_1_forward_number", "type_1_reverse_number", "type_2_forward_number", "type_2_reverse_number", "type_3_forward_number", "type_3_reverse_number"]:
        val = trips_dict.get(key)
        if val is not None:
            total += val
    return total


def calc_total_length(length_dict: dict) -> float:
    """Считает общую протяженность."""
    np = length_dict.get("length_forward") or 0
    op = length_dict.get("length_reverse") or 0
    return np + op


def parse_route_315_xls(file_path: str) -> dict:
    """
    Парсит XLS/XLSX файл приложения к контракту для маршрута 315.

    Возвращает словарь с данными для двух подмаршрутов:
    - route_315: данные для маршрута "315" (больше рейсов, меньше протяженность)
    - route_315_dot: данные для маршрута "315." (меньше рейсов, больше протяженность)
    """
    data = read_excel_file(file_path)

    # Извлекаем заголовочную информацию
    header = extract_header_info(data)

    # Извлекаем данные о протяженности (п.4 и п.4.1)
    length_data = extract_length_data(data)

    # Извлекаем данные о рейсах (п.8 и п.8.1)
    trips_data = extract_trips_data(data)

    # Считаем суммы для определения маршрутов
    # Данные из п.4 и п.8 (первый вариант)
    trips_first = calc_total_trips(trips_data["route_315"])
    length_first = calc_total_length(length_data["route_315"])

    # Данные из п.4.1 и п.8.1 (второй вариант)
    trips_second = calc_total_trips(trips_data["route_315_dot"])
    length_second = calc_total_length(length_data["route_315_dot"])

    # Определяем маршруты:
    # - где больше рейсов → "315"
    # - где меньше протяженность → "315"
    # Используем оба критерия для надежности
    first_is_315 = (trips_first > trips_second) or (length_first < length_second)
    second_is_315 = (trips_second > trips_first) or (length_second < length_first)

    if second_is_315 and not first_is_315:
        # Второй вариант (п.4.1, п.8.1) - это "315"
        route_315 = {
            "appendix_num": header["appendix_num"],
            "route": "315",
            "date_from": header["date_from"],
            "date_to": header["date_to"],
            "date_on": header["date_on"],
            "length_forward": length_data["route_315_dot"]["length_forward"],
            "length_reverse": length_data["route_315_dot"]["length_reverse"],
            **trips_data["route_315_dot"]
        }
        route_315_dot = {
            "appendix_num": header["appendix_num"],
            "route": "315.",
            "date_from": header["date_from"],
            "date_to": header["date_to"],
            "date_on": header["date_on"],
            "length_forward": length_data["route_315"]["length_forward"],
            "length_reverse": length_data["route_315"]["length_reverse"],
            **trips_data["route_315"]
        }
    else:
        # Первый вариант (п.4, п.8) - это "315"
        route_315 = {
            "appendix_num": header["appendix_num"],
            "route": "315",
            "date_from": header["date_from"],
            "date_to": header["date_to"],
            "date_on": header["date_on"],
            "length_forward": length_data["route_315"]["length_forward"],
            "length_reverse": length_data["route_315"]["length_reverse"],
            **trips_data["route_315"]
        }
        route_315_dot = {
            "appendix_num": header["appendix_num"],
            "route": "315.",
            "date_from": header["date_from"],
            "date_to": header["date_to"],
            "date_on": header["date_on"],
            "length_forward": length_data["route_315_dot"]["length_forward"],
            "length_reverse": length_data["route_315_dot"]["length_reverse"],
            **trips_data["route_315_dot"]
        }

    return {
        "route_315": route_315,
        "route_315_dot": route_315_dot
    }


def main():
    if len(sys.argv) < 2:
        # Пробуем найти файл в текущей директории
        script_dir = Path(__file__).parent
        xls_files = list(script_dir.glob("*.xls")) + list(script_dir.glob("*.xlsx"))

        # Исключаем файлы с "к параметрам" в названии
        xls_files = [f for f in xls_files if "к параметрам" not in f.name.lower()]

        if not xls_files:
            print("Использование: python parse_route_315_xls.py <путь_к_файлу.xls>")
            sys.exit(1)

        file_path = xls_files[0]
        print(f"Найден файл: {file_path.name}")
    else:
        file_path = Path(sys.argv[1])

    if not file_path.exists():
        print(f"Файл не найден: {file_path}")
        sys.exit(1)

    try:
        result = parse_route_315_xls(str(file_path))

        # Выводим результат
        print(json.dumps(result, ensure_ascii=False, indent=2))

        # Сохраняем в файл
        output_path = file_path.with_suffix(".json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"\nРезультат сохранен в: {output_path}")

    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
