"""Константы и маппинги."""

# Маппинг названий типов дней на дни недели
DAY_TYPE_TO_WEEKDAYS: dict[str, list[int]] = {
    "Рабочие дни": [1, 2, 3, 4, 5],
    "Рабочие дни кроме пятницы": [1, 2, 3, 4],
    "Пятница": [5],
    "Субботние дни": [6],
    "Субботы": [6],
    "Воскресные дни": [7],
    "Воскресные и праздничные дни": [7],
    "Выходные дни": [6, 7],
    "Выходные и праздничные дни": [6, 7],
    "Рабочие, выходные и праздничные дни": [1, 2, 3, 4, 5, 6, 7],
    "Рабочие и субботние дни": [1, 2, 3, 4, 5, 6],
    "Пятница, выходные и праздничные дни": [5, 6, 7],
}

# Расширенный маппинг для парсинга типов дней из документов ДС
# ВАЖНО: более специфичные ключи должны идти ПЕРЕД более общими (поиск останавливается на первом совпадении)
DAY_TYPE_MAPPING_EXTENDED: dict[str, list[int]] = {
    "рабочие, выходные и праздничные дни": [1, 2, 3, 4, 5, 6, 7],
    "рабочие, выходные и праздничные": [1, 2, 3, 4, 5, 6, 7],
    "рабочие и субботние дни": [1, 2, 3, 4, 5, 6],
    "рабочие и субботние": [1, 2, 3, 4, 5, 6],
    "рабочие дни кроме пятницы": [1, 2, 3, 4],
    "рабочие дни": [1, 2, 3, 4, 5],
    "рабочие": [1, 2, 3, 4, 5],
    "понедельник": [1],
    "вторник": [2],
    "среда": [3],
    "четверг": [4],
    "пятница": [5],
    "субботние дни": [6],
    "субботние": [6],
    "суббота": [6],
    "воскресные и праздничные дни": [7],
    "воскресные и праздничные": [7],
    "воскресные дни": [7],
    "воскресные": [7],
    "воскресенье": [7],
    "выходные": [6, 7],
    "выходные дни": [6, 7],
    "выходные и праздничные": [6, 7],
    "выходные и праздничные дни": [6, 7],
    "ежедневно": [1, 2, 3, 4, 5, 6, 7],
}

# Маппинг для исключения дней при паттерне "кроме X"
EXCLUDE_DAY_MAP: dict[str, int] = {
    "понедельник": 1,
    "вторник": 2,
    "среды": 3, "среда": 3,
    "четверга": 4, "четверг": 4,
    "пятниц": 5, "пятницы": 5, "пятница": 5,
    "суббот": 6, "субботы": 6, "суббота": 6,
    "воскресен": 7,
}

# Маппинг текста из change_*_no_appendix на тип дня
POINT_TEXT_TO_WEEKDAY: dict[str, int] = {
    "пятниц": 5,
    "суббот": 6,
    "воскресн": 7,
}

# Праздничные дни (повторяющиеся каждый год)
RECURRING_HOLIDAYS: list[tuple[int, int]] = [
    (1, 1),   # Новый год
    (1, 2),
    (1, 3),
    (1, 4),
    (1, 5),
    (1, 6),
    (1, 7),
    (1, 8),
    (2, 23),  # День защитника Отечества
    (3, 8),   # Международный женский день
    (5, 1),   # Праздник Весны и Труда
    (5, 9),   # День Победы
    (6, 12),  # День России
    (11, 4),  # День народного единства
]

# Переносы рабочих дней (конкретные даты)
# Ключ: дата, которая становится выходным
# Значение: дата, которая становится рабочим
WORK_DAY_TRANSFERS: dict[str, str] = {
    # Формат: "DD.MM.YYYY": "DD.MM.YYYY"
    # Заполняется из официальных постановлений
}

# Эталонные сезонные периоды по маршрутам (start_month, start_day, end_month, end_day)
# Используются для проверки соответствия периодов в ДС
ROUTE_SEASON_PERIODS: dict[str, dict[str, tuple[int, int, int, int]]] = {
    "207":  {"winter": (11, 16, 4, 14), "summer": (4, 15, 11, 15)},
    "305":  {"winter": (9, 1, 5, 31),   "summer": (6, 1, 8, 31)},
    "1КР":  {"winter": (11, 16, 4, 14), "summer": (4, 15, 11, 15)},
    "2КР":  {"winter": (11, 16, 4, 14), "summer": (4, 15, 11, 15)},
    "3КР":  {"winter": (11, 16, 4, 14), "summer": (4, 15, 11, 15)},
}

# Сезоны (для маршрутов с сезонным расписанием)
SEASON_PERIODS = {
    "winter": {
        "start_month": 11,  # ноябрь
        "start_day": 16,
        "end_month": 4,     # апрель
        "end_day": 14,
    },
    "summer": {
        "start_month": 4,   # апрель
        "start_day": 15,
        "end_month": 11,    # ноябрь
        "end_day": 15,
    },
}


def get_weekdays_for_day_type(day_type_name: str) -> list[int]:
    """Получить список дней недели для типа дня (простой поиск для калькулятора)."""
    # Точное совпадение
    if day_type_name in DAY_TYPE_TO_WEEKDAYS:
        return DAY_TYPE_TO_WEEKDAYS[day_type_name]

    # Нечёткий поиск
    name_lower = day_type_name.lower()
    for key, weekdays in DAY_TYPE_TO_WEEKDAYS.items():
        if key.lower() in name_lower or name_lower in key.lower():
            return weekdays

    # По умолчанию - все дни
    return [1, 2, 3, 4, 5, 6, 7]


def get_weekdays_for_type_extended(type_name: str) -> list[int]:
    """
    Расширенный поиск дней недели для типа дня из документов ДС.

    Поддерживает сокращённые названия, паттерн «кроме X» и аббревиатуры.
    """
    name_lower = type_name.lower().strip()

    base_days = None
    for key, days in DAY_TYPE_MAPPING_EXTENDED.items():
        if key in name_lower:
            base_days = list(days)
            break

    if base_days is None:
        # Пробуем по первым буквам
        if name_lower.startswith("пн"):
            base_days = [1]
        elif name_lower.startswith("вт"):
            base_days = [2]
        elif name_lower.startswith("ср"):
            base_days = [3]
        elif name_lower.startswith("чт"):
            base_days = [4]
        elif name_lower.startswith("пт"):
            base_days = [5]
        elif name_lower.startswith("сб"):
            base_days = [6]
        elif name_lower.startswith("вс"):
            base_days = [7]

    if base_days is None:
        return []

    # Обрабатываем паттерн "кроме X" — исключаем указанный день
    if "кроме" in name_lower:
        after_krome = name_lower.split("кроме", 1)[1]
        for day_name, day_num in EXCLUDE_DAY_MAP.items():
            if day_name in after_krome:
                base_days = [d for d in base_days if d != day_num]
                break

    return base_days


def detect_mid_season_change(
    route: str,
    date_from: "date",
) -> str | None:
    """
    Определяет, является ли изменение середино-сезонным.

    Если date_from не совпадает с началом ни одного сезона для маршрута,
    возвращает сообщение о середино-сезонном изменении.
    """
    route_upper = (route or "").strip().upper()
    ref = ROUTE_SEASON_PERIODS.get(route_upper)
    if not ref:
        return None

    month, day = date_from.month, date_from.day

    # Проверяем: date_from не является точным началом ни одного сезона
    is_season_start = any(
        (month, day) == (sm, sd)
        for sm, sd, _, _ in ref.values()
    )
    if is_season_start:
        return None

    # Находим конец текущего сезона
    for sm, sd, em, ed in ref.values():
        if sm <= em:
            if (sm, sd) <= (month, day) <= (em, ed):
                end_year = date_from.year + 1 if em < month else date_from.year
                date_from_fmt = date_from.strftime("%d.%m.%Y")
                date_to_fmt = f"{ed:02d}.{em:02d}.{end_year}"
                return (
                    f"Маршрут {route}: изменён сезонный график за {date_from.year} год — "
                    f"новые параметры с {date_from_fmt} по {date_to_fmt}"
                )
        else:
            if (month, day) >= (sm, sd) or (month, day) <= (em, ed):
                end_year = date_from.year + 1 if em < month else date_from.year
                date_from_fmt = date_from.strftime("%d.%m.%Y")
                date_to_fmt = f"{ed:02d}.{em:02d}.{end_year}"
                return (
                    f"Маршрут {route}: изменён сезонный график за {date_from.year} год — "
                    f"новые параметры с {date_from_fmt} по {date_to_fmt}"
                )
    return None


def parse_point_to_weekday(point_text: str) -> int | None:
    """
    Извлечь тип дня из текста change_*_no_appendix.

    Пример: "по графику движения соответствующему графику движения пятниц" → 5
    """
    text_lower = point_text.lower()
    for keyword, weekday in POINT_TEXT_TO_WEEKDAY.items():
        if keyword in text_lower:
            return weekday
    return None
