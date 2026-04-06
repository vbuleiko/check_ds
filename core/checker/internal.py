"""
Внутренняя проверка JSON данных ДС.

Проверяет:
- Консистентность сумм и пробегов
- Соответствие изменений и приложений
- Арифметику в приложениях
- Нумерацию приложений
- Номер ДС в приложениях
"""
from dataclasses import dataclass, field

from core.constants import ROUTE_SEASON_PERIODS


@dataclass
class CheckItem:
    """Результат одной категории проверки."""
    id: str
    label: str
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    detail: str | None = None  # Краткий результат при успехе

    @property
    def ok(self) -> bool:
        return len(self.errors) == 0

    @property
    def has_warnings(self) -> bool:
        return len(self.warnings) > 0

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "label": self.label,
            "ok": self.ok,
            "has_warnings": self.has_warnings,
            "errors": self.errors,
            "warnings": self.warnings,
            "detail": self.detail,
        }


@dataclass
class CheckResult:
    """Результат проверки."""
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    checks: list[CheckItem] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        """Проверка прошла без ошибок."""
        return len(self.errors) == 0

    def add_check(self, item: CheckItem):
        """Добавляет проверку и накапливает ошибки/предупреждения."""
        self.checks.append(item)
        self.errors.extend(item.errors)
        self.warnings.extend(item.warnings)


def check_sums_consistency(data: dict) -> CheckItem:
    """
    Проверяет, что все sum_* значения равны между собой.
    """
    item = CheckItem(id="sums", label="Суммы в рублях совпадают между собой")
    general = data.get("general", {})

    sum_fields = [
        "sum_text",
        "sum_finansirovanie_text",
        "sum_etapy",
        "sum_etapy_avans",
        "sum_finansirovanie_table",
    ]

    sum_values = {f: general.get(f) for f in sum_fields if general.get(f) is not None}

    if not sum_values:
        item.detail = "Нет данных"
    else:
        unique_sums = set(sum_values.values())
        if len(unique_sums) > 1:
            item.errors.append(
                "Значения sum_* не равны: " +
                ", ".join(f"{k}: {v}" for k, v in sum_values.items())
            )
        else:
            val = next(iter(unique_sums))
            item.detail = f"{val:,.2f} руб.".replace(",", " ") if isinstance(val, (int, float)) else str(val)

    return item


def check_probeg_consistency(data: dict) -> CheckItem:
    """
    Проверяет, что все probeg_* значения равны между собой.
    """
    item = CheckItem(id="probeg", label="Суммы в км совпадают между собой")
    general = data.get("general", {})

    probeg_fields = [
        "probeg_sravnenie",
        "probeg_etapy",
        "probeg_etapy_avans",
    ]

    probeg_values = {f: general.get(f) for f in probeg_fields if general.get(f) is not None}

    if not probeg_values:
        item.detail = "Нет данных"
    else:
        unique_probegs = set(probeg_values.values())
        if len(unique_probegs) > 1:
            item.errors.append(
                "Значения probeg_* не равны: " +
                ", ".join(f"{k}: {v}" for k, v in probeg_values.items())
            )
        else:
            val = next(iter(unique_probegs))
            item.detail = f"{val:,.2f} км".replace(",", " ") if isinstance(val, (int, float)) else str(val)

    return item


def check_changes_vs_appendices(data: dict) -> tuple[CheckItem, CheckItem]:
    """
    Проверяет соответствие изменений и приложений.

    Возвращает два CheckItem:
    - refs_item: ссылки на номера приложений
    - dates_item: соответствие дат
    """
    refs_item = CheckItem(id="appendix_refs", label="Ссылки на номера приложений корректны")
    dates_item = CheckItem(id="appendix_dates", label="Даты действия совпадают в основном файле и приложениях")

    appendices = data.get("appendices", {})

    for section_name in ["change_with_money", "change_without_money"]:
        changes = data.get(section_name, [])

        for change in changes:
            appendix_id = change.get("appendix")
            if not appendix_id:
                continue

            if appendix_id not in appendices:
                refs_item.errors.append(
                    f"В {section_name}: приложение '{appendix_id}' "
                    f"(маршрут {change.get('route')}) не найдено в appendices"
                )
                continue

            appendix_data = appendices[appendix_id]

            # Проверка appendix == appendix_num
            if appendix_id != appendix_data.get("appendix_num"):
                refs_item.errors.append(
                    f"В {section_name}, приложение {appendix_id}: "
                    f"appendix '{appendix_id}' != appendix_num '{appendix_data.get('appendix_num')}'"
                )

            # Проверка route (без учёта регистра)
            change_route = (change.get("route") or "").upper()
            appendix_route = (appendix_data.get("route") or "").upper()
            if change_route != appendix_route:
                refs_item.errors.append(
                    f"В {section_name}, приложение {appendix_id}: "
                    f"route '{change.get('route')}' != '{appendix_data.get('route')}'"
                )

            # Проверка дат
            for date_field in ["date_from", "date_to", "date_on"]:
                if change.get(date_field) != appendix_data.get(date_field):
                    dates_item.warnings.append(
                        f"В {section_name}, приложение {appendix_id}: "
                        f"{date_field} '{change.get(date_field)}' != '{appendix_data.get(date_field)}'"
                    )

    if not refs_item.errors:
        refs_item.detail = "Все ссылки корректны"
    if not dates_item.warnings:
        dates_item.detail = "Все даты совпадают"

    return refs_item, dates_item


def check_appendix_arithmetic(data: dict) -> CheckItem:
    """
    Проверяет арифметику в приложениях:
    - sum_number = forward_number + reverse_number
    - sum_probeg = forward_probeg + reverse_probeg
    - probeg = length * number
    """
    item = CheckItem(id="arithmetic", label="Арифметика в приложениях")
    appendices = data.get("appendices", {})

    for appendix_id, appendix_data in appendices.items():
        appendix_num = appendix_data.get("appendix_num", appendix_id)
        num_of_types = appendix_data.get("num_of_types", 0)

        for type_num in range(1, num_of_types + 1):
            type_name = appendix_data.get(f"type_{type_num}_name", f"type_{type_num}")

            # Проверка sum_number = forward_number + reverse_number
            forward_number = appendix_data.get(f"type_{type_num}_forward_number", 0) or 0
            reverse_number = appendix_data.get(f"type_{type_num}_reverse_number", 0) or 0
            sum_number = appendix_data.get(f"type_{type_num}_sum_number")
            expected_sum_number = round(forward_number + reverse_number, 2)

            if sum_number is not None and round(sum_number, 2) != expected_sum_number:
                item.errors.append(
                    f"Приложение № {appendix_num}, \"{type_name}\": "
                    f"Кол-во рейсов общее ({sum_number}) != "
                    f"от НП ({forward_number}) + от КП ({reverse_number}) = {expected_sum_number}"
                )

            # Проверка sum_probeg = forward_probeg + reverse_probeg
            forward_probeg = appendix_data.get(f"type_{type_num}_forward_probeg", 0) or 0
            reverse_probeg = appendix_data.get(f"type_{type_num}_reverse_probeg", 0) or 0
            sum_probeg = appendix_data.get(f"type_{type_num}_sum_probeg")
            expected_sum_probeg = round(forward_probeg + reverse_probeg, 2)

            if sum_probeg is not None and round(sum_probeg, 2) != expected_sum_probeg:
                item.errors.append(
                    f"Приложение № {appendix_num}, \"{type_name}\": "
                    f"Пробег общий ({sum_probeg}) != "
                    f"от НП ({forward_probeg}) + от КП ({reverse_probeg}) = {expected_sum_probeg}"
                )

            # Проверка probeg = length * number
            length_forward = appendix_data.get("length_forward")
            length_reverse = appendix_data.get("length_reverse")

            if length_forward is not None and forward_number > 0:
                expected_forward_probeg = round(length_forward * forward_number, 2)
                if abs(forward_probeg - expected_forward_probeg) > 0.01:
                    item.errors.append(
                        f"Приложение № {appendix_num}, \"{type_name}\": "
                        f"Пробег от НП ({forward_probeg}) != "
                        f"протяжённость ({length_forward}) × "
                        f"кол-во рейсов ({forward_number}) = {expected_forward_probeg}"
                    )

            if length_reverse is not None and reverse_number > 0:
                expected_reverse_probeg = round(length_reverse * reverse_number, 2)
                if abs(reverse_probeg - expected_reverse_probeg) > 0.01:
                    item.errors.append(
                        f"Приложение № {appendix_num}, \"{type_name}\": "
                        f"Пробег от КП ({reverse_probeg}) != "
                        f"протяжённость ({length_reverse}) × "
                        f"кол-во рейсов ({reverse_number}) = {expected_reverse_probeg}"
                    )

        # Проверка length_sum = length_forward + length_reverse
        length_sum = appendix_data.get("length_sum")
        length_forward = appendix_data.get("length_forward")
        length_reverse = appendix_data.get("length_reverse")

        if length_sum is not None:
            if length_forward is not None and length_reverse is not None:
                expected_length_sum = round(length_forward + length_reverse, 2)
                if abs(round(length_sum, 2) - expected_length_sum) > 0.01:
                    item.errors.append(
                        f"Приложение № {appendix_num}: "
                        f"Протяжённость всего ({length_sum}) != "
                        f"прямое ({length_forward}) + обратное ({length_reverse}) = {expected_length_sum}"
                    )
            elif length_forward is not None:
                if abs(round(length_sum, 2) - round(length_forward, 2)) > 0.01:
                    item.errors.append(
                        f"Приложение № {appendix_num}: "
                        f"Протяжённость всего ({length_sum}) != прямое ({length_forward})"
                    )
            elif length_reverse is not None:
                if abs(round(length_sum, 2) - round(length_reverse, 2)) > 0.01:
                    item.errors.append(
                        f"Приложение № {appendix_num}: "
                        f"Протяжённость всего ({length_sum}) != обратное ({length_reverse})"
                    )

    if not item.errors:
        item.detail = "Нарушений не найдено"

    return item


def _parse_season_date(date_str: str | None) -> tuple[int, int] | None:
    """
    Извлекает (month, day) из строки "0000-MM-DD" или "YYYY-MM-DD".
    Возвращает None если строка пустая или не распарсилась.
    """
    if not date_str:
        return None
    try:
        parts = date_str.split("-")
        if len(parts) == 3:
            return (int(parts[1]), int(parts[2]))
    except (ValueError, IndexError):
        pass
    return None


def _fmt_season_date(md: tuple[int, int] | None) -> str:
    """Форматирует (month, day) → "DD.MM"."""
    if md is None:
        return "?"
    return f"{md[1]:02d}.{md[0]:02d}"


def check_seasonal_periods(data: dict) -> CheckItem:
    """
    Проверяет сезонные периоды в приложениях:
    1. Арифметика рейсов внутри периодов
    2. Соответствие дат периодов эталонным значениям (ROUTE_SEASON_PERIODS)
    """
    item = CheckItem(id="seasonal", label="Арифметика в сезонных периодах")
    appendices = data.get("appendices", {})

    for appendix_id, appendix_data in appendices.items():
        period_winter = appendix_data.get("period_winter")
        period_summer = appendix_data.get("period_summer")

        if not period_winter and not period_summer:
            continue

        appendix_num = appendix_data.get("appendix_num", appendix_id)
        route = (appendix_data.get("route") or "").strip().upper()
        ref_periods = ROUTE_SEASON_PERIODS.get(route) or ROUTE_SEASON_PERIODS.get(route.upper())

        # --- Проверка соответствия дат эталону ---
        period_map = [("winter", period_winter, "Зимний"), ("summer", period_summer, "Летний")]
        for season_key, period_data, season_label_ru in period_map:
            if not period_data:
                continue
            if ref_periods and season_key in ref_periods:
                ref = ref_periods[season_key]  # (start_month, start_day, end_month, end_day)
                ref_from = (ref[0], ref[1])
                ref_to = (ref[2], ref[3])

                ds_from = _parse_season_date(period_data.get("date_from"))
                ds_to = _parse_season_date(period_data.get("date_to"))

                mismatch_parts = []
                if ds_from is not None and ds_from != ref_from:
                    mismatch_parts.append(
                        f"начало: в ДС {_fmt_season_date(ds_from)}, эталон {_fmt_season_date(ref_from)}"
                    )
                if ds_to is not None and ds_to != ref_to:
                    mismatch_parts.append(
                        f"конец: в ДС {_fmt_season_date(ds_to)}, эталон {_fmt_season_date(ref_to)}"
                    )

                if mismatch_parts:
                    ds_range = f"{_fmt_season_date(ds_from)} – {_fmt_season_date(ds_to)}"
                    ref_range = f"{_fmt_season_date(ref_from)} – {_fmt_season_date(ref_to)}"
                    item.warnings.append(
                        f"Приложение № {appendix_num}, маршрут {route}: "
                        f"возможна ошибка в графиках — {season_label_ru.lower()} период "
                        f"в ДС: {ds_range}, ожидается: {ref_range}"
                    )

        # --- Арифметика рейсов ---
        for period_name, period_data in [("winter", period_winter), ("summer", period_summer)]:
            if not period_data:
                continue

            num_of_types = period_data.get("num_of_types", 0)

            for type_num in range(1, num_of_types + 1):
                type_name = period_data.get(f"type_{type_num}_name", f"type_{type_num}")

                forward_number = period_data.get(f"type_{type_num}_forward_number", 0) or 0
                reverse_number = period_data.get(f"type_{type_num}_reverse_number", 0) or 0
                sum_number = period_data.get(f"type_{type_num}_sum_number")
                expected_sum = round(forward_number + reverse_number, 2)

                if sum_number is not None and round(sum_number, 2) != expected_sum:
                    item.errors.append(
                        f"Приложение № {appendix_num}, период {period_name}, \"{type_name}\": "
                        f"Кол-во рейсов ({sum_number}) != {forward_number} + {reverse_number}"
                    )

    if not item.errors and not item.warnings:
        item.detail = "Нарушений не найдено"

    return item


def _get_season_end_for_date(route: str, month: int, day: int) -> tuple[int, int] | None:
    """
    Возвращает (end_month, end_day) сезона, в который попадает дата (month, day),
    для указанного маршрута. Возвращает None если маршрут не сезонный.
    """
    ref = ROUTE_SEASON_PERIODS.get(route) or ROUTE_SEASON_PERIODS.get(route.upper())
    if not ref:
        return None

    # Зимний период может переходить через год (напр. 16.11 → 14.04)
    # Летний период находится в пределах одного года (напр. 15.04 → 15.11)
    # Для 305: зима 01.09→31.05 (переходит через год), лето 01.06→31.08

    for season_key in ("winter", "summer"):
        sm, sd, em, ed = ref[season_key]
        if sm <= em:
            # Период в пределах одного года: sm/sd — em/ed
            if (sm, sd) <= (month, day) <= (em, ed):
                return (em, ed)
        else:
            # Период переходит через год: sm/sd..12/31 + 01/01..em/ed
            if (month, day) >= (sm, sd) or (month, day) <= (em, ed):
                return (em, ed)
    return None


def check_seasonal_change_year(data: dict) -> CheckItem:
    """
    Обнаруживает изменения, которые вступают в силу в середине сезонного периода.
    Формирует предупреждение "изменён сезонный график за {год}".
    """
    item = CheckItem(id="seasonal_change", label="Изменения сезонного графика")
    appendices = data.get("appendices", {})

    for appendix_id, appendix_data in appendices.items():
        # Считаем приложение сезонным если присутствует хотя бы один ключ period_*
        has_seasonal = (
            appendix_data.get("period_winter") is not None or
            appendix_data.get("period_summer") is not None
        )
        if not has_seasonal:
            continue

        route = (appendix_data.get("route") or "").strip().upper()
        if not ROUTE_SEASON_PERIODS.get(route):
            continue

        # Берём date_from приложения (дата начала действия новых параметров)
        date_from_str = appendix_data.get("date_from") or appendix_data.get("date_on")
        if not date_from_str:
            continue

        # Пропускаем если год не указан (0000 — эталонная запись без конкретного года)
        parts = date_from_str.split("-")
        if len(parts) != 3:
            continue
        try:
            year = int(parts[0])
            month = int(parts[1])
            day = int(parts[2])
        except ValueError:
            continue

        if year == 0:
            continue  # Нет конкретного года — не определяем как изменение графика

        ref = ROUTE_SEASON_PERIODS[route]
        season_end = _get_season_end_for_date(route, month, day)
        if not season_end:
            continue

        # Проверяем: является ли date_from точным началом сезона?
        is_season_start = False
        for season_key in ("winter", "summer"):
            sm, sd, em, ed = ref[season_key]
            if (month, day) == (sm, sd):
                is_season_start = True
                break

        if not is_season_start:
            # Дата вступления в силу — середина сезона
            end_m, end_d = season_end
            # Определяем год конца сезона
            if end_m < month:
                end_year = year + 1
            else:
                end_year = year

            date_from_fmt = f"{day:02d}.{month:02d}.{year}"
            date_to_fmt = f"{end_d:02d}.{end_m:02d}.{end_year}"
            appendix_num = appendix_data.get("appendix_num", appendix_id)

            item.warnings.append(
                f"Маршрут {route}, Приложение № {appendix_num}: "
                f"изменён сезонный график за {year} год — "
                f"новые параметры действуют с {date_from_fmt} по {date_to_fmt}"
            )

    if not item.warnings:
        item.detail = "Изменений сезонного графика не обнаружено"

    return item


def check_appendix_numbering(data: dict) -> CheckItem:
    """
    Проверяет, что нумерация приложений последовательна (1..N без пропусков).
    Для ГК219, ГК220, ГК222 допускает отсутствие одного приложения-Excel (km_data).
    """
    item = CheckItem(id="appendix_numbering", label="Все приложения присутствуют")
    appendices = data.get("appendices", {})

    if not appendices:
        return item

    try:
        nums = sorted(int(k) for k in appendices.keys())
    except ValueError:
        return item  # Нечисловые ключи — пропускаем проверку

    max_num = nums[-1]
    expected = set(range(1, max_num + 1))
    present = set(nums)
    missing = expected - present

    if missing:
        # Для ГК219/220/222 одно приложение — файл Excel (km_data)
        general = data.get("general", {})
        full_number = general.get("contract_number", "")
        contract_code = ""
        if full_number and len(full_number) >= 7 and full_number.endswith("0001"):
            contract_code = full_number[-7:-4]

        km_contracts = {"219", "220", "222"}
        km_data = data.get("km_data")

        if contract_code in km_contracts and km_data:
            km_appendix_str = km_data.get("appendix_number")
            if km_appendix_str:
                try:
                    missing.discard(int(km_appendix_str))
                except ValueError:
                    pass

        for num in sorted(missing):
            item.errors.append(f"Отсутствует приложение № {num}")

    if not item.errors:
        item.detail = f"Найдено {len(appendices)} приложений"

    return item


def check_ds_numbers_in_appendices(data: dict) -> CheckItem:
    """
    Проверяет, что номер ДС в каждом приложении совпадает с основным файлом.
    """
    item = CheckItem(id="appendix_ds_numbers", label="Номер ДС совпадает во всех приложениях")
    general = data.get("general", {})
    ds_number = general.get("ds_number")
    appendices = data.get("appendices", {})

    if not ds_number or not appendices:
        return item

    for appendix_id, appendix_data in appendices.items():
        appendix_num = appendix_data.get("appendix_num", appendix_id)
        app_ds_num = appendix_data.get("ds_num")

        if app_ds_num is None:
            continue  # Поле отсутствует — пропускаем

        if str(app_ds_num) != str(ds_number):
            item.errors.append(
                f"Приложение № {appendix_num}: номер ДС '{app_ds_num}' "
                f"не соответствует основному файлу ('{ds_number}')"
            )

    if not item.errors:
        item.detail = f"Номер ДС: {ds_number}"

    return item


def check_json(data: dict) -> CheckResult:
    """
    Выполняет полную проверку JSON данных ДС.

    Args:
        data: Словарь с данными ДС

    Returns:
        CheckResult с ошибками, предупреждениями и структурированными проверками
    """
    result = CheckResult()

    result.add_check(check_sums_consistency(data))
    result.add_check(check_probeg_consistency(data))

    refs_item, dates_item = check_changes_vs_appendices(data)
    if data.get("change_with_money") or data.get("change_without_money"):
        result.add_check(refs_item)
        result.add_check(dates_item)

    appendices = data.get("appendices", {})

    result.add_check(check_appendix_arithmetic(data))

    seasonal_item = check_seasonal_periods(data)
    has_seasonal = any(
        v.get("period_winter") or v.get("period_summer")
        for v in appendices.values()
    ) if appendices else False
    if has_seasonal or seasonal_item.errors or seasonal_item.warnings:
        result.add_check(seasonal_item)

    seasonal_change_item = check_seasonal_change_year(data)
    if has_seasonal or seasonal_change_item.warnings:
        result.add_check(seasonal_change_item)

    if appendices:
        result.add_check(check_appendix_numbering(data))
        result.add_check(check_ds_numbers_in_appendices(data))

    return result
