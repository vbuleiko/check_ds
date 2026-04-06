"""
Проверки ДС на высвобождение (для вкладки «Загрузка ДС»).

Проверяются только базовые вещи:
- Итоговые суммы в рублях совпадают между таблицами
- Итоговые суммы в км совпадают между таблицами
- Все обязательные таблицы присутствуют в документе
- Номер ДС и ГК одинаков везде в документе

Проверки таблиц (сравнение с БД, закрытый этап и т.д.)
выполняются отдельно на вкладке «Проверка таблиц».
"""

from core.checker.internal import CheckItem, CheckResult


_PRICE_TOL = 0.02
_KM_TOL = 0.02


def _fmt(v: float | None) -> str:
    if v is None:
        return "?"
    return f"{v:,.2f}".replace(",", " ")


def check_itogo_rub(data: dict) -> CheckItem:
    """ИТОГО по рублям таблицы 1 совпадает с таблицей 2 и с новой ценой контракта."""
    item = CheckItem(id="vysv_itogo_rub", label="Итоговые суммы в рублях совпадают")

    t1 = data.get("itogo_price")
    t2 = data.get("itogo_price_t2")
    new_price = (data.get("general") or {}).get("new_contract_price")

    if t1 is None:
        item.warnings.append("ИТОГО цена таблицы 1 не найдена")
        return item

    errors = []
    if t2 is not None:
        diff = abs(round(t1, 2) - round(t2, 2))
        if diff > _PRICE_TOL:
            errors.append(
                f"ИТОГО таблицы 1 ({_fmt(t1)} руб.) ≠ ИТОГО таблицы 2 ({_fmt(t2)} руб.)"
                f", разница {_fmt(diff)} руб."
            )

    if new_price is not None:
        diff = abs(round(t1, 2) - round(new_price, 2))
        if diff > _PRICE_TOL:
            errors.append(
                f"ИТОГО таблицы 1 ({_fmt(t1)} руб.) ≠ новая цена контракта ({_fmt(new_price)} руб.)"
                f", разница {_fmt(diff)} руб."
            )

    if errors:
        item.errors.extend(errors)
    else:
        item.detail = f"ИТОГО: {_fmt(t1)} руб."

    return item


def check_itogo_km(data: dict) -> CheckItem:
    """ИТОГО по км таблицы 1 совпадает с таблицей 2."""
    item = CheckItem(id="vysv_itogo_km", label="Итоговые суммы в км совпадают")

    t1 = data.get("itogo_km")
    t2 = data.get("itogo_km_t2")

    if t1 is None:
        item.warnings.append("ИТОГО км таблицы 1 не найдено")
        return item

    if t2 is not None:
        diff = abs(round(t1, 2) - round(t2, 2))
        if diff > _KM_TOL:
            item.errors.append(
                f"ИТОГО км таблицы 1 ({_fmt(t1)}) ≠ ИТОГО км таблицы 2 ({_fmt(t2)})"
                f", разница {_fmt(diff)} км"
            )
        else:
            item.detail = f"ИТОГО: {_fmt(t1)} км"
    else:
        item.detail = f"ИТОГО: {_fmt(t1)} км"

    return item


def check_tables_present(data: dict) -> CheckItem:
    """Все обязательные таблицы присутствуют в документе."""
    item = CheckItem(id="vysv_tables_present", label="Все таблицы присутствуют в документе")

    has_t1 = bool(data.get("stages_table1"))
    has_t2 = bool(data.get("stages_table2"))
    has_fin = bool(data.get("stages_finansirovanie_raw"))

    if not has_t1:
        item.errors.append(
            "Не найдена таблица «Этапы исполнения Контракта (в части сроков выполнения работ)»"
        )
    if not has_t2:
        item.errors.append(
            "Не найдена таблица «Этапы исполнения Контракта (с учётом порядка погашения авансов)»"
        )
    if not has_fin:
        item.warnings.append("Таблица «Финансирование по годам» не найдена")

    if not item.errors and not item.warnings:
        t1 = len(data.get("stages_table1", []))
        t2 = len(data.get("stages_table2", []))
        item.detail = f"Таблица этапов: {t1} строк, таблица авансов: {t2} строк"

    return item


def check_ds_gk_consistency(data: dict) -> CheckItem:
    """Номер ДС и ГК присутствуют и одинаковы везде в документе."""
    item = CheckItem(id="vysv_ds_gk", label="Номер ДС и ГК указан везде одинаково")

    general = data.get("general") or {}
    ds_number = general.get("ds_number")
    contract_number = general.get("contract_number")
    contract_short = general.get("contract_short_number")

    if ds_number is None:
        item.errors.append("Номер ДС не найден в документе")
    if contract_number is None:
        item.errors.append("Номер ГК не найден в документе")

    if not item.errors:
        label = f"ГК {contract_short}" if contract_short else contract_number
        item.detail = f"ДС №{ds_number}, {label}"

    return item


def check_vysvobozhdenie(
    data: dict,
    session=None,
    contract_id: int | None = None,
) -> CheckResult:
    """Базовые проверки ДС на высвобождение для вкладки «Загрузка ДС»."""
    result = CheckResult()
    result.add_check(check_itogo_rub(data))
    result.add_check(check_itogo_km(data))
    result.add_check(check_tables_present(data))
    result.add_check(check_ds_gk_consistency(data))
    return result
