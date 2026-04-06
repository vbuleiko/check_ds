"""Helper functions and check logic for table_checks endpoints."""
from datetime import date
from decimal import Decimal
from typing import Optional

from sqlalchemy import cast, Integer
from sqlalchemy.orm import Session

from db.models import Agreement, AgreementReference, CalculatedStage, StageStatus
from core.calculator.kilometers import calculate_contract_period
from core.calculator.price import (
    get_coefficients_for_date, get_capacities, preload_price_data
)
from core.historical_stages import get_historical_stages, get_contract_config


# =============================================================================
# Вспомогательные функции
# =============================================================================

def parse_ru_number(s: str) -> float:
    """Парсит число в русском формате: '122 170 346,72' -> 122170346.72"""
    if s is None:
        return 0.0
    cleaned = str(s).strip().replace('\xa0', ' ').replace(' ', '').replace(',', '.')
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_ru_signed_number(s: str) -> float:
    """Парсит знаковое число: '+ 368 307,77' -> +368307.77, '- 775 954,19' -> -775954.19"""
    if s is None:
        return 0.0
    s = str(s).strip().replace('\xa0', ' ')
    sign = 1
    if s.startswith('-'):
        sign = -1
        s = s[1:].strip()
    elif s.startswith('+'):
        s = s[1:].strip()
    return sign * parse_ru_number(s)


def fmt(value: float) -> str:
    """Форматирует число в читаемый вид с двумя знаками после запятой."""
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def find_previous_agreement(
    db: Session,
    contract_id: int,
    current_number: str,
) -> tuple[Optional[Agreement], Optional[AgreementReference]]:
    """
    Ищет ближайший предыдущий ДС для данного контракта.
    Возвращает (Agreement|None, AgreementReference|None).
    Приоритет: реальный ДС из agreements > эталонные данные из agreement_references.
    """
    try:
        current_num_int = int(current_number)
    except ValueError:
        return None, None

    # Ищем ближайший предыдущий ДС прямым SQL-запросом.
    # CAST(number AS INTEGER) в SQLite возвращает 0 для нечисловых строк,
    # поэтому фильтр > 0 отсекает их без загрузки всей таблицы.
    prev_agreement = (
        db.query(Agreement)
        .filter(
            Agreement.contract_id == contract_id,
            cast(Agreement.number, Integer) > 0,
            cast(Agreement.number, Integer) < current_num_int,
        )
        .order_by(cast(Agreement.number, Integer).desc())
        .first()
    )
    if prev_agreement:
        return prev_agreement, None

    # Если реального ДС нет — ищем эталонные данные
    prev_ref = (
        db.query(AgreementReference)
        .filter(
            AgreementReference.contract_id == contract_id,
            cast(AgreementReference.reference_ds_number, Integer) > 0,
            cast(AgreementReference.reference_ds_number, Integer) < current_num_int,
        )
        .order_by(cast(AgreementReference.reference_ds_number, Integer).desc())
        .first()
    )
    if prev_ref:
        return None, prev_ref

    return None, None


def find_all_previous_agreements(
    db: Session,
    contract_id: int,
    current_number: str,
) -> tuple[list[Agreement], list[AgreementReference]]:
    """
    Возвращает все предыдущие ДС и эталоны, отсортированные по номеру от ближайшего к самому старому.
    """
    try:
        current_num_int = int(current_number)
    except ValueError:
        return [], []

    all_agreements = db.query(Agreement).filter(
        Agreement.contract_id == contract_id
    ).all()

    candidates = []
    for a in all_agreements:
        try:
            n = int(a.number)
            if n < current_num_int:
                candidates.append((n, a))
        except ValueError:
            pass
    candidates.sort(key=lambda x: x[0], reverse=True)

    all_refs = db.query(AgreementReference).filter(
        AgreementReference.contract_id == contract_id
    ).all()

    ref_candidates = []
    for r in all_refs:
        try:
            n = int(r.reference_ds_number)
            if n < current_num_int:
                ref_candidates.append((n, r))
        except ValueError:
            pass
    ref_candidates.sort(key=lambda x: x[0], reverse=True)

    return [a for _, a in candidates], [r for _, r in ref_candidates]


def make_check(name: str, ok: Optional[bool], expected, actual, message: str = None) -> dict:
    return {
        "name": name,
        "ok": ok,
        "expected": expected,
        "expected_fmt": fmt(expected) if isinstance(expected, float) else str(expected) if expected is not None else None,
        "actual": actual,
        "actual_fmt": fmt(actual) if isinstance(actual, float) else str(actual) if actual is not None else None,
        "message": message,
    }


def _get_itogo_row(table: list) -> Optional[list]:
    """Возвращает строку ИТОГО из таблицы (или None). Учитывает вариант 'ИТОГО:'."""
    for row in table:
        if row and str(row[0]).strip().rstrip(':').upper() == "ИТОГО":
            return row
    return None


def check_finansirovanie_table(
    table_finansirovanie: list,
    table_etapy_sroki: list,
) -> list[dict]:
    """
    Проверяет таблицу «Финансирование по годам».

    ИТОГО[1] (Финансирование по годам, руб.) должно совпадать с
    ИТОГО[5] из таблицы «2. Этапы исполнения Контракта» (Транспортная работа, руб.).
    """
    checks = []

    itogo_fin = _get_itogo_row(table_finansirovanie)
    itogo_etapy = _get_itogo_row(table_etapy_sroki)

    if itogo_fin is None:
        checks.append(make_check(
            "Итого финансирования совпадает с итого транспортной работы (таблица 2)",
            None, None, None, "Строка ИТОГО не найдена в таблице «Финансирование по годам»",
        ))
        return checks

    if itogo_etapy is None:
        checks.append(make_check(
            "Итого финансирования совпадает с итого транспортной работы (таблица 2)",
            None, None, None, "Строка ИТОГО не найдена в таблице «2. Этапы исполнения Контракта»",
        ))
        return checks

    fin_val = parse_ru_number(itogo_fin[1]) if len(itogo_fin) > 1 else 0.0
    etapy_price_val = parse_ru_number(itogo_etapy[5]) if len(itogo_etapy) > 5 else 0.0
    ok = abs(fin_val - etapy_price_val) < 1.0
    checks.append(make_check(
        "ИТОГО «Финансирование по годам, руб.» = ИТОГО «Транспортная работа, руб.» (таблица 2)",
        ok,
        etapy_price_val,
        fin_val,
        None if ok else f"Расхождение: {fmt(abs(fin_val - etapy_price_val))} руб.",
    ))

    # Арифметическая проверка: сумма строк == ИТОГО
    sum_data_values = []
    for row in table_finansirovanie[1:]:  # пропускаем заголовок
        if not row:
            continue
        if str(row[0]).strip().rstrip(':').upper() == "ИТОГО":
            continue
        sum_data_values.append(parse_ru_number(row[1]) if len(row) > 1 else 0.0)

    if sum_data_values:
        sum_fin = round(sum(sum_data_values), 2)
        ok_sum = abs(sum_fin - fin_val) < 1.0
        checks.append(make_check(
            "Сумма строк «Финансирование по годам, руб.» совпадает с ИТОГО",
            ok_sum,
            fin_val,
            sum_fin,
            None if ok_sum else f"Расхождение: {fmt(abs(sum_fin - fin_val))} руб.",
        ))

    return checks


def check_etapy_avans_table(
    table_etapy_avans: list,
    table_etapy_sroki: list,
) -> list[dict]:
    """
    Проверяет таблицу «4. Этапы исполнения Контракта (с учётом авансов)».

    Из строки ИТОГО:
    - [4] (Транспортная работа, км) == ИТОГО[4] таблицы 2
    - [5] (Стоимость транспортной работы, руб.) == ИТОГО[5] таблицы 2
    - [6] (Стоимость транспортной работы по этапу, руб. с учётом авансов) == ИТОГО[5] таблицы 2
    Все три значения должны быть одинаковы.
    """
    checks = []

    itogo_avans = _get_itogo_row(table_etapy_avans)
    itogo_etapy = _get_itogo_row(table_etapy_sroki)

    if itogo_avans is None:
        checks.append(make_check(
            "ИТОГО: проверки таблицы 4 (с учётом авансов)",
            None, None, None, "Строка ИТОГО не найдена в таблице «Этапы с учётом авансов»",
        ))
        return checks

    if itogo_etapy is None:
        checks.append(make_check(
            "ИТОГО: проверки таблицы 4 (с учётом авансов)",
            None, None, None, "Строка ИТОГО не найдена в таблице «2. Этапы исполнения Контракта»",
        ))
        return checks

    avans_km   = parse_ru_number(itogo_avans[4]) if len(itogo_avans) > 4 else 0.0
    avans_rub5 = parse_ru_number(itogo_avans[5]) if len(itogo_avans) > 5 else 0.0
    avans_rub6 = parse_ru_number(itogo_avans[6]) if len(itogo_avans) > 6 else 0.0
    etapy_km   = parse_ru_number(itogo_etapy[4]) if len(itogo_etapy) > 4 else 0.0
    etapy_rub  = parse_ru_number(itogo_etapy[5]) if len(itogo_etapy) > 5 else 0.0

    # Проверка 1: км
    ok_km = abs(avans_km - etapy_km) < 0.02
    checks.append(make_check(
        "ИТОГО «Транспортная работа, км» = ИТОГО «Транспортная работа, км» (таблица 2)",
        ok_km,
        etapy_km,
        avans_km,
        None if ok_km else f"Расхождение: {fmt(abs(avans_km - etapy_km))} км",
    ))

    # Проверка 2: Стоимость транспортной работы, руб. (столбец 5)
    ok_rub5 = abs(avans_rub5 - etapy_rub) < 1.0
    checks.append(make_check(
        "ИТОГО «Стоимость транспортной работы, руб.» = ИТОГО «Транспортная работа, руб.» (таблица 2)",
        ok_rub5,
        etapy_rub,
        avans_rub5,
        None if ok_rub5 else f"Расхождение: {fmt(abs(avans_rub5 - etapy_rub))} руб.",
    ))

    # Проверка 3: Стоимость транспортной работы по этапу, руб. (с учётом авансов, столбец 6)
    ok_rub6 = abs(avans_rub6 - etapy_rub) < 1.0
    checks.append(make_check(
        "ИТОГО «Стоимость транспортной работы по этапу, руб. (с учётом авансов)» = ИТОГО «Транспортная работа, руб.» (таблица 2)",
        ok_rub6,
        etapy_rub,
        avans_rub6,
        None if ok_rub6 else f"Расхождение: {fmt(abs(avans_rub6 - etapy_rub))} руб.",
    ))

    # Арифметические проверки: сумма строк == ИТОГО (внутри таблицы 4)
    sum_km_vals, sum_rub5_vals, sum_rub6_vals = [], [], []
    seen_stage_for_rub6 = set()  # для rub6 суммируем только уникальные значения по номеру этапа
    for row in table_etapy_avans[1:]:  # пропускаем заголовок
        if not row:
            continue
        if str(row[0]).strip().rstrip(':').upper() == "ИТОГО":
            continue
        try:
            stage_num = int(row[0])  # только строки с номером этапа
        except (ValueError, TypeError):
            continue
        sum_km_vals.append(parse_ru_number(row[4]) if len(row) > 4 else 0.0)
        sum_rub5_vals.append(parse_ru_number(row[5]) if len(row) > 5 else 0.0)
        # Столбец rub6 может содержать объединённые ячейки — суммируем только первое вхождение для каждого этапа
        if stage_num not in seen_stage_for_rub6:
            seen_stage_for_rub6.add(stage_num)
            sum_rub6_vals.append(parse_ru_number(row[6]) if len(row) > 6 else 0.0)

    if sum_km_vals:
        s_km = round(sum(sum_km_vals), 2)
        ok_s_km = abs(s_km - avans_km) < 0.02
        checks.append(make_check(
            "Сумма строк «Транспортная работа, км» совпадает с ИТОГО",
            ok_s_km, avans_km, s_km,
            None if ok_s_km else f"Расхождение: {fmt(abs(s_km - avans_km))} км",
        ))

        s_rub5 = round(sum(sum_rub5_vals), 2)
        ok_s_rub5 = abs(s_rub5 - avans_rub5) < 1.0
        checks.append(make_check(
            "Сумма строк «Стоимость транспортной работы, руб.» совпадает с ИТОГО",
            ok_s_rub5, avans_rub5, s_rub5,
            None if ok_s_rub5 else f"Расхождение: {fmt(abs(s_rub5 - avans_rub5))} руб.",
        ))

        s_rub6 = round(sum(sum_rub6_vals), 2)
        ok_s_rub6 = abs(s_rub6 - avans_rub6) < 1.0
        checks.append(make_check(
            "Сумма строк «Стоимость транспортной работы по этапу, руб. (с учётом авансов)» совпадает с ИТОГО",
            ok_s_rub6, avans_rub6, s_rub6,
            None if ok_s_rub6 else f"Расхождение: {fmt(abs(s_rub6 - avans_rub6))} руб.",
        ))

    return checks


def _find_prev_initial_km(
    prev_agreements: list[Agreement],
    prev_references: list[AgreementReference],
) -> tuple[Optional[float], Optional[str]]:
    """Ищет initial_km в цепочке предыдущих ДС (от ближайшего к старшему), потом в эталонах."""
    for a in prev_agreements:
        tbl = (a.json_data or {}).get("table_raschet_izm_objema", [])
        if len(tbl) > 1:
            return parse_ru_number(tbl[1][1]), f"ДС №{a.number}"
    for r in prev_references:
        if r.initial_km is not None:
            return r.initial_km, f"эталон ДС №{r.reference_ds_number}"
    return None, None


def _find_prev_probeg(
    prev_agreements: list[Agreement],
    prev_references: list[AgreementReference],
) -> tuple[Optional[float], Optional[str]]:
    """Ищет probeg_etapy в цепочке предыдущих ДС, потом в эталонах."""
    for a in prev_agreements:
        jd = a.json_data or {}
        # Обычный ДС: general.probeg_etapy
        val = jd.get("general", {}).get("probeg_etapy")
        if val is not None:
            return float(val), f"ДС №{a.number}"
        # Высвобождение: itogo_km из таблицы этапов
        if jd.get("type") == "vysvobozhdenie" and jd.get("itogo_km") is not None:
            return float(jd["itogo_km"]), f"ДС №{a.number} (высвоб.)"
    for r in prev_references:
        if r.probeg_etapy is not None:
            return float(r.probeg_etapy), f"эталон ДС №{r.reference_ds_number}"
    return None, None


def _find_prev_price(
    prev_agreements: list[Agreement],
    prev_references: list[AgreementReference],
) -> tuple[Optional[float], Optional[str]]:
    """Ищет цену контракта в цепочке предыдущих ДС, потом в эталонах."""
    for a in prev_agreements:
        price = _get_contract_price(a.json_data)
        if price is not None:
            return price, f"ДС №{a.number}"
    for r in prev_references:
        if r.sum_price is not None:
            return r.sum_price, f"эталон ДС №{r.reference_ds_number}"
    return None, None


def check_raschet_table(
    table: list,
    current_json: dict,
    prev_agreements: list[Agreement],
    prev_references: list[AgreementReference],
) -> list[dict]:
    """Выполняет 7 проверок для таблицы Расчет изменения объема."""
    checks = []

    if len(table) < 2:
        return checks

    row = table[1]  # единственная строка данных

    initial_km     = parse_ru_number(row[1])
    ten_pct_km     = parse_ru_number(row[2])
    prev_total_km  = parse_ru_number(row[3])
    proposed_km    = parse_ru_number(row[4])
    delta_ds_val   = parse_ru_signed_number(row[5])
    total_change_val = parse_ru_signed_number(row[6])

    # --- Проверка 1: initial_km == initial_km предыдущего ДС ---
    prev_initial, src1 = _find_prev_initial_km(prev_agreements, prev_references)
    if prev_initial is not None:
        ok = abs(initial_km - prev_initial) < 0.02
        checks.append(make_check(
            f"Объем при заключении совпадает с {src1}",
            ok,
            prev_initial,
            initial_km,
            None if ok else f"Расхождение: {fmt(abs(initial_km - prev_initial))} км",
        ))
    else:
        checks.append(make_check(
            "Объем при заключении совпадает с предыдущим ДС",
            None, None, initial_km,
            "Не найден ДС с данными initial_km (ни в загруженных ДС, ни в эталонах)",
        ))

    # --- Проверка 2: ten_pct_km == round(initial_km * 0.1, 2) ---
    expected_10 = round(initial_km * 0.1, 2)
    ok2 = abs(ten_pct_km - expected_10) < 0.02
    checks.append(make_check(
        "10% от исходного объема верно указано",
        ok2,
        expected_10,
        ten_pct_km,
        None if ok2 else f"Расхождение: {fmt(abs(ten_pct_km - expected_10))} км",
    ))

    # --- Проверка 3: prev_total_km == probeg_etapy предыдущего ДС ---
    prev_probeg, src3 = _find_prev_probeg(prev_agreements, prev_references)
    if prev_probeg is not None:
        ok3 = abs(prev_total_km - prev_probeg) < 0.02
        checks.append(make_check(
            f"«Объем с учётом ранее внесённых изменений» == пробег этапов {src3}",
            ok3,
            prev_probeg,
            prev_total_km,
            None if ok3 else f"Расхождение: {fmt(abs(prev_total_km - prev_probeg))} км",
        ))
    else:
        checks.append(make_check(
            "«Объем с учётом ранее внесённых изменений» == пробег этапов предыдущего ДС",
            None, None, prev_total_km,
            "Не найден ДС с данными probeg_etapy (ни в загруженных ДС, ни в эталонах)",
        ))

    # --- Проверка 4: proposed_km == general.probeg_etapy текущего ДС ---
    current_probeg = (current_json.get("general") or {}).get("probeg_etapy")
    if current_probeg is not None:
        ok4 = abs(proposed_km - float(current_probeg)) < 0.02
        checks.append(make_check(
            "«Предлагаемые изменения» == суммарный пробег этапов (general.probeg_etapy)",
            ok4,
            float(current_probeg),
            proposed_km,
            None if ok4 else f"Расхождение: {fmt(abs(proposed_km - float(current_probeg)))} км",
        ))
    else:
        checks.append(make_check(
            "«Предлагаемые изменения» == суммарный пробег этапов (general.probeg_etapy)",
            None, None, proposed_km,
            "В JSON нет поля general.probeg_etapy",
        ))

    # --- Проверка 5: delta_ds == proposed_km - prev_total_km ---
    expected_delta = proposed_km - prev_total_km
    ok5 = abs(delta_ds_val - expected_delta) < 0.02
    checks.append(make_check(
        "«Изменение объема этим ДС» = Предлагаемые − Предыдущий объём",
        ok5,
        expected_delta,
        delta_ds_val,
        None if ok5 else f"Расхождение: {fmt(abs(delta_ds_val - expected_delta))} км",
    ))

    # --- Проверка 6: total_change == proposed_km - initial_km ---
    expected_total = proposed_km - initial_km
    ok6 = abs(total_change_val - expected_total) < 0.02
    checks.append(make_check(
        "«Общее изменение по контракту» = Предлагаемые − Объём при заключении",
        ok6,
        expected_total,
        total_change_val,
        None if ok6 else f"Расхождение: {fmt(abs(total_change_val - expected_total))} км",
    ))

    # --- Проверка 7: |total_change| <= ten_pct_km ---
    abs_total = abs(total_change_val)
    ok7 = abs_total <= ten_pct_km + 0.02
    checks.append(make_check(
        "«Общее изменение по контракту» по модулю не превышает 10% от исходного объёма",
        ok7,
        ten_pct_km,
        abs_total,
        None if ok7 else f"Превышение на {fmt(abs_total - ten_pct_km)} км",
    ))

    return checks


def _get_contract_price(json_data: dict) -> Optional[float]:
    """Извлекает цену контракта из json_data ДС (обычного или высвобождения)."""
    if not json_data:
        return None
    # Высвобождение: new_contract_price
    general = json_data.get("general", {})
    if json_data.get("type") == "vysvobozhdenie":
        return general.get("new_contract_price")
    # Обычный ДС: sum_text (цена контракта из текста)
    return general.get("sum_text")


def check_price_change(
    current_json: dict,
    prev_agreements: list[Agreement],
    prev_references: list[AgreementReference],
) -> list[dict]:
    """
    Проверяет корректность указания «увеличить/уменьшить цену Контракта на X руб.».
    Сравнивает заявленное направление и дельту с фактической разницей цен.
    """
    checks = []

    general = current_json.get("general", {})
    direction = general.get("price_change_direction")  # "увеличить" / "уменьшить"
    stated_amount = general.get("price_change_amount")  # дельта из текста

    if direction is None and stated_amount is None:
        # В этом ДС нет пункта об изменении цены — проверку пропускаем
        return checks

    current_price = _get_contract_price(current_json)

    # Ищем цену предыдущего ДС по цепочке
    prev_price, prev_source_label = _find_prev_price(prev_agreements, prev_references)

    # --- Проверка направления ---
    if direction is not None and current_price is not None and prev_price is not None:
        actual_delta = current_price - prev_price
        expected_direction = "увеличить" if actual_delta > 0 else "уменьшить"
        ok_dir = direction == expected_direction
        checks.append(make_check(
            "Направление изменения цены (увеличить/уменьшить) указано верно",
            ok_dir,
            expected_direction,
            direction,
            None if ok_dir else (
                f"Указано «{direction}», но цена {expected_direction.replace('ить', 'илась')}: "
                f"было {fmt(prev_price)} руб. ({prev_source_label}), стало {fmt(current_price)} руб."
            ),
        ))
    elif direction is not None:
        missing = []
        if current_price is None:
            missing.append("текущая цена контракта")
        if prev_price is None:
            missing.append("цена предыдущего ДС (ни в загруженных ДС, ни в эталонах)")
        checks.append(make_check(
            "Направление изменения цены (увеличить/уменьшить) указано верно",
            None, None, direction,
            f"Не удалось проверить: не найдена {', '.join(missing)}",
        ))

    # --- Проверка суммы изменения ---
    if stated_amount is not None and current_price is not None and prev_price is not None:
        actual_abs_delta = abs(current_price - prev_price)
        ok_amount = abs(stated_amount - actual_abs_delta) < 0.02
        checks.append(make_check(
            "Сумма изменения цены контракта указана верно",
            ok_amount,
            actual_abs_delta,
            stated_amount,
            None if ok_amount else f"Расхождение: {fmt(abs(stated_amount - actual_abs_delta))} руб.",
        ))
    elif stated_amount is not None:
        checks.append(make_check(
            "Сумма изменения цены контракта указана верно",
            None, None, stated_amount,
            "Не удалось проверить: нет данных о цене текущего или предыдущего ДС",
        ))

    return checks


# Короткие названия столбцов для таблицы этапов (индексы 4 и 5)
ETAPY_SROKI_HEADER_MAP = {
    "Максимальная транспортная работа, подлежащая оплате, км.": "Транспортная работа, км",
    "Стоимость транспортной работы, руб. (за исполненные этапы по контракту указана стоимость фактически выполненных и принятых Заказчиком работ)": "Транспортная работа, руб.",
}



def parse_stage_dates(year_str: str, date_range_str: str) -> Optional[tuple[date, date]]:
    """
    Парсит год и диапазон дат из таблицы этапов.
    year_str: "2026"
    date_range_str: "01.02-28.02" -> (2026-02-01, 2026-02-28)
    """
    try:
        year = int(year_str)
        parts = date_range_str.split('-')
        if len(parts) != 2:
            return None
        d0 = parts[0].strip().split('.')
        d1 = parts[1].strip().split('.')
        d_from = date(year, int(d0[1]), int(d0[0]))
        d_to = date(year, int(d1[1]), int(d1[0]))
        return d_from, d_to
    except (ValueError, IndexError):
        return None


def calculate_price_rounded(
    route_km: dict[str, float],
    d_to: date,
    contract_number: str,
) -> tuple[Optional[float], list[str]]:
    """
    Рассчитывает суммарную цену этапа с округлением по каждому маршруту.

    Для каждого маршрута: price_route = round(km * вместимость * коэффициент, 2)
    Итого: sum(price_route)

    Ищет коэффициент: квартальный -> месячный -> любой период, покрывающий дату.

    Returns:
        (цена или None, список маршрутов без коэффициента/вместимости)
    """
    coefficients = get_coefficients_for_date(d_to, contract_number)
    capacities = get_capacities(contract_number)

    if not capacities:
        return None, []

    if not coefficients:
        return None, sorted(route_km.keys())

    total = Decimal("0")
    no_coef: list[str] = []
    for route, km in route_km.items():
        normalized = route.rstrip('.')
        coef = coefficients.get(normalized)
        cap = capacities.get(normalized)
        if coef is None or cap is None:
            no_coef.append(route)
            continue
        route_price = round(float(Decimal(str(km)) * Decimal(cap) * coef), 2)
        total += Decimal(str(route_price))

    return float(total), no_coef


def make_row_check(
    ok: Optional[bool],
    expected: Optional[float],
    actual: Optional[float],
    message: Optional[str] = None,
) -> dict:
    return {
        "ok": ok,
        "expected": expected,
        "expected_fmt": fmt(expected) if expected is not None else None,
        "actual": actual,
        "actual_fmt": fmt(actual) if actual is not None else None,
        "message": message,
    }


def find_prev_stage_row(prev_table: list, stage_num: int) -> Optional[list]:
    """Находит строку с нужным номером этапа в таблице предыдущего ДС."""
    for row in prev_table[1:]:
        try:
            if int(row[0]) == stage_num:
                return row
        except (ValueError, IndexError):
            pass
    return None


def _normalize_route(route: str) -> str:
    """Нормализует название маршрута для сравнения: убирает пробелы, приводит к верхнему регистру, удаляет точки в конце."""
    return route.replace(" ", "").upper().rstrip(".")


def _normalize_routes_dict(routes: dict) -> dict[str, float]:
    """Нормализует словарь {маршрут: км}, суммируя значения маршрутов с одинаковым нормализованным ключом."""
    result: dict[str, float] = {}
    for route, km in routes.items():
        key = _normalize_route(route)
        result[key] = result.get(key, 0.0) + (km if km is not None else 0.0)
    return result


def check_km_by_routes(
    km_data: Optional[dict],
    db: Session,
    contract_id: int,
    contract_number: str,
    prev_agreement: Optional[Agreement],
) -> list[dict]:
    """
    Проверяет объемы км по маршрутам из КМ_Приложение12/13.

    Закрытые периоды (исторические или CalculatedStage.status == CLOSED):
        сравниваем total из JSON с CalculatedStage.total_km из БД.
    Открытые периоды:
        рассчитываем из БД через calculate_contract_period(), сравниваем по каждому маршруту.
    """
    if not km_data:
        return [{"ok": None, "period_label": None, "closed": False,
                 "source": None, "errors": [],
                 "message": "Данные КМ (Приложение 12/13) отсутствуют в ДС"}]

    monthly = km_data.get("monthly", [])
    if not monthly:
        return [{"ok": None, "period_label": None, "closed": False,
                 "source": None, "errors": [],
                 "message": "Нет данных по периодам в КМ_Приложение"}]

    # Дата начала расчётных этапов и исторические константы для данного контракта
    config = get_contract_config(contract_number)
    calc_start: Optional[date] = config.get("start_date")

    # Индекс исторических этапов по (date_from, date_to)
    historical_index: dict[tuple, object] = {
        (h.date_from, h.date_to): h
        for h in get_historical_stages(contract_number)
    }

    results = []

    for period in monthly:
        date_from_str = period.get("date_from")
        date_to_str = period.get("date_to")
        json_routes: dict[str, float] = period.get("routes") or {}
        json_total: Optional[float] = period.get("total")

        try:
            d_from = date.fromisoformat(date_from_str)
            d_to = date.fromisoformat(date_to_str)
        except (ValueError, TypeError):
            results.append({
                "period_label": f"{date_from_str}\u2013{date_to_str}",
                "date_from": date_from_str, "date_to": date_to_str,
                "closed": False, "source": None, "ok": None, "errors": [],
                "message": "Ошибка разбора дат периода",
            })
            continue

        period_label = f"{d_from.strftime('%d.%m')}-{d_to.strftime('%d.%m.%Y')}"

        # Арифметическая проверка: сумма маршрутов == total (независимо от типа периода)
        if json_routes and json_total is not None:
            sum_routes = round(sum(v for v in json_routes.values() if v is not None), 2)
            json_total_rounded = round(json_total, 2)
            total_ok = abs(sum_routes - json_total_rounded) < 0.001
            total_check = {
                "ok": total_ok,
                "sum_routes": sum_routes,
                "json_total": json_total_rounded,
                "diff": round(sum_routes - json_total_rounded, 2),
            }
        elif json_total is None:
            total_check = {"ok": None, "sum_routes": None, "json_total": None, "diff": None,
                           "message": "Поле total отсутствует в JSON"}
        else:
            total_check = {"ok": None, "sum_routes": None, "json_total": None, "diff": None,
                           "message": "Нет данных маршрутов"}

        is_historical = calc_start is not None and d_to < calc_start

        if is_historical:
            # Берём эталонное значение из констант historical_stages
            hist = historical_index.get((d_from, d_to))
            expected_total = hist.max_km if hist else None
            source = "historical"
        else:
            # Ищем CalculatedStage для расчётных периодов
            cs = db.query(CalculatedStage).filter(
                CalculatedStage.contract_id == contract_id,
                CalculatedStage.date_from == d_from,
                CalculatedStage.date_to == d_to,
            ).first()
            is_closed = cs is not None and cs.status == StageStatus.CLOSED
            if not is_closed:
                # Открытый период — расчёт по маршрутам ниже
                expected_total = None
                source = "calculated"
            else:
                expected_total = round(cs.total_km, 2)
                source = "db_stage"

        is_closed_period = is_historical or source == "db_stage"

        if is_closed_period:
            json_total_r = round(json_total, 2) if json_total is not None else None

            if json_total_r is None:
                results.append({
                    "period_label": period_label,
                    "date_from": date_from_str, "date_to": date_to_str,
                    "closed": True, "source": source, "ok": None, "errors": [],
                    "total_check": total_check,
                    "message": "В JSON отсутствует поле total для периода",
                })
                continue

            if expected_total is None:
                results.append({
                    "period_label": period_label,
                    "date_from": date_from_str, "date_to": date_to_str,
                    "closed": True, "source": source, "ok": None, "errors": [],
                    "total_check": total_check,
                    "message": "Эталонное значение не найдено" + (" в исторических константах" if is_historical else " в CalculatedStage"),
                })
                continue

            ok = abs(json_total_r - expected_total) < 0.02
            results.append({
                "period_label": period_label,
                "date_from": date_from_str, "date_to": date_to_str,
                "closed": True, "source": source,
                "ok": ok, "errors": [],
                "json_total": json_total_r,
                "db_total": expected_total,
                "total_check": total_check,
                "message": None if ok else f"Расхождение: в ДС — {fmt(json_total_r)} км, ожидалось — {fmt(expected_total)} км (разница: {fmt(abs(json_total_r - expected_total))} км)",
            })

        else:
            # Открытый период — рассчитываем из БД по каждому маршруту
            try:
                calc_results = calculate_contract_period(db, contract_id, d_from, d_to)
            except Exception as e:
                results.append({
                    "period_label": period_label,
                    "date_from": date_from_str, "date_to": date_to_str,
                    "closed": False, "source": "calculated", "ok": None, "errors": [],
                    "total_check": total_check,
                    "message": f"Ошибка расчёта: {e}",
                })
                continue

            # Для ГК219: маршруты в Excel могут иметь другие названия, чем в БД
            ROUTE_ALIASES_219: dict[str, str] = {
                "110/109А": "110",
                "180А/75А": "180А",
                "555сез": "555",
            }
            if contract_number == "219":
                aliased: dict[str, float] = {}
                for r, km in json_routes.items():
                    db_route = ROUTE_ALIASES_219.get(r, r)
                    aliased[db_route] = aliased.get(db_route, 0.0) + (km or 0.0)
                json_routes = aliased

            # Нормализуем оба словаря: убираем пробелы, приводим к верхнему регистру, удаляем точки в конце
            norm_json = _normalize_routes_dict(json_routes)
            norm_calc: dict[str, float] = {}
            for route, calc in calc_results.items():
                key = _normalize_route(route)
                norm_calc[key] = round((norm_calc.get(key, 0.0) + calc.total_km), 2)

            errors = []
            for route in sorted(set(norm_json) | set(norm_calc)):
                json_km = norm_json.get(route, 0.0)
                calc_km = norm_calc.get(route, 0.0)
                if abs(json_km - calc_km) >= 0.02:
                    errors.append({
                        "route": route,
                        "json_km": json_km,
                        "expected_km": calc_km,
                        "diff": round(json_km - calc_km, 2),
                    })

            results.append({
                "period_label": period_label,
                "date_from": date_from_str, "date_to": date_to_str,
                "closed": False, "source": "calculated",
                "ok": len(errors) == 0, "errors": errors,
                "total_check": total_check,
                "message": None,
            })

    return results


def check_km_total_vs_probeg(
    km_data: Optional[dict],
    general: dict,
) -> Optional[list[dict]]:
    """
    Проверяет, что сумма значений total по всем периодам из km_data
    совпадает с probeg_sravnenie, probeg_etapy и probeg_etapy_avans из general.
    """
    if not km_data:
        return None
    monthly = km_data.get("monthly", [])
    if not monthly:
        return None

    total_sum = round(sum(float(p["total"]) for p in monthly if p.get("total") is not None), 2)

    results = []
    for field_name, label in [
        ("probeg_sravnenie", "Расчет изменения объема"),
        ("probeg_etapy", "Этапы исполнения"),
        ("probeg_etapy_avans", "Этапы с авансом"),
    ]:
        expected = general.get(field_name)
        if expected is None:
            results.append({
                "field": field_name,
                "label": label,
                "ok": None,
                "km_total": total_sum,
                "expected": None,
                "diff": None,
                "message": f"Поле {field_name} отсутствует в данных ДС",
            })
            continue

        expected_r = round(float(expected), 2)
        diff = round(total_sum - expected_r, 2)
        ok = abs(diff) < 0.02
        results.append({
            "field": field_name,
            "label": label,
            "ok": ok,
            "km_total": total_sum,
            "expected": expected_r,
            "diff": diff,
            "message": None if ok else f"Сумма КМ по периодам ({fmt(total_sum)}) \u2260 {label} ({fmt(expected_r)}), разница: {fmt(abs(diff))} км",
        })

    return results


def check_vysv_closed_stage(
    general: dict,
    db: Session,
    contract_id: int,
) -> list[dict]:
    """
    Проверяет закрытый этап из ДС на высвобождение:
    - статус этапа в БД == CLOSED
    - сумма закрытого этапа совпадает с указанной в документе
    """
    checks = []
    closed_stage = general.get("closed_stage")
    closed_amount = general.get("closed_amount")

    if closed_stage is None:
        checks.append(make_check(
            "Закрытый этап — статус в БД",
            None, None, None,
            "Номер закрытого этапа не определён из документа",
        ))
        return checks

    cs = db.query(CalculatedStage).filter(
        CalculatedStage.contract_id == contract_id,
        CalculatedStage.stage == closed_stage,
    ).first()

    if cs is None:
        checks.append(make_check(
            f"Закрытый этап {closed_stage} — статус в БД",
            None, None, None,
            f"Этап {closed_stage} не найден в БД",
        ))
        return checks

    is_closed = cs.status == StageStatus.CLOSED
    checks.append(make_check(
        f"Закрытый этап {closed_stage} — статус в БД",
        is_closed,
        "CLOSED",
        cs.status.value if cs.status else None,
        None if is_closed else f"Ожидается статус CLOSED, в БД: {cs.status}",
    ))

    if closed_amount is not None and cs.total_price is not None:
        ok_price = abs(closed_amount - cs.total_price) < 1.0
        checks.append(make_check(
            f"Закрытый этап {closed_stage} — сумма совпадает с БД",
            ok_price,
            cs.total_price,
            closed_amount,
            None if ok_price else f"Расхождение: {fmt(abs(closed_amount - cs.total_price))} руб.",
        ))
    elif closed_amount is None:
        checks.append(make_check(
            f"Закрытый этап {closed_stage} — сумма совпадает с БД",
            None, cs.total_price, None,
            "Сумма закрытого этапа не определена из документа",
        ))
    else:
        checks.append(make_check(
            f"Закрытый этап {closed_stage} — сумма совпадает с БД",
            None, None, closed_amount,
            "Сумма закрытого этапа не сохранена в БД (total_price = None)",
        ))

    return checks


def check_etapy_sroki_table(
    table: list,
    db: Session,
    contract_id: int,
    contract_number: str,
    prev_agreement: Optional[Agreement],
) -> dict:
    """
    Проверяет таблицу «Этапы исполнения Контракта (сроки)».

    Для исторических этапов (stage_num < first_calculated_stage):
        сравниваем с эталонными данными из HISTORICAL_STAGES_xxx.
    Для этапов со статусом CLOSED в БД:
        сравниваем с предыдущим ДС.
    Для открытых расчётных этапов:
        рассчитываем км и цену из БД.

    Возвращает:
        {
          "row_checks": [...],   # параллельный список к rows (null для ИТОГО)
          "itogo_check": {...},  # проверка строки ИТОГО
          "price_available": bool,  # доступны ли данные из внешней БД
        }
    """
    if len(table) < 2:
        return {"row_checks": [], "itogo_check": None, "price_available": False}

    data_rows = table[1:]  # без заголовка

    # Конфигурация контракта: с какого этапа идут расчётные
    config = get_contract_config(contract_number)
    first_calc_stage = config.get("first_stage", 1)

    # Исторические этапы по номеру: {stage_num: HistoricalStage}
    historical = {s.stage: s for s in get_historical_stages(contract_number)}

    # Таблица предыдущего ДС (для закрытых расчётных этапов)
    prev_table = (prev_agreement.json_data or {}).get("table_etapy_sroki", []) if prev_agreement else []

    # Предзагружаем данные для расчёта цены один раз
    preload_price_data(contract_number)
    capacities = get_capacities(contract_number)
    price_available = bool(capacities)

    row_checks = []

    for row in data_rows:
        # Строка ИТОГО — обрабатываем отдельно
        if str(row[0]).strip().rstrip(':').upper() == "ИТОГО":
            row_checks.append(None)
            continue

        try:
            stage_num = int(row[0])
        except (ValueError, TypeError):
            row_checks.append(None)
            continue



        year_str = str(row[1])
        date_range_str = str(row[3])
        table_km = parse_ru_number(row[4])
        table_price = parse_ru_number(row[5])

        # --- Исторический этап: сравниваем с константами из кода ---
        if stage_num < first_calc_stage:
            hist = historical.get(stage_num)
            if hist:
                ok_km = abs(table_km - hist.max_km) < 0.02
                km_check = make_row_check(
                    ok_km, hist.max_km, table_km,
                    None if ok_km else f"Расхождение: {fmt(abs(table_km - hist.max_km))} км",
                )
                if hist.price is not None:
                    ok_price = abs(table_price - hist.price) < 1.0
                    price_check = make_row_check(
                        ok_price, hist.price, table_price,
                        None if ok_price else f"Расхождение: {fmt(abs(table_price - hist.price))} руб.",
                    )
                else:
                    price_check = make_row_check(None, None, table_price, "Цена не задана в эталоне")
            else:
                price_check = make_row_check(None, None, table_price, f"Этап {stage_num} не найден в историческом эталоне")
                km_check = make_row_check(None, None, table_km, f"Этап {stage_num} не найден в историческом эталоне")

            row_checks.append({
                "stage": stage_num,
                "closed": True,
                "source": "historical",
                "km": km_check,
                "price": price_check,
            })
            continue

        # --- Расчётный этап: определяем, закрыт ли он ---
        cs = db.query(CalculatedStage).filter(
            CalculatedStage.contract_id == contract_id,
            CalculatedStage.stage == stage_num,
        ).first()
        is_closed = cs is not None and cs.status == StageStatus.CLOSED

        if is_closed:
            # Закрытый расчётный этап: сравниваем с данными вкладки Этапы (CalculatedStage)
            if cs.total_km is not None and cs.total_price is not None:
                ok_km = abs(table_km - cs.total_km) < 0.02
                km_check = make_row_check(
                    ok_km, cs.total_km, table_km,
                    None if ok_km else f"Расхождение: {fmt(abs(table_km - cs.total_km))} км",
                )
                ok_price = abs(table_price - cs.total_price) < 1.0
                price_check = make_row_check(
                    ok_price, cs.total_price, table_price,
                    None if ok_price else f"Расхождение: {fmt(abs(table_price - cs.total_price))} руб.",
                )
                source = "stages_tab"
            elif prev_table:
                # Fallback: сравниваем с предыдущим ДС
                prev_row = find_prev_stage_row(prev_table, stage_num)
                if prev_row:
                    prev_km = parse_ru_number(prev_row[4])
                    prev_price = parse_ru_number(prev_row[5])
                    ok_km = abs(table_km - prev_km) < 0.02
                    km_check = make_row_check(
                        ok_km, prev_km, table_km,
                        None if ok_km else f"Расхождение: {fmt(abs(table_km - prev_km))} км",
                    )
                    ok_price = abs(table_price - prev_price) < 1.0
                    price_check = make_row_check(
                        ok_price, prev_price, table_price,
                        None if ok_price else f"Расхождение: {fmt(abs(table_price - prev_price))} руб.",
                    )
                else:
                    msg = f"Этап {stage_num} не найден в предыдущем ДС"
                    km_check = make_row_check(None, None, table_km, msg)
                    price_check = make_row_check(None, None, table_price, msg)
                source = "prev_ds"
            else:
                msg = "Нет данных предыдущего ДС"
                km_check = make_row_check(None, None, table_km, msg)
                price_check = make_row_check(None, None, table_price, msg)
                source = "prev_ds"

            row_checks.append({
                "stage": stage_num,
                "closed": True,
                "source": source,
                "km": km_check,
                "price": price_check,
            })

        else:
            # Открытый расчётный этап — считаем из БД
            dates = parse_stage_dates(year_str, date_range_str)
            if dates is None:
                row_checks.append({
                    "stage": stage_num, "closed": False, "source": "calculated",
                    "km": make_row_check(None, None, table_km, "Ошибка разбора дат"),
                    "price": make_row_check(None, None, table_price, "Ошибка разбора дат"),
                })
                continue

            d_from, d_to = dates

            # Расчёт км
            try:
                results = calculate_contract_period(db, contract_id, d_from, d_to)
                calc_km = sum(round(r.total_km, 2) for r in results.values())
                ok_km = abs(calc_km - table_km) < 0.02
                km_check = make_row_check(
                    ok_km, calc_km, table_km,
                    None if ok_km else f"Расхождение: {fmt(abs(calc_km - table_km))} км",
                )
            except Exception as e:
                km_check = make_row_check(None, None, table_km, f"Ошибка расчёта: {e}")
                results = {}

            # Расчёт цены
            if price_available and results:
                try:
                    route_km = {route: calc.total_km for route, calc in results.items()}
                    calc_price, no_coef_routes = calculate_price_rounded(route_km, d_to, contract_number)
                    if calc_price is not None:
                        ok_price = abs(calc_price - table_price) < 1.0
                        msg = None if ok_price else f"Расхождение: {fmt(abs(calc_price - table_price))} руб."
                        if no_coef_routes:
                            note = f"Без коэф.: {', '.join(sorted(no_coef_routes))}"
                            msg = f"{msg}; {note}" if msg else note
                        price_check = make_row_check(ok_price, calc_price, table_price, msg)
                    else:
                        all_routes = sorted(route_km.keys())
                        price_check = make_row_check(
                            None, None, table_price,
                            f"Нет коэффициентов для периода. Маршруты: {', '.join(all_routes)}"
                        )
                except Exception as e:
                    price_check = make_row_check(None, None, table_price, f"Ошибка расчёта: {e}")
            else:
                msg = "Нет данных внешней БД (вместимость/коэффициенты)" if not price_available else "Нет данных для расчёта"
                price_check = make_row_check(None, None, table_price, msg)

            row_checks.append({
                "stage": stage_num,
                "closed": False,
                "source": "calculated",
                "km": km_check,
                "price": price_check,
            })

    # Проверка строки ИТОГО (арифметика)
    itogo_check = None
    data_km_values = []
    data_price_values = []
    itogo_km = None
    itogo_price = None

    for row in data_rows:
        if str(row[0]).strip().rstrip(':').upper() == "ИТОГО":
            itogo_km = parse_ru_number(row[4])
            itogo_price = parse_ru_number(row[5])
        else:
            try:
                int(row[0])  # только числовые строки
                data_km_values.append(parse_ru_number(row[4]))
                data_price_values.append(parse_ru_number(row[5]))
            except (ValueError, TypeError):
                pass

    if itogo_km is not None and data_km_values:
        sum_km = round(sum(data_km_values), 2)
        sum_price = round(sum(data_price_values), 2)
        itogo_check = {
            "km": make_row_check(
                abs(sum_km - itogo_km) < 0.02, sum_km, itogo_km,
                None if abs(sum_km - itogo_km) < 0.02 else f"Расхождение: {fmt(abs(sum_km - itogo_km))} км",
            ),
            "price": make_row_check(
                abs(sum_price - itogo_price) < 1.0, sum_price, itogo_price,
                None if abs(sum_price - itogo_price) < 1.0 else f"Расхождение: {fmt(abs(sum_price - itogo_price))} руб.",
            ),
        }

    return {
        "row_checks": row_checks,
        "itogo_check": itogo_check,
        "price_available": price_available,
    }
