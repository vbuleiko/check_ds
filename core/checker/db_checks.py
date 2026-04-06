"""Проверки ДС против данных в БД.

Проверяет соответствие параметров изменений без изменения объёмов
с ранее действующими параметрами маршрутов в БД.
"""
from datetime import date, timedelta

from sqlalchemy import case
from sqlalchemy.orm import Session

from core.checker.internal import CheckItem
from core.utils import parse_date as _parse_date
from db.models import RouteParams, SeasonType


def _get_active_params(
    session: Session,
    contract_id: int,
    route: str,
    check_date: date,
    season_type: SeasonType,
) -> RouteParams | None:
    """Возвращает активные RouteParams для маршрута на указанную дату и сезон."""
    q = session.query(RouteParams).filter(
        RouteParams.contract_id == contract_id,
        RouteParams.route == route,
        RouteParams.date_from <= check_date,
    ).filter(
        (RouteParams.date_to.is_(None)) | (RouteParams.date_to >= check_date)
    )

    if season_type == SeasonType.ALL_YEAR:
        q = q.filter(RouteParams.season == SeasonType.ALL_YEAR)
    else:
        season_priority = case(
            (RouteParams.season == season_type, 1),
            else_=0,
        )
        q = q.filter(
            (RouteParams.season == season_type) | (RouteParams.season == SeasonType.ALL_YEAR)
        ).order_by(
            RouteParams.date_from.desc(),
            season_priority.desc(),
            RouteParams.id.desc(),
        )
        return q.first()

    return q.order_by(RouteParams.date_from.desc(), RouteParams.id.desc()).first()


def check_without_money_vs_db(data: dict, session: Session, contract_id: int) -> CheckItem:
    """
    Проверяет, что параметры маршрутов из change_without_money совпадают
    с ранее действующими параметрами в БД.

    Для каждой записи change_without_money находит соответствующее приложение,
    затем ищет активные RouteParams на день до начала нового периода.
    Если параметры (протяжённость, кол-во рейсов) отличаются — ошибка.
    """
    item = CheckItem(
        id="without_money_params",
        label="Соответствие параметров маршрутов (без изменения объёмов)",
    )

    changes = data.get("change_without_money", [])
    if not changes:
        item.detail = "Нет изменений без изменения объёмов"
        return item

    appendices = data.get("appendices", {})
    checked_count = 0

    for change in changes:
        appendix_ref = str(change.get("appendix", "")).strip()
        route = change.get("route", "").strip()

        if not appendix_ref or not route:
            continue

        # Найти данные приложения по номеру
        app = None
        for app_data in appendices.values():
            if str(app_data.get("appendix_num", "")).strip() == appendix_ref:
                app = app_data
                break

        if not app:
            continue  # Отсутствие приложения уже проверяется в check_changes_vs_appendices

        date_from = _parse_date(app.get("date_from") or app.get("date_on"))
        if not date_from:
            continue

        # Дата, на которую ищем ранее действующие параметры
        check_date = date_from - timedelta(days=1)

        appendix_label = f"Маршрут {route}, Приложение № {appendix_ref}"

        # Определяем сезонные варианты
        has_winter = bool(app.get("period_winter") and app["period_winter"].get("num_of_types", 0) > 0)
        has_summer = bool(app.get("period_summer") and app["period_summer"].get("num_of_types", 0) > 0)

        if has_winter or has_summer:
            seasons = []
            if has_winter:
                seasons.append((SeasonType.WINTER, app["period_winter"]))
            if has_summer:
                seasons.append((SeasonType.SUMMER, app["period_summer"]))
        else:
            seasons = [(SeasonType.ALL_YEAR, app)]

        for season_type, period_data in seasons:
            active = _get_active_params(session, contract_id, route, check_date, season_type)

            if active is None:
                # Новый маршрут или нет данных — пропускаем
                continue

            checked_count += 1
            season_label = f" (сезон: {season_type.value})" if season_type != SeasonType.ALL_YEAR else ""

            # Сравниваем протяжённость (хранится в основных данных приложения, не в сезонных)
            for field_name, direction_label in [("length_forward", "прямое"), ("length_reverse", "обратное")]:
                app_val = app.get(field_name)
                db_val = getattr(active, field_name)
                if app_val is not None and db_val is not None:
                    if abs(float(app_val) - float(db_val)) > 0.001:
                        item.errors.append(
                            f"{appendix_label}: протяжённость ({direction_label}) изменилась "
                            f"({db_val} → {app_val} км), "
                            f"но ДС указан как «без изменения объёмов»"
                        )

            # Сравниваем количество рейсов по типам дней
            num_types = period_data.get("num_of_types", 0)

            app_trips: dict[str, tuple[str, int, int]] = {}  # lower_name -> (orig_name, fwd, rev)
            for i in range(1, num_types + 1):
                type_name = (period_data.get(f"type_{i}_name") or "").strip()
                fwd = int(period_data.get(f"type_{i}_forward_number") or 0)
                rev = int(period_data.get(f"type_{i}_reverse_number") or 0)
                if type_name:
                    app_trips[type_name.lower()] = (type_name, fwd, rev)

            db_trips: dict[str, tuple[int, int]] = {}
            for trip in active.trips:
                db_trips[trip.day_type_name.strip().lower()] = (trip.forward_number, trip.reverse_number)

            for name_lower, (name_orig, app_fwd, app_rev) in app_trips.items():
                if name_lower not in db_trips:
                    item.errors.append(
                        f"{appendix_label}{season_label}: тип дня «{name_orig}» "
                        f"отсутствует в предыдущих параметрах, "
                        f"но ДС указан как «без изменения объёмов»"
                    )
                    continue

                db_fwd, db_rev = db_trips[name_lower]
                if app_fwd != db_fwd or app_rev != db_rev:
                    item.errors.append(
                        f"{appendix_label}{season_label}: кол-во рейсов «{name_orig}» изменилось "
                        f"(было: {db_fwd}/{db_rev}, стало: {app_fwd}/{app_rev}), "
                        f"но ДС указан как «без изменения объёмов»"
                    )

    if not item.errors:
        if checked_count > 0:
            item.detail = f"Параметры совпадают с предыдущими (проверено {checked_count} маршрутов)"
        else:
            item.detail = "Нет данных для сравнения (новые маршруты или первый ДС в системе)"

    return item
