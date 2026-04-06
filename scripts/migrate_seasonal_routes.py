"""
Миграция: установить корректный сезон (WINTER/SUMMER) для RouteParams
сезонных маршрутов, которые хранятся как ALL_YEAR.

Маршруты: 207, 305, 1КР, 2КР, 3КР
Логика: определяем сезон по значению date_from каждой записи.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from db.database import SessionLocal
from db.models import RouteParams, SeasonType
from core.constants import ROUTE_SEASON_PERIODS


def get_season_for_date(route: str, month: int, day: int) -> SeasonType | None:
    """Определяет сезон для заданной даты (month, day) по маршруту."""
    ref = ROUTE_SEASON_PERIODS.get(route)
    if not ref:
        return None

    for season_key, (sm, sd, em, ed) in ref.items():
        if sm <= em:
            # Период в пределах одного года
            if (sm, sd) <= (month, day) <= (em, ed):
                return SeasonType.WINTER if season_key == "winter" else SeasonType.SUMMER
        else:
            # Период переходит через год
            if (month, day) >= (sm, sd) or (month, day) <= (em, ed):
                return SeasonType.WINTER if season_key == "winter" else SeasonType.SUMMER

    return None


def run_migration(dry_run: bool = True):
    db = SessionLocal()
    try:
        seasonal_routes = list(ROUTE_SEASON_PERIODS.keys())

        rows = (
            db.query(RouteParams)
            .filter(
                RouteParams.route.in_(seasonal_routes),
                RouteParams.season == SeasonType.ALL_YEAR,
            )
            .order_by(RouteParams.route, RouteParams.date_from)
            .all()
        )

        print(f"Найдено записей ALL_YEAR для сезонных маршрутов: {len(rows)}")
        print()

        updated = 0
        skipped = 0

        for rp in rows:
            month = rp.date_from.month
            day = rp.date_from.day
            new_season = get_season_for_date(rp.route, month, day)

            if new_season is None:
                print(f"  ПРОПУСК  id={rp.id} route={rp.route} date_from={rp.date_from} — не удалось определить сезон")
                skipped += 1
                continue

            season_label = "ЗИМА (winter)" if new_season == SeasonType.WINTER else "ЛЕТО (summer)"
            print(f"  {'[DRY]' if dry_run else 'UPDATE'} id={rp.id:4d} route={rp.route:4s} "
                  f"date_from={rp.date_from} -> {season_label}")

            if not dry_run:
                rp.season = new_season
            updated += 1

        print()
        print(f"Итого: обновить {updated}, пропустить {skipped}")

        if not dry_run:
            db.commit()
            print("Изменения сохранены в БД.")
        else:
            print("Это dry-run. Для применения запустите с аргументом --apply")

    finally:
        db.close()


if __name__ == "__main__":
    dry_run = "--apply" not in sys.argv
    run_migration(dry_run=dry_run)
