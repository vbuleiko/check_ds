"""
Исторические данные этапов исполнения контрактов.

Периоды до января 2026 включительно - константы, не подлежащие изменению.
Источник: "Этапы исполнения Контракта.docx"
"""
from datetime import date
from dataclasses import dataclass
from typing import Optional


@dataclass
class HistoricalStage:
    """Исторический этап контракта."""
    stage: int
    year: int
    period: str
    date_from: date
    date_to: date
    max_km: float
    price: Optional[float] = None  # Заполняется позже


# Контракт 222: исторические этапы (1-18)
# Данные из "Этапы исполнения Контракта.docx"
HISTORICAL_STAGES_222: list[HistoricalStage] = [
    HistoricalStage(
        stage=1, year=2022, period="Q2",
        date_from=date(2022, 4, 1), date_to=date(2022, 6, 30),
        max_km=2_366_166.18, price=414_728_738.16,
    ),
    HistoricalStage(
        stage=2, year=2022, period="Q3",
        date_from=date(2022, 7, 1), date_to=date(2022, 9, 30),
        max_km=4_842_645.47, price=867_734_117.72,
    ),
    HistoricalStage(
        stage=3, year=2022, period="Q4",
        date_from=date(2022, 10, 1), date_to=date(2022, 12, 31),
        max_km=5_165_950.31, price=927_773_040.97,
    ),
    HistoricalStage(
        stage=4, year=2023, period="Q1",
        date_from=date(2023, 1, 1), date_to=date(2023, 3, 31),
        max_km=4_996_294.13, price=1_007_422_437.09,
    ),
    HistoricalStage(
        stage=5, year=2023, period="Q2",
        date_from=date(2023, 4, 1), date_to=date(2023, 6, 30),
        max_km=5_088_313.42, price=1_023_541_327.71,
    ),
    HistoricalStage(
        stage=6, year=2023, period="Q3",
        date_from=date(2023, 7, 1), date_to=date(2023, 9, 30),
        max_km=5_267_062.13, price=1_050_320_091.00,
    ),
    HistoricalStage(
        stage=7, year=2023, period="Q4",
        date_from=date(2023, 10, 1), date_to=date(2023, 12, 31),
        max_km=5_108_109.39, price=999_616_045.26,
    ),
    HistoricalStage(
        stage=8, year=2024, period="Q1",
        date_from=date(2024, 1, 1), date_to=date(2024, 3, 31),
        max_km=5_002_759.48, price=992_635_270.89,
    ),
    HistoricalStage(
        stage=9, year=2024, period="Q2",
        date_from=date(2024, 4, 1), date_to=date(2024, 6, 30),
        max_km=5_004_649.90, price=982_713_255.23,
    ),
    HistoricalStage(
        stage=10, year=2024, period="Q3",
        date_from=date(2024, 7, 1), date_to=date(2024, 9, 30),
        max_km=5_140_704.77, price=942_579_160.84,
    ),
    HistoricalStage(
        stage=11, year=2024, period="Q4",
        date_from=date(2024, 10, 1), date_to=date(2024, 12, 31),
        max_km=5_121_175.22, price=1_004_053_058.12,
    ),
    HistoricalStage(
        stage=12, year=2025, period="Q1",
        date_from=date(2025, 1, 1), date_to=date(2025, 3, 31),
        max_km=4_974_088.13, price=1_109_419_806.95,
    ),
    HistoricalStage(
        stage=13, year=2025, period="Q2",
        date_from=date(2025, 4, 1), date_to=date(2025, 6, 30),
        max_km=4_978_870.90, price=1_136_639_743.93,
    ),
    HistoricalStage(
        stage=14, year=2025, period="Q3",
        date_from=date(2025, 7, 1), date_to=date(2025, 9, 30),
        max_km=5_054_348.09, price=1_121_556_687.50,
    ),
    HistoricalStage(
        stage=15, year=2025, period="Октябрь",
        date_from=date(2025, 10, 1), date_to=date(2025, 10, 31),
        max_km=1_715_016.71, price=396_048_604.03,
    ),
    HistoricalStage(
        stage=16, year=2025, period="Ноябрь",
        date_from=date(2025, 11, 1), date_to=date(2025, 11, 30),
        max_km=1_636_343.17, price=372_047_877.60,
    ),
    HistoricalStage(
        stage=17, year=2025, period="Декабрь",
        date_from=date(2025, 12, 1), date_to=date(2025, 12, 31),
        max_km=1_710_060.20, price=388_889_864.08,
    ),
]

# Контракт 219: исторические этапы (1-17)
HISTORICAL_STAGES_219: list[HistoricalStage] = [
    HistoricalStage(
        stage=1, year=2022, period="Q2",
        date_from=date(2022, 4, 1), date_to=date(2022, 6, 30),
        max_km=2_381_546.40, price=386_240_798.42,
    ),
    HistoricalStage(
        stage=2, year=2022, period="Q3",
        date_from=date(2022, 7, 1), date_to=date(2022, 9, 30),
        max_km=4_759_835.25, price=788_862_616.56,
    ),
    HistoricalStage(
        stage=3, year=2022, period="Q4",
        date_from=date(2022, 10, 1), date_to=date(2022, 12, 31),
        max_km=5_211_799.08, price=860_057_962.09,
    ),
    HistoricalStage(
        stage=4, year=2023, period="Q1",
        date_from=date(2023, 1, 1), date_to=date(2023, 3, 31),
        max_km=5_075_082.98, price=922_923_092.66,
    ),
    HistoricalStage(
        stage=5, year=2023, period="Q2",
        date_from=date(2023, 4, 1), date_to=date(2023, 6, 30),
        max_km=5_177_971.13, price=936_068_275.92,
    ),
    HistoricalStage(
        stage=6, year=2023, period="Q3",
        date_from=date(2023, 7, 1), date_to=date(2023, 9, 30),
        max_km=5_270_382.44, price=946_625_805.46,
    ),
    HistoricalStage(
        stage=7, year=2023, period="Q4",
        date_from=date(2023, 10, 1), date_to=date(2023, 12, 31),
        max_km=5_285_042.22, price=934_065_204.40,
    ),
    HistoricalStage(
        stage=8, year=2024, period="Q1",
        date_from=date(2024, 1, 1), date_to=date(2024, 3, 31),
        max_km=5_157_633.27, price=939_924_175.10,
    ),
    HistoricalStage(
        stage=9, year=2024, period="Q2",
        date_from=date(2024, 4, 1), date_to=date(2024, 6, 30),
        max_km=5_199_011.50, price=937_492_775.47,
    ),
    HistoricalStage(
        stage=10, year=2024, period="Q3",
        date_from=date(2024, 7, 1), date_to=date(2024, 9, 30),
        max_km=5_318_462.92, price=933_934_210.29,
    ),
    HistoricalStage(
        stage=11, year=2024, period="Q4",
        date_from=date(2024, 10, 1), date_to=date(2024, 12, 31),
        max_km=5_292_007.32, price=978_029_209.30,
    ),
    HistoricalStage(
        stage=12, year=2025, period="Q1",
        date_from=date(2025, 1, 1), date_to=date(2025, 3, 31),
        max_km=5_077_684.11, price=1_044_046_543.49,
    ),
    HistoricalStage(
        stage=13, year=2025, period="Q2",
        date_from=date(2025, 4, 1), date_to=date(2025, 6, 30),
        max_km=5_131_460.98, price=1_062_164_060.51,
    ),
    HistoricalStage(
        stage=14, year=2025, period="Q3",
        date_from=date(2025, 7, 1), date_to=date(2025, 9, 30),
        max_km=5_266_295.58, price=1_055_359_722.27,
    ),
    HistoricalStage(
        stage=15, year=2025, period="Октябрь",
        date_from=date(2025, 10, 1), date_to=date(2025, 10, 31),
        max_km=1_782_680.28, price=373_305_617.04,
    ),
    HistoricalStage(
        stage=16, year=2025, period="Ноябрь",
        date_from=date(2025, 11, 1), date_to=date(2025, 11, 30),
        max_km=1_683_953.85, price=346_774_605.81,
    ),
    HistoricalStage(
        stage=17, year=2025, period="Декабрь",
        date_from=date(2025, 12, 1), date_to=date(2025, 12, 31),
        max_km=1_765_989.90, price=365_006_868.36,
    ),
]

# Контракт 220: исторические этапы (1-17)
HISTORICAL_STAGES_220: list[HistoricalStage] = [
    HistoricalStage(
        stage=1, year=2022, period="Q2",
        date_from=date(2022, 4, 1), date_to=date(2022, 6, 30),
        max_km=1_330_322.17, price=152_463_525.16,
    ),
    HistoricalStage(
        stage=2, year=2022, period="Q3",
        date_from=date(2022, 7, 1), date_to=date(2022, 9, 30),
        max_km=4_422_786.55, price=510_604_450.81,
    ),
    HistoricalStage(
        stage=3, year=2022, period="Q4",
        date_from=date(2022, 10, 1), date_to=date(2022, 12, 31),
        max_km=4_930_925.04, price=569_950_459.83,
    ),
    HistoricalStage(
        stage=4, year=2023, period="Q1",
        date_from=date(2023, 1, 1), date_to=date(2023, 3, 31),
        max_km=4_816_169.91, price=617_830_318.60,
    ),
    HistoricalStage(
        stage=5, year=2023, period="Q2",
        date_from=date(2023, 4, 1), date_to=date(2023, 6, 30),
        max_km=4_896_785.45, price=625_264_280.02,
    ),
    HistoricalStage(
        stage=6, year=2023, period="Q3",
        date_from=date(2023, 7, 1), date_to=date(2023, 9, 30),
        max_km=4_967_614.22, price=630_353_486.53,
    ),
    HistoricalStage(
        stage=7, year=2023, period="Q4",
        date_from=date(2023, 10, 1), date_to=date(2023, 12, 31),
        max_km=4_949_809.35, price=619_788_330.95,
    ),
    HistoricalStage(
        stage=8, year=2024, period="Q1",
        date_from=date(2024, 1, 1), date_to=date(2024, 3, 31),
        max_km=4_886_989.21, price=631_856_830.75,
    ),
    HistoricalStage(
        stage=9, year=2024, period="Q2",
        date_from=date(2024, 4, 1), date_to=date(2024, 6, 30),
        max_km=4_880_488.20, price=628_259_096.65,
    ),
    HistoricalStage(
        stage=10, year=2024, period="Q3",
        date_from=date(2024, 7, 1), date_to=date(2024, 9, 30),
        max_km=4_906_849.23, price=608_157_157.87,
    ),
    HistoricalStage(
        stage=11, year=2024, period="Q4",
        date_from=date(2024, 10, 1), date_to=date(2024, 12, 31),
        max_km=4_905_759.28, price=645_649_773.75,
    ),
    HistoricalStage(
        stage=12, year=2025, period="Q1",
        date_from=date(2025, 1, 1), date_to=date(2025, 3, 31),
        max_km=4_782_652.71, price=690_466_669.68,
    ),
    HistoricalStage(
        stage=13, year=2025, period="Q2",
        date_from=date(2025, 4, 1), date_to=date(2025, 6, 30),
        max_km=4_895_097.40, price=710_578_075.12,
    ),
    HistoricalStage(
        stage=14, year=2025, period="Q3",
        date_from=date(2025, 7, 1), date_to=date(2025, 9, 30),
        max_km=5_003_974.83, price=709_113_628.54,
    ),
    HistoricalStage(
        stage=15, year=2025, period="Октябрь",
        date_from=date(2025, 10, 1), date_to=date(2025, 10, 31),
        max_km=1_688_393.83, price=247_885_638.44,
    ),
    HistoricalStage(
        stage=16, year=2025, period="Ноябрь",
        date_from=date(2025, 11, 1), date_to=date(2025, 11, 30),
        max_km=1_627_511.86, price=236_403_209.40,
    ),
    HistoricalStage(
        stage=17, year=2025, period="Декабрь",
        date_from=date(2025, 12, 1), date_to=date(2025, 12, 31),
        max_km=1_685_839.09, price=245_193_921.91,
    ),
]

# Контракт 252: исторические этапы (1-16)
HISTORICAL_STAGES_252: list[HistoricalStage] = [
    HistoricalStage(
        stage=1, year=2022, period="Q3",
        date_from=date(2022, 7, 15), date_to=date(2022, 9, 30),
        max_km=326_244.78, price=45_366_106.71,
    ),
    HistoricalStage(
        stage=2, year=2022, period="Q4",
        date_from=date(2022, 10, 1), date_to=date(2022, 12, 31),
        max_km=380_596.92, price=52_899_541.46,
    ),
    HistoricalStage(
        stage=3, year=2023, period="Q1",
        date_from=date(2023, 1, 1), date_to=date(2023, 3, 31),
        max_km=360_826.65, price=55_264_933.30,
    ),
    HistoricalStage(
        stage=4, year=2023, period="Q2",
        date_from=date(2023, 4, 1), date_to=date(2023, 6, 30),
        max_km=371_736.66, price=56_727_428.87,
    ),
    HistoricalStage(
        stage=5, year=2023, period="Q3",
        date_from=date(2023, 7, 1), date_to=date(2023, 9, 30),
        max_km=382_646.67, price=58_308_623.38,
    ),
    HistoricalStage(
        stage=6, year=2023, period="Q4",
        date_from=date(2023, 10, 1), date_to=date(2023, 12, 31),
        max_km=404_498.88, price=60_769_026.64,
    ),
    HistoricalStage(
        stage=7, year=2024, period="Q1",
        date_from=date(2024, 1, 1), date_to=date(2024, 3, 31),
        max_km=395_564.64, price=62_005_944.98,
    ),
    HistoricalStage(
        stage=8, year=2024, period="Q2",
        date_from=date(2024, 4, 1), date_to=date(2024, 6, 30),
        max_km=395_564.64, price=61_623_635.89,
    ),
    HistoricalStage(
        stage=9, year=2024, period="Q3",
        date_from=date(2024, 7, 1), date_to=date(2024, 9, 30),
        max_km=414_777.00, price=61_114_158.93,
    ),
    HistoricalStage(
        stage=10, year=2024, period="Q4",
        date_from=date(2024, 10, 1), date_to=date(2024, 12, 31),
        max_km=411_737.69, price=64_325_305.72,
    ),
    HistoricalStage(
        stage=11, year=2025, period="Q1",
        date_from=date(2025, 1, 1), date_to=date(2025, 3, 31),
        max_km=390_726.48, price=65_475_813.16,
    ),
    HistoricalStage(
        stage=12, year=2025, period="Q2",
        date_from=date(2025, 4, 1), date_to=date(2025, 6, 30),
        max_km=396_820.96, price=68_992_990.16,
    ),
    HistoricalStage(
        stage=13, year=2025, period="Q3",
        date_from=date(2025, 7, 1), date_to=date(2025, 9, 30),
        max_km=415_425.90, price=72_347_512.28,
    ),
    HistoricalStage(
        stage=14, year=2025, period="Октябрь",
        date_from=date(2025, 10, 1), date_to=date(2025, 10, 31),
        max_km=141_686.80, price=24_636_206.99,
    ),
    HistoricalStage(
        stage=15, year=2025, period="Ноябрь",
        date_from=date(2025, 11, 1), date_to=date(2025, 11, 30),
        max_km=129_809.69, price=22_571_039.73,
    ),
    HistoricalStage(
        stage=16, year=2025, period="Декабрь",
        date_from=date(2025, 12, 1), date_to=date(2025, 12, 31),
        max_km=139_805.59, price=24_309_106.10,
    ),
]

# Граница расчётов для ГК222: с этой даты данные считаются системой
CALCULATION_START_DATE = date(2026, 1, 1)
LAST_HISTORICAL_STAGE = 17

# Конфигурация контрактов: начальная дата расчётов
CONTRACT_CONFIG = {
    "222": {
        "start_date": date(2026, 1, 1),  # Январь 2026 стал расчётным
        "first_stage": 18,               # Первый расчётный этап
        "has_historical": True,          # Есть исторические данные
    },
    "219": {
        "start_date": date(2026, 1, 1),  # После исторических этапов
        "first_stage": 18,               # Первый расчётный этап (после 17 исторических)
        "has_historical": True,
    },
    "220": {
        "start_date": date(2026, 1, 1),  # После исторических этапов
        "first_stage": 18,               # Первый расчётный этап (после 17 исторических)
        "has_historical": True,
    },
    "252": {
        "start_date": date(2026, 1, 1),  # После исторических этапов
        "first_stage": 17,               # Первый расчётный этап (после 16 исторических)
        "has_historical": True,
    },
}


def get_contract_config(contract_number: str) -> dict:
    """Получить конфигурацию контракта."""
    return CONTRACT_CONFIG.get(contract_number, {
        "start_date": date(2022, 4, 1),
        "first_stage": 1,
        "has_historical": False,
    })

# Плановые этапы из docx (для сверки расчётов)
# Этапы 19-48: февраль 2026 - июль 2028
PLANNED_STAGES_222: list[dict] = [
    {"stage": 19, "year": 2026, "period": "Февраль",  "date_from": "2026-02-01", "date_to": "2026-02-28", "max_km": 1_559_886.96},
    {"stage": 20, "year": 2026, "period": "Март",     "date_from": "2026-03-01", "date_to": "2026-03-31", "max_km": 1_725_418.96},
    {"stage": 21, "year": 2026, "period": "Апрель",   "date_from": "2026-04-01", "date_to": "2026-04-30", "max_km": 1_658_370.83},
    {"stage": 22, "year": 2026, "period": "Май",      "date_from": "2026-05-01", "date_to": "2026-05-31", "max_km": 1_691_276.77},
    {"stage": 23, "year": 2026, "period": "Июнь",     "date_from": "2026-06-01", "date_to": "2026-06-30", "max_km": 1_656_225.03},
    {"stage": 24, "year": 2026, "period": "Июль",     "date_from": "2026-07-01", "date_to": "2026-07-31", "max_km": 1_721_153.84},
    {"stage": 25, "year": 2026, "period": "Август",   "date_from": "2026-08-01", "date_to": "2026-08-31", "max_km": 1_706_204.91},
    {"stage": 26, "year": 2026, "period": "Сентябрь", "date_from": "2026-09-01", "date_to": "2026-09-30", "max_km": 1_663_689.10},
    {"stage": 27, "year": 2026, "period": "Октябрь",  "date_from": "2026-10-01", "date_to": "2026-10-31", "max_km": 1_713_689.77},
    {"stage": 28, "year": 2026, "period": "Ноябрь",   "date_from": "2026-11-01", "date_to": "2026-11-30", "max_km": 1_648_781.75},
    {"stage": 29, "year": 2026, "period": "Декабрь",  "date_from": "2026-12-01", "date_to": "2026-12-31", "max_km": 1_721_153.84},
    {"stage": 30, "year": 2027, "period": "Январь",   "date_from": "2027-01-01", "date_to": "2027-01-31", "max_km": 1_661_399.70},
    {"stage": 31, "year": 2027, "period": "Февраль",  "date_from": "2027-02-01", "date_to": "2027-02-28", "max_km": 1_541_357.92},
    {"stage": 32, "year": 2027, "period": "Март",     "date_from": "2027-03-01", "date_to": "2027-03-31", "max_km": 1_713_668.98},
    {"stage": 33, "year": 2027, "period": "Апрель",   "date_from": "2027-04-01", "date_to": "2027-04-30", "max_km": 1_663_709.89},
    {"stage": 34, "year": 2027, "period": "Май",      "date_from": "2027-05-01", "date_to": "2027-05-31", "max_km": 1_691_276.77},
    {"stage": 35, "year": 2027, "period": "Июнь",     "date_from": "2027-06-01", "date_to": "2027-06-30", "max_km": 1_656_225.03},
    {"stage": 36, "year": 2027, "period": "Июль",     "date_from": "2027-07-01", "date_to": "2027-07-31", "max_km": 1_713_689.77},
    {"stage": 37, "year": 2027, "period": "Август",   "date_from": "2027-08-01", "date_to": "2027-08-31", "max_km": 1_713_668.98},
    {"stage": 38, "year": 2027, "period": "Сентябрь", "date_from": "2027-09-01", "date_to": "2027-09-30", "max_km": 1_663_689.10},
    {"stage": 39, "year": 2027, "period": "Октябрь",  "date_from": "2027-10-01", "date_to": "2027-10-31", "max_km": 1_706_225.70},
    {"stage": 40, "year": 2027, "period": "Ноябрь",   "date_from": "2027-11-01", "date_to": "2027-11-30", "max_km": 1_656_245.82},
    {"stage": 41, "year": 2027, "period": "Декабрь",  "date_from": "2027-12-01", "date_to": "2027-12-31", "max_km": 1_721_153.84},
    {"stage": 42, "year": 2028, "period": "Январь",   "date_from": "2028-01-01", "date_to": "2028-01-31", "max_km": 1_668_863.77},
    {"stage": 43, "year": 2028, "period": "Февраль",  "date_from": "2028-02-01", "date_to": "2028-02-29", "max_km": 1_598_801.87},
    {"stage": 44, "year": 2028, "period": "Март",     "date_from": "2028-03-01", "date_to": "2028-03-31", "max_km": 1_713_710.56},
    {"stage": 45, "year": 2028, "period": "Апрель",   "date_from": "2028-04-01", "date_to": "2028-04-30", "max_km": 981_391.76},
    {"stage": 46, "year": 2028, "period": "Май",      "date_from": "2028-05-01", "date_to": "2028-05-31", "max_km": 1_015_469.59},
    {"stage": 47, "year": 2028, "period": "Июнь",     "date_from": "2028-06-01", "date_to": "2028-06-30", "max_km": 678_718.92},
    {"stage": 48, "year": 2028, "period": "Июль",     "date_from": "2028-07-01", "date_to": "2028-07-31", "max_km": 317_305.88},
]

# Общие итоги контракта 222 (из docx)
CONTRACT_222_TOTALS = {
    "total_km": 121_394_392.53,
    "total_price": 25_150_730_894.07,  # Заполнится позже
}


def get_historical_stages(contract_number: str) -> list[HistoricalStage]:
    """Получить исторические этапы контракта."""
    if contract_number == "222":
        return HISTORICAL_STAGES_222
    if contract_number == "219":
        return HISTORICAL_STAGES_219
    if contract_number == "220":
        return HISTORICAL_STAGES_220
    if contract_number == "252":
        return HISTORICAL_STAGES_252
    return []


def get_historical_km_total(contract_number: str) -> float:
    """Получить сумму км по историческим этапам."""
    stages = get_historical_stages(contract_number)
    return sum(s.max_km for s in stages)


def get_historical_price_total(contract_number: str) -> float:
    """Получить сумму цен по историческим этапам."""
    stages = get_historical_stages(contract_number)
    return sum(s.price for s in stages if s.price)
