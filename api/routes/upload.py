"""API для загрузки и обработки архивов ДС."""
import asyncio
import json
import logging
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

from fastapi import APIRouter, UploadFile, File, Depends, HTTPException
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from db.database import get_db
from db.models import Contract, Agreement, AgreementStatus
from core.config import settings
from core.parser.km_excel import parse_km_excel, find_km_excel_in_directory
from core.checker.internal import check_json
from core.checker.db_checks import check_without_money_vs_db
from core.parser.vysvobozhdenie_parser import parse_vysvobozhdenie
from core.checker.vysvobozhdenie_checker import check_vysvobozhdenie

try:
    import rarfile
    RAR_SUPPORT = True
except ImportError:
    RAR_SUPPORT = False

def extract_rar_with_7z(archive_path: Path, dest_dir: Path) -> bool:
    """Распаковывает RAR архив с помощью 7z."""
    sevenzip = settings.sevenzip_path
    if not sevenzip:
        return False

    try:
        result = subprocess.run(
            [sevenzip, "x", "-y", f"-o{dest_dir}", str(archive_path)],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            logger.warning("7z extraction failed (code %d): %s", result.returncode, result.stderr[:500])
        return result.returncode == 0
    except Exception as e:
        logger.error("7z extraction error: %s", e)
        return False

from core.parser.extract_contract_info import process_archive

router = APIRouter()


def _save_agreement(
    db,
    contract,
    ds_number: str,
    json_data: dict,
    check_result,
) -> "Agreement":
    """
    Находит или создаёт Agreement и сохраняет результаты проверки.
    Если контракт не определён — объект не сохраняется в БД.
    """
    agreement = None
    if contract:
        agreement = db.query(Agreement).filter(
            Agreement.contract_id == contract.id,
            Agreement.number == ds_number,
        ).first()

    if agreement:
        agreement.status = AgreementStatus.CHECKED
        agreement.json_data = json_data
        agreement.check_errors = check_result.errors if check_result.errors else None
        agreement.check_warnings = check_result.warnings if check_result.warnings else None
        agreement.updated_at = datetime.now(timezone.utc)
    else:
        agreement = Agreement(
            contract_id=contract.id if contract else None,
            number=ds_number,
            status=AgreementStatus.CHECKED,
            json_data=json_data,
            check_errors=check_result.errors if check_result.errors else None,
            check_warnings=check_result.warnings if check_result.warnings else None,
        )
        if contract:
            db.add(agreement)

    if contract:
        db.commit()
        db.refresh(agreement)

    return agreement


def extract_km_data_from_archive(archive_path: Path) -> dict | None:
    """
    Извлекает данные км из архива.

    Ищет xlsx файл с листом "Приложение 3 (помесячно)" рядом с основным
    документом или папками приложений.

    Returns:
        dict с данными км или None если файл не найден
    """
    suffix = archive_path.suffix.lower()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Распаковываем архив
        try:
            if suffix == ".zip":
                with zipfile.ZipFile(archive_path, "r") as zf:
                    zf.extractall(tmpdir_path)
            elif suffix == ".rar":
                extracted = False
                # Сначала пробуем 7z (более надёжно)
                if extract_rar_with_7z(archive_path, tmpdir_path):
                    extracted = True
                elif RAR_SUPPORT:
                    # Fallback на rarfile
                    try:
                        with rarfile.RarFile(archive_path, "r") as rf:
                            rf.extractall(tmpdir_path)
                        extracted = True
                    except Exception as e:
                        logger.warning("rarfile extraction failed: %s", e)
                if not extracted:
                    return None
            else:
                return None
        except Exception as e:
            logger.error("Ошибка распаковки архива для км: %s", e)
            return None

        # Ищем файл км
        km_file = find_km_excel_in_directory(tmpdir_path)
        if not km_file:
            return None

        # Парсим
        try:
            km_data = parse_km_excel(km_file)
            return km_data.to_dict()
        except Exception as e:
            logger.error("Ошибка парсинга файла км: %s", e)
            return None


@router.post("/upload")
async def upload_archive(
    file: UploadFile = File(...),
    contract_number: str = None,
    db: Session = Depends(get_db)
):
    """
    Загружает архив ДС, парсит и проверяет.
    Возвращает SSE-поток с прогрессом обработки.
    """
    # Проверяем формат файла
    filename = file.filename.lower()
    if not (filename.endswith(".zip") or filename.endswith(".rar")):
        raise HTTPException(400, "Поддерживаются только ZIP и RAR архивы")

    # Сохраняем файл сразу (до генератора, чтобы не потерять)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_filename = f"{timestamp}_{file.filename}"
    file_path = settings.upload_dir / safe_filename

    file_bytes = await file.read()
    await asyncio.get_running_loop().run_in_executor(None, file_path.write_bytes, file_bytes)

    def sse(event: str, data: dict) -> str:
        return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"

    async def generate():
        loop = asyncio.get_running_loop()

        try:
            yield sse("progress", {"step": "upload", "status": "done", "label": "Файл загружен на сервер"})

            # Парсинг архива
            yield sse("progress", {"step": "extract", "status": "progress", "label": "Распаковка и парсинг архива..."})
            try:
                json_data = await loop.run_in_executor(None, lambda: process_archive(file_path, verbose=False))
            except Exception as e:
                yield sse("error", {"message": f"Ошибка парсинга архива: {str(e)}"})
                return
            yield sse("progress", {"step": "extract", "status": "done", "label": "Архив распакован и разобран"})

            # Извлечение км
            if json_data.get("km_data"):
                # Уже извлечено из основного документа (ГК252 — этапы из docx)
                yield sse("progress", {"step": "km", "status": "done", "label": "Данные км извлечены из ДС"})
            else:
                yield sse("progress", {"step": "km", "status": "progress", "label": "Поиск данных км (Приложение №12/13)..."})
                try:
                    km_data = await loop.run_in_executor(None, lambda: extract_km_data_from_archive(file_path))
                    if km_data:
                        json_data["km_data"] = km_data
                        yield sse("progress", {"step": "km", "status": "done", "label": "Данные км извлечены"})
                    else:
                        yield sse("progress", {"step": "km", "status": "skipped", "label": "Файл км не найден"})
                except Exception as e:
                    logger.warning("Ошибка извлечения данных км: %s", e)
                    yield sse("progress", {"step": "km", "status": "skipped", "label": "Файл км не найден"})

            # Проверка данных
            yield sse("progress", {"step": "check", "status": "progress", "label": "Проверка данных..."})
            check_result = check_json(json_data)
            yield sse("progress", {"step": "check", "status": "done", "label": "Проверка завершена"})

            # Определяем контракт
            resolved_contract_number = contract_number
            if not resolved_contract_number:
                full_number = json_data.get("general", {}).get("contract_number", "")
                if full_number and len(full_number) >= 7 and full_number.endswith("0001"):
                    resolved_contract_number = full_number[-7:-4]

            contract = None
            if resolved_contract_number:
                contract = db.query(Contract).filter(Contract.number == resolved_contract_number).first()

            # Проверка change_without_money против БД (только если контракт известен)
            if contract and json_data.get("change_without_money"):
                check_result.add_check(check_without_money_vs_db(json_data, db, contract.id))

            # Сохранение
            yield sse("progress", {"step": "save", "status": "progress", "label": "Сохранение результатов..."})
            ds_number = json_data.get("general", {}).get("ds_number", "?")
            agreement = _save_agreement(db, contract, ds_number, json_data, check_result)
            yield sse("progress", {"step": "save", "status": "done", "label": "Результаты сохранены"})

            # Проверяем наличие высвобождения в данных
            vysvobozhdenie_data = json_data.get("vysvobozhdenie")
            
            yield sse("result", {
                "success": True,
                "agreement_id": agreement.id if contract else None,
                "contract_number": resolved_contract_number,
                "ds_number": ds_number,
                "is_valid": check_result.is_valid,
                "errors": check_result.errors,
                "warnings": check_result.warnings,
                "checks": [c.to_dict() for c in check_result.checks],
                "data": json_data,
                "has_vysvobozhdenie": vysvobozhdenie_data is not None,
                "vysvobozhdenie": vysvobozhdenie_data,
            })

        except Exception as e:
            yield sse("error", {"message": f"Ошибка обработки архива: {str(e)}"})

        finally:
            if file_path.exists():
                file_path.unlink()

    return StreamingResponse(generate(), media_type="text/event-stream")


@router.post("/upload-vysvobozhdenie")
async def upload_vysvobozhdenie(
    file: UploadFile = File(...),
    db: Session = Depends(get_db)
):
    """
    Загружает ДС на высвобождение (.doc/.docx), парсит и проверяет.
    Возвращает SSE-поток с прогрессом обработки.
    """
    filename = file.filename
    filename_lower = filename.lower()
    if not (filename_lower.endswith(".doc") or filename_lower.endswith(".docx")):
        raise HTTPException(400, "Поддерживаются только файлы .doc и .docx")

    file_bytes = await file.read()

    def sse(event: str, data: dict) -> str:
        return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"

    async def generate():
        loop = asyncio.get_running_loop()

        try:
            yield sse("progress", {"step": "upload", "status": "done", "label": "Файл загружен на сервер"})

            # Парсинг документа
            yield sse("progress", {"step": "parse", "status": "progress", "label": "Парсинг документа..."})
            try:
                json_data = await loop.run_in_executor(
                    None, lambda: parse_vysvobozhdenie(file_bytes, filename)
                )
            except Exception as e:
                yield sse("error", {"message": f"Ошибка парсинга документа: {str(e)}"})
                return
            yield sse("progress", {"step": "parse", "status": "done", "label": "Документ разобран"})

            # Определяем контракт
            general = json_data.get("general", {})
            contract_short = general.get("contract_short_number")
            contract = None
            if contract_short:
                contract = db.query(Contract).filter(Contract.number == contract_short).first()

            # Проверка данных
            yield sse("progress", {"step": "check", "status": "progress", "label": "Проверка данных..."})
            check_result = check_vysvobozhdenie(
                json_data,
                session=db if contract else None,
                contract_id=contract.id if contract else None,
            )
            yield sse("progress", {"step": "check", "status": "done", "label": "Проверка завершена"})

            # Сохранение
            yield sse("progress", {"step": "save", "status": "progress", "label": "Сохранение результатов..."})
            ds_number = general.get("ds_number", "?")
            agreement = _save_agreement(db, contract, ds_number, json_data, check_result)
            yield sse("progress", {"step": "save", "status": "done", "label": "Результаты сохранены"})

            yield sse("result", {
                "success": True,
                "agreement_id": agreement.id if contract else None,
                "contract_number": contract_short,
                "ds_number": ds_number,
                "is_valid": check_result.is_valid,
                "errors": check_result.errors,
                "warnings": check_result.warnings,
                "checks": [c.to_dict() for c in check_result.checks],
                "data": json_data,
            })

        except Exception as e:
            yield sse("error", {"message": f"Ошибка обработки документа: {str(e)}"})

    return StreamingResponse(generate(), media_type="text/event-stream")


@router.post("/km-excel")
async def upload_km_excel(file: UploadFile = File(...)):
    """
    Принимает xlsx-файл (Приложение №12/13 с км),
    парсит его и возвращает JSON.
    """
    filename = file.filename.lower()
    if not (filename.endswith(".xlsx") or filename.endswith(".xls")):
        raise HTTPException(400, "Поддерживаются только .xlsx и .xls файлы")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir) / file.filename
        with open(tmp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)

        try:
            km_data = parse_km_excel(tmp_path)
            return km_data.to_dict()
        except Exception as e:
            raise HTTPException(422, f"Ошибка парсинга файла: {e}")


