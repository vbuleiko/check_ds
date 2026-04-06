"""Работа с архивами (ZIP, RAR)."""
import tempfile
import zipfile
from pathlib import Path
from typing import Iterator
from contextlib import contextmanager

try:
    import rarfile
    RAR_SUPPORT = True
except ImportError:
    RAR_SUPPORT = False


@contextmanager
def extract_archive(archive_path: str | Path) -> Iterator[Path]:
    """
    Распаковывает архив во временную директорию.

    Поддерживает ZIP и RAR форматы.

    Args:
        archive_path: Путь к архиву

    Yields:
        Path: Путь к временной директории с распакованными файлами
    """
    archive_path = Path(archive_path)
    suffix = archive_path.suffix.lower()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        if suffix == ".zip":
            with zipfile.ZipFile(archive_path, "r") as zf:
                zf.extractall(tmpdir_path)

        elif suffix == ".rar":
            if not RAR_SUPPORT:
                raise ImportError("rarfile не установлен. pip install rarfile")
            with rarfile.RarFile(archive_path, "r") as rf:
                rf.extractall(tmpdir_path)

        else:
            raise ValueError(f"Неподдерживаемый формат архива: {suffix}")

        yield tmpdir_path


def list_archive_files(archive_path: str | Path) -> list[str]:
    """
    Возвращает список файлов в архиве.

    Args:
        archive_path: Путь к архиву

    Returns:
        Список путей к файлам внутри архива
    """
    archive_path = Path(archive_path)
    suffix = archive_path.suffix.lower()

    if suffix == ".zip":
        with zipfile.ZipFile(archive_path, "r") as zf:
            return zf.namelist()

    elif suffix == ".rar":
        if not RAR_SUPPORT:
            raise ImportError("rarfile не установлен. pip install rarfile")
        with rarfile.RarFile(archive_path, "r") as rf:
            return rf.namelist()

    else:
        raise ValueError(f"Неподдерживаемый формат архива: {suffix}")


def find_files_by_extension(
    directory: Path,
    extensions: list[str],
    recursive: bool = True
) -> list[Path]:
    """
    Находит файлы с указанными расширениями.

    Args:
        directory: Директория для поиска
        extensions: Список расширений (например, [".docx", ".doc"])
        recursive: Искать рекурсивно

    Returns:
        Список путей к найденным файлам
    """
    extensions = [ext.lower() for ext in extensions]
    pattern = "**/*" if recursive else "*"

    files = []
    for ext in extensions:
        files.extend(directory.glob(f"{pattern}{ext}"))

    # Фильтруем временные файлы
    files = [f for f in files if not f.name.startswith("~$")]

    return sorted(files)
