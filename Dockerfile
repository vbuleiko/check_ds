FROM python:3.11-slim

# Системные зависимости: p7zip для RAR-архивов
RUN apt-get update && apt-get install -y --no-install-recommends \
    p7zip-full \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Устанавливаем зависимости Python (отдельным слоем для кеширования)
COPY pyproject.toml .
RUN pip install --no-cache-dir hatchling && \
    pip install --no-cache-dir .

# Копируем код приложения
COPY api/ api/
COPY core/ core/
COPY db/ db/
COPY schemas/ schemas/
COPY services/ services/
COPY scripts/ scripts/
COPY templates/ templates/
COPY static/ static/
COPY alembic.ini .

# Папки для данных создаём заранее (будут перекрыты volume-ами)
RUN mkdir -p /app/uploads /app/old

EXPOSE 8000

CMD ["uvicorn", "api.main:app", "--host", "0.0.0.0", "--port", "8000"]
