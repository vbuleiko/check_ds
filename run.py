#!/usr/bin/env python3
"""Скрипт запуска веб-сервиса."""
import uvicorn

if __name__ == "__main__":
    import os
    uvicorn.run(
        "api.main:app",
        host="0.0.0.0",
        port=8000,
        reload=os.environ.get("DEBUG", "").lower() == "true",
    )
