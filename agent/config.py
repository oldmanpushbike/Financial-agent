# -*- coding: utf-8 -*-
"""Centralized config: loads .env and exposes settings."""
import os
from pathlib import Path

from dotenv import load_dotenv

# FS/data/.env
_ENV_PATH = Path(__file__).resolve().parent.parent / "data" / ".env"
load_dotenv(_ENV_PATH, override=True)

ZHIPU_API_KEY: str = os.environ["ZHIPU_API_KEY"]
ZHIPU_MODEL: str = os.environ.get("ZHIPU_MODEL", "glm-4.7")
DB_PATH: str = str(Path(__file__).resolve().parent.parent / "data" / "financial.db")

# Baidu Qianfan Embedding
BAIDU_API_KEY: str = os.environ.get("BAIDU_API_KEY", "")
BAIDU_EMBEDDING_MODEL: str = os.environ.get("BAIDU_EMBEDDING_MODEL", "embedding-v1")
