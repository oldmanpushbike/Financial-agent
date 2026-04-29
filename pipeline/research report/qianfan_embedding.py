# -*- coding: utf-8 -*-
"""千帆 Embedding + Rerank API 封装：直接用 API Key 鉴权。

供 build_vectordb.py 和 rag_search.py 共用。
"""
from __future__ import annotations

import json
from pathlib import Path

import requests
from dotenv import load_dotenv
import os

_ENV_PATH = Path(__file__).resolve().parent.parent.parent / "data" / ".env"
load_dotenv(_ENV_PATH, override=True)

_API_KEY = os.environ.get("BAIDU_API_KEY", "")
_MODEL = os.environ.get("BAIDU_EMBEDDING_MODEL", "embedding-v1")
_EMBED_URL = "https://qianfan.baidubce.com/v2/embeddings"
_RERANK_URL = "https://qianfan.baidubce.com/v2/rerank"
_RERANK_MODEL = "bce-reranker-base"
_MAX_TEXT_LEN = 1000  # 千帆 embedding-v1 单条文本最大长度


def embed_texts(texts: list[str], batch_size: int = 16) -> list[list[float]]:
    """调用千帆 Embedding API 批量获取向量。

    千帆 /v2/embeddings 单次最多 16 条文本，自动分批。
    返回与 texts 等长的 embedding 列表。
    """
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {_API_KEY}",
    }
    all_embeddings: list[list[float]] = []

    for start in range(0, len(texts), batch_size):
        batch = [t[:_MAX_TEXT_LEN] for t in texts[start:start + batch_size]]
        payload = json.dumps({"model": _MODEL, "input": batch})
        resp = requests.post(_EMBED_URL, headers=headers, data=payload, timeout=60)
        if resp.status_code != 200:
            print(f"千帆 API 错误 {resp.status_code}: {resp.text}")
            print(f"批次文本数: {len(batch)}, 最长文本长度: {max(len(t) for t in batch)}")
            resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"千帆 Embedding API 错误: {data['error']}")
        sorted_data = sorted(data["data"], key=lambda x: x["index"])
        all_embeddings.extend([item["embedding"] for item in sorted_data])

    return all_embeddings


def embed_query(text: str) -> list[float]:
    """单条文本 embedding，用于检索查询。"""
    return embed_texts([text])[0]


def rerank(query: str, documents: list[str], top_n: int | None = None) -> list[dict]:
    """调用千帆 Rerank API 对文档重排序。

    返回按相关性降序排列的列表，每项为:
        {"index": int, "relevance_score": float}
    """
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {_API_KEY}",
    }
    body: dict = {
        "model": _RERANK_MODEL,
        "query": query,
        "documents": documents,
    }
    if top_n is not None:
        body["top_n"] = top_n

    resp = requests.post(_RERANK_URL, headers=headers, data=json.dumps(body), timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"千帆 Rerank API 错误: {data['error']}")
    return data["results"]
