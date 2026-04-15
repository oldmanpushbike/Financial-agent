# -*- coding: utf-8 -*-
"""混合检索工具：BGE 向量 + BM25 关键词 + Cross-Encoder 重排。

三阶段流程:
  1. 双路召回: ChromaDB 向量检索 (top 20) + BM25 关键词检索 (top 20)
  2. 合并去重: 按 chunk ID 去重
  3. Cross-Encoder 重排: bge-reranker-v2-m3 打分，取 top N
"""
from __future__ import annotations

import json
import io
import contextlib
from pathlib import Path
from typing import Optional

import jieba
from rank_bm25 import BM25Okapi

_PROJECT_ROOT = Path(__file__).resolve().parent.parent  # FS/
_CHUNKS_DIR = _PROJECT_ROOT / "data" / "chunks"
_CHROMA_DIR = _PROJECT_ROOT / "data" / "chroma_db"
_COLLECTION_NAME = "research_reports"

# ── 全局缓存 ──
_bm25_index: Optional[BM25Okapi] = None
_bm25_chunks: list[dict] = []  # [{id, text, metadata}, ...]
_bm25_corpus: list[list[str]] = []  # 分词后的语料

_embedding_model = None
_reranker_model = None


def _enrich_text(text: str, metadata: dict) -> str:
    """为表格 chunk 拼接 table_title/heading 前缀，增强语义检索效果。

    普通文本 chunk 原样返回。
    """
    prefix_parts = []
    table_title = metadata.get("table_title", "")
    heading = metadata.get("heading", "")
    if table_title:
        prefix_parts.append(table_title)
    if heading and heading != table_title:
        prefix_parts.append(heading)
    if prefix_parts:
        return " ".join(prefix_parts) + "\n" + text
    return text


# ── APPEND_MARKER_1 ──


def _load_bm25_index():
    """从 data/chunks/ 加载全部 chunk，构建 BM25 索引。"""
    global _bm25_index, _bm25_chunks, _bm25_corpus
    if _bm25_index is not None:
        return

    chunks = []
    for jsonl_file in sorted(_CHUNKS_DIR.rglob("*.jsonl")):
        with open(jsonl_file, "r", encoding="utf-8") as f:
            for i, line in enumerate(f):
                line = line.strip()
                if not line:
                    continue
                obj = json.loads(line)
                text = obj["text"]
                meta = obj["metadata"]
                # 表格 chunk：将 table_title + heading 拼到文本前面，增强语义匹配
                enriched = _enrich_text(text, meta)
                chunks.append({
                    "id": f"{jsonl_file.stem}_{i}",
                    "text": text,           # 原始文本（给 LLM 用）
                    "search_text": enriched, # 增强文本（给 BM25 用）
                    "metadata": meta,
                })

    _bm25_chunks = chunks
    # jieba 分词构建语料 — 用增强文本
    _bm25_corpus = [list(jieba.cut(c["search_text"])) for c in chunks]
    _bm25_index = BM25Okapi(_bm25_corpus)

# ── APPEND_MARKER_2 ──


def _vector_search(query: str, n: int = 20) -> list[dict]:
    """BGE 向量检索，返回 [{id, text, metadata}, ...]。"""
    import chromadb

    model = _get_embedding_model()
    q_with_inst = f"为这个句子生成表示以用于检索中文相关段落：{query}"
    vec = model.encode(q_with_inst, normalize_embeddings=True).tolist()

    client = chromadb.PersistentClient(path=str(_CHROMA_DIR))
    col = client.get_or_create_collection(
        name=_COLLECTION_NAME, metadata={"hnsw:space": "cosine"}
    )
    results = col.query(query_embeddings=[vec], n_results=n)

    hits = []
    for i in range(len(results["ids"][0])):
        hits.append({
            "id": results["ids"][0][i],
            "text": results["documents"][0][i],
            "metadata": results["metadatas"][0][i],
        })
    return hits


def _bm25_search(query: str, n: int = 20) -> list[dict]:
    """BM25 关键词检索，返回 [{id, text, metadata}, ...]。"""
    _load_bm25_index()
    tokens = list(jieba.cut(query))
    scores = _bm25_index.get_scores(tokens)

    # 取 top n
    top_indices = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)[:n]
    hits = []
    for idx in top_indices:
        if scores[idx] <= 0:
            break
        chunk = _bm25_chunks[idx]
        hits.append({
            "id": chunk["id"],
            "text": chunk["text"],
            "metadata": chunk["metadata"],
        })
    return hits

# ── APPEND_MARKER_3 ──


def hybrid_search(query: str, n_results: int = 5) -> list[dict]:
    """三阶段混合检索：向量 + BM25 召回 → 去重 → Cross-Encoder 重排。

    返回 top n_results 个 chunk，每个为:
        {"id": str, "text": str, "metadata": dict, "score": float}
    """
    # Stage 1: 双路召回
    vec_hits = _vector_search(query, n=20)
    bm25_hits = _bm25_search(query, n=20)

    # Stage 2: 合并去重
    seen = set()
    candidates = []
    for hit in vec_hits + bm25_hits:
        if hit["id"] not in seen:
            seen.add(hit["id"])
            candidates.append(hit)

    if not candidates:
        return []

    # Stage 3: Cross-Encoder 重排 — 用增强文本提升表格 chunk 的匹配度
    reranker = _get_reranker()
    pairs = [(query, _enrich_text(c["text"], c["metadata"])[:512]) for c in candidates]
    scores = reranker.predict(pairs)

    for i, c in enumerate(candidates):
        c["score"] = float(scores[i])

    candidates.sort(key=lambda x: x["score"], reverse=True)
    return candidates[:n_results]


def _get_embedding_model():
    """懒加载 BGE embedding 模型。"""
    global _embedding_model
    if _embedding_model is None:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            from sentence_transformers import SentenceTransformer
            from modelscope import snapshot_download
            model_dir = snapshot_download("BAAI/bge-large-zh-v1.5")
            _embedding_model = SentenceTransformer(model_dir)
    return _embedding_model


def _get_reranker():
    """懒加载 Cross-Encoder reranker。"""
    global _reranker_model
    if _reranker_model is None:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            from sentence_transformers import CrossEncoder
            from modelscope import snapshot_download
            model_dir = snapshot_download("BAAI/bge-reranker-v2-m3")
            _reranker_model = CrossEncoder(model_dir)
    return _reranker_model
