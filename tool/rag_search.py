# -*- coding: utf-8 -*-
"""混合检索工具：BGE 向量 + BM25 关键词 + Cross-Encoder 重排。

三阶段流程:
  1. 双路召回: ChromaDB 向量检索 (top 20) + BM25 关键词检索 (top 20)
  2. 合并去重: 按 chunk ID 去重
  3. Cross-Encoder 重排: bge-reranker-v2-m3 打分，取 top N
"""
from __future__ import annotations

import json
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
_bm25_chunks: list[dict] = []  # child chunks [{id, text, metadata}, ...]
_bm25_corpus: list[list[str]] = []  # 分词后的语料
_parent_lookup: dict[str, dict] = {}  # parent_id → {text, metadata}


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
    """从 data/chunks/ 加载 child chunks 构建 BM25 索引，同时加载 parent chunks 做 lookup。"""
    global _bm25_index, _bm25_chunks, _bm25_corpus, _parent_lookup
    if _bm25_index is not None:
        return

    # 加载 parent chunks
    for jsonl_file in sorted(_CHUNKS_DIR.rglob("*.jsonl")):
        if "_children" in jsonl_file.name:
            continue
        stem = jsonl_file.stem
        with open(jsonl_file, "r", encoding="utf-8") as f:
            for i, line in enumerate(f):
                line = line.strip()
                if not line:
                    continue
                obj = json.loads(line)
                pid = f"{stem}_{i}"
                _parent_lookup[pid] = {"text": obj["text"], "metadata": obj["metadata"]}

    # 加载 child chunks 构建 BM25
    chunks = []
    for jsonl_file in sorted(_CHUNKS_DIR.rglob("*_children.jsonl")):
        with open(jsonl_file, "r", encoding="utf-8") as f:
            for i, line in enumerate(f):
                line = line.strip()
                if not line:
                    continue
                obj = json.loads(line)
                text = obj["text"]
                meta = obj["metadata"]
                enriched = _enrich_text(text, meta)
                chunks.append({
                    "id": f"{jsonl_file.stem}_{i}",
                    "text": text,
                    "search_text": enriched,
                    "metadata": meta,
                })

    _bm25_chunks = chunks
    _bm25_corpus = [list(jieba.cut(c["search_text"])) for c in chunks]
    _bm25_index = BM25Okapi(_bm25_corpus)

# ── APPEND_MARKER_2 ──


def _vector_search(query: str, n: int = 20) -> list[dict]:
    """千帆 API 向量检索，返回 [{id, text, metadata}, ...]。"""
    import chromadb
    import sys
    from pathlib import Path

    # 添加 pipeline 路径以导入 qianfan_embedding
    _pipeline_dir = str(Path(__file__).resolve().parent.parent / "pipeline" / "research report")
    if _pipeline_dir not in sys.path:
        sys.path.insert(0, _pipeline_dir)
    from qianfan_embedding import embed_query

    vec = embed_query(query)

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
    """三阶段混合检索：向量 + BM25 召回 → 去重 → 千帆 Rerank → 映射回 parent。

    返回 top n_results 个 chunk，每个为:
        {"id": str, "text": str, "metadata": dict, "score": float}
    其中 text 为 parent 的完整文本。
    """
    import sys
    from pathlib import Path

    _pipeline_dir = str(Path(__file__).resolve().parent.parent / "pipeline" / "research report")
    if _pipeline_dir not in sys.path:
        sys.path.insert(0, _pipeline_dir)
    from qianfan_embedding import rerank

    _load_bm25_index()

    # Stage 1: 双路召回（child chunks）
    vec_hits = _vector_search(query, n=20)
    bm25_hits = _bm25_search(query, n=20)

    # Stage 2: 合并去重（按 child id）
    seen = set()
    candidates = []
    for hit in vec_hits + bm25_hits:
        if hit["id"] not in seen:
            seen.add(hit["id"])
            candidates.append(hit)

    if not candidates:
        return []

    # Stage 3: 千帆 Rerank API 重排（用 child 短文本）
    documents = [_enrich_text(c["text"], c["metadata"])[:512] for c in candidates]
    rerank_results = rerank(query, documents, top_n=min(len(candidates), n_results * 3))

    # Stage 4: 映射回 parent，按 parent_id 去重
    seen_parents = set()
    ranked = []
    for item in rerank_results:
        idx = item["index"]
        c = candidates[idx]
        parent_id = c["metadata"].get("parent_id", "")
        if parent_id in seen_parents:
            continue
        seen_parents.add(parent_id)

        parent = _parent_lookup.get(parent_id)
        if parent:
            ranked.append({
                "id": parent_id,
                "text": parent["text"],
                "metadata": parent["metadata"],
                "score": float(item["relevance_score"]),
            })
        else:
            c["score"] = float(item["relevance_score"])
            ranked.append(c)

        if len(ranked) >= n_results:
            break

    return ranked
