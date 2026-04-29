# -*- coding: utf-8 -*-
"""将切分后的 chunks 写入本地 ChromaDB 向量数据库。

Embedding: 百度千帆 Embedding API (通过 qianfan_embedding 模块调用)
存储位置: data/chroma_db/

用法:
  写入全部 chunks:
    python build_vectordb.py

  写入单个 jsonl:
    python build_vectordb.py "data/chunks/个股研报/xxx.jsonl"
"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import chromadb

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent  # FS/
_CHUNKS_DIR = _PROJECT_ROOT / "data" / "chunks"
_CHROMA_DIR = _PROJECT_ROOT / "data" / "chroma_db"
_COLLECTION_NAME = "research_reports"

_BATCH_SIZE = 16  # 千帆 API 单次最多 16 条


# 延迟导入，避免循环依赖
def _embed_texts(texts: list[str]) -> list[list[float]]:
    from qianfan_embedding import embed_texts
    return embed_texts(texts, batch_size=_BATCH_SIZE)


def _load_chunks(jsonl_path: Path) -> list[dict]:
    """从 jsonl 文件加载 chunks。"""
    chunks = []
    with open(jsonl_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                chunks.append(json.loads(line))
    return chunks


def _enrich_text(text: str, metadata: dict) -> str:
    """为表格 chunk 拼接 table_title/heading 前缀，增强向量语义。"""
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


# ── APPEND_PLACEHOLDER ──


def get_collection(client: chromadb.ClientAPI | None = None) -> chromadb.Collection:
    """获取或创建 collection（不绑定 embedding function，全部手动传 embeddings）。"""
    if client is None:
        client = chromadb.PersistentClient(path=str(_CHROMA_DIR))
    return client.get_or_create_collection(
        name=_COLLECTION_NAME,
        metadata={"hnsw:space": "cosine"},
    )


def query(text: str, n_results: int = 5) -> dict:
    """用千帆 API 对 query 做 embedding，然后检索 ChromaDB。"""
    from qianfan_embedding import embed_query
    query_embedding = embed_query(text)

    collection = get_collection()
    return collection.query(
        query_embeddings=[query_embedding],
        n_results=n_results,
    )


def build_from_jsonl(jsonl_path: Path, collection: chromadb.Collection) -> int:
    """将单个 jsonl 文件的 chunks 写入 ChromaDB collection，返回写入数量。"""
    chunks = _load_chunks(jsonl_path)
    if not chunks:
        return 0

    texts = [c["text"] for c in chunks]
    metadatas = [c["metadata"] for c in chunks]
    embed_texts_list = [_enrich_text(c["text"], c["metadata"]) for c in chunks]

    stem = jsonl_path.stem
    ids = [f"{stem}_{i}" for i in range(len(chunks))]

    embeddings = _embed_texts(embed_texts_list)

    for start in range(0, len(texts), _BATCH_SIZE):
        end = min(start + _BATCH_SIZE, len(texts))
        collection.add(
            ids=ids[start:end],
            embeddings=embeddings[start:end],
            documents=texts[start:end],
            metadatas=metadatas[start:end],
        )

    return len(texts)


def build_all() -> None:
    """遍历 data/chunks/ 下所有 children jsonl，写入 ChromaDB。"""
    client = chromadb.PersistentClient(path=str(_CHROMA_DIR))

    # 全量重建：先删后建
    try:
        client.delete_collection(_COLLECTION_NAME)
    except Exception:
        pass
    collection = get_collection(client)

    total = 0
    files = sorted(_CHUNKS_DIR.rglob("*_children.jsonl"))
    print(f"共 {len(files)} 个 children jsonl 文件待处理")

    t0 = time.time()
    for i, jsonl_file in enumerate(files, 1):
        n = build_from_jsonl(jsonl_file, collection)
        total += n
        if i % 50 == 0 or i == len(files):
            print(f"  [{i}/{len(files)}] 已写入 {total} 个 child chunk")

    elapsed = time.time() - t0
    print(f"完成，共 {total} 个 child chunk 写入 {_CHROMA_DIR}，耗时 {elapsed:.1f}s")


def build_single(jsonl_path: Path) -> None:
    """将单个 children jsonl 写入 ChromaDB（追加模式）。"""
    if not jsonl_path.name.endswith("_children.jsonl"):
        children_path = jsonl_path.with_name(jsonl_path.stem + "_children.jsonl")
        if children_path.exists():
            jsonl_path = children_path

    collection = get_collection()

    stem = jsonl_path.stem
    existing = collection.get(where={"source_pdf": {"$ne": ""}})
    ids_to_delete = [id_ for id_ in existing["ids"] if id_.startswith(f"{stem}_")]
    if ids_to_delete:
        collection.delete(ids=ids_to_delete)

    n = build_from_jsonl(jsonl_path, collection)
    print(f"写入 {n} 个 child chunk → {_CHROMA_DIR}")


def main() -> None:
    if len(sys.argv) < 2:
        build_all()
    else:
        path = Path(sys.argv[1])
        if not path.exists():
            print(f"文件不存在: {path}", file=sys.stderr)
            sys.exit(1)
        build_single(path)


if __name__ == "__main__":
    main()
