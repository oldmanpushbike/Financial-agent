# -*- coding: utf-8 -*-
"""将切分后的 chunks 写入本地 ChromaDB 向量数据库。

Embedding 模型: BAAI/bge-large-zh-v1.5 (1024维, sentence-transformers)
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
from sentence_transformers import SentenceTransformer

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent  # FS/
_CHUNKS_DIR = _PROJECT_ROOT / "data" / "chunks"
_CHROMA_DIR = _PROJECT_ROOT / "data" / "chroma_db"
_COLLECTION_NAME = "research_reports"

# BGE-large-zh-v1.5: 1024维, 中文 RAG 标杆模型
_MODEL_NAME = "BAAI/bge-large-zh-v1.5"
_BATCH_SIZE = 64  # embedding 批次大小

_model: SentenceTransformer | None = None


def _get_model() -> SentenceTransformer:
    global _model
    if _model is None:
        from modelscope import snapshot_download
        # 从 ModelScope 下载模型到本地缓存
        model_dir = snapshot_download(_MODEL_NAME)
        print(f"加载 embedding 模型: {model_dir} ...")
        _model = SentenceTransformer(model_dir)
    return _model


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
    """用 BGE 模型对 query 做 embedding，然后检索 ChromaDB。

    BGE 系列 query 端加指令前缀可提升检索效果。
    """
    model = _get_model()
    # BGE 推荐的中文 query 指令前缀
    query_with_instruction = f"为这个句子生成表示以用于检索中文相关段落：{text}"
    query_embedding = model.encode(
        query_with_instruction, normalize_embeddings=True
    ).tolist()

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

    model = _get_model()

    texts = [c["text"] for c in chunks]
    metadatas = [c["metadata"] for c in chunks]
    # 用增强文本做 embedding（表格 chunk 拼接 table_title），提升语义检索效果
    embed_texts = [_enrich_text(c["text"], c["metadata"]) for c in chunks]

    # 生成唯一 ID: {文件stem}_{chunk序号}
    stem = jsonl_path.stem
    ids = [f"{stem}_{i}" for i in range(len(chunks))]

    # 分批 embedding + 写入
    written = 0
    for start in range(0, len(texts), _BATCH_SIZE):
        end = min(start + _BATCH_SIZE, len(texts))
        batch_embed_texts = embed_texts[start:end]
        batch_texts = texts[start:end]
        batch_ids = ids[start:end]
        batch_metas = metadatas[start:end]

        # embedding 用增强文本，document 存原始文本
        embeddings = model.encode(batch_embed_texts, normalize_embeddings=True)

        collection.add(
            ids=batch_ids,
            embeddings=embeddings.tolist(),
            documents=batch_texts,
            metadatas=batch_metas,
        )
        written += len(batch_texts)

    return written


def build_all() -> None:
    """遍历 data/chunks/ 下所有 jsonl，写入 ChromaDB。"""
    client = chromadb.PersistentClient(path=str(_CHROMA_DIR))

    # 全量重建：先删后建
    try:
        client.delete_collection(_COLLECTION_NAME)
    except Exception:
        pass
    collection = get_collection(client)

    total = 0
    files = sorted(_CHUNKS_DIR.rglob("*.jsonl"))
    print(f"共 {len(files)} 个 jsonl 文件待处理")

    t0 = time.time()
    for i, jsonl_file in enumerate(files, 1):
        n = build_from_jsonl(jsonl_file, collection)
        total += n
        if i % 50 == 0 or i == len(files):
            print(f"  [{i}/{len(files)}] 已写入 {total} 个 chunk")

    elapsed = time.time() - t0
    print(f"完成，共 {total} 个 chunk 写入 {_CHROMA_DIR}，耗时 {elapsed:.1f}s")


def build_single(jsonl_path: Path) -> None:
    """将单个 jsonl 写入 ChromaDB（追加模式）。"""
    collection = get_collection()

    # 先删除该文件已有的 chunk（按 id 前缀）
    stem = jsonl_path.stem
    existing = collection.get(where={"source_pdf": {"$ne": ""}})
    ids_to_delete = [id_ for id_ in existing["ids"] if id_.startswith(f"{stem}_")]
    if ids_to_delete:
        collection.delete(ids=ids_to_delete)

    n = build_from_jsonl(jsonl_path, collection)
    print(f"写入 {n} 个 chunk → {_CHROMA_DIR}")


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
