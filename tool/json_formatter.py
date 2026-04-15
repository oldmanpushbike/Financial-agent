# -*- coding: utf-8 -*-
"""JSON 格式化工具：将 agent 每轮输出统一为标准 JSON 结构。

输出格式（单轮）:
  有 RAG:  {"Q": "...", "A": {"content": "...", "image": [...], "references": [...]}}
  无 RAG:  {"Q": "...", "A": {"content": "...", "image": [...]}}

image / references 字段仅在有值时出现。
"""
from __future__ import annotations

import json
from typing import Any


def format_turn(
    question: str,
    content: str,
    plot_path: str = "",
    rag_result: str = "",
) -> dict:
    """从单轮会话的各字段组装标准 JSON dict。"""
    # 将 content 中的换行符替换为更紧凑的格式
    content = content.replace("\n\n", "；").replace("\n", "；").strip("；")
    a: dict[str, Any] = {"content": content}

    # image（支持逗号分隔的多图片路径）
    if plot_path:
        paths = [p.strip() for p in plot_path.split(",") if p.strip()]
        if paths:
            a["image"] = paths

    # references（从 rag_result JSON 中提取）
    if rag_result:
        try:
            rag_obj = json.loads(rag_result)
            refs = rag_obj.get("references", [])
            if refs:
                a["references"] = refs
        except (json.JSONDecodeError, TypeError):
            pass

    return {"Q": question, "A": a}


def format_history(turns: list[dict]) -> str:
    """多轮会话格式化为 JSON 字符串。"""
    return json.dumps(turns, ensure_ascii=False, indent=2)
