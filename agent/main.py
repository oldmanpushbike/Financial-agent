# -*- coding: utf-8 -*-
"""CLI 入口：多轮对话，支持 --thread-id 短期记忆。"""
from __future__ import annotations

import argparse
import json
import logging
import os
import uuid
import sys
import warnings

from langchain_core.messages import HumanMessage

from agent.graph import build_graph


def _suppress_noise():
    """静默模型加载、jieba 等无关日志。"""
    warnings.filterwarnings("ignore")
    os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")
    # modelscope snapshot_download 的进度条和 INFO/WARNING
    os.environ.setdefault("MODELSCOPE_LOG_LEVEL", str(logging.ERROR))
    # 全面压制相关 logger
    for name in (
        "modelscope", "sentence_transformers", "transformers",
        "torch", "jieba", "chromadb", "httpx", "urllib3",
    ):
        logging.getLogger(name).setLevel(logging.ERROR)
    # 压制 root logger 的 INFO（兜底）
    logging.basicConfig(level=logging.ERROR, force=True)
    # 压制 tqdm 进度条（Loading weights 等）
    os.environ.setdefault("TQDM_DISABLE", "1")
    # jieba 静默模式
    import jieba
    jieba.setLogLevel(logging.ERROR)


def main():
    _suppress_noise()

    parser = argparse.ArgumentParser(description="FS 财务数据分析 Agent (CLI)")
    parser.add_argument("--thread-id", default=None, help="会话线程 ID（用于短期记忆）")
    args = parser.parse_args()

    thread_id = args.thread_id or str(uuid.uuid4())
    config = {"configurable": {"thread_id": thread_id}}

    graph = build_graph()
    print(f"FS Agent 已启动 (thread={thread_id})")
    print("输入问题开始对话，exit/quit 退出。")
    print("知识库加载中...", end="", flush=True)

    # 预热：触发模型懒加载，避免首次查询时大量日志
    _warmup_done = False

    while True:
        try:
            if not _warmup_done:
                # 首次提示后等待用户输入，模型会在第一次查询时加载
                _warmup_done = True
                print(" 就绪\n")
            user_input = input("You> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit"):
            print("再见！")
            break

        result = graph.invoke(
            {"messages": [HumanMessage(content=user_input)]},
            config=config,
        )

        # 只输出 formatted_answer（JSON 格式）
        fa = result.get("formatted_answer", "")
        if fa:
            try:
                obj = json.loads(fa)
                # json.dumps 会把 content 里的换行转义为 \\n，
                # 终端显示时还原为真实换行
                raw = json.dumps(obj, ensure_ascii=False, indent=2)
                print(f"\nAgent>\n{raw}\n")
            except (json.JSONDecodeError, TypeError):
                print(f"\nAgent> {fa}\n")
        else:
            # 兜底：没有 formatted_answer 时用 final_answer
            answer = result.get("final_answer", "")
            if answer:
                print(f"\nAgent> {answer}\n")
            else:
                last = result["messages"][-1]
                if hasattr(last, "content"):
                    print(f"\nAgent> {last.content}\n")


if __name__ == "__main__":
    main()
