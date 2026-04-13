# -*- coding: utf-8 -*-
"""CLI 入口：多轮对话，支持 --thread-id 短期记忆。"""
from __future__ import annotations

import argparse
import uuid
import sys

from langchain_core.messages import HumanMessage

from agent.graph import build_graph


def main():
    parser = argparse.ArgumentParser(description="FS 财务数据分析 Agent (CLI)")
    parser.add_argument("--thread-id", default=None, help="会话线程 ID（用于短期记忆）")
    args = parser.parse_args()

    thread_id = args.thread_id or str(uuid.uuid4())
    config = {"configurable": {"thread_id": thread_id}}

    graph = build_graph()
    print(f"FS Agent 已启动 (thread={thread_id})")
    print("输入问题开始对话，exit/quit 退出。\n")

    while True:
        try:
            user_input = input("You> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit"):
            print("再见！")
            break

        events = graph.stream(
            {"messages": [HumanMessage(content=user_input)]},
            config=config,
            stream_mode="values",
        )

        for event in events:
            last = event["messages"][-1]
            if hasattr(last, "content") and last.content and last.type == "ai":
                # 跳过纯 tool_call 消息（content 为空或只含 tool_calls）
                if not getattr(last, "tool_calls", None):
                    print(f"Agent> {last.content}\n")


if __name__ == "__main__":
    main()
