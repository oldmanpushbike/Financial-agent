# -*- coding: utf-8 -*-
"""Agent state definition — 包含 planner 意图解析结果。"""
from __future__ import annotations

from typing import Annotated, Optional

from langgraph.graph.message import add_messages
from typing_extensions import TypedDict


class IntentFlags(TypedDict, total=False):
    """Planner 解析出的意图标志。"""
    needs_clarification: bool
    needs_sql: bool
    needs_rag: bool
    needs_plot: bool
    clarification_message: str  # 当 needs_clarification=True 时的反问内容


class SubTask(TypedDict, total=False):
    """Planner 拆分出的子任务。"""
    id: str           # "task_1", "task_2"
    query: str        # 子问题文本
    tool: str         # "sql" | "rag" | "visualize"
    depends_on: str   # "" 或依赖的 task id
    status: str       # "pending" | "done" | "failed"
    result: str       # 执行结果


class AgentState(TypedDict):
    messages: Annotated[list, add_messages]
    intent: IntentFlags
    sub_tasks: list[SubTask]  # 多意图拆分的子任务列表
    sql_result: str       # sql_query_tool 的输出
    rag_result: str       # hybrid_rag_tool 的输出
    plot_path: str        # data_visualizer_tool 保存的图片路径
    final_answer: str     # synthesis 节点的最终输出
    formatted_answer: str # json_formatter 格式化后的 JSON 字符串
