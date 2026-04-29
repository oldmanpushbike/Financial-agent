# -*- coding: utf-8 -*-
"""Agent state definition — 统一任务列表架构。"""
from __future__ import annotations

from typing import Annotated

from langgraph.graph.message import add_messages
from typing_extensions import TypedDict


class SubTask(TypedDict, total=False):
    """Planner 输出的子任务。tool: sql | rag | visualize | clarify | chat"""
    id: str
    query: str
    tool: str
    depends_on: str
    status: str
    result: str


class AgentState(TypedDict):
    messages: Annotated[list, add_messages]
    sub_tasks: list[SubTask]
    sql_result: str
    rag_result: str
    plot_path: str
    final_answer: str
    formatted_answer: str
