# -*- coding: utf-8 -*-
"""Routing conditions：基于 Planner 输出的 IntentFlags 决定流转方向。"""
from __future__ import annotations

from typing import Literal

from agent.state import AgentState

# ──────────────────────────────────────────────────────
# Condition 1-4: Planner 之后的主路由
# ──────────────────────────────────────────────────────

RouteAfterPlanner = Literal["clarify", "sql", "rag", "sql_and_rag"]


def route_after_planner(state: AgentState) -> RouteAfterPlanner:
    """Planner → 主路由。优先级：clarification > both > sql > rag > clarify(兜底)。"""
    intent = state.get("intent", {})

    # Condition 1: needs_clarification — 最优先
    if intent.get("needs_clarification"):
        return "clarify"

    needs_sql = intent.get("needs_sql", False)
    needs_rag = intent.get("needs_rag", False)

    # Condition 4: route_to_both
    if needs_sql and needs_rag:
        return "sql_and_rag"

    # Condition 2: route_to_sql
    if needs_sql:
        return "sql"

    # Condition 3: route_to_rag
    if needs_rag:
        return "rag"

    # 兜底：无法判断意图 → 澄清
    return "clarify"


# ──────────────────────────────────────────────────────
# Condition 5: SQL 之后 — 是否需要可视化
# ──────────────────────────────────────────────────────

# SQL 失败标志关键词
_SQL_FAIL_MARKERS = ("SQL 生成失败", "无法生成有效SQL", "SQL 执行错误", "查询无结果")


def sql_succeeded(state: AgentState) -> bool:
    """判断 sql_result 是否为有效成功结果。"""
    sql_result = state.get("sql_result", "")
    if not sql_result:
        return False
    return not any(marker in sql_result for marker in _SQL_FAIL_MARKERS)


RouteAfterSQL = Literal["visualize", "synthesis"]


def route_after_sql(state: AgentState) -> RouteAfterSQL:
    """SQL 节点完成后：如果 needs_plot=True 且查询成功有数据 → 可视化，否则 → synthesis。"""
    intent = state.get("intent", {})

    if intent.get("needs_plot") and sql_succeeded(state):
        return "visualize"
    return "synthesis"


# ──────────────────────────────────────────────────────
# Condition 6: 所有证据收集完毕 → synthesis（隐式，通过 edge 直连）
# route_to_synthesis 不需要条件函数，直接用 add_edge 即可。
# ──────────────────────────────────────────────────────
