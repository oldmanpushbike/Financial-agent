# -*- coding: utf-8 -*-
"""Routing conditions — 统一任务列表架构下只保留辅助判断。"""
from __future__ import annotations

from agent.state import AgentState

_SQL_FAIL_MARKERS = ("SQL 生成失败", "无法生成有效SQL", "SQL 执行错误", "查询无结果")


def sql_succeeded(state: AgentState) -> bool:
    """判断 sql_result 是否为有效成功结果。"""
    sql_result = state.get("sql_result", "")
    if not sql_result:
        return False
    return not any(marker in sql_result for marker in _SQL_FAIL_MARKERS)
