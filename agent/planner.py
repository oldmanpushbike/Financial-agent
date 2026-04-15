# -*- coding: utf-8 -*-
"""Planner 节点：用 LLM 解析用户意图，支持单一意图和多意图拆分两种模式。"""
from __future__ import annotations

import json
import re
from typing import Tuple

from langchain_core.messages import HumanMessage, SystemMessage

from agent.llm import get_llm
from agent.state import AgentState, IntentFlags, SubTask

PLANNER_SYSTEM_PROMPT = """你是一个财务问答系统的意图规划器（Planner）。分析用户问题，判断是"单一问题"还是"复合问题"，输出对应的 JSON。

## 判断规则：
- 单一问题：只涉及一个查询意图（如"金花股份2022年营收多少"）
- 复合问题：包含2个或以上独立子问题（如"金花股份营收多少？行业趋势如何？"）

## 输出格式（严格 JSON，不要输出其他任何内容）：

### 单一问题 → mode="single"：
```json
{
  "mode": "single",
  "intent": {
    "needs_clarification": false,
    "needs_sql": true,
    "needs_rag": false,
    "needs_plot": false,
    "clarification_message": ""
  }
}
```

### 复合问题 → mode="multi"：
```json
{
  "mode": "multi",
  "tasks": [
    {"id": "task_1", "query": "子问题1", "tool": "sql", "depends_on": ""},
    {"id": "task_2", "query": "子问题2", "tool": "rag", "depends_on": ""}
  ]
}
```

## intent 标志说明（单一问题时使用）：
- needs_clarification: 问题缺少关键信息（公司名、年份、指标），需要反问。此时提供 clarification_message。
- needs_sql: 涉及结构化财务数据查询（营收、净利润、资产负债率、排名等数值）
- needs_rag: 涉及研报分析、行业趋势、定性信息检索
- needs_plot: 明确要求图表/可视化（仅在 needs_sql=true 时可为 true）

## tool 类型说明（复合问题时使用）：
- sql：查询具体财务数据（营收、净利润、资产负债率等数值问题）
- rag：查询研报分析、行业趋势、定性信息
- visualize：生成图表（必须依赖一个 sql 任务，depends_on 不能为空）

## 依赖规则：
- sql 和 rag 任务的 depends_on 为 ""（独立执行）
- visualize 任务必须 depends_on 某个 sql 任务的 id
- 任务按执行顺序排列，被依赖的任务排在前面
- id 格式为 task_1, task_2, task_3...

## 示例：

用户：金花股份2022年营收多少？
```json
{"mode": "single", "intent": {"needs_clarification": false, "needs_sql": true, "needs_rag": false, "needs_plot": false, "clarification_message": ""}}
```

用户：中药行业发展趋势如何？
```json
{"mode": "single", "intent": {"needs_clarification": false, "needs_sql": false, "needs_rag": true, "needs_plot": false, "clarification_message": ""}}
```

用户：金花股份2022年营收多少？中药行业发展趋势如何？
```json
{"mode": "multi", "tasks": [{"id": "task_1", "query": "金花股份2022年营收多少", "tool": "sql", "depends_on": ""}, {"id": "task_2", "query": "中药行业发展趋势如何", "tool": "rag", "depends_on": ""}]}
```

用户：金花股份近三年营收趋势，画个图。中药行业前景如何？
```json
{"mode": "multi", "tasks": [{"id": "task_1", "query": "金花股份近三年营收数据", "tool": "sql", "depends_on": ""}, {"id": "task_2", "query": "画金花股份近三年营收趋势图", "tool": "visualize", "depends_on": "task_1"}, {"id": "task_3", "query": "中药行业前景如何", "tool": "rag", "depends_on": ""}]}
```

用户：请问
```json
{"mode": "single", "intent": {"needs_clarification": true, "needs_sql": false, "needs_rag": false, "needs_plot": false, "clarification_message": "请输入您想查询的具体问题。"}}
```
"""


def _parse_plan(raw: str) -> tuple[str, IntentFlags, list[SubTask]]:
    """解析 LLM 输出，返回 (mode, intent, sub_tasks)。"""
    m = re.search(r"\{[\s\S]+\}", raw)
    if not m:
        return _fallback_clarify("无法解析规划结果，请重新描述您的问题。")

    try:
        d = json.loads(m.group())
    except (json.JSONDecodeError, TypeError):
        return _fallback_clarify("无法解析规划结果，请重新描述您的问题。")

    mode = d.get("mode", "single")

    if mode == "multi":
        tasks_raw = d.get("tasks", [])
        if not tasks_raw:
            return _fallback_clarify("未能拆分出有效子任务，请重新描述。")
        sub_tasks = []
        for t in tasks_raw:
            sub_tasks.append(SubTask(
                id=str(t.get("id", "")),
                query=str(t.get("query", "")),
                tool=str(t.get("tool", "")),
                depends_on=str(t.get("depends_on", "")),
                status="pending",
                result="",
            ))
        # 校验：至少有一个有效任务
        valid = [t for t in sub_tasks if t["query"] and t["tool"] in ("sql", "rag", "visualize")]
        if not valid:
            return _fallback_clarify("未能拆分出有效子任务，请重新描述。")
        return "multi", IntentFlags(), valid

    # single 模式
    intent_raw = d.get("intent", d)  # 兼容直接输出 intent 字段或整个 dict
    intent = IntentFlags(
        needs_clarification=bool(intent_raw.get("needs_clarification", False)),
        needs_sql=bool(intent_raw.get("needs_sql", False)),
        needs_rag=bool(intent_raw.get("needs_rag", False)),
        needs_plot=bool(intent_raw.get("needs_plot", False)),
        clarification_message=str(intent_raw.get("clarification_message", "")),
    )
    return "single", intent, []


def _fallback_clarify(msg: str) -> tuple[str, IntentFlags, list[SubTask]]:
    """解析失败时的兜底：返回 clarify。"""
    return "single", IntentFlags(
        needs_clarification=True,
        needs_sql=False,
        needs_rag=False,
        needs_plot=False,
        clarification_message=msg,
    ), []


_planner_llm = None


def _get_planner_llm():
    global _planner_llm
    if _planner_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _planner_llm = ChatOpenAI(
            model="glm-4-flash",
            api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.1,
            streaming=False,
        )
    return _planner_llm


def planner_node(state: AgentState) -> dict:
    """Planner 节点：解析用户最新消息的意图，支持单一/多意图两种模式。

    每轮对话开始时清空上一轮的中间结果，防止跨轮 state 残留。
    """
    user_msg = ""
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            user_msg = m.content
            break
        if isinstance(m, dict) and m.get("role") == "user":
            user_msg = m.get("content", "")
            break

    _reset = {
        "sql_result": "",
        "rag_result": "",
        "plot_path": "",
        "final_answer": "",
        "formatted_answer": "",
        "sub_tasks": [],
    }

    if not user_msg:
        return {
            **_reset,
            "intent": IntentFlags(
                needs_clarification=True,
                needs_sql=False,
                needs_rag=False,
                needs_plot=False,
                clarification_message="请输入您的问题。",
            ),
        }

    llm = _get_planner_llm()
    resp = llm.invoke([
        SystemMessage(content=PLANNER_SYSTEM_PROMPT),
        HumanMessage(content=user_msg),
    ])

    mode, intent, sub_tasks = _parse_plan(resp.content)
    return {**_reset, "intent": intent, "sub_tasks": sub_tasks}
