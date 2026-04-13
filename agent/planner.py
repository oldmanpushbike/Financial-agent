# -*- coding: utf-8 -*-
"""Planner 节点：用 LLM 解析用户意图，输出 IntentFlags。"""
from __future__ import annotations

import json
import re

from langchain_core.messages import HumanMessage, SystemMessage

from agent.llm import get_llm
from agent.state import AgentState, IntentFlags

PLANNER_SYSTEM_PROMPT = """你是一个财务问答系统的意图规划器（Planner）。你的唯一任务是分析用户问题，输出 JSON 意图标志。

## 你必须判断以下标志（全部为 bool）：

1. **needs_clarification**: 用户问题缺少关键实体（如公司名/代码、年份、指标），或者逻辑不清晰，无法进行任何有效查询。此时你还需要提供 clarification_message 字段，内容为你要反问用户的具体问题。
2. **needs_sql**: 问题涉及结构化财务数据的精确查询（营业收入、净利润、资产负债率、排名等数值型问题）。
3. **needs_rag**: 问题涉及定性分析、原因探究、研报总结、行业趋势等非结构化信息检索。
4. **needs_plot**: 问题明确要求生成图表、可视化、趋势图、对比图等。此标志依附于 needs_sql，仅在 needs_sql=true 时才可能为 true。

## 组合规则：
- needs_clarification=true 时，其余标志全部为 false。
- needs_sql 和 needs_rag 可以同时为 true（复杂综合问题）。
- needs_plot 仅在 needs_sql=true 时才可为 true。

## 输出格式（严格 JSON，不要输出其他任何内容）：
```json
{
  "needs_clarification": false,
  "needs_sql": true,
  "needs_rag": false,
  "needs_plot": false,
  "clarification_message": ""
}
```
"""


def _parse_intent(raw: str) -> IntentFlags:
    """从 LLM 输出中提取 JSON 意图。"""
    # 尝试提取 JSON 块
    m = re.search(r"\{[^{}]+\}", raw, re.DOTALL)
    if m:
        try:
            d = json.loads(m.group())
            return IntentFlags(
                needs_clarification=bool(d.get("needs_clarification", False)),
                needs_sql=bool(d.get("needs_sql", False)),
                needs_rag=bool(d.get("needs_rag", False)),
                needs_plot=bool(d.get("needs_plot", False)),
                clarification_message=str(d.get("clarification_message", "")),
            )
        except (json.JSONDecodeError, TypeError):
            pass
    # 回退：无法解析 → 要求澄清
    return IntentFlags(
        needs_clarification=True,
        needs_sql=False,
        needs_rag=False,
        needs_plot=False,
        clarification_message="抱歉，我无法理解您的问题，请重新描述。",
    )


_planner_llm = None


def _get_planner_llm():
    global _planner_llm
    if _planner_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        # planner 用轻量模型，速度优先
        _planner_llm = ChatOpenAI(
            model="glm-4-flash",
            api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.1,
            streaming=False,
        )
    return _planner_llm


def planner_node(state: AgentState) -> dict:
    """Planner 节点：解析用户最新消息的意图。"""
    # 取最后一条用户消息
    user_msg = ""
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            user_msg = m.content
            break
        if isinstance(m, dict) and m.get("role") == "user":
            user_msg = m.get("content", "")
            break

    if not user_msg:
        return {
            "intent": IntentFlags(
                needs_clarification=True,
                needs_sql=False,
                needs_rag=False,
                needs_plot=False,
                clarification_message="请输入您的问题。",
            )
        }

    llm = _get_planner_llm()
    resp = llm.invoke([
        SystemMessage(content=PLANNER_SYSTEM_PROMPT),
        HumanMessage(content=user_msg),
    ])
    intent = _parse_intent(resp.content)
    return {"intent": intent}
