# -*- coding: utf-8 -*-
"""LangGraph StateGraph — Planner → 条件路由 → Tool 节点 → Synthesis → END。

当前版本：sql_query_tool 已实现，rag / visualizer / synthesis 用 stub 占位。
"""
from __future__ import annotations

from langchain_core.messages import AIMessage, SystemMessage

from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import END, StateGraph

from agent.conditions import route_after_planner, route_after_sql
from agent.llm import get_llm
from agent.planner import planner_node
from agent.state import AgentState
from tool.sql_query import (
    DB_SCHEMA,
    FEW_SHOT_EXAMPLES,
    SQL_SYSTEM_PROMPT,
    sql_query_tool,
)

# ═══════════════════════════════════════════════════════
# 节点实现
# ═══════════════════════════════════════════════════════

# ── Clarify 节点 (Condition 1) ────────────────────────

def clarify_node(state: AgentState) -> dict:
    """直接把 planner 的反问返回给用户。"""
    intent = state.get("intent", {})
    msg = intent.get("clarification_message", "请提供更多信息以便我回答您的问题。")
    return {"messages": [AIMessage(content=msg)]}


# ── SQL 节点 (Condition 2) ────────────────────────────

_sql_llm = None


def _get_sql_llm():
    global _sql_llm
    if _sql_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _sql_llm = ChatOpenAI(
            model="glm-4-flash",
            api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.3,
            streaming=False,
        )
    return _sql_llm


def sql_node(state: AgentState) -> dict:
    """用 LLM 生成 SQL → 调用 sql_query_tool 执行 → 写入 state.sql_result。"""
    user_msg = ""
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            user_msg = m.content
            break

    llm = _get_sql_llm()
    try:
        resp = llm.invoke([
            SystemMessage(content=SQL_SYSTEM_PROMPT),
            {"role": "user", "content": user_msg},
        ])
        sql = resp.content.strip()
    except Exception as e:
        return {
            "sql_result": f"SQL 生成失败: {e}",
            "messages": [AIMessage(content=f"SQL 生成失败: {e}")],
        }

    # 检查 LLM 是否返回了有效 SQL
    if not sql.upper().startswith("SELECT"):
        return {
            "sql_result": f"无法生成有效SQL。模型返回: {sql[:200]}",
            "messages": [AIMessage(content=f"无法为该问题生成有效的SQL查询。数据库中可能不包含所需字段。\n模型返回: {sql[:200]}")],
        }

    # 执行
    result = sql_query_tool.invoke({"sql": sql})
    return {
        "sql_result": f"SQL: {sql}\n\n结果:\n{result}",
        "messages": [AIMessage(content=f"[SQL 查询完成]\nSQL: {sql}\n结果:\n{result}")],
    }


# ── RAG 节点 (Condition 3) — stub ─────────────────────

def rag_node(state: AgentState) -> dict:
    """Stub: hybrid_rag_tool 尚未实现，返回占位结果。"""
    result = "[RAG stub] 研报检索功能尚未实现，后续将接入 hybrid_rag_tool。"
    return {
        "rag_result": result,
        "messages": [AIMessage(content=result)],
    }


# ── SQL+RAG 并行节点 (Condition 4) ────────────────────
# LangGraph 的并行分支通过 fan-out 实现：
# 两个独立节点分别执行，结果合并到 state。

def sql_branch_node(state: AgentState) -> dict:
    """并行分支中的 SQL 部分（复用 sql_node 逻辑）。"""
    return sql_node(state)


def rag_branch_node(state: AgentState) -> dict:
    """并行分支中的 RAG 部分（复用 rag_node 逻辑）。"""
    return rag_node(state)


# ── Visualizer 节点 (Condition 5) ─────────────────────

_viz_llm = None


def _get_viz_llm():
    global _viz_llm
    if _viz_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _viz_llm = ChatOpenAI(
            model="glm-4-flash",
            api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.1,
            streaming=False,
        )
    return _viz_llm


_VIZ_SYSTEM_PROMPT = """你是一个数据可视化助手。根据用户问题和 SQL 查询结果，决定最佳的图表类型并组织数据。

## 图表类型选择规则（必须严格遵守）：
- **line（折线图）**：用于时间序列、趋势分析、多年变化、增长率走势。关键词：趋势、变化、走势、近N年、历年。
- **bar（柱状图）**：用于不同实体之间的对比、排名、Top N。关键词：对比、排名、最高、最低、Top。
- **pie（饼图）**：用于占比、构成、结构分析。关键词：占比、构成、结构、比例、分布。

## 数据单位说明：
- 数据库中金额字段单位为万元（如 total_operating_revenue、net_profit、asset_total_assets 等），不做换算，直接使用原值。
- 百分比类字段（如 yoy_growth、roe、asset_liability_ratio）单位为 %。
- y_label 应标注正确单位：金额写"万元"，百分比写"%"。
- title 中应包含单位信息。

## 输出格式（严格 JSON，不要输出其他任何内容）：
```json
{
  "chart_type": "line",
  "title": "金花股份2020-2022年营业总收入趋势（亿元）",
  "y_label": "亿元",
  "data": {
    "labels": ["2020", "2021", "2022"],
    "values": [12.5, 14.3, 16.8]
  }
}
```

多系列数据时用 datasets:
```json
{
  "chart_type": "bar",
  "title": "营业收入与净利润对比（万元）",
  "y_label": "万元",
  "data": {
    "labels": ["2020", "2021", "2022"],
    "datasets": [
      {"label": "营业收入", "values": [100, 150, 180]},
      {"label": "净利润", "values": [20, 30, 35]}
    ]
  }
}
```
"""


def visualizer_node(state: AgentState) -> dict:
    """用 LLM 从 sql_result 中提取可视化参数，调用 data_visualizer_tool 生成图表。"""
    from tool.visualizer import data_visualizer_tool
    import json, re

    sql_result = state.get("sql_result", "")
    user_msg = ""
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            user_msg = m.content
            break

    llm = _get_viz_llm()
    resp = llm.invoke([
        SystemMessage(content=_VIZ_SYSTEM_PROMPT),
        {"role": "user", "content": f"用户问题：{user_msg}\n\nSQL查询结果：\n{sql_result}"},
    ])

    # 解析 LLM 输出
    m = re.search(r"\{[\s\S]+\}", resp.content)
    if not m:
        return {
            "plot_path": "",
            "messages": [AIMessage(content="无法生成图表：未能从数据中提取可视化参数。")],
        }

    try:
        parsed = json.loads(m.group())
        chart_type = parsed.get("chart_type", "bar")
        title = parsed.get("title", "图表")
        data = parsed.get("data", {})
        y_label = parsed.get("y_label", "")
    except (json.JSONDecodeError, TypeError):
        return {
            "plot_path": "",
            "messages": [AIMessage(content="无法生成图表：数据解析失败。")],
        }

    result = data_visualizer_tool.invoke({
        "chart_type": chart_type,
        "title": title,
        "y_label": y_label,
        "data_json": json.dumps(data, ensure_ascii=False),
    })

    plot_path = ""
    if "已保存" in result:
        # 提取路径
        plot_path = result.split("已保存: ", 1)[-1].strip()

    return {
        "plot_path": plot_path,
        "messages": [AIMessage(content=result)],
    }


# ── Synthesis 节点 (Condition 6) ──────────────────────

_synthesis_llm = None


def _get_synthesis_llm():
    global _synthesis_llm
    if _synthesis_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _synthesis_llm = ChatOpenAI(
            model="glm-4-flash",
            api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.2,
            streaming=True,
        )
    return _synthesis_llm


def synthesis_node(state: AgentState) -> dict:
    """汇总所有证据，生成最终回答。"""
    from agent.conditions import sql_succeeded

    sql_result = state.get("sql_result", "")
    rag_result = state.get("rag_result", "")
    plot_path = state.get("plot_path", "")

    has_valid_sql = sql_succeeded(state)
    has_rag = bool(rag_result) and "stub" not in rag_result

    # ── 快速兜底：SQL 失败且无 RAG → 直接返回明确失败信息，不调 LLM ──
    if sql_result and not has_valid_sql and not has_rag:
        fail_msg = (
            "抱歉，无法回答该问题。\n"
            "原因：数据库中不包含该问题所需的数据字段。"
            "当前数据库仅覆盖核心业绩指标表、利润表、资产负债表、现金流量表的结构化数据。"
        )
        return {
            "final_answer": fail_msg,
            "messages": [AIMessage(content=fail_msg)],
        }

    # ── 正常路径：收集有效证据 ──
    evidence_parts = []
    if sql_result and has_valid_sql:
        evidence_parts.append(f"【结构化查询结果】\n{sql_result}")
    if has_rag:
        evidence_parts.append(f"【研报检索结果】\n{rag_result}")
    if plot_path:
        evidence_parts.append(f"【图表路径】{plot_path}")

    evidence = "\n\n".join(evidence_parts) if evidence_parts else "无可用证据。"

    # 取用户最新问题
    user_msg = ""
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            user_msg = m.content
            break

    sys_prompt = (
        "你是一个专业的财务分析师。请**仅根据**以下收集到的证据回答用户问题，不要参考对话历史中的其他问题。\n"
        "用简洁准确的中文回答。如果证据不足以回答，请诚实说明。不要编造数据。\n\n"
        "【重要：数据库字段单位说明】\n"
        "- 金额类字段（如 total_operating_revenue、net_profit、asset_total_assets、liability_total_liabilities 等）单位均为**万元**。\n"
        "- 百分比类字段（如 yoy_growth、roe、asset_liability_ratio、gross_profit_margin 等）单位为 **%**。\n"
        "- 每股类字段（如 eps、net_asset_per_share、operating_cf_per_share）单位为 **元/股**。\n"
        "回答时必须带上正确的单位。\n\n"
        f"【证据】\n{evidence}"
    )

    llm = _get_synthesis_llm()
    resp = llm.invoke([
        SystemMessage(content=sys_prompt),
        {"role": "user", "content": user_msg},
    ])

    return {
        "final_answer": resp.content,
        "messages": [AIMessage(content=resp.content)],
    }


# ═══════════════════════════════════════════════════════
# 构建 StateGraph
# ═══════════════════════════════════════════════════════

def build_graph():
    g = StateGraph(AgentState)

    # ── 添加节点 ──
    g.add_node("planner", planner_node)
    g.add_node("clarify", clarify_node)
    g.add_node("sql", sql_node)
    g.add_node("rag", rag_node)
    g.add_node("sql_branch", sql_branch_node)
    g.add_node("rag_branch", rag_branch_node)
    g.add_node("visualizer", visualizer_node)
    g.add_node("synthesis", synthesis_node)

    # ── 入口 ──
    g.set_entry_point("planner")

    # ── Planner 之后的主路由 (Conditions 1-4) ──
    g.add_conditional_edges(
        "planner",
        route_after_planner,
        {
            "clarify": "clarify",
            "sql": "sql",
            "rag": "rag",
            "sql_and_rag": "sql_branch",
        },
    )

    # ── Condition 1: clarify → END ──
    g.add_edge("clarify", END)

    # ── Condition 2: sql → Condition 5 (需要画图?) ──
    g.add_conditional_edges(
        "sql",
        route_after_sql,
        {
            "visualize": "visualizer",
            "synthesis": "synthesis",
        },
    )

    # ── Condition 3: rag → synthesis ──
    g.add_edge("rag", "synthesis")

    # ── Condition 4: 并行分支 → synthesis ──
    # sql_branch 完成后 → rag_branch → synthesis
    # （LangGraph 无原生 fan-out+join，这里用串行模拟；
    #   后续可改为 parallel node 或 Send API）
    g.add_edge("sql_branch", "rag_branch")
    g.add_edge("rag_branch", "synthesis")

    # ── Condition 5: visualizer → synthesis ──
    g.add_edge("visualizer", "synthesis")

    # ── Condition 6: synthesis → END ──
    g.add_edge("synthesis", END)

    # ── 编译 ──
    memory = MemorySaver()
    return g.compile(checkpointer=memory)
