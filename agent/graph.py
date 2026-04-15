# -*- coding: utf-8 -*-
"""LangGraph StateGraph — Planner → 条件路由 → Tool 节点 → Synthesis → END。

支持单一意图（原有路径）和多意图拆分（executor 路径）。
"""
from __future__ import annotations

import json
import re
from html.parser import HTMLParser

from langchain_core.messages import AIMessage, SystemMessage

from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import END, StateGraph

from agent.conditions import route_after_planner, route_after_sql
from agent.llm import get_llm
from agent.planner import planner_node


# ═══════════════════════════════════════════════════════
# HTML 表格 → 纯文本
# ═══════════════════════════════════════════════════════

def _html_table_to_text(html: str) -> str:
    """将 <table> HTML 转为可读纯文本，每行用 | 分隔单元格。"""
    if "<table>" not in html:
        return html

    class _Parser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows: list[list[str]] = []
            self._row: list[str] = []
            self._cell = ""
            self._in_cell = False

        def handle_starttag(self, tag, attrs):
            if tag == "tr":
                self._row = []
            elif tag in ("td", "th"):
                self._cell = ""
                self._in_cell = True

        def handle_endtag(self, tag):
            if tag in ("td", "th"):
                self._in_cell = False
                self._row.append(self._cell.strip())
            elif tag == "tr":
                if self._row:
                    self.rows.append(self._row)

        def handle_data(self, data):
            if self._in_cell:
                self._cell += data

    p = _Parser()
    p.feed(html)
    if not p.rows:
        return html
    return "\n".join(" | ".join(row) for row in p.rows)
from agent.state import AgentState
from tool.sql_query import (
    DB_SCHEMA,
    FEW_SHOT_EXAMPLES,
    SQL_SYSTEM_PROMPT,
    sql_query_tool,
    _clean_sql,
)

# ═══════════════════════════════════════════════════════
# LLM 懒加载
# ═══════════════════════════════════════════════════════

_sql_llm = None
_rag_llm = None
_viz_llm = None
_synthesis_llm = None


def _get_sql_llm():
    global _sql_llm
    if _sql_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _sql_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.3, streaming=False,
        )
    return _sql_llm

_trace_llm = None


def _get_trace_llm():
    global _trace_llm
    if _trace_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _trace_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0, streaming=False,
        )
    return _trace_llm


def _get_rag_llm():
    global _rag_llm
    if _rag_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _rag_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.1, streaming=False,
        )
    return _rag_llm


def _get_viz_llm():
    global _viz_llm
    if _viz_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _viz_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.1, streaming=False,
        )
    return _viz_llm


def _get_synthesis_llm():
    global _synthesis_llm
    if _synthesis_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _synthesis_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.2, streaming=True,
        )
    return _synthesis_llm


# ═══════════════════════════════════════════════════════
# 可复用工具函数（供节点和 executor 共用）
# ═══════════════════════════════════════════════════════

def _execute_sql(query: str) -> str:
    """生成 SQL + 执行，返回结果字符串。"""
    # 去掉可视化相关词汇，避免干扰 SQL 生成模型
    clean_query = re.sub(r"[，,]?\s*(做|进行|生成)?\s*(可视化|绘图|画图|图表|画个图|做个图)[。？]?", "", query).strip()
    if not clean_query:
        clean_query = query
    llm = _get_sql_llm()
    try:
        resp = llm.invoke([
            SystemMessage(content=SQL_SYSTEM_PROMPT),
            {"role": "user", "content": clean_query},
        ])
        sql = _clean_sql(resp.content.strip())
    except Exception as e:
        return f"SQL 生成失败: {e}"

    if not sql.upper().startswith("SELECT"):
        return f"无法生成有效SQL。模型返回: {sql[:200]}"

    result = sql_query_tool.invoke({"sql": sql})
    return f"SQL: {sql}\n\n结果:\n{result}"


_RAG_SYSTEM_PROMPT = """你是一个专业的中药行业研报分析师。请严格根据以下检索到的研报片段回答用户问题。

## 规则：
1. 只使用提供的研报内容回答，不要编造任何数据或信息
2. 如果检索内容不足以回答问题，明确说明"根据现有研报资料，暂无相关信息"
3. 回答要专业、简洁，引用具体数据时标注来源编号（用 [编号] 标注）
4. 如果涉及多篇研报的信息，综合分析后给出结论
5. 如果片段中包含表格数据，必须完整提取表格中的所有条目，不要遗漏

## 表格理解说明：
研报片段中的表格已转为纯文本格式，每行用 | 分隔列。第一行通常是表头。
例如：
序号 | 公司 | 药品 | 注册分类
1 | 华润三九 | 益气清肺颗粒 | 中药3.2类
2 | 以岭药业 | 芪防鼻通片 | 中药1.1类
→ 这个表格包含2条记录，回答时必须列出全部2条，不能只提其中一部分。

## 检索到的研报片段：
{context}
"""


def _execute_rag(query: str) -> str:
    """混合检索 + LLM 生成，返回 JSON 结果字符串（含 content + references）。"""
    from tool.rag_search import hybrid_search

    try:
        hits = hybrid_search(query, n_results=5)
    except Exception as e:
        return json.dumps({"content": f"研报检索失败: {e}", "references": []}, ensure_ascii=False)

    if not hits:
        return json.dumps({"content": "研报检索无结果。", "references": []}, ensure_ascii=False)

    _RERANK_THRESHOLD = 0.5
    filtered = [h for h in hits if h.get("score", 0) >= _RERANK_THRESHOLD]
    if not filtered:
        return json.dumps({"content": "根据现有研报资料，暂无与该问题高度相关的信息。", "references": []}, ensure_ascii=False)
    hits = filtered

    # 构建 context — 表格 chunk 转纯文本且不截断
    context_parts = []
    for i, h in enumerate(hits, 1):
        title = h["metadata"].get("title", "")
        heading = h["metadata"].get("heading", "")
        label = f"[{i}] 《{title}》"
        if heading:
            label += f" - {heading}"
        chunk_text = h["text"]
        if "<table>" in chunk_text:
            chunk_text = _html_table_to_text(chunk_text)
        else:
            chunk_text = chunk_text[:800]
        context_parts.append(f"{label}\n{chunk_text}")
    context = "\n\n---\n\n".join(context_parts)

    llm = _get_rag_llm()
    try:
        resp = llm.invoke([
            SystemMessage(content=_RAG_SYSTEM_PROMPT.format(context=context)),
            {"role": "user", "content": query},
        ])
        content = resp.content
    except Exception as e:
        content = f"LLM 生成失败: {e}"

    # 溯源：用 LLM 判断哪些 chunk 真正被引用在回答中
    references = _trace_references(content, hits)

    return json.dumps({"content": content, "references": references}, ensure_ascii=False)


def _trace_references(answer: str, hits: list[dict]) -> list[dict]:
    """用强模型判断哪些 chunk 真正被回答引用，只保留被引用的。"""
    if not hits or not answer:
        return []

    llm = _get_trace_llm()
    # 一次性把所有 chunk 编号送给 LLM，让它返回被引用的编号列表
    chunk_list = []
    for i, h in enumerate(hits, 1):
        title = h["metadata"].get("title", "")
        chunk_text = h["text"]
        if "<table>" in chunk_text:
            chunk_text = _html_table_to_text(chunk_text)
        else:
            chunk_text = chunk_text[:600]
        chunk_list.append(f"[{i}] 《{title}》\n{chunk_text}")
    chunks_text = "\n\n".join(chunk_list)

    prompt = (
        "以下是一段回答和若干研报片段。请判断回答中实际引用或参考了哪些片段的内容。\n\n"
        "只返回被引用片段的编号，用逗号分隔，如: 1,3\n"
        "如果没有任何片段被引用，返回: 无\n\n"
        f"【回答】\n{answer}\n\n"
        f"【研报片段】\n{chunks_text}"
    )

    try:
        resp = llm.invoke([{"role": "user", "content": prompt}])
        text = resp.content.strip()
    except Exception:
        return []

    if text == "无":
        return []

    # 解析编号
    indices = set()
    for part in re.findall(r"\d+", text):
        idx = int(part)
        if 1 <= idx <= len(hits):
            indices.add(idx)

    refs = []
    for idx in sorted(indices):
        h = hits[idx - 1]
        meta = h["metadata"]
        # 表格转纯文本摘要
        ref_text = h["text"]
        if "<table>" in ref_text:
            ref_text = _html_table_to_text(ref_text)
        ref_text = ref_text[:300]
        refs.append({
            "paper_path": meta.get("source_pdf", ""),
            "text": ref_text,
            "paper_image": "",
        })
    return refs


# ── PLACEHOLDER_VIZ_PROMPT ──

_VIZ_SYSTEM_PROMPT = """你是一个数据可视化助手。根据用户问题和 SQL 查询结果，决定最佳的图表类型并组织数据。

## 图表类型选择规则（必须严格遵守）：
- **line（折线图）**：用于时间序列、趋势分析、多年变化、增长率走势。关键词：趋势、变化、走势、近N年、历年。
- **bar（柱状图）**：用于不同实体之间的对比、排名、Top N。关键词：对比、排名、最高、最低、Top。
- **pie（饼图）**：用于占比、构成、结构分析。关键词：占比、构成、结构、比例、分布。

## 数据单位说明：
- 数据库中金额字段单位为万元，不做换算，直接使用原值。
- 百分比类字段单位为 %。
- y_label 应标注正确单位：金额写"万元"，百分比写"%"。
- title 中应包含单位信息。

## 输出格式（严格 JSON，不要输出其他任何内容）：
```json
{
  "chart_type": "line",
  "title": "金花股份2020-2022年营业总收入趋势（万元）",
  "y_label": "万元",
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


def _execute_viz(query: str, sql_result: str) -> str:
    """基于 SQL 结果生成图表，返回图片路径或错误信息。"""
    from tool.visualizer import data_visualizer_tool

    llm = _get_viz_llm()
    resp = llm.invoke([
        SystemMessage(content=_VIZ_SYSTEM_PROMPT),
        {"role": "user", "content": f"用户问题：{query}\n\nSQL查询结果：\n{sql_result}"},
    ])

    m = re.search(r"\{[\s\S]+\}", resp.content)
    if not m:
        return ""

    try:
        parsed = json.loads(m.group())
        chart_type = parsed.get("chart_type", "bar")
        title = parsed.get("title", "图表")
        data = parsed.get("data", {})
        y_label = parsed.get("y_label", "")
    except (json.JSONDecodeError, TypeError):
        return ""

    result = data_visualizer_tool.invoke({
        "chart_type": chart_type,
        "title": title,
        "y_label": y_label,
        "data_json": json.dumps(data, ensure_ascii=False),
    })

    if "已保存" in result:
        return result.split("已保存: ", 1)[-1].strip()
    return ""


# ═══════════════════════════════════════════════════════
# 图节点实现（单一意图路径，复用工具函数）
# ═══════════════════════════════════════════════════════

def _get_user_msg(state: AgentState) -> str:
    """从 state 中取最新用户消息。"""
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            return m.content
    return ""


def clarify_node(state: AgentState) -> dict:
    intent = state.get("intent", {})
    msg = intent.get("clarification_message", "请提供更多信息以便我回答您的问题。")
    return {"final_answer": msg, "messages": [AIMessage(content=msg)]}


def sql_node(state: AgentState) -> dict:
    user_msg = _get_user_msg(state)
    result = _execute_sql(user_msg)
    return {
        "sql_result": result,
        "messages": [AIMessage(content=f"[SQL 查询完成]\n{result}")],
    }


def rag_node(state: AgentState) -> dict:
    user_msg = _get_user_msg(state)
    result = _execute_rag(user_msg)
    try:
        content = json.loads(result).get("content", result)
    except (json.JSONDecodeError, TypeError):
        content = result
    return {
        "rag_result": result,
        "messages": [AIMessage(content=f"[RAG 检索完成] {content[:200]}...")],
    }


def sql_branch_node(state: AgentState) -> dict:
    return sql_node(state)


def rag_branch_node(state: AgentState) -> dict:
    return rag_node(state)


def visualizer_node(state: AgentState) -> dict:
    user_msg = _get_user_msg(state)
    sql_result = state.get("sql_result", "")
    plot_path = _execute_viz(user_msg, sql_result)
    msg = f"图表已保存: {plot_path}" if plot_path else "无法生成图表。"
    return {"plot_path": plot_path, "messages": [AIMessage(content=msg)]}

# ── PLACEHOLDER_EXECUTOR ──


# ═══════════════════════════════════════════════════════
# Executor 节点（多意图拆分路径）
# ═══════════════════════════════════════════════════════

def executor_node(state: AgentState) -> dict:
    """按依赖顺序执行所有子任务，聚合结果到 state 字段。"""
    sub_tasks = [dict(t) for t in state["sub_tasks"]]  # 深拷贝
    results_map: dict[str, str] = {}

    for task in sub_tasks:
        tid = task["id"]
        tool = task["tool"]
        query = task["query"]
        dep = task.get("depends_on", "")
        dep_result = results_map.get(dep, "") if dep else ""

        try:
            if tool == "sql":
                task["result"] = _execute_sql(query)
            elif tool == "rag":
                task["result"] = _execute_rag(query)
            elif tool == "visualize":
                task["result"] = _execute_viz(query, dep_result)
            else:
                task["result"] = f"未知工具类型: {tool}"
            task["status"] = "done"
        except Exception as e:
            task["result"] = f"执行失败: {e}"
            task["status"] = "failed"

        results_map[tid] = task["result"]

    # 聚合到 state 字段供 synthesis 使用
    sql_parts = [t["result"] for t in sub_tasks if t["tool"] == "sql" and t["result"]]
    rag_parts = [t["result"] for t in sub_tasks if t["tool"] == "rag" and t["result"]]
    viz_paths = [t["result"] for t in sub_tasks if t["tool"] == "visualize" and t["result"]]

    # 合并多个 RAG 结果
    merged_rag = ""
    if rag_parts:
        merged_rag = _merge_rag_results(rag_parts)

    return {
        "sub_tasks": sub_tasks,
        "sql_result": "\n---\n".join(sql_parts),
        "rag_result": merged_rag,
        "plot_path": ",".join(viz_paths),
        "messages": [AIMessage(content=f"[Executor] 完成 {len(sub_tasks)} 个子任务")],
    }


def _merge_rag_results(rag_jsons: list[str]) -> str:
    """合并多个 RAG JSON 结果为一个。"""
    all_content = []
    all_refs = []
    for rj in rag_jsons:
        try:
            obj = json.loads(rj)
            all_content.append(obj.get("content", ""))
            all_refs.extend(obj.get("references", []))
        except (json.JSONDecodeError, TypeError):
            all_content.append(rj)
    return json.dumps({
        "content": "\n\n".join(all_content),
        "references": all_refs,
    }, ensure_ascii=False)


# ═══════════════════════════════════════════════════════
# Synthesis 节点
# ═══════════════════════════════════════════════════════

def synthesis_node(state: AgentState) -> dict:
    """汇总所有证据，生成最终回答。支持单一意图和多意图两种模式。"""
    from agent.conditions import sql_succeeded

    sub_tasks = state.get("sub_tasks", [])

    # ── 多意图模式：按子任务组织证据 ──
    if sub_tasks:
        evidence_parts = []
        for i, task in enumerate(sub_tasks, 1):
            tool_label = {"sql": "数据查询", "rag": "研报检索", "visualize": "图表生成"}.get(task["tool"], task["tool"])
            result_text = task.get("result", "")
            # RAG 结果提取 content
            if task["tool"] == "rag":
                try:
                    result_text = json.loads(result_text).get("content", result_text)
                except (json.JSONDecodeError, TypeError):
                    pass
            evidence_parts.append(f"【子问题{i}】{task['query']}\n【工具】{tool_label}\n【结果】\n{result_text}")

        evidence = "\n\n---\n\n".join(evidence_parts)
        user_msg = _get_user_msg(state)

        sys_prompt = (
            "你是一个专业的财务分析师。用户提出了一个包含多个子问题的复合问题。\n"
            "以下是各子问题的查询结果，请综合所有结果，给出完整、连贯的回答。\n"
            "用简洁准确的中文回答。如果某个子问题的证据不足，请诚实说明。不要编造数据。\n"
            "金额单位为万元，百分比单位为%，每股单位为元/股。\n\n"
            f"【证据】\n{evidence}"
        )

        llm = _get_synthesis_llm()
        resp = llm.invoke([
            SystemMessage(content=sys_prompt),
            {"role": "user", "content": user_msg},
        ])
        return {"final_answer": resp.content, "messages": [AIMessage(content=resp.content)]}

    # ── 单一意图模式（原有逻辑） ──
    sql_result = state.get("sql_result", "")
    rag_result = state.get("rag_result", "")
    plot_path = state.get("plot_path", "")

    has_valid_sql = sql_succeeded(state)
    has_rag = bool(rag_result) and "stub" not in rag_result and "检索失败" not in rag_result

    if sql_result and not has_valid_sql and not has_rag:
        fail_msg = (
            "抱歉，无法回答该问题。\n"
            "原因：数据库中不包含该问题所需的数据字段。"
            "当前数据库仅覆盖核心业绩指标表、利润表、资产负债表、现金流量表的结构化数据。"
        )
        return {"final_answer": fail_msg, "messages": [AIMessage(content=fail_msg)]}

    evidence_parts = []
    if sql_result and has_valid_sql:
        evidence_parts.append(f"【结构化查询结果】\n{sql_result}")
    if has_rag:
        rag_text = rag_result
        try:
            rag_text = json.loads(rag_result).get("content", rag_result)
        except (json.JSONDecodeError, TypeError):
            pass
        evidence_parts.append(f"【研报检索结果】\n{rag_text}")
    if plot_path:
        evidence_parts.append(f"【图表路径】{plot_path}")

    evidence = "\n\n".join(evidence_parts) if evidence_parts else "无可用证据。"
    user_msg = _get_user_msg(state)

    sys_prompt = (
        "你是一个专业的财务分析师。请**仅根据**以下收集到的证据回答用户问题，不要参考对话历史中的其他问题。\n"
        "用简洁准确的中文回答。如果证据不足以回答，请诚实说明。不要编造数据。\n\n"
        "【重要规则】\n"
        "- 如果证据中包含 SQL 查询结果的数据表格，你必须用具体数字回答，不要说'请参考图表'。\n"
        "- 如果有图表路径，在回答末尾提及'详见图表'即可，但主体回答必须包含具体数据。\n"
        "- 金额类字段单位均为**万元**。\n"
        "- 百分比类字段单位为 **%**。\n"
        "- 每股类字段单位为 **元/股**。\n"
        "回答时必须带上正确的单位。\n\n"
        f"【证据】\n{evidence}"
    )

    llm = _get_synthesis_llm()
    resp = llm.invoke([
        SystemMessage(content=sys_prompt),
        {"role": "user", "content": user_msg},
    ])
    return {"final_answer": resp.content, "messages": [AIMessage(content=resp.content)]}


# ═══════════════════════════════════════════════════════
# Formatter 节点
# ═══════════════════════════════════════════════════════

def formatter_node(state: AgentState) -> dict:
    """将本轮结果格式化为标准 JSON。"""
    from tool.json_formatter import format_turn

    user_msg = _get_user_msg(state)
    turn = format_turn(
        question=user_msg,
        content=state.get("final_answer", ""),
        plot_path=state.get("plot_path", ""),
        rag_result=state.get("rag_result", ""),
    )
    return {"formatted_answer": json.dumps(turn, ensure_ascii=False)}


# ═══════════════════════════════════════════════════════
# 构建 StateGraph
# ═══════════════════════════════════════════════════════

def build_graph():
    g = StateGraph(AgentState)

    g.add_node("planner", planner_node)
    g.add_node("clarify", clarify_node)
    g.add_node("sql", sql_node)
    g.add_node("rag", rag_node)
    g.add_node("sql_branch", sql_branch_node)
    g.add_node("rag_branch", rag_branch_node)
    g.add_node("visualizer", visualizer_node)
    g.add_node("executor", executor_node)
    g.add_node("synthesis", synthesis_node)
    g.add_node("formatter", formatter_node)

    g.set_entry_point("planner")

    # Planner 之后的主路由
    g.add_conditional_edges(
        "planner",
        route_after_planner,
        {
            "clarify": "clarify",
            "sql": "sql",
            "rag": "rag",
            "sql_and_rag": "sql_branch",
            "executor": "executor",
        },
    )

    # clarify → formatter → END
    g.add_edge("clarify", "formatter")

    # sql → 是否需要可视化
    g.add_conditional_edges("sql", route_after_sql, {
        "visualize": "visualizer",
        "synthesis": "synthesis",
    })

    # rag → synthesis
    g.add_edge("rag", "synthesis")

    # sql_and_rag 并行分支（串行模拟）
    g.add_edge("sql_branch", "rag_branch")
    g.add_edge("rag_branch", "synthesis")

    # visualizer → synthesis
    g.add_edge("visualizer", "synthesis")

    # executor → synthesis
    g.add_edge("executor", "synthesis")

    # synthesis → formatter → END
    g.add_edge("synthesis", "formatter")
    g.add_edge("formatter", END)

    memory = MemorySaver()
    return g.compile(checkpointer=memory)
