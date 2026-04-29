# -*- coding: utf-8 -*-
"""LangGraph StateGraph — 统一任务列表架构。

流程：Planner → Executor → Synthesis → Formatter → END
"""
from __future__ import annotations

import json
import re
from html.parser import HTMLParser

from langchain_core.messages import AIMessage, SystemMessage

from langgraph.checkpoint.memory import MemorySaver
from langgraph.graph import END, StateGraph

from agent.planner import planner_node
from agent.state import AgentState
from tool.sql_query import (
    SQL_SYSTEM_PROMPT,
    sql_query_tool,
    _clean_sql,
)


# ═══════════════════════════════════════════════════════
# HTML 表格 → 纯文本
# ═══════════════════════════════════════════════════════

def _html_table_to_text(html: str) -> str:
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


# ═══════════════════════════════════════════════════════
# LLM 懒加载
# ═══════════════════════════════════════════════════════

_sql_llm = None
_rag_llm = None
_viz_llm = None
_synthesis_llm = None
_trace_llm = None


def _get_sql_llm():
    global _sql_llm
    if _sql_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _sql_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0.3, streaming=False, max_retries=3,
        )
    return _sql_llm


def _get_trace_llm():
    global _trace_llm
    if _trace_llm is None:
        from langchain_openai import ChatOpenAI
        from agent.config import ZHIPU_API_KEY
        _trace_llm = ChatOpenAI(
            model="glm-4-flash", api_key=ZHIPU_API_KEY,
            base_url="https://open.bigmodel.cn/api/paas/v4/",
            temperature=0, streaming=False, max_retries=3,
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
            temperature=0.1, streaming=False, max_retries=3,
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
            temperature=0.1, streaming=False, max_retries=3,
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
            temperature=0.2, streaming=True, max_retries=3,
        )
    return _synthesis_llm


# ═══════════════════════════════════════════════════════
# 列名→表名映射（用于 SQL 重试时精准提示）
# ═══════════════════════════════════════════════════════

_COLUMN_TABLE_MAP: dict[str, list[str]] = {}


def _build_column_map():
    if _COLUMN_TABLE_MAP:
        return
    from tool.sql_query import DB_SCHEMA
    current_table = ""
    for line in DB_SCHEMA.splitlines():
        line = line.strip()
        if line.startswith("### "):
            m = re.match(r"###\s+\d+\.\s+(\w+)", line)
            if m:
                current_table = m.group(1)
        elif current_table and "（" in line:
            for part in line.split(","):
                col = part.strip().split("（")[0].strip()
                if col and not col.startswith("#"):
                    _COLUMN_TABLE_MAP.setdefault(col, [])
                    if current_table not in _COLUMN_TABLE_MAP[col]:
                        _COLUMN_TABLE_MAP[col].append(current_table)


def _get_column_hint(error_msg: str) -> str:
    _build_column_map()
    m = re.search(r"no such column:\s*(?:\w+\.)?(\w+)", error_msg)
    if not m:
        return ""
    col = m.group(1)
    tables = _COLUMN_TABLE_MAP.get(col, [])
    if tables:
        return f"字段 {col} 不在你查询的表中，它只存在于 {'/'.join(tables)} 表，请用 JOIN 关联后再引用该字段。"
    return ""


# ═══════════════════════════════════════════════════════
# 可复用工具函数
# ═══════════════════════════════════════════════════════

def _clean_query_for_sql(query: str) -> str:
    clean = re.sub(r"[，,]?\s*(做|进行|生成|给出|画出|绘制)?\s*(可视化|绘图|画图|图表|画个图|做个图|折线图|柱状图|饼图|趋势图|雷达图|直方图|双条形图|散点图|箱线图|表格图)[。？]?", "", query).strip()
    return clean or query


def _try_direct_sql(query: str) -> str:
    """Phase 1: 直接生成 SQL 并执行，含重试。返回结果或错误。"""
    clean_query = _clean_query_for_sql(query)
    llm = _get_sql_llm()
    messages = [
        SystemMessage(content=SQL_SYSTEM_PROMPT),
        {"role": "user", "content": clean_query},
    ]

    last_sql = ""
    last_result = ""
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            resp = llm.invoke(messages)
            sql = _clean_sql(resp.content.strip())
        except Exception as e:
            return f"SQL 生成失败: {e}"

        if not sql.upper().startswith("SELECT"):
            if attempt < max_attempts - 1:
                messages.append({"role": "assistant", "content": resp.content})
                messages.append({"role": "user", "content": "请只返回一条SQL SELECT语句，不要包含任何解释。"})
                continue
            return f"无法生成有效SQL。模型返回: {sql[:200]}"

        last_sql = sql
        result = sql_query_tool.invoke({"sql": sql})
        last_result = result

        if "SQL 执行错误" in result and attempt < max_attempts - 1:
            hint = _get_column_hint(result)
            retry_msg = f"上面的SQL执行报错：{result}\n"
            if hint:
                retry_msg += f"【提示】{hint}\n"
            retry_msg += "请修正SQL语句，只返回修正后的SQL。"
            messages.append({"role": "assistant", "content": sql})
            messages.append({"role": "user", "content": retry_msg})
            continue

        if result == "查询无结果。" and attempt < max_attempts - 1:
            retry_msg = (
                f"上面的SQL查询无结果。可能原因：\n"
                f"1. stock_code 错误（请改用 stock_abbr 匹配）\n"
                f"2. 2022年只有FY数据，没有Q1/HY/Q3\n"
                f"3. 条件过于严格（如HAVING COUNT(*)=4但实际最多3年）\n"
                f"请检查并修正SQL，只返回修正后的SQL。"
            )
            messages.append({"role": "assistant", "content": sql})
            messages.append({"role": "user", "content": retry_msg})
            continue

        return f"SQL: {sql}\n\n结果:\n{result}"

    return f"SQL: {last_sql}\n\n结果:\n{last_result}"


# ── SQL 分解执行（Phase 2 fallback） ──

_DECOMPOSE_PROMPT = """你是 SQL 分解助手。用户的查询需要跨表数据，直接生成的 SQL 执行失败了。
请将问题拆成多条单表查询，每条只查一张表的字段。

## 数据库表结构：
{schema}

## 核心原则：
每个字段只存在于它所属的表中！拆分时必须确保每条 SQL 只使用该表拥有的字段。

## 示例：
问题：经营性现金流量净额与净利润的比值
→ 拆分为：
1. 从 cash_flow_sheet 查 operating_cf_net_amount（经营性现金流量净额只在现金流量表）
2. 从 income_sheet 查 net_profit（净利润只在利润表）

问题：销售费用率低于均值的公司的ROE
→ 拆分为：
1. 从 income_sheet 查 operating_expense_selling_expenses 和 total_operating_revenue
2. 从 core_performance_indicators_sheet 查 roe

## 规则：
1. 每条 SQL 只查一张表，绝对不能用 JOIN，也不能引用其他表的字段
2. 每条 SQL 必须 SELECT stock_code, stock_abbr 以便后续关联
3. 每条 SQL 必须包含 report_year 和 report_period 过滤条件
4. 只返回 JSON 数组，不要解释：
[{{"sql": "SELECT ...", "purpose": "查询说明"}}]

## 失败的原始 SQL 和错误信息：
{error_context}
"""

_SYNTHESIZE_PROMPT = """你是数据分析助手。用户问题被拆成了多条单表查询，以下是各查询结果。
请根据这些数据回答用户问题，进行必要的计算（比率、排名、对比、筛选等）。

## 规则：
1. 用 stock_code 关联不同表的数据（同一公司在不同表中 stock_code 相同）
2. 金额单位万元，百分比单位%，每股单位元/股
3. 只用提供的数据，不要编造
4. 必须输出一个合并后的结果表格（用 | 分隔列），然后再给出文字总结
5. 表格第一行是表头，后续每行一条数据
6. 如果需要计算复合增长率(CAGR)：CAGR = (终值/初值)^(1/年数) - 1，用百分比表示

## 查询结果：
{results}
"""


def _decompose_and_execute(query: str, error_context: str) -> str:
    """Phase 2: 将跨表查询拆成多条单表 SQL，逐条执行后 LLM 综合。"""
    from tool.sql_query import DB_SCHEMA

    llm = _get_sql_llm()

    prompt = _DECOMPOSE_PROMPT.format(schema=DB_SCHEMA, error_context=error_context)
    try:
        resp = llm.invoke([
            SystemMessage(content=prompt),
            {"role": "user", "content": _clean_query_for_sql(query)},
        ])
    except Exception as e:
        return f"SQL 分解失败: {e}"

    m = re.search(r"\[[\s\S]+\]", resp.content)
    if not m:
        return error_context

    try:
        sub_queries = json.loads(m.group())
    except (json.JSONDecodeError, TypeError):
        return error_context

    if not sub_queries or not isinstance(sub_queries, list):
        return error_context

    sql_parts = []
    result_parts = []
    all_ok = True
    for i, sq in enumerate(sub_queries, 1):
        sql = _clean_sql(str(sq.get("sql", "")))
        purpose = sq.get("purpose", "")
        if not sql.upper().startswith("SELECT"):
            continue
        res = sql_query_tool.invoke({"sql": sql})
        if "SQL 执行错误" in res:
            hint = _get_column_hint(res)
            if hint:
                retry_msg = f"SQL执行报错：{res}\n【提示】{hint}\n请修正SQL，只返回修正后的SQL。"
                try:
                    fix_resp = llm.invoke([
                        SystemMessage(content=SQL_SYSTEM_PROMPT),
                        {"role": "user", "content": _clean_query_for_sql(query)},
                        {"role": "assistant", "content": sql},
                        {"role": "user", "content": retry_msg},
                    ])
                    fixed_sql = _clean_sql(fix_resp.content.strip())
                    if fixed_sql.upper().startswith("SELECT"):
                        res2 = sql_query_tool.invoke({"sql": fixed_sql})
                        if "SQL 执行错误" not in res2:
                            sql = fixed_sql
                            res = res2
                except Exception:
                    pass
        sql_parts.append(f"-- 子查询{i}: {purpose}\n{sql}")
        if "SQL 执行错误" in res:
            all_ok = False
            result_parts.append(f"【子查询{i}】{purpose}\n执行失败: {res}")
        else:
            result_parts.append(f"【子查询{i}】{purpose}\n{res}")

    if not result_parts:
        return error_context

    results_text = "\n\n".join(result_parts)

    if not all_ok and all("执行失败" in r for r in result_parts):
        return error_context

    synth_prompt = _SYNTHESIZE_PROMPT.format(results=results_text)
    try:
        synth_resp = llm.invoke([
            SystemMessage(content=synth_prompt),
            {"role": "user", "content": _clean_query_for_sql(query)},
        ])
        answer = synth_resp.content.strip()
    except Exception as e:
        answer = results_text

    all_sql = "\n".join(sql_parts)
    return f"SQL(分解执行):\n{all_sql}\n\n结果:\n{answer}"


def _execute_sql(query: str) -> str:
    """两阶段 SQL 执行：先直接生成，失败则分解为多条单表查询。"""
    result = _try_direct_sql(query)
    if "SQL 执行错误" not in result:
        return result
    return _decompose_and_execute(query, result)


_RAG_SYSTEM_PROMPT = """你是一个专业的医药/CXO行业研报分析师。请严格根据以下检索到的研报片段回答用户问题。

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
    from tool.rag_search import hybrid_search

    search_query = query
    if len(query) < 15:
        search_query = f"医药CXO行业 {query}"

    try:
        hits = hybrid_search(search_query, n_results=5)
    except Exception as e:
        return json.dumps({"content": f"研报检索失败: {e}", "references": []}, ensure_ascii=False)

    if not hits and search_query != query:
        try:
            hits = hybrid_search(query, n_results=5)
        except Exception:
            pass

    if not hits:
        return json.dumps({"content": "研报检索无结果。", "references": []}, ensure_ascii=False)

    _RERANK_THRESHOLD = 0.15
    filtered = [h for h in hits if h.get("score", 0) >= _RERANK_THRESHOLD]
    if not filtered:
        return json.dumps({"content": "根据现有研报资料，暂无与该问题高度相关的信息。", "references": []}, ensure_ascii=False)
    hits = filtered

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

    references = _trace_references(content, hits)
    return json.dumps({"content": content, "references": references}, ensure_ascii=False)


def _trace_references(answer: str, hits: list[dict]) -> list[dict]:
    if not hits or not answer:
        return []

    llm = _get_trace_llm()
    chunk_list = []
    for i, h in enumerate(hits, 1):
        title = h["metadata"].get("title", "")
        chunk_text = h["text"]
        if "<table>" in chunk_text:
            chunk_text = _html_table_to_text(chunk_text)
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

    indices = set()
    for part in re.findall(r"\d+", text):
        idx = int(part)
        if 1 <= idx <= len(hits):
            indices.add(idx)

    refs = []
    for idx in sorted(indices):
        h = hits[idx - 1]
        meta = h["metadata"]
        ref_text = h["text"]
        if "<table>" in ref_text:
            ref_text = _html_table_to_text(ref_text)
        ref_text = ref_text[:300]
        ref = {
            "paper_path": meta.get("source_pdf", "").replace("./research report/", "./附件5：研报数据/"),
            "text": ref_text,
        }
        table_title = meta.get("table_title", "")
        if table_title:
            ref["paper_image"] = table_title
        refs.append(ref)
    return refs


_VIZ_SYSTEM_PROMPT = """你是一个数据可视化助手。根据用户问题和 SQL 查询结果，决定最佳的图表类型并组织数据。

## 图表类型选择规则（必须严格遵守）：

### 默认图表（根据数据特征自动选择）：
- **line（折线图）**：用于时间序列、趋势分析、多年变化、增长率走势。关键词：趋势、变化、走势、近N年、历年。
- **bar（柱状图）**：用于不同实体之间的对比、排名、Top N。关键词：对比、排名、最高、最低、Top。
- **pie（饼图）**：用于占比、构成、结构分析。关键词：占比、构成、结构、比例、分布。

### 特殊图表（仅当用户明确要求时才使用，不要自行推断）：
- **radar（雷达图）**：用户明确说"雷达图"时才用。需要至少3个维度。
- **histogram（直方图）**：用户明确说"直方图"时才用。展示数值分布频率。
- **double_bar（双条形图）**：用户明确说"双条形图"或"水平双条形图"时才用。水平方向，恰好2组数据对比。
- **scatter（散点图）**：用户明确说"散点图"时才用。展示两个变量之间的关系。
- **table（表格）**：用户明确说"表格"或"生成表格"时才用。以表格形式展示数据。
- **boxplot（箱线图）**：用户明确说"箱线图"时才用。展示数据分布的四分位数。

## 数据单位说明：
- 数据库中金额字段单位为万元，不做换算，直接使用原值。
- 百分比类字段单位为 %。
- y_label 应标注正确单位：金额写"万元"，百分比写"%"。
- title 中应包含单位信息。

## 输出格式（严格 JSON，不要输出其他任何内容）：

line/bar/pie 单系列：
```json
{"chart_type": "line", "title": "标题（万元）", "y_label": "万元", "data": {"labels": ["2020", "2021"], "values": [12.5, 14.3]}}
```

line/bar 多系列：
```json
{"chart_type": "bar", "title": "标题（万元）", "y_label": "万元", "data": {"labels": ["2020", "2021"], "datasets": [{"label": "系列1", "values": [100, 150]}, {"label": "系列2", "values": [20, 30]}]}}
```

radar（雷达图）：
```json
{"chart_type": "radar", "title": "标题", "y_label": "", "data": {"labels": ["维度1", "维度2", "维度3"], "values": [80, 60, 90]}}
```
多系列雷达图用 datasets，格式同多系列 bar。

histogram（直方图）：
```json
{"chart_type": "histogram", "title": "标题", "y_label": "频次", "data": {"values": [1.2, 3.4, 2.1, 5.6, 3.3], "bins": 10, "x_label": "万元"}}
```

double_bar（双条形图，水平方向）：
```json
{"chart_type": "double_bar", "title": "标题（万元）", "y_label": "万元", "data": {"labels": ["公司A", "公司B"], "datasets": [{"label": "指标1", "values": [100, 200]}, {"label": "指标2", "values": [80, 150]}]}}
```

scatter（散点图）：
```json
{"chart_type": "scatter", "title": "标题", "y_label": "Y轴", "data": {"x_values": [1, 2, 3], "y_values": [4, 5, 6], "labels": ["A", "B", "C"], "x_label": "X轴", "y_label": "Y轴"}}
```

table（表格）：
```json
{"chart_type": "table", "title": "标题", "y_label": "", "data": {"headers": ["公司", "营收", "净利润"], "rows": [["公司A", "100", "20"], ["公司B", "200", "30"]]}}
```

boxplot（箱线图）：
```json
{"chart_type": "boxplot", "title": "标题（万元）", "y_label": "万元", "data": {"labels": ["2022", "2023"], "datasets": [{"values": [10, 20, 30, 40]}, {"values": [15, 25, 35, 45]}]}}
```
"""


def _parse_sql_table(sql_result: str) -> tuple[list[str], list[list[str]]]:
    """从 SQL 结果文本中解析表头和数据行。"""
    lines = []
    in_result = False
    for line in sql_result.split("\n"):
        if line.startswith("结果:") or line.startswith("结果："):
            in_result = True
            continue
        if in_result and line.strip():
            lines.append(line)
    if not lines:
        for line in sql_result.split("\n"):
            if "\t" in line:
                lines.append(line)
    if not lines:
        return [], []
    headers = lines[0].split("\t")
    rows = [line.split("\t") for line in lines[1:] if line.strip()]
    return headers, rows


def _fix_truncated_json(raw: str) -> str:
    """尝试修复 LLM 输出的非标准 JSON。"""
    raw = raw.replace("None", "null").replace("NaN", "null")
    depth_sq = raw.count("[") - raw.count("]")
    depth_br = raw.count("{") - raw.count("}")
    fixed = raw + "]" * depth_sq + "}" * depth_br
    return fixed


def _parse_json_lenient(raw: str) -> dict | None:
    """宽容地解析 JSON，处理 None/NaN 和截断。"""
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        pass
    fixed = _fix_truncated_json(raw)
    try:
        return json.loads(fixed)
    except (json.JSONDecodeError, TypeError):
        pass
    try:
        decoder = json.JSONDecoder()
        obj, _ = decoder.raw_decode(fixed)
        return obj
    except (json.JSONDecodeError, TypeError):
        pass
    return None


def _execute_viz(query: str, sql_result: str) -> str:
    from tool.visualizer import data_visualizer_tool

    llm = _get_viz_llm()
    resp = llm.invoke([
        SystemMessage(content=_VIZ_SYSTEM_PROMPT),
        {"role": "user", "content": f"用户问题：{query}\n\nSQL查询结果：\n{sql_result}"},
    ])

    m = re.search(r"\{[\s\S]+\}", resp.content)
    if not m:
        return ""

    raw_json = m.group()
    parsed = _parse_json_lenient(raw_json)

    if not parsed:
        chart_type_m = re.search(r'"chart_type"\s*:\s*"(\w+)"', raw_json)
        title_m = re.search(r'"title"\s*:\s*"([^"]+)"', raw_json)
        chart_type = chart_type_m.group(1) if chart_type_m else "bar"
        title = title_m.group(1) if title_m else "图表"
        y_label = ""
        data = {}
    else:
        chart_type = parsed.get("chart_type", "bar")
        title = parsed.get("title", "图表")
        data = parsed.get("data", {})
        y_label = parsed.get("y_label", "")

    if chart_type in ("scatter", "histogram", "boxplot"):
        headers, rows = _parse_sql_table(sql_result)
        if headers and rows:
            if chart_type == "scatter" and len(headers) >= 4:
                x_col, y_col = 2, 3
                labels_col = 1
                x_vals, y_vals, labels = [], [], []
                for row in rows:
                    try:
                        xv = float(row[x_col])
                        yv = float(row[y_col])
                    except (ValueError, IndexError):
                        continue
                    x_vals.append(xv)
                    y_vals.append(yv)
                    labels.append(row[labels_col] if labels_col < len(row) else "")
                if x_vals:
                    data = {
                        "x_values": x_vals, "y_values": y_vals, "labels": labels,
                        "x_label": data.get("x_label", headers[x_col]),
                        "y_label": data.get("y_label", headers[y_col]),
                    }
            elif chart_type == "histogram":
                val_col = len(headers) - 1
                vals = []
                for row in rows:
                    try:
                        vals.append(float(row[val_col]))
                    except (ValueError, IndexError):
                        continue
                if vals:
                    data = {"values": vals, "bins": min(20, max(5, len(vals) // 3)),
                            "x_label": data.get("x_label", headers[val_col])}

    if not data:
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
# 辅助函数
# ═══════════════════════════════════════════════════════

def _get_user_msg(state: AgentState) -> str:
    for m in reversed(state["messages"]):
        if hasattr(m, "type") and m.type == "human":
            return m.content
    return ""


def _merge_rag_results(rag_jsons: list[str]) -> str:
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
# Executor 节点
# ═══════════════════════════════════════════════════════

def executor_node(state: AgentState) -> dict:
    """按依赖顺序执行所有子任务。clarify/chat 直接用 query 作为 result。"""
    sub_tasks = [dict(t) for t in state["sub_tasks"]]
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
                viz_input = dep_result
                if not viz_input or "\t" not in viz_input:
                    sql_results = [t["result"] for t in sub_tasks
                                   if t["tool"] == "sql" and t.get("result") and "\t" in t["result"]]
                    if sql_results:
                        viz_input = "\n---\n".join(sql_results)
                task["result"] = _execute_viz(query, viz_input)
            elif tool in ("clarify", "chat"):
                task["result"] = query
            else:
                task["result"] = f"未知工具类型: {tool}"
            task["status"] = "done"
        except Exception as e:
            task["result"] = f"执行失败: {e}"
            task["status"] = "failed"

        results_map[tid] = task["result"]

    sql_parts = [t["result"] for t in sub_tasks if t["tool"] == "sql" and t["result"]]
    rag_parts = [t["result"] for t in sub_tasks if t["tool"] == "rag" and t["result"]]
    viz_paths = [t["result"] for t in sub_tasks if t["tool"] == "visualize" and t["result"]]

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


# ═══════════════════════════════════════════════════════
# Synthesis 节点
# ═══════════════════════════════════════════════════════

_SYNTHESIS_PLACEHOLDER = "<!-- SYNTHESIS_NEXT -->"


def synthesis_node(state: AgentState) -> dict:
    """汇总所有证据，生成最终回答。"""
    sub_tasks = state.get("sub_tasks", [])
    tools_used = [t.get("tool", "") for t in sub_tasks]

    # clarify 或 chat 单任务：直接返回，不走 LLM
    if len(sub_tasks) == 1 and tools_used[0] in ("clarify", "chat"):
        msg = sub_tasks[0].get("result", sub_tasks[0].get("query", ""))
        return {"final_answer": msg, "messages": [AIMessage(content=msg)]}

    # 组织证据
    evidence_parts = []
    for i, task in enumerate(sub_tasks, 1):
        tool_label = {"sql": "数据查询", "rag": "研报检索", "visualize": "图表生成"}.get(task["tool"], task["tool"])
        result_text = task.get("result", "")
        if task["tool"] == "rag":
            try:
                result_text = json.loads(result_text).get("content", result_text)
            except (json.JSONDecodeError, TypeError):
                pass
        if task["tool"] in ("clarify", "chat"):
            continue
        evidence_parts.append(f"【子问题{i}】{task['query']}\n【工具】{tool_label}\n【结果】\n{result_text}")

    if not evidence_parts:
        fallback = "抱歉，无法回答该问题。"
        return {"final_answer": fallback, "messages": [AIMessage(content=fallback)]}

    evidence = "\n\n---\n\n".join(evidence_parts)
    user_msg = _get_user_msg(state)
    plot_path = state.get("plot_path", "")

    sys_prompt = (
        "你是一个专业的财务分析师。请**仅根据**以下收集到的证据回答用户问题。\n"
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
    if plot_path:
        sys_prompt += f"\n\n【图表路径】{plot_path}"

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
    g.add_node("executor", executor_node)
    g.add_node("synthesis", synthesis_node)
    g.add_node("formatter", formatter_node)

    g.set_entry_point("planner")
    g.add_edge("planner", "executor")
    g.add_edge("executor", "synthesis")
    g.add_edge("synthesis", "formatter")
    g.add_edge("formatter", END)

    memory = MemorySaver()
    return g.compile(checkpointer=memory)
