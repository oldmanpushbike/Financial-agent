# -*- coding: utf-8 -*-
"""SQL 生成 + 只读查询工具，参考 B黎_1/agent/tools/sql_generator.py 改写为 StructuredTool."""
from __future__ import annotations

import re
import sqlite3
from typing import Optional

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

from agent.config import DB_PATH

# ──────────────────────────────────────────────
# DB Schema 描述（供 LLM 生成 SQL 时参考）
# ──────────────────────────────────────────────
DB_SCHEMA = """
数据库为66家中药上市公司的财务数据，覆盖2022-2025年，共4张表。
公共字段：stock_code（证券代码TEXT）、stock_abbr（股票简称TEXT）、report_year（年份INT）、report_period（报告期TEXT：'Q1'一季度、'HY'半年度、'Q3'前三季度、'FY'年度）。
所有金额字段单位为万元，百分比字段单位为%，每股字段单位为元/股。

### 1. core_performance_indicators_sheet（核心业绩指标表，583行）
eps（每股收益）, total_operating_revenue（营业总收入）, operating_revenue_yoy_growth（营收同比增长率%）, operating_revenue_qoq_growth（营收环比增长率%）, net_profit_10k_yuan（净利润万元）, net_profit_yoy_growth（净利润同比增长率%）, net_profit_qoq_growth（净利润环比增长率%）, net_asset_per_share（每股净资产）, roe（净资产收益率%）, operating_cf_per_share（每股经营现金流）, net_profit_excl_non_recurring（扣非净利润万元）, net_profit_excl_non_recurring_yoy（扣非净利润同比%）, gross_profit_margin（销售毛利率%）, net_profit_margin（销售净利率%）, roe_weighted_excl_non_recurring（加权ROE扣非%）

### 2. income_sheet（利润表，564行）
net_profit（净利润）, net_profit_yoy_growth（净利润同比%）, other_income（其他收益）, total_operating_revenue（营业总收入/营业收入）, operating_revenue_yoy_growth（营收同比%）, operating_expense_cost_of_sales（营业成本）, operating_expense_selling_expenses（销售费用）, operating_expense_administrative_expenses（管理费用）, operating_expense_financial_expenses（财务费用）, operating_expense_rnd_expenses（研发费用）, operating_expense_taxes_and_surcharges（税金及附加）, total_operating_expenses（营业总支出）, operating_profit（营业利润）, total_profit（利润总额）, asset_impairment_loss（资产减值损失）, credit_impairment_loss（信用减值损失）

### 3. balance_sheet（资产负债表，567行）
asset_cash_and_cash_equivalents（货币资金）, asset_accounts_receivable（应收账款）, asset_inventory（存货）, asset_trading_financial_assets（交易性金融资产）, asset_construction_in_progress（在建工程）, asset_total_assets（资产总额）, asset_total_assets_yoy_growth（总资产同比%）, liability_accounts_payable（应付账款）, liability_advance_from_customers（预收账款）, liability_total_liabilities（负债总额）, liability_total_liabilities_yoy_growth（负债同比%）, liability_contract_liabilities（合同负债）, liability_short_term_loans（短期借款）, asset_liability_ratio（资产负债率%）, equity_unappropriated_profit（未分配利润）, equity_total_equity（股东权益总额）

### 4. cash_flow_sheet（现金流量表，583行）
net_cash_flow（现金净流量）, net_cash_flow_yoy_growth（现金流同比%）, operating_cf_net_amount（经营性现金流量净额）, operating_cf_ratio_of_net_cf（经营现金流占比%）, operating_cf_cash_from_sales（销售商品收到的现金）, investing_cf_net_amount（投资性现金流量净额）, investing_cf_ratio_of_net_cf（投资现金流占比%）, investing_cf_cash_for_investments（投资支付的现金）, investing_cf_cash_from_investment_recovery（收回投资收到的现金）, financing_cf_cash_from_borrowing（取得借款收到的现金）, financing_cf_cash_for_debt_repayment（偿还债务支付的现金）, financing_cf_net_amount（筹资性现金流量净额）, financing_cf_ratio_of_net_cf（筹资现金流占比%）
""".strip()

FEW_SHOT_EXAMPLES = """
## 基础查询
User: 金花股份2022年年度的总资产是多少？
SQL: SELECT stock_abbr, asset_total_assets FROM balance_sheet WHERE stock_abbr = '金花股份' AND report_year = 2022 AND report_period = 'FY';

User: 2025年第三季度所有公司的股票代码、简称和营业总收入，按营业总收入从高到低排序。
SQL: SELECT stock_code, stock_abbr, total_operating_revenue FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3' ORDER BY total_operating_revenue DESC;

## 条件筛选（亏损=净利润<0）
User: 哪些企业是亏钱的？（指净利润为负）
SQL: SELECT stock_code, stock_abbr, net_profit FROM income_sheet WHERE net_profit < 0 AND report_year = 2025 AND report_period = 'Q3';

User: 2025年前三季度营业总收入超过200亿元（即2000000万元）的企业有哪些？
SQL: SELECT stock_code, stock_abbr, total_operating_revenue FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3' AND total_operating_revenue > 2000000;

## 排名 Top N
User: 2025年第三季度营业总收入排名前五的中药公司有哪些？
SQL: SELECT stock_code, stock_abbr, total_operating_revenue FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3' ORDER BY total_operating_revenue DESC LIMIT 5;

User: 2025年第三季度"股东权益-未分配利润"金额排名前五的公司，其2025年的净利润占未分配利润的比例。
SQL: SELECT b.stock_code, b.stock_abbr, b.equity_unappropriated_profit, i.net_profit, ROUND(i.net_profit * 100.0 / b.equity_unappropriated_profit, 2) AS ratio_pct FROM balance_sheet b JOIN income_sheet i ON b.stock_code = i.stock_code AND b.report_year = i.report_year AND b.report_period = i.report_period WHERE b.report_year = 2025 AND b.report_period = 'Q3' ORDER BY b.equity_unappropriated_profit DESC LIMIT 5;

## 多表关联
User: 2025年第三季度经营性现金流量净额大于净利润的公司。
SQL: SELECT c.stock_code, c.stock_abbr, i.net_profit, c.operating_cf_net_amount FROM cash_flow_sheet c JOIN income_sheet i ON c.stock_code = i.stock_code AND c.report_year = i.report_year AND c.report_period = i.report_period WHERE c.report_year = 2025 AND c.report_period = 'Q3' AND c.operating_cf_net_amount > i.net_profit;

User: 2025年第三季度短期借款超过货币资金的公司。
SQL: SELECT stock_code, stock_abbr, liability_short_term_loans, asset_cash_and_cash_equivalents FROM balance_sheet WHERE report_year = 2025 AND report_period = 'Q3' AND liability_short_term_loans > asset_cash_and_cash_equivalents;

## 行业均值与对比
User: 2025年第三季度资产负债率超过60%的公司。
SQL: SELECT stock_code, stock_abbr, ROUND(asset_liability_ratio, 2) AS asset_liability_ratio FROM balance_sheet WHERE report_year = 2025 AND report_period = 'Q3' AND asset_liability_ratio > 60;

User: 2025年第三季度销售毛利率和销售净利率均高于行业均值的公司。
SQL: SELECT stock_code, stock_abbr, gross_profit_margin, net_profit_margin FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3' AND gross_profit_margin > (SELECT AVG(gross_profit_margin) FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3') AND net_profit_margin > (SELECT AVG(net_profit_margin) FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3');

## 计算字段（毛利率、占比等）
User: 2025年第三季度研发费用占营业总收入比例排名前五的公司。
SQL: SELECT i.stock_code, i.stock_abbr, i.operating_expense_rnd_expenses, i.total_operating_revenue, ROUND(i.operating_expense_rnd_expenses * 100.0 / i.total_operating_revenue, 2) AS rnd_ratio_pct FROM income_sheet i WHERE i.report_year = 2025 AND i.report_period = 'Q3' AND i.total_operating_revenue > 0 ORDER BY rnd_ratio_pct DESC LIMIT 5;

## 跨期对比与趋势
User: 对比白云山2022年至2025年第三季度的营业总收入同比增长率。
SQL: SELECT report_year, report_period, operating_revenue_yoy_growth FROM core_performance_indicators_sheet WHERE stock_abbr = '白云山' AND report_period = 'Q3' AND report_year BETWEEN 2022 AND 2025 ORDER BY report_year;

User: 金花股份近几年的利润总额变化趋势。
SQL: SELECT report_year, total_profit FROM income_sheet WHERE stock_abbr = '金花股份' AND report_period = 'FY' ORDER BY report_year;

User: 2022-2025年第三季度加权ROE扣非连续四期均超过10%的公司。
SQL: SELECT stock_code, stock_abbr, GROUP_CONCAT(report_year || ':' || roe_weighted_excl_non_recurring) AS roe_by_year FROM core_performance_indicators_sheet WHERE report_period = 'Q3' AND report_year BETWEEN 2022 AND 2025 GROUP BY stock_code, stock_abbr HAVING MIN(roe_weighted_excl_non_recurring) > 10;

## 复合增长率（CAGR）
User: 2022年至2025年第三季度66家公司营业总收入复合增长率排名前三。
SQL: SELECT stock_code, stock_abbr, ROUND((POWER(t2.rev / t1.rev, 1.0/3) - 1) * 100, 2) AS cagr_pct FROM (SELECT stock_code, stock_abbr, total_operating_revenue AS rev FROM core_performance_indicators_sheet WHERE report_year = 2022 AND report_period = 'Q3') t1 JOIN (SELECT stock_code, total_operating_revenue AS rev FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3') t2 ON t1.stock_code = t2.stock_code WHERE t1.rev > 0 ORDER BY cagr_pct DESC LIMIT 3;

## 营业总收入与净利润均排名前N的交集
User: 2025年第三季度营业总收入和净利润均排名前五的公司。
SQL: SELECT a.stock_code, a.stock_abbr, a.total_operating_revenue, b.net_profit FROM (SELECT stock_code, stock_abbr, total_operating_revenue FROM core_performance_indicators_sheet WHERE report_year = 2025 AND report_period = 'Q3' ORDER BY total_operating_revenue DESC LIMIT 5) a INNER JOIN (SELECT stock_code, net_profit FROM income_sheet WHERE report_year = 2025 AND report_period = 'Q3' ORDER BY net_profit DESC LIMIT 5) b ON a.stock_code = b.stock_code;
""".strip()

# ──────────────────────────────────────────────
# SQL 生成系统提示词
# ──────────────────────────────────────────────
SQL_SYSTEM_PROMPT = f"""你是一个SQL查询生成助手，专门为66家中药上市公司财报"智能问数"系统生成SQL语句。数据库为SQLite。

## 数据库表结构：
{DB_SCHEMA}

## 核心规则：
1. 只返回一条SQL SELECT语句，以分号结尾，不要包含任何解释或markdown代码块。
2. 忽略用户消息中关于"可视化""绘图""画图""图表"等请求，只关注数据查询部分。
3. 仅允许访问上述四张表。
3. 报告期映射："第一季度/一季度"→'Q1'，"上半年/半年度"→'HY'，"第三季度/前三季度"→'Q3'，"年度/全年/年报"→'FY'。report_year为四位数字。
4. 当问题给出证券代码（6位数字）时优先用stock_code过滤，否则用stock_abbr模糊匹配（LIKE '%关键词%'）。
5. "亏钱/亏损"→net_profit < 0；"盈利"→net_profit > 0。
6. "营业总收入超过N亿元"→total_operating_revenue > N*10000（万元）；"超过N万元"直接比较。
7. 需要计算比率时用ROUND(..., 2)保留两位小数，注意分母不为0（加WHERE或NULLIF）。
8. 多表关联时用 JOIN ON stock_code AND report_year AND report_period。
9. 行业均值用子查询 (SELECT AVG(...) FROM ... WHERE 同期条件)。
10. 排名/Top N 用 ORDER BY ... DESC/ASC LIMIT N。
11. "复合增长率"用 POWER(终值/初值, 1.0/年数) - 1。
12. 如问题要求多个字段或多行，一次性查询齐全。
13. "销售毛利率"如果核心指标表中有gross_profit_margin字段可直接用；否则用利润表计算：(total_operating_revenue - operating_expense_cost_of_sales) / total_operating_revenue * 100。
14. stock_abbr匹配时注意简称可能不完全一致，如"三金"应匹配"三金药业"，"999"应匹配"华润三九"，用LIKE。

## Few-shot 示例：
{FEW_SHOT_EXAMPLES}
"""

# ──────────────────────────────────────────────
# 只读安全校验
# ──────────────────────────────────────────────
_WRITE_PATTERN = re.compile(
    r"\b(INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|REPLACE|TRUNCATE|ATTACH|DETACH|PRAGMA)\b",
    re.IGNORECASE,
)

MAX_ROWS = 100


def _validate_readonly(sql: str) -> Optional[str]:
    """返回 None 表示安全，否则返回拒绝原因。"""
    if _WRITE_PATTERN.search(sql):
        return f"拒绝执行写操作：检测到非 SELECT 关键词。SQL: {sql}"
    return None


def _clean_sql(sql_text: str) -> str:
    if sql_text.startswith("```"):
        lines = sql_text.split("\n")
        if lines[0].strip().startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        sql_text = "\n".join(lines).strip()
    return " ".join(sql_text.split())


# ──────────────────────────────────────────────
# 启发式 SQL 回退（参考 sql_generator.py）
# ──────────────────────────────────────────────
_METRIC_MAP = {
    "总资产": ("balance_sheet", "asset_total_assets"),
    "资产总计": ("balance_sheet", "asset_total_assets"),
    "营业总收入": ("income_sheet", "total_operating_revenue"),
    "营业收入": ("income_sheet", "total_operating_revenue"),
    "营收": ("income_sheet", "total_operating_revenue"),
    "净利润": ("income_sheet", "net_profit"),
    "利润总额": ("income_sheet", "total_profit"),
    "资产负债率": ("balance_sheet", "asset_liability_ratio"),
    "总负债": ("balance_sheet", "liability_total_liabilities"),
    "负债合计": ("balance_sheet", "liability_total_liabilities"),
    "现金流": ("cash_flow_sheet", "net_cash_flow"),
    "净现金流": ("cash_flow_sheet", "net_cash_flow"),
}


def _heuristic_sql(q: str) -> str:
    q = (q or "").strip()
    if not q:
        return ""

    year_m = re.search(r"(20\d{2})", q)
    year = int(year_m.group(1)) if year_m else None

    period = None
    for keys, val in [
        (["一季", "第一季度", "Q1"], "Q1"),
        (["半年", "半年度", "上半年", "HY"], "HY"),
        (["三季", "第三季度", "前三季度", "Q3"], "Q3"),
        (["年度", "年报", "全年", "FY"], "FY"),
    ]:
        if any(k in q for k in keys):
            period = val
            break

    if year is None:
        return ""
    if period is None:
        period = "FY"

    code_m = re.search(r"\b(\d{6})\b", q)
    stock_code = code_m.group(1) if code_m else None

    stock_abbr = None
    if stock_code is None:
        abbr_m = re.search(r"([\u4e00-\u9fff]{2,10})", q)
        if abbr_m:
            stock_abbr = abbr_m.group(1).strip()
            for prefix in ["查询", "请问", "帮我", "麻烦", "一下", "请", "给我"]:
                if stock_abbr.startswith(prefix):
                    stock_abbr = stock_abbr[len(prefix):]
            stock_abbr = stock_abbr.strip("的")

    selected, table = [], None
    for kw, (t, col) in _METRIC_MAP.items():
        if kw in q and col not in selected:
            table = table or t
            if t == table:
                selected.append(col)

    if not selected or table is None:
        return ""

    where = []
    if stock_code:
        where.append(f"stock_code = '{stock_code}'")
    elif stock_abbr:
        where.append(f"stock_abbr = '{stock_abbr}'")
    else:
        return ""

    where.append(f"report_year = {year}")
    where.append(f"report_period = '{period}'")
    return f"SELECT {', '.join(selected)} FROM {table} WHERE {' AND '.join(where)};"


# ──────────────────────────────────────────────
# 执行 SQL（只读）
# ──────────────────────────────────────────────
def _execute_readonly_sql(sql: str) -> str:
    rejection = _validate_readonly(sql)
    if rejection:
        return rejection

    sql = _clean_sql(sql)

    try:
        conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
        conn.row_factory = sqlite3.Row
        cursor = conn.execute(sql)
        rows = cursor.fetchmany(MAX_ROWS)
        if not rows:
            conn.close()
            return "查询无结果。"
        cols = rows[0].keys()
        lines = ["\t".join(cols)]
        for r in rows:
            lines.append("\t".join(str(r[c]) for c in cols))
        conn.close()
        return "\n".join(lines)
    except Exception as e:
        return f"SQL 执行错误: {e}"


# ──────────────────────────────────────────────
# StructuredTool 定义
# ──────────────────────────────────────────────
class SQLQueryInput(BaseModel):
    sql: str = Field(description="要执行的 SQL SELECT 查询语句")


sql_query_tool = StructuredTool.from_function(
    func=_execute_readonly_sql,
    name="sql_query",
    description="在财务数据库上执行只读 SQL 查询，返回结果表格。仅支持 SELECT 语句。",
    args_schema=SQLQueryInput,
)


# 导出给 agent 使用的工具列表
ALL_TOOLS = [sql_query_tool]
