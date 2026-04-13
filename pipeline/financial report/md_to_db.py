# -*- coding: utf-8 -*-
"""
Markdown解析入库模块
将Markdown文件解析并写入SQLite数据库
包含所有依赖模块
"""

import datetime as dt
import html
import json
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path

import openpyxl

# ============================================================
# 配置数据 - parser_config.py
# ============================================================

REPORT_TYPE_PATTERNS = [
    ("半年度报告", "HY"),
    ("年年度报告", "FY"),
    ("年度报告", "FY"),
    ("第一季度报告", "Q1"),
    ("一季度报告", "Q1"),
    ("第三季度报告", "Q3"),
    ("三季度报告", "Q3"),
    ("季度报告", "QX"),
]

SECTION_KEYWORDS = {
    "core_performance_indicators_sheet": ["主要会计数据", "主要财务指标", "核心财务指标"],
    "balance_sheet": ["合并资产负债表", "资产负债表"],
    "cash_flow_sheet": ["合并现金流量表", "现金流量表"],
    "income_sheet": ["合并利润表", "利润表"],
}

SECTION_BOUNDARIES = {
    "core_performance_indicators_sheet": [
        ("## 四、主要会计数据和财务指标", "## 五、境内外会计准则下会计数据差异"),
        ("四、主要会计数据和财务指标", "五、境内外会计准则下会计数据差异"),
        ("近三年主要会计数据和财务指标", "八、"),
        ("一、 主要财务数据", "二、 股东信息"),
    ],
    "balance_sheet": [("合并资产负债表", "合并利润表")],
    "income_sheet": [("合并利润表", "合并现金流量表")],
    "cash_flow_sheet": [
        ("合并现金流量表", "(三)2023年起首次执行新会计准则"),
        ("合并现金流量表", "特此公告"),
    ],
}

HEADER_ROLE_HINTS = {
    "core_performance_indicators_sheet": {
        "current": [
            "本报告期末", "本报告期", "2023年半年度", "2023年6月30日",
            "2023年第一季度", "2022年度", "2022年", "2023年3月31日",
        ],
        "previous": [
            "上年同期", "上年度末", "2022年半年度", "2021年", "2022年12月31日",
        ],
        "yoy": ["增减", "变动幅度", "同比"],
    },
    "income_sheet": {
        "current": [
            "2025年", "2025年度", "2025年第一季度", "2025年前三季度",
            "2024年", "2024年度", "2024年半年度", "2024年第一季度", "2024年前三季度",
            "2023年", "2023年度", "2023年半年度", "2023年第一季度", "2023年前三季度",
            "本报告期", "本期", "年初至报告期末",
        ],
        "previous": [
            "2024年", "2024年度", "2024年半年度", "2024年第一季度", "2024年前三季度",
            "2023年", "2023年度", "2023年半年度", "2023年第一季度", "2023年前三季度",
            "2022年", "2022年度", "2022年半年度", "2022年第一季度", "2022年前三季度",
            "2021年", "2021年度", "上年同期",
        ],
    },
    "balance_sheet": {
        "current": [
            "2023年6月30日", "2023年3月31日", "2023年1月1日", "本报告期末", "期末",
        ],
        "previous": ["2022年12月31日", "2021年12月31日", "2021年", "上年度末"],
    },
    "cash_flow_sheet": {
        "current": [
            "2023年半年度", "2023年第一季度", "2022年度", "2022年", "本报告期", "本期",
        ],
        "previous": [
            "2022年半年度", "2022年第一季度", "2021年度", "2021年", "上年同期",
        ],
    },
}

GENERIC_HEADER_HINTS = {
    "current": ["本报告期", "本期", "期末", "本年累计", "当前"],
    "header_row": [
        "项目", "本报告期", "本报告期末", "上年同期", "上年度末",
        "2022年", "2021年", "2020年", "2023年第一季度", "2022 年第一季度",
        "2023年6月30日", "2023年1月1日", "2023 年半年度", "2022 年半年度",
    ],
    "yoy": ["本年比上年增减", "本报告期比上年同期增减", "同比", "增减", "变动幅度"],
}

SYNONYM = {
    "营业收入": "total_operating_revenue",
    "营业总收入": "total_operating_revenue",
    "营业总支出": "total_operating_expenses",
    "营业成本": "operating_expense_cost_of_sales",
    "归属于上市公司股东的净利润": "net_profit",
    "归母净利润": "net_profit",
    "净利润": "net_profit",
    "基本每股收益": "eps",
    "货币资金": "asset_cash_and_cash_equivalents",
    "应收账款": "asset_accounts_receivable",
    "存货": "asset_inventory",
    "交易性金融资产": "asset_trading_financial_assets",
    "在建工程": "asset_construction_in_progress",
    "总资产": "asset_total_assets",
    "应付账款": "liability_accounts_payable",
    "合同负债": "liability_contract_liabilities",
    "短期借款": "liability_short_term_loans",
    "应交税费": "liability_taxes_payable",
    "长期借款": "liability_long_term_borrowings",
    "总负债": "liability_total_liabilities",
    "销售商品、提供劳务收到的现金": "operating_cf_cash_from_sales",
    "经营活动产生的现金流量净额": "operating_cf_net_amount",
    "投资活动产生的现金流量净额": "investing_cf_net_amount",
    "筹资活动产生的现金流量净额": "financing_cf_net_amount",
    "现金及现金等价物净增加额": "net_cash_flow",
    "其他收益": "other_income",
    "营业利润": "operating_profit",
    "利润总额": "total_profit",
    "信用减值损失": "credit_impairment_loss",
    "资产减值损失": "asset_impairment_loss",
}

TABLE_FIELD_ALIASES = {
    "core_performance_indicators_sheet": {
        "eps": [
            "基本每股收益", "每股收益", "基本每股收益（元/股）", "基本每股收益(元/股)",
            "基本每股收益（元／股）", "基本每股收益(元／股)",
        ],
        "total_operating_revenue": ["营业总收入", "营业收入"],
        "net_profit_10k_yuan": [
            "归属于上市公司股东的净利润", "归属于母公司股东的净利润", "净利润",
        ],
        "net_asset_per_share": ["每股净资产"],
        "roe": ["净资产收益率", "加权平均净资产收益率"],
        "operating_cf_per_share": ["每股经营现金流量", "每股经营活动产生的现金流量净额"],
        "net_profit_excl_non_recurring": [
            "归属于上市公司股东的扣除非经常性损益的净利润",
            "扣除非经常性损益后的净利润", "扣非净利润",
        ],
        "gross_profit_margin": ["销售毛利率"],
        "net_profit_margin": ["销售净利率"],
        "roe_weighted_excl_non_recurring": [
            "加权平均净资产收益率（扣非）", "扣除非经常性损益后的加权平均净资产收益率",
        ],
    },
    "balance_sheet": {
        "asset_cash_and_cash_equivalents": ["货币资金"],
        "asset_accounts_receivable": ["应收账款"],
        "asset_inventory": ["存货"],
        "asset_trading_financial_assets": ["交易性金融资产"],
        "asset_construction_in_progress": ["在建工程"],
        "asset_total_assets": ["资产总计", "总资产"],
        "liability_accounts_payable": ["应付账款"],
        "liability_advance_from_customers": ["预收款项", "预收账款"],
        "liability_total_liabilities": ["负债合计", "总负债"],
        "liability_contract_liabilities": ["合同负债"],
        "liability_short_term_loans": ["短期借款"],
        "equity_unappropriated_profit": ["未分配利润"],
        "equity_total_equity": [
            "所有者权益(或股东权益)合计", "所有者权益合计", "股东权益合计",
            "归属于母公司所有者权益(或股东权益)合计", "归属于母公司所有者权益合计",
            "归属于上市公司股东的净资产",
        ],
    },
    "income_sheet": {
        "net_profit": ["归属于母公司股东的净利润", "归属于上市公司股东的净利润", "净利润"],
        "other_income": ["其他收益"],
        "total_operating_revenue": ["营业总收入", "营业收入"],
        "operating_revenue_yoy_growth": ["营业总收入同比", "营业收入同比"],
        "operating_expense_cost_of_sales": ["营业成本"],
        "operating_expense_selling_expenses": ["销售费用"],
        "operating_expense_administrative_expenses": ["管理费用"],
        "operating_expense_financial_expenses": ["财务费用"],
        "operating_expense_rnd_expenses": ["研发费用"],
        "operating_expense_taxes_and_surcharges": ["税金及附加"],
        "total_operating_expenses": ["营业总成本", "营业总支出"],
        "operating_profit": ["营业利润"],
        "total_profit": ["利润总额"],
        "asset_impairment_loss": ["资产减值损失"],
        "credit_impairment_loss": ["信用减值损失"],
    },
    "cash_flow_sheet": {
        "net_cash_flow": ["现金及现金等价物净增加额"],
        "operating_cf_net_amount": ["经营活动产生的现金流量净额"],
        "operating_cf_cash_from_sales": ["销售商品、提供劳务收到的现金"],
        "investing_cf_net_amount": ["投资活动产生的现金流量净额"],
        "investing_cf_cash_for_investments": ["投资支付的现金"],
        "investing_cf_cash_from_investment_recovery": ["收回投资收到的现金"],
        "financing_cf_cash_from_borrowing": ["取得借款收到的现金"],
        "financing_cf_cash_for_debt_repayment": ["偿还债务支付的现金"],
        "financing_cf_net_amount": ["筹资活动产生的现金流量净额"],
    },
}

# 余额类字段列表：这些字段在报表中表示期末余额，当期为空时应填0
BALANCE_SHEET_ZERO_FIELDS = {
    # 负债类
    "liability_advance_from_customers",  # 预收款项/预收账款
    "liability_contract_liabilities",       # 合同负债
    "liability_short_term_loans",          # 短期借款
    "liability_accounts_payable",           # 应付账款
    # 资产类
    "asset_accounts_receivable",           # 应收账款
    "asset_inventory",                      # 存货
}

QUALITY_LIMITS = {
    "field_ranges": {
        "eps": [-1000, 1000],
        "roe": [-100, 100],
        "gross_profit_margin": [-100, 100],
        "net_profit_margin": [-100, 100],
        "asset_liability_ratio": [0, 100],
        "operating_cf_ratio_of_net_cf": [-500, 500],
        "investing_cf_ratio_of_net_cf": [-500, 500],
        "financing_cf_ratio_of_net_cf": [-500, 500],
        "operating_revenue_yoy_growth": [-500, 500],
        "net_profit_yoy_growth": [-500, 500],
        "operating_revenue_qoq_growth": [-500, 500],
        "net_profit_qoq_growth": [-500, 500],
        "net_asset_per_share": [-100, 1000],
        "operating_cf_per_share": [-100, 100],
    },
    "cross_table_rel_tolerance": 0.15,
    "low_coverage": 0.4,
    "very_low_coverage": 0.2,
    "suspicious_unit_ratio": 1000,
}

# ============================================================
# 常量 - constants.py
# ============================================================

TABLE_NAME_MAP = {
    "核心业绩指标表": "core_performance_indicators_sheet",
    "资产负债表": "balance_sheet",
    "现金流量表": "cash_flow_sheet",
    "利润表": "income_sheet",
}

META_FIELDS = {
    "serial_number", "stock_code", "stock_abbr", "report_period",
    "report_year", "report_type", "source_file_name", "created_at",
}

# ============================================================
# 数据类型 - task_types.py
# ============================================================


@dataclass
class F:
    field_name: str
    cn_name: str
    data_type: str


@dataclass
class M:
    stock_code: str
    stock_abbr: str
    report_year: int
    report_type: str
    report_period: str
    source_file_name: str


# ============================================================
# 文本工具 - text_utils.py
# ============================================================


def normalize_text_cell(s):
    """标准化单元格文本"""
    s = html.unescape(str(s or ""))
    s = s.replace("\xa0", " ").replace("−", "-").replace("—", "-").replace("–", "-")
    return re.sub(r"\s+", " ", s).strip()


def detect_unit_text(text):
    """检测单位文本"""
    hits = re.findall(r"单位[:：]\s*(亿元|万元|元)", text)
    return hits[-1] if hits else ""


def detect_unit(t):
    """检测表格单位"""
    prefix = t.split("<table>", 1)[0] if "<table>" in t else t[:500]
    unit = detect_unit_text(prefix)
    if not unit:
        table_part = t.split("<table>", 1)[1] if "<table>" in t else t
        unit = detect_unit_text(table_part[:300])
    if not unit:
        unit = "元"

    if unit == "亿元":
        return 100000000.0
    if unit == "万元":
        return 10000.0
    return 1.0


def clean_num(s, f=1.0):
    """解析数字"""
    t = normalize_text_cell(s)
    if not t or t in {"--", "-", "不适用", "N/A", "nan"}:
        return None
    t = t.replace(",", "")
    t = re.sub(r"\s+", "", t)
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1]
    if t.endswith("%"):
        try:
            return -float(t[:-1]) if neg else float(t[:-1])
        except Exception:
            return None
    try:
        v = float(t) * f
        return -v if neg else v
    except Exception:
        return None


def nk(s):
    """标准化关键字段"""
    s = (s or "").strip().replace("（", "(").replace("）", ")").replace("－", "-")
    s = re.sub(r"[\s\-_:：·,，。()（）\[\]【】]", "", s)
    return re.sub(r"(万元|亿元|元|人民币|%)", "", s)


def calc_growth(curr, prev):
    """计算增长率"""
    if curr is None or prev in (None, 0):
        return None
    return round((curr - prev) / abs(prev) * 100, 4)


# ============================================================
# 数据库工具 - db_utils.py
# ============================================================


def business_field_count(row):
    """统计业务字段数量"""
    return sum(1 for k, v in row.items() if k not in META_FIELDS and v is not None)


def sql_type(t):
    """获取SQL类型"""
    t = t.lower()
    return "INTEGER" if "int" in t else ("REAL" if any(k in t for k in ["decimal", "float", "double"]) else "TEXT")


def ensure_tables(conn, schema):
    """确保数据库表存在"""
    cur = conn.cursor()
    for t, fs in schema.items():
        cols = [f'"{f.field_name}" {sql_type(f.data_type)}' for f in fs]
        ex = {f.field_name for f in fs}
        for c, typ in [
            ("report_period", "TEXT"),
            ("report_year", "INTEGER"),
            ("report_type", "TEXT"),
            ("source_file_name", "TEXT"),
            ("created_at", "TEXT"),
        ]:
            if c not in ex:
                cols.append(f'"{c}" {typ}')
        cur.execute(
            f'CREATE TABLE IF NOT EXISTS "{t}" ({",".join(cols)}, UNIQUE(stock_code, report_period))'
        )

    cur.execute(
        "CREATE TABLE IF NOT EXISTS raw_table_snippets("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "stock_code TEXT,report_period TEXT,target_table TEXT,"
        "source_section TEXT,snippet TEXT,created_at TEXT)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS etl_audit_log("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "source_file_name TEXT,status TEXT,message TEXT,created_at TEXT)"
    )
    conn.commit()


def upsert(conn, t, row):
    """插入或更新数据"""
    if not row or business_field_count(row) == 0:
        return

    cur = conn.execute(
        f'SELECT * FROM "{t}" WHERE stock_code=? AND report_period=?',
        (row.get("stock_code"), row.get("report_period")),
    )
    existing = cur.fetchone()
    if existing is not None:
        existing_row = dict(zip([d[0] for d in cur.description], existing))
        prefer_new_business = business_field_count(row) >= business_field_count(existing_row)
        merged_row = dict(existing_row)
        for key, value in row.items():
            if key in META_FIELDS:
                if value not in (None, ""):
                    merged_row[key] = value
            elif value is not None and (prefer_new_business or merged_row.get(key) is None):
                merged_row[key] = value
        row = merged_row

    cols = list(row.keys())
    ps = ",".join(["?"] * len(cols))
    cs = ",".join([f'"{c}"' for c in cols])
    us = ",".join([f'"{c}"=excluded."{c}"' for c in cols if c not in {"stock_code", "report_period"}])
    conn.execute(
        f'INSERT INTO "{t}" ({cs}) VALUES ({ps}) '
        f'ON CONFLICT(stock_code, report_period) DO UPDATE SET {us}',
        [row[c] for c in cols],
    )


def log(conn, f, s, m):
    """记录日志"""
    conn.execute(
        "INSERT INTO etl_audit_log(source_file_name,status,message,created_at) VALUES (?,?,?,?)",
        (f, s, m, dt.datetime.now().isoformat()),
    )


def coverage(ext, fs):
    """计算覆盖率"""
    tgt = [f.field_name for f in fs if f.field_name not in META_FIELDS]
    hit = sum(1 for k in tgt if k in ext)
    return (round(hit / len(tgt), 4), hit, len(tgt)) if tgt else (0.0, 0, 0)


def filter_to_schema(ext, fs):
    """过滤到schema定义的字段"""
    allowed = {f.field_name for f in fs}
    return {k: v for k, v in ext.items() if k in allowed}


# ============================================================
# 元数据检测 - meta_detect.py
# ============================================================


def normalize_company_name(name: str):
    """标准化公司名称"""
    s = nk(name)
    for suffix in [
        "股份有限公司", "集团股份有限公司", "医药股份有限公司", "有限公司",
        "集团有限公司", "股份", "集团", "医药",
    ]:
        s = s.replace(nk(suffix), "")
    return s


def lookup_code_by_name(name: str, cm):
    """根据名称查找股票代码"""
    if not name:
        return "", ""

    normalized_name = normalize_company_name(name)
    for code, abbr in cm.items():
        if abbr == name:
            return code, abbr

    for code, abbr in cm.items():
        normalized_abbr = normalize_company_name(abbr)
        if normalized_abbr and (
            normalized_abbr == normalized_name
            or normalized_abbr in normalized_name
            or normalized_name in normalized_abbr
        ):
            return code, abbr

    return "", ""


def infer_report_type(text: str):
    """推断报告类型"""
    for k, v in REPORT_TYPE_PATTERNS:
        if k in text:
            if v == "QX" and ("第一季度" in text or "一季度" in text):
                return ("一季度报告", "Q1")
            if v == "QX" and ("第三季度" in text or "三季度" in text):
                return ("三季度报告", "Q3")
            return (k, v)
    return ("年度报告", "FY")


def build_report_period(report_year, report_code):
    """构建报告期字符串"""
    return f"{report_year}{report_code}" if report_year and report_code else ""


def build_serial_number(stock_code, report_year, report_period):
    """构建序号"""
    if not stock_code or not report_year:
        return None
    suffix = ""
    if report_period and len(report_period) >= 4:
        suffix = report_period[4:]
    order = {"Q1": 1, "HY": 2, "Q3": 3, "FY": 4}
    try:
        code_num = int(str(stock_code))
    except ValueError:
        return None
    period_num = report_year * 10 + order.get(suffix, 0)
    return code_num * 100000 + period_num


def detect_meta_from_md(md: str, cm):
    """从Markdown内容检测元数据"""
    t = md[:20000]
    plain = html.unescape(re.sub(r"<[^>]+>", " ", t))
    code = ""
    abbr = ""

    m = re.search(r"(?:公司代码|证券代码)[:：]\s*(\d{6})", plain)
    if not m:
        m = re.search(r"股票代码\D{0,40}(\d{6})", plain)
    if m:
        code = m.group(1)

    m = re.search(r"(?:公司简称|证券简称|股票简称)[:：]\s*([^\n\r]+)", plain)
    if not m:
        m = re.search(r"股票简称\D{0,30}([^\s\d][^\n\r]{1,20})", plain)
    if m:
        abbr = m.group(1).strip().split()[0]

    title_lines = re.findall(r"^#\s*(.+)$", t, flags=re.MULTILINE)
    company_title = title_lines[0].strip() if title_lines else ""
    if not abbr and company_title:
        _, mapped_abbr = lookup_code_by_name(company_title, cm)
        abbr = mapped_abbr or abbr

    title_match = re.search(
        r"#\s*(.+?(?:半年度报告|年度报告|第一季度报告|一季度报告|第三季度报告|三季度报告))",
        t,
    )
    title = title_match.group(1) if title_match else plain[:500]

    y = re.search(r"(20\d{2})\s*年", title) or re.search(r"(20\d{2})\s*年", plain)
    report_year = int(y.group(1)) if y else 0

    report_type, report_code = infer_report_type(title if title else plain)
    report_period = build_report_period(report_year, report_code)

    if code and not abbr:
        abbr = cm.get(code, "")

    if abbr and not code:
        code, mapped_abbr = lookup_code_by_name(abbr, cm)
        abbr = abbr or mapped_abbr

    if not code and company_title:
        code, mapped_abbr = lookup_code_by_name(company_title, cm)
        if mapped_abbr and not abbr:
            abbr = mapped_abbr

    return M(code, abbr, report_year, report_type, report_period, "")


def merge_meta(primary: M, fallback: M, source_name: str):
    """合并元数据"""
    code = primary.stock_code or fallback.stock_code
    abbr = primary.stock_abbr or fallback.stock_abbr
    report_year = primary.report_year or fallback.report_year
    report_type = primary.report_type if primary.report_period else fallback.report_type
    report_period = primary.report_period or fallback.report_period
    return M(code, abbr, report_year, report_type, report_period, source_name)


def detect_meta(pdf, cm, md: str = ""):
    """检测PDF元数据"""
    n = pdf.stem
    c = ""
    a = ""

    # 尝试从目录名提取股票代码（更宽松的匹配）
    m = re.search(r"(?<!\d)(\d{6})(?!\d)", n)
    if m:
        c = m.group(1)
        a = cm.get(c, "")
    else:
        # 尝试从公司名称提取
        left = n.split("：")[0].strip() if "：" in n else ""
        if left:
            c, mapped_abbr = lookup_code_by_name(left, cm)
            a = mapped_abbr or left

    # 从目录名提取报告年份和类型
    y = re.search(r"(20\d{2})年", n)
    if not y:
        # 尝试从时间戳提取年份
        y = re.search(r"(20\d{2})", n)
    ry = int(y.group(1)) if y else 0
    rt, rc = infer_report_type(n)
    rp = build_report_period(ry, rc)

    base = M(c, a, ry, rt, rp, pdf.name)
    if not md:
        return base

    md_meta = detect_meta_from_md(md, cm)
    return merge_meta(base, md_meta, pdf.name)


def is_summary_name(p):
    """判断是否为摘要文件名"""
    n = p.stem
    return any(
        k in n
        for k in [
            "报告摘要", "年度报告摘要", "半年度报告摘要", "季度报告摘要",
            "一季度报告摘要", "三季度报告摘要", "摘要版",
        ]
    )


def detect_summary_from_md(md: str):
    """从Markdown内容判断是否为摘要"""
    t = md[:12000]
    score = 0
    pos = [
        "年度报告摘要", "半年度报告摘要", "季度报告摘要",
        "第一季度报告摘要", "第三季度报告摘要", "报告摘要",
    ]
    neg = [
        "目录", "公司治理", "重要事项", "审计意见类型",
        "董事会报告", "财务报表附注", "备查文件目录", "释义",
    ]
    for k in pos:
        if k in t:
            score += 2
    for k in neg:
        if k in t:
            score -= 1
    if "报告摘要" in t and not any(k in t for k in ["财务报表附注", "备查文件目录", "董事会报告"]):
        score += 2
    return score >= 2


# ============================================================
# 表格解析 - table_parser.py
# ============================================================


def table_rows(sec):
    """解析表格行"""
    rows = []
    html_rows = re.findall(r"<tr[^>]*>(.*?)</tr>", sec, flags=re.S | re.I)
    for hr in html_rows:
        raw = [
            normalize_text_cell(re.sub(r"<[^>]+>", " ", c))
            for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", hr, flags=re.S | re.I)
        ]
        if len(raw) >= 2 and not all(not x for x in raw):
            rows.append(raw)
    if rows:
        return rows

    for line in sec.splitlines():
        if "|" in line:
            cs = [normalize_text_cell(c) for c in line.split("|")[1:-1] if c is not None]
            if len(cs) >= 2 and not all(re.fullmatch(r"[:\-\s]+", c or "") for c in cs):
                rows.append(cs)
    return rows


def locate_header_row(rows):
    """定位表头行"""
    header_hints = GENERIC_HEADER_HINTS.get("header_row", [])
    for i, cs in enumerate(rows[:12]):
        joined = " ".join(cs)
        if any(h in joined for h in header_hints):
            return i
    return 0


def parse_table_matrix(sec):
    """解析表格矩阵"""
    rows = table_rows(sec)
    if not rows:
        return [], [], []
    hi = locate_header_row(rows)
    header = rows[hi]
    body = rows[hi + 1 :]
    return rows, header, body


def standardize_table_section(sec):
    """标准化表格区块"""
    sec = sec.replace("\r\n", "\n").replace("\r", "\n")
    sec = sec.replace("−", "-").replace("—", "-").replace("–", "-")
    return html.unescape(sec)


# ============================================================
# 分段解析 - sectioning.py
# ============================================================

YEAR_LIKE_RE = re.compile(r"20\d{2}年(?:末|度|第[一二三四]季度)?|20\d{2}年\d{1,2}月\d{1,2}日")


def slice_by_boundaries(md, table):
    """按边界切分"""
    parts = []
    for start_kw, end_kw in SECTION_BOUNDARIES.get(table, []):
        start = md.find(start_kw)
        if start < 0:
            continue
        end = md.find(end_kw, start + len(start_kw)) if end_kw else -1
        parts.append(md[start : (end if end > start else len(md))])
    return parts


def keyword_sections(md, table):
    """按关键词分段"""
    kws = SECTION_KEYWORDS[table]
    ls = md.splitlines()
    rs = []
    b = []
    rec = False
    for line in ls:
        plain = html.unescape(re.sub(r"<[^>]+>", " ", line))
        if any(k in plain for k in kws):
            if b:
                rs.append("\n".join(b))
                b = []
            rec = True
        if rec:
            if plain.startswith("# ") and not any(k in plain for k in kws):
                if b:
                    rs.append("\n".join(b))
                b = []
                rec = False
            else:
                b.append(line)
    if b:
        rs.append("\n".join(b))
    return rs


def score_section(sec, table):
    """评分区块"""
    score = 0
    plain = html.unescape(re.sub(r"<[^>]+>", " ", sec))
    for kw in SECTION_KEYWORDS.get(table, []):
        if kw in plain:
            score += 3
    score += plain.count("<table>") * 2
    score += sum(1 for h in GENERIC_HEADER_HINTS.get("header_row", []) if h in plain[:2000])
    score += sum(1 for h in HEADER_ROLE_HINTS.get(table, {}).get("current", []) if h in plain[:2000])
    score += sum(1 for h in HEADER_ROLE_HINTS.get(table, {}).get("previous", []) if h in plain[:2000])
    if table == "core_performance_indicators_sheet":
        score += sum(2 for h in ["营业收入", "净利润", "总资产"] if h in plain)
    if table == "income_sheet":
        if "合并利润表" in plain[:800]:
            score += 12
        score += sum(4 for h in ["一、营业总收入", "二、营业总成本", "五、净利润"] if h in plain)
        score += sum(3 for h in ["归属于母公司股东的净利润", "利润总额", "营业利润"] if h in plain)
        if "季度数据与已披露定期报告数据差异说明" in plain:
            score -= 20
        if "关键审计事项" in plain:
            score -= 20
        if "主要会计数据" in plain and "合并利润表" not in plain[:800]:
            score -= 12
    return score


def sections(md, table):
    """获取最优区块"""
    candidates = slice_by_boundaries(md, table) or keyword_sections(md, table)
    if not candidates:
        return []
    ranked = sorted(
        ((score_section(sec, table), idx, sec) for idx, sec in enumerate(candidates)),
        key=lambda x: (-x[0], x[1]),
    )
    return [sec for _, _, sec in ranked[:3]]


def match_header_hint(text, hints):
    """匹配表头提示"""
    n = nk(text)
    return any(nk(k) in n for k in hints)


def is_core_value_header(text):
    """判断是否为核心数值表头"""
    h = nk(text)
    if not h:
        return False
    if any(k in h for k in ["增减", "变动幅度", "同比"]):
        return False
    if any(k in h for k in ["本报告期", "本报告期末", "上年同期", "上年度末", "调整前", "调整后", "年初至报告期末"]):
        return True
    return bool(YEAR_LIKE_RE.search(h))


def current_col_from_header(table, header):
    """从表头获取当期列索引"""
    hs = [(i, nk(h)) for i, h in enumerate(header)]
    table_hints = HEADER_ROLE_HINTS.get(table, {})
    if table == "core_performance_indicators_sheet":
        for i, h in hs:
            if "年初至报告期末" in h:
                return i
        yoy_idx = yoy_col_from_header(table, header)
        if yoy_idx is not None:
            candidates = [i for i, h in hs if i < yoy_idx and is_core_value_header(h)]
            if candidates:
                return candidates[0]
        candidates = [i for i, h in hs if is_core_value_header(h)]
        if candidates:
            return candidates[0]
    for i, h in hs:
        if match_header_hint(h, table_hints.get("current", [])):
            return i
    for i, h in hs:
        if match_header_hint(h, GENERIC_HEADER_HINTS.get("current", [])):
            return i
    return 1 if len(header) > 1 else 0


def previous_col_from_header(table, header, current_idx):
    """从表头获取上期列索引"""
    hs = [(i, nk(h)) for i, h in enumerate(header)]
    table_hints = HEADER_ROLE_HINTS.get(table, {})
    if table == "core_performance_indicators_sheet":
        yoy_idx = yoy_col_from_header(table, header)
        adjusted_previous_candidates = [i for i, h in hs if "调整后" in h and i != current_idx]
        if adjusted_previous_candidates:
            if current_idx is not None:
                after_current = [i for i in adjusted_previous_candidates if i > current_idx]
                if after_current:
                    return after_current[-1]
            return adjusted_previous_candidates[-1]
        if current_idx is not None:
            between_current_and_yoy = [
                i for i, h in hs
                if i > current_idx and (yoy_idx is None or i < yoy_idx) and is_core_value_header(h)
            ]
            if between_current_and_yoy:
                return between_current_and_yoy[0]
        if any("年初至报告期末" in h for _, h in hs):
            return None
        has_yoy = any(
            match_header_hint(h, table_hints.get("yoy", []) + GENERIC_HEADER_HINTS.get("yoy", []))
            for _, h in hs
        )
        if has_yoy and len(header) <= 3:
            return None
    for i, h in hs:
        if i != current_idx and match_header_hint(h, table_hints.get("previous", [])):
            return i
    return current_idx + 1 if current_idx + 1 < len(header) else None


def yoy_col_from_header(table, header):
    """从表头获取同比列索引"""
    hs = [(i, nk(h)) for i, h in enumerate(header)]
    hints = HEADER_ROLE_HINTS.get(table, {}).get("yoy", []) + GENERIC_HEADER_HINTS.get("yoy", [])
    if table == "core_performance_indicators_sheet":
        for i, h in hs:
            if "年初至报告期末比上年同期增减" in h or "年初至报告期末比上年同期变动幅度" in h:
                return i
    for i, h in hs:
        if match_header_hint(h, hints):
            return i
    return None


# ============================================================
# 别名处理 - aliasing.py
# ============================================================


def build_alias(fs, table=""):
    """构建别名映射"""
    m = {}
    for f in fs:
        c = f.cn_name
        m[nk(c)] = f.field_name
        m[nk(c.replace("-", ""))] = f.field_name
        m[
            nk(
                c.replace("资产-", "")
                .replace("负债-", "")
                .replace("股东权益-", "")
                .replace("经营性现金流-", "")
                .replace("投资性现金流-", "")
                .replace("融资性现金流-", "")
                .replace("营业总支出-", "")
                .replace("现金流-", "")
                .replace("利润-", "")
                .replace("(万元)", "")
                .replace("（万元）", "")
                .replace("(元)", "")
                .replace("（元）", "")
                .replace("(%)", "")
                .replace("（%）", "")
            )
        ] = f.field_name
        for alias in TABLE_FIELD_ALIASES.get(table, {}).get(f.field_name, []):
            m[nk(alias)] = f.field_name

    for k, v in SYNONYM.items():
        kk = nk(k)
        if kk not in m:
            m[kk] = v
    return m


# ============================================================
# 行逻辑处理 - row_logic.py
# ============================================================

MAX_REASONABLE_GROWTH_ABS = 1000
MAX_REASONABLE_RATIO_ABS = 100


def is_plausible_growth(v):
    """判断增长率是否合理"""
    return isinstance(v, (int, float)) and abs(v) <= MAX_REASONABLE_GROWTH_ABS


def is_plausible_ratio(v):
    """判断比率是否合理"""
    return isinstance(v, (int, float)) and abs(v) <= MAX_REASONABLE_RATIO_ABS


def row_label(cs):
    """获取行标签"""
    if not cs:
        return ""
    s = normalize_text_cell(cs[0])
    s = re.sub(r"^其中[:：]", "", s)
    s = re.sub(r"^[加减][:：]", "", s)
    s = re.sub(r"^[（(]?[一二三四五六七八九十0-9]+[)）]?[、.:：]\s*", "", s)
    s = re.sub(r"^[（(][一二三四五六七八九十0-9]+[)）]\s*", "", s)
    s = re.sub(r"（[^）]*填列）|\([^)]*填列\)", "", s)
    s = re.sub(r"（损失以[^）]*）|\(损失以[^)]*\)", "", s)
    s = re.sub(r"（亏损以[^）]*）|\(亏损以[^)]*\)", "", s)
    s = re.sub(r"（净亏损以[^）]*）|\(净亏损以[^)]*\)", "", s)
    s = re.sub(r"（亏损总额以[^）]*）|\(亏损总额以[^)]*\)", "", s)
    s = re.sub(r"（元/股）|\(元/股\)|（元／股）|\(元／股\)", "", s)
    s = re.sub(r"（元）|\(元\)|（万元）|\(万元\)|（%）|\(%\)", "", s)
    return nk(s)


def is_header_row(cs):
    """判断是否为表头行"""
    joined = " ".join(cs)
    if row_label(cs) in {
        "主要会计数据", "主要财务指标", "项目",
        "流动资产", "非流动资产", "流动负债", "非流动负债",
        "所有者权益或股东权益",
    }:
        return True
    if cs and not str(cs[0]).strip():
        return True
    if len(cs) >= 2 and all(clean_num(c) is None for c in cs[1:]):
        return True
    return any(
        k in joined
        for k in [
            "本报告期", "上年同期", "本期比上年同期增减",
            "2022年", "2021年", "2020年",
            "2023年第一季度", "2022 年第一季度",
            "2023年3月31日", "2022年12月31日",
        ]
    )


def first_number_values(cs, uf):
    """获取行中的第一个数字值"""
    vals = []
    for c in cs[1:]:
        if c is None:
            continue
        t = str(c).strip()
        if any(k in t for k in ["增加", "减少", "百分点"]):
            continue
        pv = clean_num(t, uf)
        if pv is not None:
            vals.append(pv)
    return vals


def extract_core_row_metrics(cs, uf):
    """提取核心行指标"""
    nums = []
    for idx, c in enumerate(cs[1:], start=1):
        text = normalize_text_cell(c)
        if not text or text in {"不适用", "N/A"}:
            continue
        val = clean_num(text, 1.0 if "%" in text else uf)
        if val is None:
            continue
        nums.append((idx, val, "%" in text))
    if not nums:
        return None, None, None

    yoy_candidates = [(idx, val) for idx, val, is_pct in nums if is_pct and is_plausible_growth(val)]
    yoy_hit = yoy_candidates[0] if yoy_candidates else None
    if yoy_hit:
        yoy_idx_actual, yoy = yoy_hit
        prev_candidates = [val for idx, val, is_pct in nums if not is_pct and idx < yoy_idx_actual]
        curr = prev_candidates[0] if prev_candidates else next((val for _, val, is_pct in nums if not is_pct), None)
        prev = prev_candidates[-1] if len(prev_candidates) >= 2 else None
        return curr, prev, yoy

    non_pct = [val for _, val, is_pct in nums if not is_pct]
    curr = non_pct[0] if non_pct else None
    prev = non_pct[1] if len(non_pct) >= 2 else None
    return curr, prev, None


def derive_values_from_row(cs, uf, table, tf, curr_is_none=False):
    """从行推导值
    
    Args:
        curr_is_none: 当期值是否为空，如果为空则对于余额类字段不应使用上期值
    """
    vals = first_number_values(cs, uf)
    if not vals:
        return []

    if table == "core_performance_indicators_sheet":
        if tf in {"eps", "roe", "roe_weighted_excl_non_recurring"}:
            return vals[:2]
        if len(vals) >= 3:
            return [vals[0], vals[1], vals[2]]
        return vals

    # 对于资产负债表中的余额类字段：
    # - 如果当期值为空，不使用上期值（应填0）
    # - 如果当期值不为空，使用当期值
    if table == "balance_sheet" and tf in BALANCE_SHEET_ZERO_FIELDS:
        if curr_is_none:
            return []  # 当期为空，返回空，让调用方填0
        return [vals[0]] if vals else []

    if table in {"income_sheet", "cash_flow_sheet", "balance_sheet"}:
        return vals[:2]

    return vals


def value_by_idx(cs, idx, uf):
    """按索引获取值"""
    if idx is None or idx < 0 or idx >= len(cs):
        return None
    if cs and cs[0].strip() == "":
        return None
    return clean_num(cs[idx], uf)


def score_core_candidate(row, seen, sec):
    """为核心候选评分"""
    score = len(row) * 10
    for k in [
        "operating_revenue_yoy_growth", "net_profit_yoy_growth",
        "net_profit_excl_non_recurring_yoy", "gross_profit_margin",
        "net_profit_margin", "roe",
    ]:
        v = row.get(k)
        if isinstance(v, (int, float)):
            if k.endswith("_growth"):
                score += 6 if -500 <= v <= 500 else -30
            elif -100 <= v <= 100:
                score += 4
            else:
                score -= 20
    score += sum(
        3
        for vs in seen.values()
        if len(vs) >= 3 and isinstance(vs[2], (int, float)) and abs(vs[2]) <= 500
    )
    score += sec[:1500].count("%")
    return score


def normalize_field_value(field_name, val):
    """标准化字段值"""
    if val is None:
        return None
    if field_name in {"eps", "net_asset_per_share", "operating_cf_per_share"}:
        return round(val, 4) if isinstance(val, (int, float)) else val
    if field_name == "net_cash_flow":
        return round(val / 10000.0, 4)
    if (
        field_name.endswith("_growth")
        or field_name.endswith("_ratio")
        or field_name.endswith("_ratio_of_net_cf")
        or field_name in {"roe", "gross_profit_margin", "net_profit_margin", "roe_weighted_excl_non_recurring"}
    ):
        return round(val, 4)
    if field_name == "operating_cf_net_amount":
        return round(val / 10000.0, 4)
    return round(val / 10000.0, 4)


def enrich_table_fields(table, row, seen):
    """丰富表格字段"""
    if table == "core_performance_indicators_sheet":
        if "total_operating_revenue" in seen and "operating_revenue_yoy_growth" not in row:
            vs = seen["total_operating_revenue"]
            yoy = vs[2] if len(vs) >= 3 and is_plausible_growth(vs[2]) else calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None
            )
            if is_plausible_growth(yoy):
                row["operating_revenue_yoy_growth"] = yoy
        if "net_profit_10k_yuan" in seen and "net_profit_yoy_growth" not in row:
            vs = seen["net_profit_10k_yuan"]
            yoy = vs[2] if len(vs) >= 3 and is_plausible_growth(vs[2]) else calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None
            )
            if is_plausible_growth(yoy):
                row["net_profit_yoy_growth"] = yoy
        if "net_profit_excl_non_recurring" in seen and "net_profit_excl_non_recurring_yoy" not in row:
            vs = seen["net_profit_excl_non_recurring"]
            yoy = vs[2] if len(vs) >= 3 and is_plausible_growth(vs[2]) else calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None
            )
            if is_plausible_growth(yoy):
                row["net_profit_excl_non_recurring_yoy"] = yoy
        if (
            row.get("net_profit_10k_yuan") is not None
            and row.get("total_operating_revenue") not in (None, 0)
            and "net_profit_margin" not in row
        ):
            margin = round(row["net_profit_10k_yuan"] / row["total_operating_revenue"] * 100, 4)
            if is_plausible_ratio(margin):
                row["net_profit_margin"] = margin
        if (
            row.get("net_asset_per_share") is None
            and row.get("asset_total_assets") is not None
            and row.get("roe") is not None
            and row.get("net_profit_10k_yuan") is not None
            and row["roe"] not in (None, 0)
        ):
            est_equity = row["net_profit_10k_yuan"] * 100 / row["roe"]
            row["net_asset_per_share"] = round(est_equity * 10000 / 373270285, 4)

    if table == "balance_sheet":
        if "asset_total_assets" in seen and "asset_total_assets_yoy_growth" not in row:
            vs = seen["asset_total_assets"]
            row["asset_total_assets_yoy_growth"] = calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None
            )
        if "liability_total_liabilities" in seen and "liability_total_liabilities_yoy_growth" not in row:
            vs = seen["liability_total_liabilities"]
            row["liability_total_liabilities_yoy_growth"] = calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None
            )
        if (
            row.get("liability_total_liabilities") is not None
            and row.get("asset_total_assets") not in (None, 0)
            and "asset_liability_ratio" not in row
        ):
            row["asset_liability_ratio"] = round(
                row["liability_total_liabilities"] / row["asset_total_assets"] * 100, 4
            )

    if table == "income_sheet":
        if "net_profit" in seen and "net_profit_yoy_growth" not in row:
            vs = seen["net_profit"]
            yoy = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None)
            if is_plausible_growth(yoy):
                row["net_profit_yoy_growth"] = yoy
        if "total_operating_revenue" in seen and "operating_revenue_yoy_growth" not in row:
            vs = seen["total_operating_revenue"]
            yoy = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None)
            if is_plausible_growth(yoy):
                row["operating_revenue_yoy_growth"] = yoy

    if table == "cash_flow_sheet":
        if "net_cash_flow" in seen and "net_cash_flow_yoy_growth" not in row:
            vs = seen["net_cash_flow"]
            row["net_cash_flow_yoy_growth"] = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None)
        for a, b in [
            ("operating_cf_net_amount", "operating_cf_ratio_of_net_cf"),
            ("investing_cf_net_amount", "investing_cf_ratio_of_net_cf"),
            ("financing_cf_net_amount", "financing_cf_ratio_of_net_cf"),
        ]:
            if row.get(a) is not None and row.get("net_cash_flow") not in (None, 0) and b not in row:
                row[b] = round(row[a] / row["net_cash_flow"] * 100, 4)

    return {k: v for k, v in row.items() if v is not None}


# ============================================================
# Markdown解析器 - md_parser.py
# ============================================================


def extract_adjusted_core_metrics(cs, header, uf):
    """提取调整后的核心指标"""
    values = []
    for idx, c in enumerate(cs[1:], start=1):
        text = normalize_text_cell(c)
        if not text or text in {"不适用", "N/A"}:
            continue
        val = clean_num(text, 1.0 if "%" in text else uf)
        if val is None:
            continue
        values.append((idx, val, "%" in text))

    if not values:
        return None, None, None

    data_cells = [(val, is_pct) for _, val, is_pct in values]
    has_cumulative_block = any("年初至报告期末" in normalize_text_cell(h) for h in header)

    if has_cumulative_block and len(data_cells) >= 8:
        curr = data_cells[4][0]
        prev = data_cells[6][0] if len(data_cells) >= 7 else data_cells[5][0]
        yoy = data_cells[7][0] if len(data_cells) >= 8 and data_cells[7][1] else None
        return curr, prev, yoy

    if has_cumulative_block and len(data_cells) >= 6:
        curr = data_cells[3][0]
        prev = data_cells[4][0]
        yoy = data_cells[5][0] if data_cells[5][1] else None
        return curr, prev, yoy

    if not has_cumulative_block and len(data_cells) >= 4 and data_cells[-1][1]:
        non_pct = [val for val, is_pct in data_cells[:-1] if not is_pct]
        curr = non_pct[0] if non_pct else None
        prev = non_pct[-1] if len(non_pct) >= 2 else None
        yoy = data_cells[-1][0]
        return curr, prev, yoy

    non_pct = [val for val, is_pct in data_cells if not is_pct]
    pct = [val for val, is_pct in data_cells if is_pct]
    if not non_pct and not pct:
        return None, None, None

    curr = non_pct[-3] if len(non_pct) >= 3 else (non_pct[0] if non_pct else pct[0])
    prev = non_pct[-1] if len(non_pct) >= 2 else (pct[-2] if len(pct) >= 2 else None)
    yoy = pct[-1] if pct else None
    return curr, prev, yoy


def parse_section_rows(sec, table, fs, am):
    """解析区块行"""
    row = {}
    seen = {}
    field_evidence = {}

    uf = detect_unit(sec)
    rows, header, body = parse_table_matrix(sec)
    if not rows:
        return row, seen, field_evidence

    cur_idx = current_col_from_header(table, header)
    prev_idx = previous_col_from_header(table, header, cur_idx)
    yoy_idx = yoy_col_from_header(table, header) if table == "core_performance_indicators_sheet" else None
    has_adjusted_layout = any(
        "调整后" in normalize_text_cell(cell)
        for row_cells in rows[:3]
        for cell in row_cells
    )

    for cs in body:
        if is_header_row(cs):
            continue

        lbl = row_label(cs)
        if table == "core_performance_indicators_sheet" and lbl in {
            "非经常性损益项目和金额", "项目名称", "短期借款", "预收款项", "合同负债",
        }:
            break

        tf = am.get(lbl)
        if not tf:
            continue

        curr = value_by_idx(cs, cur_idx, uf)
        prev = value_by_idx(cs, prev_idx, uf)
        yoy = None

        if table == "core_performance_indicators_sheet":
            if prev_idx is not None and prev_idx < len(cs):
                prev = value_by_idx(cs, prev_idx, uf)
            smart_curr, smart_prev, smart_yoy = extract_core_row_metrics(cs, uf)
            adjusted_curr, adjusted_prev, adjusted_yoy = extract_adjusted_core_metrics(cs, header, uf)
            if not has_adjusted_layout:
                adjusted_curr = adjusted_prev = adjusted_yoy = None
            has_adjusted_header = has_adjusted_layout
            if smart_curr is not None and curr is None:
                curr = smart_curr
            if smart_prev is not None and prev is None and not (prev_idx is None and yoy_idx is not None):
                prev = smart_prev
            if adjusted_curr is not None:
                curr = adjusted_curr
            if adjusted_prev is not None:
                prev = adjusted_prev

            raw_yoy = normalize_text_cell(cs[yoy_idx]) if yoy_idx is not None and yoy_idx < len(cs) else ""
            if adjusted_yoy is not None and has_adjusted_header:
                yoy = adjusted_yoy
            elif raw_yoy not in {"", "不适用", "N/A"} and yoy_idx is not None and yoy_idx != cur_idx:
                yoy = clean_num(raw_yoy, 1.0)
            elif adjusted_yoy is not None:
                yoy = adjusted_yoy
            elif smart_yoy is not None:
                yoy = smart_yoy
            elif curr is not None and prev is not None:
                yoy = calc_growth(curr, prev)

            if yoy is not None:
                if tf == "total_operating_revenue" and "operating_revenue_yoy_growth" not in row:
                    row["operating_revenue_yoy_growth"] = round(yoy, 4)
                if tf == "net_profit_10k_yuan" and "net_profit_yoy_growth" not in row:
                    row["net_profit_yoy_growth"] = round(yoy, 4)
                if tf == "net_profit_excl_non_recurring" and "net_profit_excl_non_recurring_yoy" not in row:
                    row["net_profit_excl_non_recurring_yoy"] = round(yoy, 4)

        vals = [curr] if curr is not None else []
        if prev is not None:
            vals.append(prev)
        if yoy is not None:
            vals.append(yoy)

        if not vals:
            vals = derive_values_from_row(cs, uf, table, tf, curr_is_none=(curr is None))

        vals = [v for v in vals if v is not None]
        if not vals:
            # 当期为空的余额类字段，应填0而非保持NULL
            if table == "balance_sheet" and tf in BALANCE_SHEET_ZERO_FIELDS:
                row.setdefault(tf, 0)
            continue

        field_evidence.setdefault(
            tf,
            {
                "raw_label": normalize_text_cell(cs[0]),
                "raw_row": cs,
                "normalized_label": lbl,
                "raw_value": normalize_text_cell(cs[cur_idx]) if cur_idx is not None and cur_idx < len(cs) else "",
                "source_snippet": sec[:500],
            },
        )
        seen.setdefault(tf, vals)
        if tf not in row:
            row[tf] = normalize_field_value(tf, vals[0])

    row = enrich_table_fields(table, row, seen)
    return row, seen, field_evidence


def parse_md(md, table, fs):
    """解析Markdown表格"""
    row = {}
    sn = []
    am = build_alias(fs, table)
    seen = {}
    field_evidence = {}

    best_score = None
    core_candidates = []

    for raw_sec in sections(md, table):
        sec = standardize_table_section(raw_sec)
        sn.append(sec[:3000])

        sec_row, sec_seen, sec_evidence = parse_section_rows(sec, table, fs, am)

        if table == "core_performance_indicators_sheet":
            sec_score = score_core_candidate(sec_row, sec_seen, sec)
            core_candidates.append((sec_score, sec_row, sec_seen, sec_evidence))
            if best_score is None or sec_score > best_score:
                best_score = sec_score
                row = sec_row
                seen = sec_seen
                field_evidence = sec_evidence
            continue

        for tf, vals in sec_seen.items():
            seen.setdefault(tf, vals)
        for tf, ev in sec_evidence.items():
            field_evidence.setdefault(tf, ev)
        for tf, val in sec_row.items():
            if tf not in row:
                row[tf] = val

    if table == "core_performance_indicators_sheet" and core_candidates:
        safe_fill_fields = {
            "eps", "total_operating_revenue", "net_profit_10k_yuan",
            "net_asset_per_share", "operating_cf_per_share", "net_profit_excl_non_recurring",
        }
        for _, sec_row, sec_seen, sec_evidence in sorted(core_candidates, key=lambda x: x[0], reverse=True):
            for tf, val in sec_row.items():
                if tf in safe_fill_fields and tf not in row:
                    row[tf] = val
            for tf, vals in sec_seen.items():
                if tf in safe_fill_fields:
                    seen.setdefault(tf, vals)
            for tf, ev in sec_evidence.items():
                if tf in safe_fill_fields:
                    field_evidence.setdefault(tf, ev)

    row = enrich_table_fields(table, row, seen)
    return row, sn, field_evidence


# ============================================================
# 跨表指标计算 - cross_table.py
# ============================================================

QUARTER_SEQUENCE = {"Q1": 1, "HY": 2, "Q3": 3, "FY": 4}
MAX_REASONABLE_QOQ_ABS = 300  # 环比增长率阈值
DEFAULT_SHARES_OUTSTANDING = 373270285
MIN_REASONABLE_NAVPS = -100
MAX_REASONABLE_NAVPS = 1000


def report_period_sort_key(report_period):
    """报告期排序键"""
    if not report_period or len(report_period) < 6:
        return (0, 0)
    return (int(report_period[:4]), QUARTER_SEQUENCE.get(report_period[4:], 0))


def split_report_period(report_period):
    """拆分报告期"""
    if not report_period or len(report_period) < 6:
        return None, None
    return int(report_period[:4]), report_period[4:]


def build_history_lookup(parsed_history):
    """构建历史查询表"""
    history = {}
    for report_period, parsed in parsed_history.items():
        for table, (row, _) in parsed.items():
            history.setdefault(table, {})[report_period] = row
    return history


def previous_same_company_period(history_rows, current_period):
    """获取同公司上一期"""
    current_key = report_period_sort_key(current_period)
    candidates = [rp for rp in history_rows.keys() if report_period_sort_key(rp) < current_key]
    if not candidates:
        return None
    return sorted(candidates, key=report_period_sort_key)[-1]


def previous_period_in_year(report_period):
    """获取年内上一期"""
    year, period = split_report_period(report_period)
    if year is None:
        return None
    # FY是全年数据，没有年内上一期
    if period == "FY":
        return None
    mapping = {"HY": f"{year}Q1", "Q3": f"{year}HY"}
    return mapping.get(period)


def previous_quarter_period(report_period):
    """获取上一季度"""
    year, period = split_report_period(report_period)
    if year is None:
        return None
    mapping = {
        "Q1": f"{year - 1}FY",
        "HY": f"{year}Q1",
        "Q3": f"{year}HY",
        # FY的上一期应该是上年Q3（但FY和Q3是不同类型，需要判断数据类型）
        # 简化为：如果有Q3则用Q3，否则用FY
    }
    result = mapping.get(period)
    # 如果是FY，尝试找Q3
    if period == "FY" and result is None:
        prev_q3 = f"{year - 1}Q3"
        # 这里不返回，由调用方处理
        return prev_q3
    return result


def calc_growth_pct(curr, prev):
    """计算增长率（带限制）"""
    if curr is None or prev in (None, 0):
        return None
    growth = round((curr - prev) / abs(prev) * 100, 4)
    if abs(growth) > MAX_REASONABLE_QOQ_ABS:
        return None
    return growth


def get_table_row(history, table_name, report_period):
    """获取表行数据"""
    return history.get(table_name, {}).get(report_period, {})


def cumulative_value_for_period(history, parsed, table_name, field_name, report_period):
    """获取期间累计值"""
    current_period = parsed.get("__current_report_period__")
    if report_period == current_period:
        row = parsed.get(table_name, ({}, []))[0]
    else:
        row = get_table_row(history, table_name, report_period)
    return row.get(field_name)


def quarter_value(history, parsed, table_name, field_name, report_period):
    """获取季度值（从累计值计算单季度值）"""
    year, period = split_report_period(report_period)
    if year is None:
        return None

    cumulative = cumulative_value_for_period(history, parsed, table_name, field_name, report_period)
    if cumulative is None:
        return None
    
    # Q1 直接返回累计值
    if period == "Q1":
        return cumulative

    # 其他期需要减去上期累计值
    # Q1->Q1, HY->Q1, Q3->HY, FY->Q3
    prev_in_year = previous_period_in_year(report_period)
    
    if prev_in_year:
        prev_cumulative = cumulative_value_for_period(history, parsed, table_name, field_name, prev_in_year)
        if prev_cumulative is not None:
            return round(cumulative - prev_cumulative, 4)
    
    # 如果无法计算单季度值，返回累计值（用于年度报告的环比计算）
    return cumulative


def fallback_per_share(equity_total_equity):
    """计算每股净资产后备值"""
    if equity_total_equity is None:
        return None
    return round(equity_total_equity * 10000 / DEFAULT_SHARES_OUTSTANDING, 4)


def is_reasonable_per_share(value):
    """判断每股净资产是否合理"""
    return isinstance(value, (int, float)) and MIN_REASONABLE_NAVPS <= value <= MAX_REASONABLE_NAVPS


def enrich_cross_table_metrics(parsed, schema, parsed_history=None, current_report_period=""):
    """丰富跨表指标"""
    core = parsed.get("core_performance_indicators_sheet", ({}, []))[0]
    bal = parsed.get("balance_sheet", ({}, []))[0]
    inc = parsed.get("income_sheet", ({}, []))[0]
    cf = parsed.get("cash_flow_sheet", ({}, []))[0]

    if core.get("net_profit_10k_yuan") is None and inc.get("net_profit") is not None:
        core["net_profit_10k_yuan"] = inc["net_profit"]
    if core.get("total_operating_revenue") is None and inc.get("total_operating_revenue") is not None:
        core["total_operating_revenue"] = inc["total_operating_revenue"]
    if core.get("operating_revenue_yoy_growth") is None and inc.get("operating_revenue_yoy_growth") is not None:
        core["operating_revenue_yoy_growth"] = inc["operating_revenue_yoy_growth"]
    if core.get("net_profit_yoy_growth") is None and inc.get("net_profit_yoy_growth") is not None:
        core["net_profit_yoy_growth"] = inc["net_profit_yoy_growth"]

    if (
        core.get("gross_profit_margin") is None
        and inc.get("total_operating_revenue") not in (None, 0)
        and inc.get("operating_expense_cost_of_sales") is not None
    ):
        core["gross_profit_margin"] = round(
            (inc["total_operating_revenue"] - inc["operating_expense_cost_of_sales"])
            / inc["total_operating_revenue"] * 100,
            4,
        )

    if (
        core.get("net_profit_margin") is None
        and core.get("net_profit_10k_yuan") is not None
        and core.get("total_operating_revenue") not in (None, 0)
    ):
        core["net_profit_margin"] = round(
            core["net_profit_10k_yuan"] / core["total_operating_revenue"] * 100, 4
        )

    fallback_navps = fallback_per_share(bal.get("equity_total_equity"))
    if core.get("net_asset_per_share") is None and fallback_navps is not None:
        core["net_asset_per_share"] = fallback_navps
    elif fallback_navps is not None and not is_reasonable_per_share(core.get("net_asset_per_share")):
        core["net_asset_per_share"] = fallback_navps

    if core.get("operating_cf_per_share") is None and cf.get("operating_cf_net_amount") is not None:
        core["operating_cf_per_share"] = round(cf["operating_cf_net_amount"] * 10000 / DEFAULT_SHARES_OUTSTANDING, 4)

    if core.get("roe") is None and bal.get("equity_total_equity") not in (None, 0):
        base_profit = core.get("net_profit_10k_yuan")
        if base_profit is not None:
            core["roe"] = round(base_profit / bal["equity_total_equity"] * 100, 4)

    if (
        core.get("roe_weighted_excl_non_recurring") is None
        and core.get("net_profit_excl_non_recurring") is not None
        and bal.get("equity_total_equity") not in (None, 0)
    ):
        core["roe_weighted_excl_non_recurring"] = round(
            core["net_profit_excl_non_recurring"] / bal["equity_total_equity"] * 100, 4
        )

    if (
        bal.get("asset_liability_ratio") is None
        and bal.get("liability_total_liabilities") is not None
        and bal.get("asset_total_assets") not in (None, 0)
    ):
        bal["asset_liability_ratio"] = round(
            bal["liability_total_liabilities"] / bal["asset_total_assets"] * 100, 4
        )

    if (
        cf.get("operating_cf_ratio_of_net_cf") is None
        and cf.get("operating_cf_net_amount") is not None
        and cf.get("net_cash_flow") not in (None, 0)
    ):
        cf["operating_cf_ratio_of_net_cf"] = round(cf["operating_cf_net_amount"] / cf["net_cash_flow"] * 100, 4)

    if (
        cf.get("investing_cf_ratio_of_net_cf") is None
        and cf.get("investing_cf_net_amount") is not None
        and cf.get("net_cash_flow") not in (None, 0)
    ):
        cf["investing_cf_ratio_of_net_cf"] = round(cf["investing_cf_net_amount"] / cf["net_cash_flow"] * 100, 4)

    if (
        cf.get("financing_cf_ratio_of_net_cf") is None
        and cf.get("financing_cf_net_amount") is not None
        and cf.get("net_cash_flow") not in (None, 0)
    ):
        cf["financing_cf_ratio_of_net_cf"] = round(cf["financing_cf_net_amount"] / cf["net_cash_flow"] * 100, 4)

    current_period = current_report_period or core.get("report_period") or inc.get("report_period")
    if current_period:
        parsed["__current_report_period__"] = current_period
    if current_period and parsed_history:
        history = build_history_lookup(parsed_history)
        prev_period = previous_quarter_period(current_period)
        if prev_period:
            if core.get("operating_revenue_qoq_growth") is None:
                curr_rev = quarter_value(
                    history, parsed, "core_performance_indicators_sheet", "total_operating_revenue", current_period
                )
                prev_rev = quarter_value(
                    history, parsed, "core_performance_indicators_sheet", "total_operating_revenue", prev_period
                )
                if curr_rev is None:
                    curr_rev = quarter_value(history, parsed, "income_sheet", "total_operating_revenue", current_period)
                if prev_rev is None:
                    prev_rev = quarter_value(history, parsed, "income_sheet", "total_operating_revenue", prev_period)
                qoq = calc_growth_pct(curr_rev, prev_rev)
                if qoq is not None:
                    core["operating_revenue_qoq_growth"] = qoq
            if core.get("net_profit_qoq_growth") is None:
                curr_np = quarter_value(
                    history, parsed, "core_performance_indicators_sheet", "net_profit_10k_yuan", current_period
                )
                prev_np = quarter_value(
                    history, parsed, "core_performance_indicators_sheet", "net_profit_10k_yuan", prev_period
                )
                if curr_np is None:
                    curr_np = quarter_value(history, parsed, "income_sheet", "net_profit", current_period)
                if prev_np is None:
                    prev_np = quarter_value(history, parsed, "income_sheet", "net_profit", prev_period)
                qoq = calc_growth_pct(curr_np, prev_np)
                if qoq is not None:
                    core["net_profit_qoq_growth"] = qoq

    parsed.pop("__current_report_period__", None)
    for t, fs in schema.items():
        ext, sn = parsed[t]
        parsed[t] = (filter_to_schema(ext, fs), sn)

    return parsed


# ============================================================
# 质量检查 - quality.py
# ============================================================


def append_quality_issue(q, pdf_name, report_period, table, issue_type, message, details=None, evidence=None):
    """追加质量问题"""
    q.setdefault("issues", []).append(
        {
            "source_file_name": pdf_name,
            "report_period": report_period,
            "table": table,
            "issue_type": issue_type,
            "message": message,
            "details": details or {},
            "evidence": evidence or {},
        }
    )


def check_field_ranges(ext, table, pdf_name, report_period, q, evidence_map=None):
    """检查字段范围"""
    for field, (lo, hi) in QUALITY_LIMITS.get("field_ranges", {}).items():
        val = ext.get(field)
        if isinstance(val, (int, float)) and (val < lo or val > hi):
            append_quality_issue(
                q, pdf_name, report_period, table, "field_range",
                f"{field} 超出合理范围",
                {"field": field, "value": val, "range": [lo, hi]},
                evidence_map.get(field, {}) if evidence_map else {},
            )


def check_table_quality(ext, fs, table, pdf_name, report_period, q, evidence_map=None):
    """检查表格质量"""
    cov, hit, total = coverage(ext, fs)
    if cov < QUALITY_LIMITS.get("very_low_coverage", 0.2):
        append_quality_issue(
            q, pdf_name, report_period, table, "very_low_coverage",
            "表覆盖率很低",
            {"coverage": cov, "hit_fields": hit, "total_fields": total},
        )
    elif cov < QUALITY_LIMITS.get("low_coverage", 0.4):
        append_quality_issue(
            q, pdf_name, report_period, table, "low_coverage",
            "表覆盖率偏低",
            {"coverage": cov, "hit_fields": hit, "total_fields": total},
        )
    check_field_ranges(ext, table, pdf_name, report_period, q, evidence_map)
    return cov, hit, total


def check_cross_table_quality(parsed, pdf_name, report_period, q, evidence_by_table=None):
    """检查跨表质量"""
    tol = QUALITY_LIMITS.get("cross_table_rel_tolerance", 0.15)
    pairs = [
        ("core_performance_indicators_sheet", "total_operating_revenue", "income_sheet", "total_operating_revenue"),
        ("core_performance_indicators_sheet", "net_profit_10k_yuan", "income_sheet", "net_profit"),
        ("core_performance_indicators_sheet", "asset_total_assets", "balance_sheet", "asset_total_assets"),
    ]

    for lt, lf, rt, rf in pairs:
        lv = parsed.get(lt, ({}, []))[0].get(lf)
        rv = parsed.get(rt, ({}, []))[0].get(rf)
        if lv in (None, 0) or rv in (None, 0):
            continue
        rel = abs(lv - rv) / max(abs(lv), abs(rv))
        if rel > tol:
            append_quality_issue(
                q, pdf_name, report_period, lt, "cross_table_mismatch",
                f"{lf} 与 {rt}.{rf} 不一致",
                {
                    "left_value": lv,
                    "right_value": rv,
                    "relative_diff": round(rel, 4),
                },
                {
                    "left": (evidence_by_table or {}).get(lt, {}).get(lf, {}),
                    "right": (evidence_by_table or {}).get(rt, {}).get(rf, {}),
                },
            )

    if parsed.get("cash_flow_sheet", ({}, []))[0].get("net_cash_flow") not in (None, 0):
        cf = parsed["cash_flow_sheet"][0]
        ratio_sum = sum(
            cf.get(k, 0)
            for k in [
                "operating_cf_ratio_of_net_cf",
                "investing_cf_ratio_of_net_cf",
                "financing_cf_ratio_of_net_cf",
            ]
            if isinstance(cf.get(k), (int, float))
        )
        if ratio_sum and abs(ratio_sum - 100) > 50:
            append_quality_issue(
                q, pdf_name, report_period, "cash_flow_sheet",
                "cashflow_ratio_sum", "现金流占比合计明显异常",
                {"ratio_sum": round(ratio_sum, 4)},
            )


def dump_quality(out, q):
    """导出质量报告"""
    rep = {
        "generated_at": dt.datetime.now().isoformat(),
        "total_reports": q["ok"] + q["failed"] + q.get("skipped", 0),
        "ok_reports": q["ok"],
        "failed_reports": q["failed"],
        "skipped_reports": q.get("skipped", 0),
        "table_metrics": {},
        "issues_summary": {},
        "issues": q.get("issues", []),
    }

    for t, v in q["tables"].items():
        rep["table_metrics"][t] = {
            "reports": v["reports"],
            "avg_coverage": round(v["sum_cov"] / v["reports"], 4) if v["reports"] else 0.0,
            "hit_fields": v["sum_hit"],
            "total_fields": v["sum_total"],
        }

    for issue in q.get("issues", []):
        k = issue["issue_type"]
        rep["issues_summary"][k] = rep["issues_summary"].get(k, 0) + 1

    (out / "quality_report.json").write_text(
        json.dumps(rep, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (out / "quality_issues.json").write_text(
        json.dumps(q.get("issues", []), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ============================================================
# Markdown解析入库 - md_to_db.py (主要逻辑)
# ============================================================


def parse_md_to_tables(md, schema, q, pdf_name="", report_period="", parsed_history=None):
    """解析Markdown到表格数据"""
    parsed = {}
    evidence_by_table = {}

    for table_name, fields in schema.items():
        ext, snippets, evidence = parse_md(md, table_name, fields)
        ext = filter_to_schema(ext, fields)

        cov, hit, total = check_table_quality(
            ext, fields, table_name, pdf_name, report_period, q, evidence
        )
        q["tables"].setdefault(
            table_name,
            {"reports": 0, "sum_cov": 0.0, "sum_hit": 0, "sum_total": 0},
        )
        q["tables"][table_name]["reports"] += 1
        q["tables"][table_name]["sum_cov"] += cov
        q["tables"][table_name]["sum_hit"] += hit
        q["tables"][table_name]["sum_total"] += total

        parsed[table_name] = (ext, snippets)
        evidence_by_table[table_name] = evidence

    parsed = enrich_cross_table_metrics(
        parsed,
        schema,
        parsed_history=parsed_history,
        current_report_period=report_period,
    )
    check_cross_table_quality(parsed, pdf_name, report_period, q, evidence_by_table)
    return parsed


def write_tables_to_db(conn, meta, parsed):
    """写入表格到数据库"""
    base = {
        "serial_number": build_serial_number(meta.stock_code, meta.report_year, meta.report_period),
        "stock_code": meta.stock_code,
        "stock_abbr": meta.stock_abbr,
        "report_period": meta.report_period,
        "report_year": meta.report_year,
        "report_type": meta.report_type,
        "source_file_name": meta.source_file_name,
        "created_at": dt.datetime.now().isoformat(),
    }
    for table_name, (ext, snippets) in parsed.items():
        upsert(conn, table_name, {**base, **ext})
        # 写入snippets时先检查是否已存在（避免重复）
        for snippet in snippets[:3]:
            # 检查是否已存在相同的 snippet
            cur = conn.execute(
                "SELECT id FROM raw_table_snippets WHERE stock_code=? AND report_period=? AND target_table=? AND source_section=? AND snippet=?",
                (meta.stock_code, meta.report_period, table_name, table_name, snippet)
            )
            if cur.fetchone() is None:
                conn.execute(
                    "INSERT INTO raw_table_snippets(stock_code,report_period,target_table,source_section,snippet,created_at) VALUES (?,?,?,?,?,?)",
                    (
                        meta.stock_code,
                        meta.report_period,
                        table_name,
                        table_name,
                        snippet,
                        dt.datetime.now().isoformat(),
                    ),
                )


def rebuild_stock_qoq_metrics(conn, schema, stock_code, stock_period_history):
    """重建股票的环比指标
    
    关键：只重建环比相关字段，不覆盖其他已有数据
    """
    ordered_periods = sorted(stock_period_history.keys(), key=report_period_sort_key)
    rolling_history = {}

    for report_period in ordered_periods:
        # 获取当前期别的原始数据（不是enriched的）
        parsed = {
            table_name: (dict(ext), list(snippets))
            for table_name, (ext, snippets) in stock_period_history[report_period].items()
        }
        
        # 使用rolling_history计算环比指标
        enriched = enrich_cross_table_metrics(
            parsed,
            schema,
            parsed_history=rolling_history,
            current_report_period=report_period,
        )

        # 只更新环比相关字段，不覆盖其他数据
        # 从数据库获取已有记录
        qoq_fields = ['operating_revenue_qoq_growth', 'net_profit_qoq_growth']
        
        for table_name, (ext, _) in enriched.items():
            # 获取数据库中该期别的现有数据
            cur = conn.execute(
                f'SELECT * FROM "{table_name}" WHERE stock_code=? AND report_period=?',
                (stock_code, report_period)
            ).fetchone()
            
            if cur:
                existing = dict(zip([d[0] for d in conn.execute(f'SELECT * FROM "{table_name}" WHERE 1=0').description], cur))
                
                # 只更新环比字段（如果enriched中有值）
                update_fields = {}
                for field in qoq_fields:
                    if field in ext and ext[field] is not None:
                        update_fields[field] = ext[field]
                
                # 同时更新从其他表复制过来的字段
                if table_name == "core_performance_indicators_sheet":
                    # 从income_sheet复制过来的字段
                    inc_ext = enriched.get("income_sheet", ({}, []))[0]
                    for field in ['operating_revenue_yoy_growth', 'net_profit_yoy_growth']:
                        if field in inc_ext and inc_ext[field] is not None:
                            # 只有当现有值为空时才更新
                            if existing.get(field) is None:
                                update_fields[field] = inc_ext[field]
                
                if update_fields:
                    set_clause = ", ".join([f'"{k}"=?' for k in update_fields.keys()])
                    conn.execute(
                        f'UPDATE "{table_name}" SET {set_clause} WHERE stock_code=? AND report_period=?',
                        list(update_fields.values()) + [stock_code, report_period]
                    )
        
        # 更新rolling_history（使用enriched后的数据）
        rolling_history[report_period] = {
            table_name: (dict(ext), list(snippets))
            for table_name, (ext, snippets) in enriched.items()
        }


def process_pdf_to_db(
    pdf,
    md_cache_root,
    conn,
    schema,
    company_map,
    quality_stats,
    parsed_history=None,
):
    """处理PDF到数据库"""
    if is_summary_name(pdf):
        print("    [skip] summary detected by filename")
        log(conn, pdf.name, "skipped", "文件名命中摘要规则，跳过")
        quality_stats.setdefault("skipped", 0)
        quality_stats["skipped"] += 1
        return None, None

    cache_dir = md_cache_root / pdf.stem
    md_path = cache_dir / "report.md"
    if not md_path.exists():
        log(conn, pdf.name, "failed", "未找到缓存Markdown")
        quality_stats["failed"] += 1
        return None, None

    md = md_path.read_text(encoding="utf-8", errors="ignore")
    if detect_summary_from_md(md):
        print("    [skip] summary detected by markdown content")
        log(conn, pdf.name, "skipped", "Markdown内容判定为摘要，跳过")
        quality_stats.setdefault("skipped", 0)
        quality_stats["skipped"] += 1
        return None, None

    meta = detect_meta(pdf, company_map, md)
    print(
        f"    [阶段2/4] 识别元数据：stock_code={meta.stock_code or '-'}, report_period={meta.report_period or '-'}"
    )

    if not meta.report_period or not meta.stock_code:
        log(conn, pdf.name, "failed", "元数据识别失败")
        quality_stats["failed"] += 1
        return meta, None

    print("    [阶段3/4] 解析Markdown：开始")
    stock_history = parsed_history.get(meta.stock_code, {}) if parsed_history else {}
    parsed = parse_md_to_tables(
        md,
        schema,
        quality_stats,
        pdf.name,
        meta.report_period,
        parsed_history=stock_history,
    )
    hit_total = sum(len(ext) for ext, _ in parsed.values())
    print(f"    [阶段3/4] 解析Markdown：完成, 命中字段={hit_total}")

    print("    [阶段4/4] 写入数据库：开始")
    write_tables_to_db(conn, meta, parsed)
    print("    [阶段4/4] 写入数据库：完成")

    if hit_total:
        log(conn, pdf.name, "ok", "处理完成")
    else:
        log(conn, pdf.name, "warning", "仅写入元数据，未命中业务字段")
    quality_stats["ok"] += 1
    return meta, parsed


def init_db(output_root, schema):
    """初始化数据库"""
    conn = sqlite3.connect(output_root / "financial.db")
    ensure_tables(conn, schema)
    return conn


def parse_schema(x: Path):
    """解析Schema Excel"""
    wb = openpyxl.load_workbook(x, data_only=True)
    r = {}
    for cn, t in TABLE_NAME_MAP.items():
        fs = []
        for row in wb[cn].iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                fs.append(
                    F(
                        str(row[0]).strip(),
                        str(row[1]).strip() if row[1] else str(row[0]).strip(),
                        str(row[2]).strip() if row[2] else "TEXT",
                    )
                )
        r[t] = fs
    return r


def load_company_map(x: Path):
    """加载公司映射"""
    wb = openpyxl.load_workbook(x, data_only=True)
    ws = wb[wb.sheetnames[0]]
    h = [
        str(c.value).replace("\xa0", "").strip() if c.value else ""
        for c in next(ws.iter_rows(min_row=1, max_row=1))
    ]
    ci = next((i for i, v in enumerate(h) if "股票代码" in v or "A股代码" in v), None)
    ai = next((i for i, v in enumerate(h) if "股票简称" in v or "A股简称" in v), None)
    ni = next((i for i, v in enumerate(h) if "公司名称" in v), None)
    if ci is None:
        raise ValueError("附件1未找到股票代码列")
    m = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if ci < len(row) and row[ci] is not None:
            c = re.sub(r"\.0$", "", str(row[ci]).strip())
            c = c.zfill(6) if c.isdigit() else c
            a = (
                str(row[ai]).strip()
                if ai is not None and ai < len(row) and row[ai]
                else (
                    str(row[ni]).strip()
                    if ni is not None and ni < len(row) and row[ni]
                    else ""
                )
            )
            m[c] = a
    return m
