# -*- coding: utf-8 -*-
"""
财务报告 Markdown 解析入库工具 (整合版)
将所有模块整合为一个文件，保持原有功能不变
"""

import datetime as dt
import html
import json
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl


# ============================================================
# 数据类型定义 (data_types.py)
# ============================================================

@dataclass
class Field:
    """字段定义"""
    field_name: str
    cn_name: str
    data_type: str


@dataclass
class Meta:
    """元数据"""
    stock_code: str
    stock_abbr: str
    report_year: int
    report_type: str
    report_period: str
    source_file_name: str


# ============================================================
# 解析器配置 (parser_config.py)
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
        "current": ["本报告期末", "本报告期", "本期", "年初至报告期末"],
        "previous": ["上年同期", "上年度末", "上期"],
        "yoy": ["增减", "变动幅度", "同比"],
    },
    "income_sheet": {
        "current": ["本期", "本报告期", "年初至报告期末", "年初至报告期期末"],
        "previous": ["上年同期", "上期"],
    },
    "balance_sheet": {
        "current": ["本报告期末", "期末", "期末数"],
        "previous": ["上年度末", "年初", "期初"],
    },
    "cash_flow_sheet": {
        "current": ["本期", "本报告期", "本报告期", "年初至报告期末", "年初至报告期期末"],
        "previous": ["上年同期", "上期", "上年半年度"],
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
# 公共函数和常量 (common.py)
# ============================================================

RE_NORMALIZE_WHITESPACE = re.compile(r"\s+")
RE_REMOVE_UNITS = re.compile(r"(万元|亿元|元|人民币|%)")
RE_CLEAN_FIELD = re.compile(r"[\s\-_:：·,，。()（）\[\]【】]")
RE_NUMBER_CLEAN = re.compile(r",")
RE_HTML_TAG = re.compile(r"<[^>]+>")
YEAR_LIKE_RE = re.compile(r"20\d{2}年(?:末|度|第[一二三四]季度)?|20\d{2}年\d{1,2}月\d{1,2}日")

ROW_LABEL_PATTERNS = [
    (re.compile(r"^其中[:：]"), ""),
    (re.compile(r"^[加减][:：]"), ""),
    (re.compile(r"^[（(]?[一二三四五六七八九十0-9]+[)）]?[、.:：]\s*"), ""),
    (re.compile(r"^[（(][一二三四五六七八九十0-9]+[)）]\s*"), ""),
    (re.compile(r"（[^）]*填列）|\([^)]*填列\)"), ""),
    (re.compile(r"（损失以[^）]*）|\(损失以[^)]*\)"), ""),
    (re.compile(r"（亏损以[^）]*）|\(亏损以[^)]*\)"), ""),
    (re.compile(r"（净亏损以[^）]*）|\(净亏损以[^)]*\)"), ""),
    (re.compile(r"（亏损总额以[^）]*）|\(亏损总额以[^)]*\)"), ""),
    (re.compile(r"（元/股）|\(元/股\)|（元／股）|\(元／股\)"), ""),
    (re.compile(r"（元）|\(元\)|（万元）|\(万元\)|（%）|\(%\)"), ""),
]

INVALID_NUMERIC_TEXTS = {"--", "-", "不适用", "N/A", "nan", ""}

CUMULATIVE_FIELDS = {
    "total_operating_revenue", "net_profit_10k_yuan", "net_profit",
    "net_profit_excl_non_recurring", "operating_cf_net_amount",
    "investing_cf_net_amount", "financing_cf_net_amount",
    "net_cash_flow", "operating_cf_cash_from_sales",
    "other_income", "operating_profit", "total_profit",
    "asset_impairment_loss", "credit_impairment_loss",
    "total_operating_expenses", "operating_expense_cost_of_sales",
    "operating_expense_selling_expenses", "operating_expense_administrative_expenses",
    "operating_expense_financial_expenses", "operating_expense_rnd_expenses",
    "operating_expense_taxes_and_surcharges",
    "investing_cf_cash_for_investments", "investing_cf_cash_from_investment_recovery",
    "financing_cf_cash_from_borrowing", "financing_cf_cash_for_debt_repayment",
}

PERIOD_END_FIELDS = {
    "asset_total_assets", "liability_total_liabilities",
    "equity_total_equity", "equity_unappropriated_profit",
    "asset_cash_and_cash_equivalents", "asset_accounts_receivable",
    "asset_inventory", "asset_trading_financial_assets",
    "asset_construction_in_progress", "liability_accounts_payable",
    "liability_advance_from_customers", "liability_contract_liabilities",
    "liability_short_term_loans", "liability_taxes_payable",
    "liability_long_term_borrowings",
}

RATIO_FIELDS = {
    "eps", "roe", "roe_weighted_excl_non_recurring",
    "net_asset_per_share", "operating_cf_per_share",
    "gross_profit_margin", "net_profit_margin",
    "asset_liability_ratio",
    "operating_revenue_yoy_growth", "net_profit_yoy_growth",
    "operating_revenue_qoq_growth", "net_profit_qoq_growth",
    "net_profit_excl_non_recurring_yoy",
    "net_cash_flow_yoy_growth",
    "asset_total_assets_yoy_growth", "liability_total_liabilities_yoy_growth",
    "operating_cf_ratio_of_net_cf", "investing_cf_ratio_of_net_cf",
    "financing_cf_ratio_of_net_cf",
}

PER_SHARE_FIELDS = {"eps", "net_asset_per_share", "operating_cf_per_share"}

ROE_FIELDS = {"roe", "roe_weighted_excl_non_recurring"}

BALANCE_SHEET_ZERO_FIELDS = {
    "liability_advance_from_customers", "liability_contract_liabilities",
    "liability_short_term_loans", "liability_accounts_payable",
    "asset_accounts_receivable", "asset_inventory",
}

FIELD_REASONABLE_LIMITS = {"eps": (-10, 20)}

MAX_REASONABLE_GROWTH_ABS = 500
MAX_REASONABLE_RATIO_ABS = 100

QUALITY_LIMITS = {
    "field_ranges": {
        "eps": [-1000, 1000], "roe": [-100, 100],
        "gross_profit_margin": [-100, 100], "net_profit_margin": [-100, 100],
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

DEFAULT_SHARES_OUTSTANDING = 373270285
MIN_REASONABLE_NAVPS = -100
MAX_REASONABLE_NAVPS = 1000
MAX_REASONABLE_QOQ_ABS = 300
QUARTER_SEQUENCE = {"Q1": 1, "HY": 2, "Q3": 3, "FY": 4}


def normalize_text_cell(s) -> str:
    """标准化单元格文本"""
    s = html.unescape(str(s or ""))
    s = s.replace("\xa0", " ").replace("−", "-").replace("—", "-").replace("–", "-")
    return RE_NORMALIZE_WHITESPACE.sub(" ", s).strip()


def normalize_key(s) -> str:
    """标准化关键字段（用于匹配）"""
    s = (s or "").strip().replace("（", "(").replace("）", ")").replace("－", "-")
    s = RE_CLEAN_FIELD.sub("", s)
    return RE_REMOVE_UNITS.sub("", s)


def clean_num(s, factor: float = 1.0) -> Optional[float]:
    """解析数字"""
    t = normalize_text_cell(s)
    if not t or t in INVALID_NUMERIC_TEXTS:
        return None
    t = RE_NUMBER_CLEAN.sub("", t)
    t = re.sub(r"\s+", "", t)
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1]
    if t.endswith("%"):
        try:
            return -float(t[:-1]) if neg else float(t[:-1])
        except ValueError:
            return None
    try:
        v = float(t) * factor
        return -v if neg else v
    except ValueError:
        return None


def row_label(cs: List[str]) -> str:
    """获取行标签"""
    if not cs:
        return ""
    s = normalize_text_cell(cs[0])
    for pattern, replacement in ROW_LABEL_PATTERNS:
        s = pattern.sub(replacement, s)
    return normalize_key(s)


def calc_growth(curr: Optional[float], prev: Optional[float], field_name: str = None) -> Optional[float]:
    """计算增长率"""
    if curr is None or prev in (None, 0):
        return None
    growth = round((curr - prev) / abs(prev) * 100, 4)
    return growth


def detect_unit(t: str) -> float:
    """检测表格单位"""
    prefix = t.split("<table>", 1)[0] if "<table>" in t else t[:500]
    hits = re.findall(r"单位[:：]\s*(亿元|万元|元)", prefix)
    if not hits:
        table_part = t.split("<table>", 1)[1] if "<table>" in t else t
        hits = re.findall(r"单位[:：]\s*(亿元|万元|元)", table_part[:300])
    unit = hits[-1] if hits else "元"
    if unit == "亿元":
        return 100000000.0
    if unit == "万元":
        return 10000.0
    return 1.0


def is_header_row(cs: List[str]) -> bool:
    """判断是否为表头行"""
    if not cs:
        return False
    joined = " ".join(cs)
    if row_label(cs) in {
        "主要会计数据", "主要财务指标", "项目",
        "流动资产", "非流动资产", "流动负债", "非流动负债",
        "所有者权益或股东权益",
    }:
        return True
    if not str(cs[0]).strip():
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


def get_field_type(field_name: str) -> str:
    """获取字段类型"""
    if field_name in CUMULATIVE_FIELDS:
        return "cumulative"
    elif field_name in PERIOD_END_FIELDS:
        return "period_end"
    elif field_name in RATIO_FIELDS:
        return "ratio"
    else:
        return "other"


def is_plausible_growth(v: float) -> bool:
    """判断增长率是否合理"""
    return isinstance(v, (int, float)) and abs(v) <= MAX_REASONABLE_GROWTH_ABS


def is_plausible_ratio(v: float) -> bool:
    """判断比率是否合理"""
    return isinstance(v, (int, float)) and abs(v) <= MAX_REASONABLE_RATIO_ABS


# ============================================================
# 数据校验模块 (data_validation.py)
# 负责入库前的多维度数据校验
# ============================================================

from enum import IntEnum
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Tuple
from datetime import datetime


class ValidationLevel(IntEnum):
    """校验级别"""
    PASS = 0       # 通过
    WARNING = 1    # 警告（轻微问题，自动修正）
    ERROR = 2      # 错误（异常问题，标记但仍入库）
    BLOCK = 3      # 阻断（严重问题，拒绝入库）


class ValidationDimension(IntEnum):
    """校验维度"""
    CROSS_TABLE = 1       # 跨表一致性
    FIELD_REASONABLE = 2  # 字段合理性
    COMPLETENESS = 3       # 完整性
    TIME_SEQUENCE = 4     # 时序一致性
    FORMAT = 5            # 格式校验
    BUSINESS_LOGIC = 6    # 业务逻辑
    CUMULATIVE_RELATION = 7  # 累计与单季度关系


@dataclass
class ValidationIssue:
    """校验问题"""
    dimension: ValidationDimension
    level: ValidationLevel
    table: str
    field: str
    message: str
    value: Any = None
    expected: Any = None
    actual: Any = None
    suggestion: str = ""
    auto_fixed: bool = False
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())


@dataclass
class ValidationResult:
    """校验结果"""
    passed: bool
    issues: List[ValidationIssue] = field(default_factory=list)
    auto_fixed_count: int = 0
    warning_count: int = 0
    error_count: int = 0
    block_count: int = 0

    def add_issue(self, issue: ValidationIssue):
        self.issues.append(issue)
        if issue.auto_fixed:
            self.auto_fixed_count += 1
        if issue.level == ValidationLevel.WARNING:
            self.warning_count += 1
        elif issue.level == ValidationLevel.ERROR:
            self.error_count += 1
        elif issue.level == ValidationLevel.BLOCK:
            self.block_count += 1

    def can_continue(self) -> bool:
        """是否可以继续处理"""
        return self.block_count == 0


# ============================================================
# 校验配置
# ============================================================

VALIDATION_CONFIG = {
    # 跨表一致性容差（相对��差）
    "cross_table_tolerance": 0.15,  # 15%
    "strict_tolerance": 0.05,      # 严格5%
    "loose_tolerance": 0.25,      # 宽松25%

    # 字段合理范围
    "field_ranges": {
        "eps": [-10, 20],
        "roe": [-100, 100],
        "roe_weighted_excl_non_recurring": [-100, 100],
        "gross_profit_margin": [-100, 100],
        "net_profit_margin": [-100, 100],
        "asset_liability_ratio": [0, 100],
        "operating_cf_ratio_of_net_cf": [-200, 200],
        "investing_cf_ratio_of_net_cf": [-200, 200],
        "financing_cf_ratio_of_net_cf": [-200, 200],
        "operating_revenue_yoy_growth": [-500, 500],
        "net_profit_yoy_growth": [-500, 500],
        "operating_revenue_qoq_growth": [-500, 500],
        "net_profit_qoq_growth": [-500, 500],
        "net_profit_excl_non_recurring_yoy": [-500, 500],
        "net_cash_flow_yoy_growth": [-500, 500],
        "asset_total_assets_yoy_growth": [-500, 500],
        "liability_total_liabilities_yoy_growth": [-500, 500],
        "net_asset_per_share": [-100, 1000],
        "operating_cf_per_share": [-100, 100],
    },

    # 业务逻辑限制
    "business_rules": {
        "net_profit_max_ratio": 1.5,  # 净利润不超过营业收入的150%
        "negative_allowed_fields": [],  # 允许负值的字段
        "must_positive_fields": [
            "asset_total_assets", "equity_total_equity",
            "asset_cash_and_cash_equivalents", "net_asset_per_share"
        ],
    },

    # 累计关系校验
    "cumulative_rules": {
        "Q1_min": 0,
        "HY_min": 0,
        "Q3_min": 0,
        "FY_min": 0,
    },

    # 格式校验
    "format_rules": {
        "stock_code_length": 6,
        "year_range": (1990, 2030),
        "valid_periods": ["FY", "Q1", "HY", "Q3"],
    },

    # 阻断阈值
    "block_threshold": {
        "missing_critical_fields": True,  # 缺少关键字段是否阻断
        "balance_sheet_unbalanced": True,  # 资产负债表不平衡是否阻断
        "duplicate_record": True,          # 重复记录是否阻断
    },
}


# ============================================================
# 工具函数
# ============================================================

def rel_diff(v1: float, v2: float) -> float:
    """计算相对差值"""
    if v1 is None or v2 is None:
        return None
    if v1 == 0 and v2 == 0:
        return 0.0
    if v1 == 0 or v2 == 0:
        return 1.0
    return abs(v1 - v2) / max(abs(v1), abs(v2))


def is_reasonable_value(val: Any, lo: float, hi: float) -> bool:
    """判断值是否在合理范围内"""
    if val is None:
        return True
    if not isinstance(val, (int, float)):
        return False
    return lo <= val <= hi


def safe_div(a: float, b: float, default: float = None) -> float:
    """安全除法"""
    if b is None or b == 0:
        return default
    return a / b if isinstance(a, (int, float)) else default


# ============================================================
# 1. 跨表一致性校验
# ============================================================

def validate_cross_table_consistency(
    parsed: Dict[str, Any],
    config: Dict = None
) -> List[ValidationIssue]:
    """跨表一致性校验"""
    if config is None:
        config = VALIDATION_CONFIG
    tolerance = config.get("cross_table_tolerance", 0.15)

    issues = []
    core = parsed.get("core_performance_indicators_sheet", ({}, []))[0]
    bal = parsed.get("balance_sheet", ({}, []))[0]
    inc = parsed.get("income_sheet", ({}, []))[0]
    cf = parsed.get("cash_flow_sheet", ({}, []))[0]

    # 1.1 营业收入一致性
    core_rev = core.get("total_operating_revenue")
    inc_rev = inc.get("total_operating_revenue")
    if core_rev and inc_rev:
        rd = rel_diff(core_rev, inc_rev)
        if rd is not None and rd > tolerance:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.CROSS_TABLE,
                level=ValidationLevel.WARNING,
                table="core/income",
                field="total_operating_revenue",
                message=f"核心指标表与利润表的营业收入不一致，相对差={rd:.2%}",
                actual=core_rev,
                expected=inc_rev,
            ))

    # 1.2 净利润一致性
    core_np = core.get("net_profit_10k_yuan")
    inc_np = inc.get("net_profit")
    if core_np and inc_np:
        rd = rel_diff(core_np, inc_np)
        if rd is not None and rd > tolerance:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.CROSS_TABLE,
                level=ValidationLevel.WARNING,
                table="core/income",
                field="net_profit",
                message=f"核心指标表与利润表的净利润不一致，相对差={rd:.2%}",
                actual=core_np,
                expected=inc_np,
            ))

    # 1.3 总资产一致性
    core_ta = core.get("asset_total_assets")
    bal_ta = bal.get("asset_total_assets")
    if core_ta and bal_ta:
        rd = rel_diff(core_ta, bal_ta)
        if rd is not None and rd > tolerance:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.CROSS_TABLE,
                level=ValidationLevel.WARNING,
                table="core/balance",
                field="asset_total_assets",
                message=f"核心指标表与资产负债表的资产不一致，相对差={rd:.2%}",
                actual=core_ta,
                expected=bal_ta,
            ))

    # 1.4 资产负债表平衡验证
    # 资产总计 = 负债合计 + 所有者权益合计
    total_assets = bal.get("asset_total_assets")
    total_liab = bal.get("liability_total_liabilities")
    total_equity = bal.get("equity_total_equity")
    if all(v is not None for v in [total_assets, total_liab, total_equity]):
        calculated_assets = total_liab + total_equity
        rd = rel_diff(total_assets, calculated_assets)
        if rd is not None and rd > tolerance:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.CROSS_TABLE,
                level=ValidationLevel.ERROR,
                table="balance_sheet",
                field="asset_total_assets",
                message=f"资产负债表不平衡：资产总计={total_assets}，负债+权益={calculated_assets}，差={total_assets - calculated_assets:.4f}",
                actual=total_assets,
                expected=calculated_assets,
            ))

    # 1.5 现金流占比合理性
    net_cf = cf.get("net_cash_flow")
    if net_cf and net_cf != 0:
        for field_name, ratio_field in [
            ("operating_cf_net_amount", "operating_cf_ratio_of_net_cf"),
            ("investing_cf_net_amount", "investing_cf_ratio_of_net_cf"),
            ("financing_cf_net_amount", "financing_cf_ratio_of_net_cf"),
        ]:
            cf_val = cf.get(field_name)
            ratio_val = cf.get(ratio_field)
            if cf_val is not None and ratio_val is not None:
                calculated_ratio = (cf_val / net_cf) * 100
                rd = rel_diff(calculated_ratio, ratio_val)
                if rd is not None and rd > tolerance:
                    issues.append(ValidationIssue(
                        dimension=ValidationDimension.CROSS_TABLE,
                        level=ValidationLevel.WARNING,
                        table="cash_flow_sheet",
                        field=ratio_field,
                        message=f"现金流占比计算不一致：{ratio_field}={ratio_val}，计算值={calculated_ratio:.2f}%",
                        actual=ratio_val,
                        expected=calculated_ratio,
                    ))

    return issues


# ============================================================
# 2. 字段合理性校验
# ============================================================

def validate_field_reasonableness(
    parsed: Dict[str, Any],
    config: Dict = None
) -> List[ValidationIssue]:
    """字段合理性校验"""
    if config is None:
        config = VALIDATION_CONFIG

    issues = []
    ranges = config.get("field_ranges", {})

    for table_name, (data, _) in parsed.items():
        for field_name, value in data.items():
            if field_name in META_FIELDS:
                continue

            # 检查是否在合理范围内
            if field_name in ranges:
                lo, hi = ranges[field_name]
                if not is_reasonable_value(value, lo, hi):
                    issues.append(ValidationIssue(
                        dimension=ValidationDimension.FIELD_REASONABLE,
                        level=ValidationLevel.WARNING,
                        table=table_name,
                        field=field_name,
                        message=f"字段值超出合理范围：{field_name}={value}，合理范围=[{lo}, {hi}]",
                        value=value,
                        expected=f"[{lo}, {hi}]",
                    ))

    return issues


# ============================================================
# 3. 完整性校验
# ============================================================

def validate_completeness(
    meta: Meta,
    parsed: Dict[str, Any],
    company_map: Dict[str, str],
    config: Dict = None
) -> List[ValidationIssue]:
    """完整性校验"""
    if config is None:
        config = VALIDATION_CONFIG

    issues = []

    # 3.1 必填字段检查
    if not meta.stock_code:
        issues.append(ValidationIssue(
            dimension=ValidationDimension.COMPLETENESS,
            level=ValidationLevel.BLOCK,
            table="meta",
            field="stock_code",
            message="股票代码为空",
        ))

    if not meta.report_period:
        issues.append(ValidationIssue(
            dimension=ValidationDimension.COMPLETENESS,
            level=ValidationLevel.BLOCK,
            table="meta",
            field="report_period",
            message="报告期为空",
        ))

    if not meta.report_year or meta.report_year == 0:
        issues.append(ValidationIssue(
            dimension=ValidationDimension.COMPLETENESS,
            level=ValidationLevel.ERROR,
            table="meta",
            field="report_year",
            message="报告年份为空或无效",
        ))

    # 3.2 公司映射检查
    if meta.stock_code and meta.stock_code not in company_map:
        issues.append(ValidationIssue(
            dimension=ValidationDimension.COMPLETENESS,
            level=ValidationLevel.WARNING,
            table="meta",
            field="stock_code",
            message=f"股票代码 {meta.stock_code} 不在已知公司列表中",
            actual=meta.stock_code,
        ))

    # 3.3 关键字段完整性
    core = parsed.get("core_performance_indicators_sheet", ({}, []))[0]
    critical_fields = {
        "total_operating_revenue": "营业收入",
        "net_profit_10k_yuan": "净利润",
        "asset_total_assets": "总资产",
    }

    for field, name in critical_fields.items():
        if field not in core or core[field] is None:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.COMPLETENESS,
                level=ValidationLevel.WARNING,
                table="core_performance_indicators_sheet",
                field=field,
                message=f"关键字段缺失：{name}",
            ))

    # 3.4 表覆盖率检查
    for table_name, (data, _) in parsed.items():
        if len(data) < 3:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.COMPLETENESS,
                level=ValidationLevel.WARNING,
                table=table_name,
                field="*",
                message=f"表数据过少，仅有 {len(data)} 个字段",
            ))

    return issues


# ============================================================
# 4. 时序一致性校验
# ============================================================

def validate_time_sequence(
    meta: Meta,
    conn: sqlite3.Connection,
    config: Dict = None
) -> List[ValidationIssue]:
    """时序一致性校验"""
    if config is None:
        config = VALIDATION_CONFIG

    issues = []

    # 4.1 检查同一公司同一报告期是否已有记录
    if meta.stock_code and meta.report_year and meta.report_period:
        cur = conn.execute(
            'SELECT serial_number FROM "core_performance_indicators_sheet" WHERE stock_code=? AND report_year=? AND report_period=?',
            (meta.stock_code, meta.report_year, meta.report_period)
        ).fetchone()

        if cur:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.TIME_SEQUENCE,
                level=ValidationLevel.WARNING,
                table="core_performance_indicators_sheet",
                field="serial_number",
                message=f"发现重复记录：股票={meta.stock_code}，年份={meta.report_year}，期别={meta.report_period}",
                actual=cur[0],
            ))

    # 4.2 报告期格式校验
    valid_periods = config.get("format_rules", {}).get("valid_periods", [])
    if meta.report_period and meta.report_period not in valid_periods:
        issues.append(ValidationIssue(
            dimension=ValidationDimension.TIME_SEQUENCE,
            level=ValidationLevel.ERROR,
            table="meta",
            field="report_period",
            message=f"报告期格式无效：{meta.report_period}，期望={valid_periods}",
            actual=meta.report_period,
        ))

    # 4.3 年份范围校验
    year_range = config.get("format_rules", {}).get("year_range", (1990, 2030))
    if meta.report_year and (meta.report_year < year_range[0] or meta.report_year > year_range[1]):
        issues.append(ValidationIssue(
            dimension=ValidationDimension.TIME_SEQUENCE,
            level=ValidationLevel.ERROR,
            table="meta",
            field="report_year",
            message=f"报告年份超出范围：{meta.report_year}，期望范围={year_range}",
            actual=meta.report_year,
        ))

    return issues


# ============================================================
# 5. 格式校验
# ============================================================

def validate_format(
    meta: Meta,
    config: Dict = None
) -> List[ValidationIssue]:
    """格式校验"""
    if config is None:
        config = VALIDATION_CONFIG

    issues = []

    # 5.1 股票代码格式
    if meta.stock_code:
        if not meta.stock_code.isdigit():
            issues.append(ValidationIssue(
                dimension=ValidationDimension.FORMAT,
                level=ValidationLevel.ERROR,
                table="meta",
                field="stock_code",
                message=f"股票代码格式错误（应为数字）：{meta.stock_code}",
                actual=meta.stock_code,
            ))
        elif len(meta.stock_code) != 6:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.FORMAT,
                level=ValidationLevel.ERROR,
                table="meta",
                field="stock_code",
                message=f"股票代码长度错误（应为6位）：{meta.stock_code}",
                actual=meta.stock_code,
            ))

    return issues


# ============================================================
# 6. 业务逻辑校验
# ============================================================

def validate_business_logic(
    parsed: Dict[str, Any],
    config: Dict = None
) -> List[ValidationIssue]:
    """业务逻辑校验"""
    if config is None:
        config = VALIDATION_CONFIG

    issues = []
    core = parsed.get("core_performance_indicators_sheet", ({}, []))[0]
    inc = parsed.get("income_sheet", ({}, []))[0]
    bal = parsed.get("balance_sheet", ({}, []))[0]
    business_rules = config.get("business_rules", {})

    # 6.1 净利润不超过营业收入
    if core.get("total_operating_revenue") and core.get("net_profit_10k_yuan"):
        rev = core["total_operating_revenue"]
        profit = core["net_profit_10k_yuan"]
        max_ratio = business_rules.get("net_profit_max_ratio", 1.5)
        if profit > rev * max_ratio:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.BUSINESS_LOGIC,
                level=ValidationLevel.WARNING,
                table="core_performance_indicators_sheet",
                field="net_profit_10k_yuan",
                message=f"净利润异常大于营业收入：净利润={profit:.4f}，营业收入={rev:.4f}，比例={profit/rev:.2%}",
                actual=profit,
                expected=f"<= {rev * max_ratio:.4f}",
            ))

    # 6.2 必须为正的字段检查
    must_positive = business_rules.get("must_positive_fields", [])
    for field in must_positive:
        for table_name, (data, _) in parsed.items():
            if field in data and data[field] is not None and data[field] < 0:
                issues.append(ValidationIssue(
                    dimension=ValidationDimension.BUSINESS_LOGIC,
                    level=ValidationLevel.WARNING,
                    table=table_name,
                    field=field,
                    message=f"{field} 不应为负值：{data[field]}",
                    actual=data[field],
                ))

    # 6.3 每股指标合理性
    if core.get("eps") and core.get("net_profit_10k_yuan"):
        eps = core["eps"]
        profit = core["net_profit_10k_yuan"]
        if abs(eps) > 50:  # EPS异常大
            issues.append(ValidationIssue(
                dimension=ValidationDimension.BUSINESS_LOGIC,
                level=ValidationLevel.WARNING,
                table="core_performance_indicators_sheet",
                field="eps",
                message=f"每股收益异常：EPS={eps}元",
                actual=eps,
            ))

    # 6.4 ROE与净利润/净资产的关系验证
    if core.get("roe") and core.get("net_profit_10k_yuan") and bal.get("equity_total_equity"):
        profit = core["net_profit_10k_yuan"]
        equity = bal["equity_total_equity"]
        calculated_roe = (profit / equity) * 100 if equity else None
        actual_roe = core["roe"]

        if calculated_roe and abs(actual_roe - calculated_roe) > 5:
            issues.append(ValidationIssue(
                dimension=ValidationDimension.BUSINESS_LOGIC,
                level=ValidationLevel.WARNING,
                table="core_performance_indicators_sheet",
                field="roe",
                message=f"ROE计算不一致：报表值={actual_roe:.2f}%，计算值={calculated_roe:.2f}%",
                actual=actual_roe,
                expected=calculated_roe,
            ))

    return issues


# ============================================================
# 7. 累计与单季度关系校验
# ============================================================

def validate_cumulative_relation(
    meta: Meta,
    conn: sqlite3.Connection,
    parsed: Dict[str, Any],
    config: Dict = None
) -> List[ValidationIssue]:
    """累计与单季度关系校验"""
    if config is None:
        config = VALIDATION_CONFIG

    issues = []

    # 仅对季��/半年报进行累计关系校验
    if meta.report_period in ("Q1", "HY", "Q3"):
        # 获取历史同期数据
        history_periods = get_previous_periods(meta.report_period)
        for table_name in ["income_sheet", "cash_flow_sheet", "core_performance_indicators_sheet"]:
            data = parsed.get(table_name, ({}, []))[0]

            for field, value in data.items():
                if value is None or field.endswith("_growth") or field.endswith("_ratio"):
                    continue

                # 检查是否有历史数据
                prev_data = get_previous_period_data(
                    conn, meta.stock_code, meta.report_year, history_periods, table_name, field
                )

                if prev_data:
                    # 累计值校验（HY >= Q1，Q3 >= HY）
                    if meta.report_period == "HY" and prev_data.get("Q1"):
                        if value < prev_data["Q1"]:
                            issues.append(ValidationIssue(
                                dimension=ValidationDimension.CUMULATIVE_RELATION,
                                level=ValidationLevel.WARNING,
                                table=table_name,
                                field=field,
                                message=f"半年累计值异常：{field}={value}，一季度={prev_data['Q1']}",
                                actual=value,
                                expected=f">= {prev_data['Q1']}",
                            ))

    return issues


def get_previous_periods(current_period: str) -> List[str]:
    """获取历史期间列表"""
    mapping = {
        "Q1": [],
        "HY": ["Q1"],
        "Q3": ["Q1", "HY"],
        "FY": [],
    }
    return mapping.get(current_period, [])


def get_previous_period_data(
    conn: sqlite3.Connection,
    stock_code: str,
    report_year: int,
    periods: List[str],
    table_name: str,
    field: str
) -> Dict[str, float]:
    """获取历史期间的数据"""
    result = {}
    for period in periods:
        cur = conn.execute(
            f'SELECT "{field}" FROM "{table_name}" WHERE stock_code=? AND report_year=? AND report_period=?',
            (stock_code, report_year, period)
        ).fetchone()
        if cur and cur[0] is not None:
            result[period] = cur[0]
    return result


# ============================================================
# 综合校验入口
# ============================================================

def validate_before_insert(
    meta: Meta,
    parsed: Dict[str, Any],
    company_map: Dict[str, str],
    conn: sqlite3.Connection,
    config: Dict = None
) -> ValidationResult:
    """入库前综合校验

    Args:
        meta: 元数据
        parsed: 解析后的表格数据
        company_map: 公司映射表
        conn: 数据库连接
        config: 校验配置

    Returns:
        ValidationResult: 校验结果
    """
    if config is None:
        config = VALIDATION_CONFIG

    result = ValidationResult(passed=True)

    # 按顺序执行各维度校验
    validation_funcs = [
        ("跨表一致性", validate_cross_table_consistency),
        ("字段合理性", lambda p, c: validate_field_reasonableness(p, c)),
        ("完整性", lambda m, p, cm, c: validate_completeness(m, p, cm, c)),
        ("时序一致性", lambda m, c: validate_time_sequence(m, c)),
        ("格式", lambda m, c: validate_format(m, c)),
        ("业务逻辑", lambda p, c: validate_business_logic(p, c)),
        ("累计关系", lambda m, cn, p, c: validate_cumulative_relation(m, cn, p, c)),
    ]

    for name, func in validation_funcs:
        try:
            # 根据函数签名调用
            import inspect
            sig = inspect.signature(func)
            params = list(sig.parameters.keys())

            if len(params) == 2:
                issues = func(parsed, config)
            elif len(params) == 3:
                issues = func(meta, parsed, company_map, config)
            elif len(params) == 4:
                issues = func(meta, conn, parsed, config)
            else:
                issues = []

            for issue in issues:
                result.add_issue(issue)

        except Exception as e:
            result.add_issue(ValidationIssue(
                dimension=ValidationDimension.FIELD_REASONABLE,
                level=ValidationLevel.WARNING,
                table="validation",
                field="*",
                message=f"校验过程异常：{str(e)}",
            ))

    # 判断是否通过（无阻断级别问题）
    result.passed = result.can_continue()

    return result


def apply_auto_fixes(
    parsed: Dict[str, Any],
    issues: List[ValidationIssue]
) -> Tuple[Dict[str, Any], List[ValidationIssue]]:
    """应用自动修正

    Args:
        parsed: 解析后的数据
        issues: 校验问题列表

    Returns:
        (修正后的数据, 修正记录)
    """
    fixed_data = {table: (dict(data), list(snippets)) for table, (data, snippets) in parsed.items()}
    fix_records = []

    for issue in issues:
        if issue.auto_fixed:
            fix_records.append(issue)

    return fixed_data, fix_records


def log_validation_issues(
    conn: sqlite3.Connection,
    meta: Meta,
    result: ValidationResult,
    quality_stats: Dict
) -> None:
    """记录校验问题到数据库和统计"""
    # 记录到数据库
    for issue in result.issues:
        level_map = {
            ValidationLevel.PASS: "ok",
            ValidationLevel.WARNING: "warning",
            ValidationLevel.ERROR: "error",
            ValidationLevel.BLOCK: "blocked",
        }
        status = level_map.get(issue.level, "unknown")
        message = f"[{issue.dimension.name}] {issue.message}"
        if issue.auto_fixed:
            message += " [已自动修正]"
        log(conn, meta.source_file_name or "unknown", status, message)

    # 统计
    if result.block_count > 0:
        quality_stats["blocked"] = quality_stats.get("blocked", 0) + 1
    if result.warning_count > 0:
        quality_stats["warnings"] = quality_stats.get("warnings", 0) + result.warning_count
    if result.error_count > 0:
        quality_stats["errors"] = quality_stats.get("errors", 0) + result.error_count


# ============================================================
# 数据验证 (validation.py)
# ============================================================

def is_reasonable_value_check(field_name: str, value: Any) -> bool:
    """检查字段值是否在合理范围内"""
    if value is None:
        return True
    if not isinstance(value, (int, float)):
        return True
    if field_name not in FIELD_REASONABLE_LIMITS:
        return True
    lo, hi = FIELD_REASONABLE_LIMITS[field_name]
    return lo <= value <= hi


def is_reasonable_per_share(value: float) -> bool:
    """判断每股净资产是否合理"""
    return isinstance(value, (int, float)) and MIN_REASONABLE_NAVPS <= value <= MAX_REASONABLE_NAVPS


def is_reasonable_quarter_value(val: float, field_name: str) -> bool:
    """判断单季度值是否合理"""
    if val is None:
        return False
    if not isinstance(val, (int, float)):
        return False
    return True


def normalize_field_value(field_name: str, val: float) -> Optional[float]:
    """标准化字段值"""
    if val is None:
        return None
    if field_name in PER_SHARE_FIELDS:
        if isinstance(val, (int, float)) and abs(val) > 100:
            val = val / 10000.0
        return round(val, 4) if isinstance(val, (int, float)) else val
    if field_name in ROE_FIELDS:
        if isinstance(val, (int, float)):
            if abs(val) > 10000:
                val = val / 100.0
            elif abs(val) > 200:
                val = val / 100.0
        return round(val, 4) if isinstance(val, (int, float)) else val
    if field_name == "net_cash_flow":
        return round(val / 10000.0, 4)
    if (
        field_name.endswith("_growth")
        or field_name.endswith("_ratio")
        or field_name.endswith("_ratio_of_net_cf")
        or field_name in {"gross_profit_margin", "net_profit_margin"}
    ):
        return round(val, 4)
    if field_name == "operating_cf_net_amount":
        return round(val / 10000.0, 4)
    return round(val / 10000.0, 4)


# ============================================================
# 数据库工具 (db_utils.py)
# ============================================================

def business_field_count(row: Dict) -> int:
    """统计业务字段数量"""
    return sum(1 for k, v in row.items() if k not in META_FIELDS and v is not None)


def sql_type(t: str) -> str:
    """获取SQL类型"""
    t = t.lower()
    return "INTEGER" if "int" in t else ("REAL" if any(k in t for k in ["decimal", "float", "double"]) else "TEXT")


def ensure_tables(conn: sqlite3.Connection, schema: Dict) -> None:
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
            f'CREATE TABLE IF NOT EXISTS "{t}" ({",".join(cols)}, UNIQUE(stock_code, report_year, report_period))'
        )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS raw_table_snippets("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "stock_code TEXT,report_year INTEGER,report_period TEXT,target_table TEXT,"
        "source_section TEXT,snippet TEXT,created_at TEXT)"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS etl_audit_log("
        "id INTEGER PRIMARY KEY AUTOINCREMENT,"
        "source_file_name TEXT,status TEXT,message TEXT,created_at TEXT)"
    )
    conn.commit()


def upsert(conn: sqlite3.Connection, t: str, row: Dict) -> None:
    """插入或更新数据"""
    if not row or business_field_count(row) == 0:
        return
    cur = conn.execute(
        f'SELECT * FROM "{t}" WHERE stock_code=? AND report_year=? AND report_period=?',
        (row.get("stock_code"), row.get("report_year"), row.get("report_period")),
    )
    existing = cur.fetchone()
    if existing is not None:
        existing_row = dict(zip([d[0] for d in cur.description], existing))
        prefer_new_business = business_field_count(row) >= business_field_count(existing_row)
        merged_row = dict(existing_row)
        for key, value in row.items():
            if key in META_FIELDS:
                if key == "report_period":
                    if value:
                        merged_row[key] = value
                elif key == "source_file_name":
                    merged_row[key] = value
                elif value not in (None, ""):
                    merged_row[key] = value
            elif value is not None:
                if prefer_new_business or merged_row.get(key) is None:
                    merged_row[key] = value
        row = merged_row
    cols = list(row.keys())
    ps = ",".join(["?"] * len(cols))
    cs = ",".join([f'"{c}"' for c in cols])
    us = ",".join([f'"{c}"=excluded."{c}"' for c in cols if c not in {"stock_code", "report_year", "report_period"}])
    conn.execute(
        f'INSERT INTO "{t}" ({cs}) VALUES ({ps}) '
        f'ON CONFLICT(stock_code, report_year, report_period) DO UPDATE SET {us}',
        [row[c] for c in cols],
    )


def log(conn: sqlite3.Connection, f: str, s: str, m: str) -> None:
    """记录日志"""
    conn.execute(
        "INSERT INTO etl_audit_log(source_file_name,status,message,created_at) VALUES (?,?,?,?)",
        (f, s, m, dt.datetime.now().isoformat()),
    )


def coverage(ext: Dict, fs: List) -> tuple:
    """计算覆盖率"""
    tgt = [f.field_name for f in fs if f.field_name not in META_FIELDS]
    hit = sum(1 for k in tgt if k in ext)
    return (round(hit / len(tgt), 4), hit, len(tgt)) if tgt else (0.0, 0, 0)


def filter_to_schema(ext: Dict, fs: List) -> Dict:
    """过滤到schema定义的字段"""
    allowed = {f.field_name for f in fs}
    return {k: v for k, v in ext.items() if k in allowed}


# ============================================================
# 质量检查 (quality.py)
# ============================================================

def append_quality_issue(q: Dict, pdf_name: str, report_period: str, table: str, issue_type: str, message: str, details: Dict = None, evidence: Dict = None) -> None:
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


def check_field_ranges(ext: Dict, table: str, pdf_name: str, report_period: str, q: Dict, evidence_map: Dict = None) -> None:
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


def check_table_quality(ext: Dict, fs, table: str, pdf_name: str, report_period: str, q: Dict, evidence_map: Dict = None) -> tuple:
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


def check_cross_table_quality(parsed: Dict, pdf_name: str, report_period: str, q: Dict, evidence_by_table: Dict = None) -> None:
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
                {"left_value": lv, "right_value": rv, "relative_diff": round(rel, 4)},
                {
                    "left": (evidence_by_table or {}).get(lt, {}).get(lf, {}),
                    "right": (evidence_by_table or {}).get(rt, {}).get(rf, {}),
                },
            )
    if parsed.get("cash_flow_sheet", ({}, []))[0].get("net_cash_flow") not in (None, 0):
        cf = parsed["cash_flow_sheet"][0]
        ratio_sum = sum(
            cf.get(k, 0)
            for k in ["operating_cf_ratio_of_net_cf", "investing_cf_ratio_of_net_cf", "financing_cf_ratio_of_net_cf"]
            if isinstance(cf.get(k), (int, float))
        )
        if ratio_sum and abs(ratio_sum - 100) > 50:
            append_quality_issue(
                q, pdf_name, report_period, "cash_flow_sheet",
                "cashflow_ratio_sum", "现金流占比合计明显异常",
                {"ratio_sum": round(ratio_sum, 4)},
            )


# ============================================================
# 元数据检测 (meta_detect.py)
# ============================================================

def normalize_company_name(name: str) -> str:
    """标准化公司名称"""
    s = normalize_key(name)
    for suffix in [
        "股份有限公司", "集团股份有限公司", "医药股份有限公司", "有限公司",
        "集团有限公司", "股份", "集团", "医药",
    ]:
        s = s.replace(normalize_key(suffix), "")
    return s


def lookup_code_by_name(name: str, cm: Dict[str, str]) -> Tuple[str, str]:
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


def infer_report_type(text: str) -> Tuple[str, str]:
    """推断报告类型"""
    for k, v in REPORT_TYPE_PATTERNS:
        if k in text:
            if v == "QX" and ("第一季度" in text or "一季度" in text):
                return ("一季度报告", "Q1")
            if v == "QX" and ("第三季度" in text or "三季度" in text):
                return ("三季度报告", "Q3")
            return (k, v)
    return ("年度报告", "FY")


def build_report_period(report_year: int, report_code: str) -> str:
    """构建报告期字符串"""
    return report_code if report_code else ""


def build_serial_number(stock_code: str, report_year: int, report_period: str) -> int:
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


def detect_meta_from_md(md: str, cm: Dict[str, str]) -> Meta:
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

    if not code:
        m = re.search(r"编制单位[：:]\s*([^\n\r]+)", plain)
        if m:
            unit_name = m.group(1).strip()
            code, mapped_abbr = lookup_code_by_name(unit_name, cm)
            if mapped_abbr and not abbr:
                abbr = mapped_abbr

    if not code and abbr:
        code, mapped_abbr = lookup_code_by_name(abbr, cm)

    return Meta(code, abbr, report_year, report_type, report_period, "")


def merge_meta(primary: Meta, fallback: Meta, source_name: str) -> Meta:
    """合并元数据"""
    code = primary.stock_code or fallback.stock_code
    abbr = primary.stock_abbr or fallback.stock_abbr
    if fallback.report_year > 0 and (primary.report_year == 0 or primary.report_year != fallback.report_year):
        report_year = fallback.report_year
    else:
        report_year = primary.report_year or fallback.report_year
    report_period = fallback.report_period or primary.report_period
    report_type = fallback.report_type if fallback.report_period else primary.report_type
    return Meta(code, abbr, report_year, report_type, report_period, source_name)


def detect_meta(pdf: Path, cm: Dict[str, str], md: str = "") -> Meta:
    """检测PDF元数据"""
    n = pdf.stem
    c = ""
    a = ""

    m = re.search(r"(?<!\d)(\d{6})(?!\d)", n)
    if m:
        c = m.group(1)
        a = cm.get(c, "")
    else:
        left = n.split("：")[0].strip() if "：" in n else ""
        if left:
            c, mapped_abbr = lookup_code_by_name(left, cm)
            a = mapped_abbr or left

    y = re.search(r"(20\d{2})年", n)
    if not y:
        y = re.search(r"(20\d{2})", n)
    ry = int(y.group(1)) if y else 0
    rt, rc = infer_report_type(n)
    rp = build_report_period(ry, rc)

    base = Meta(c, a, ry, rt, rp, pdf.name)
    if not md:
        return base

    md_meta = detect_meta_from_md(md, cm)
    return merge_meta(base, md_meta, pdf.name)


def is_summary_name(p: Path) -> bool:
    """判断是否为摘要文件名"""
    n = p.stem
    return any(
        k in n
        for k in [
            "报告摘要", "年度报告摘要", "半年度报告摘要", "季度报告摘要",
            "一季度报告摘要", "三季度报告摘要", "摘要版",
        ]
    )


def detect_summary_from_md(md: str) -> bool:
    """从Markdown内容判断是否为摘要"""
    t = md[:12000]
    score = 0
    pos = [
        "年度报告摘要", "半年度报告摘要", "季度报告摘要",
        "第一季度报告摘要", "第三季度报告摘要", "报告摘要",
    ]
    neg = ["目录", "公司治理", "重要事项", "审计意见类型", "董事会报告", "财务报表附注", "备查文件目录", "释义"]
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
# 表头分析 (header_analyzer.py)
# ============================================================

COLUMN_TYPE_PATTERNS = {
    "yoy": ["增减", "变动幅度", "同比", "本报告期.*比上年", "比上年同期增减", "比上年.*增减"],
    "cumulative": ["年初至报告期末", "年初至本报告期末", "本年累计", "年初至报告期期末"],
    "period_end": ["本报告期末", "期末", "期末数", "报告期末"],
    "period": ["本报告期", "本期", "单季度", "1-3月", "4-6月", "7-9月"],
    "previous": ["上年同期", "上年度末", "上期", "期初", "年初", "上年半年度", "上年第一季度"],
    "year": [r"20\d{2}年"],
}


@dataclass
class HeaderColumn:
    """表头列信息"""
    index: int
    raw: str
    norm: str
    col_type: str = "unknown"
    is_value: bool = False
    relative_time: str = ""


def analyze_header_columns(header: List[str]) -> List[HeaderColumn]:
    """分析表头列的类型"""
    cols = []
    for i, h in enumerate(header):
        raw = h.strip()
        norm = normalize_key(raw)
        col = HeaderColumn(index=i, raw=raw, norm=norm)

        detected_type = "unknown"

        if norm in {"项目", ""} or not norm:
            col.col_type = "label"
            col.is_value = False
            cols.append(col)
            continue

        for pattern in COLUMN_TYPE_PATTERNS["yoy"]:
            if re.search(pattern, norm):
                col.col_type = "yoy"
                col.is_value = True
                break

        if col.col_type != "unknown":
            cols.append(col)
            continue

        for pattern in COLUMN_TYPE_PATTERNS["cumulative"]:
            if pattern in norm:
                col.col_type = "cumulative"
                col.is_value = True
                break

        if col.col_type != "unknown":
            col.relative_time = "current"
            cols.append(col)
            continue

        for pattern in COLUMN_TYPE_PATTERNS["period_end"]:
            if pattern in norm:
                col.col_type = "period_end"
                col.is_value = True
                break

        if col.col_type != "unknown":
            col.relative_time = "current"
            cols.append(col)
            continue

        for pattern in COLUMN_TYPE_PATTERNS["period"]:
            if re.search(pattern, norm):
                col.col_type = "period"
                col.is_value = True
                break

        if col.col_type != "unknown":
            col.relative_time = "current"
            cols.append(col)
            continue

        for pattern in COLUMN_TYPE_PATTERNS["previous"]:
            if re.search(pattern, norm):
                col.col_type = "previous"
                col.is_value = True
                col.relative_time = "previous"
                break

        if col.col_type != "unknown":
            cols.append(col)
            continue

        for pattern in COLUMN_TYPE_PATTERNS["year"]:
            if re.search(pattern, norm):
                col.col_type = "year"
                col.is_value = True
                break

        if col.col_type != "unknown":
            cols.append(col)
            continue

        col.is_value = True
        cols.append(col)

    return cols


def is_core_value_header(text: str, table: str = "core_performance_indicators_sheet") -> bool:
    """判断是否为核心数值表头"""
    h = normalize_key(text)
    if not h:
        return False
    if any(k in h for k in ["增减", "变动幅度", "同比"]):
        return False
    if table == "core_performance_indicators_sheet":
        if any(k in h for k in ["本报告期", "本报告期末", "本期", "上期", "上年同期", "上年度末", "调整前", "调整后", "年初至报告期末", "年初至报告期期末"]):
            return True
        return False
    if any(k in h for k in ["本报告期", "本报告期末", "本期", "上期", "上年同期", "上年度末", "调整前", "调整后", "年初至报告期末", "年初至报告期期末"]):
        return True
    return bool(YEAR_LIKE_RE.search(h))


def match_header_hint(text: str, hints: List[str]) -> bool:
    """匹配表头提示"""
    n = normalize_key(text)
    return any(normalize_key(k) in n for k in hints)


def current_col_from_header(table: str, header: List[str]) -> int:
    """从表头获取当期列索引"""
    hs = [(i, h, normalize_key(h)) for i, h in enumerate(header)]

    if table == "core_performance_indicators_sheet":
        has_cumulative = any("年初至报告期末" in h[2] for h in hs)
        has_cumulative_excl = any("年初至报告期末" in h[2] and "本报告期末" not in h[2] for h in hs)
        has_report_period_only = any("本报告期" in h[2] and "年初至" not in h[2] and "本报告期末" not in h[2] for h in hs)
        has_period_end = any("本报告期末" in h[2] for h in hs)

        if has_period_end:
            for i, orig, norm in hs:
                if "本报告期末" in norm:
                    return i

        if has_cumulative and has_report_period_only:
            for i, orig, norm in hs:
                if "年初至报告期末" in norm and "本报告期末" not in norm:
                    return i

        if has_report_period_only:
            for i, orig, norm in hs:
                if "本报告期" in norm and "年初至" not in norm:
                    return i

        if has_cumulative:
            for i, orig, norm in hs:
                if "年初至报告期末" in norm:
                    return i

        candidates = [i for i, orig, norm in hs if is_core_value_header(norm, table)]
        if candidates:
            return candidates[0]

    table_hints = HEADER_ROLE_HINTS.get(table, {})
    for i, orig, norm in hs:
        if match_header_hint(norm, table_hints.get("current", [])):
            return i
    for i, orig, norm in hs:
        if match_header_hint(norm, GENERIC_HEADER_HINTS.get("current", [])):
            return i
    return 1 if len(header) > 1 else 0


def previous_col_from_header(table: str, header: List[str], current_idx: Optional[int]) -> Optional[int]:
    """从表头获取上期列索引"""
    hs = [(i, h, normalize_key(h)) for i, h in enumerate(header)]
    table_hints = HEADER_ROLE_HINTS.get(table, {})

    if table == "core_performance_indicators_sheet":
        adjusted_candidates = [i for i, orig, norm in hs if "调整后" in norm and i != current_idx]
        if adjusted_candidates:
            if current_idx is not None:
                after_current = [i for i in adjusted_candidates if i > current_idx]
                if after_current:
                    return after_current[-1]
            return adjusted_candidates[-1]

        if current_idx is not None and current_idx < len(header):
            curr_norm = normalize_key(header[current_idx])
            if "本报告期末" in curr_norm:
                for i, orig, norm in hs:
                    if "上年度末" in norm:
                        return i

        if current_idx is not None and current_idx < len(header):
            curr_norm = normalize_key(header[current_idx])
            if "本报告期" in curr_norm and "年初至" not in curr_norm:
                for i, orig, norm in hs:
                    if "上年同期" in norm:
                        return i

        if current_idx is not None and current_idx > 0:
            left_candidates = [i for i, orig, norm in hs
                             if 0 <= i < current_idx and is_core_value_header(norm, table)]
            if left_candidates:
                return left_candidates[-1]

        if current_idx is not None and current_idx < len(header):
            curr_norm = normalize_key(header[current_idx])
            if "年初至报告期末" in curr_norm:
                for i, orig, norm in hs:
                    if "上年同期" in norm:
                        return i
                left_candidates = [i for i, orig, norm in hs
                                 if i < current_idx and is_core_value_header(norm, table)]
                if left_candidates:
                    return left_candidates[-1]

        has_yoy = any(
            match_header_hint(norm, table_hints.get("yoy", []) + GENERIC_HEADER_HINTS.get("yoy", []))
            for _, _, norm in hs
        )
        if has_yoy and len(header) <= 3:
            return None

    for i, orig, norm in hs:
        if i != current_idx and match_header_hint(norm, table_hints.get("previous", [])):
            return i

    if current_idx is not None and current_idx > 0:
        for i in range(current_idx - 1, -1, -1):
            if i > 0 or (len(hs) > 1 and hs[0][2] and "项目" not in hs[0][2]):
                return i
    return None


def yoy_col_from_header(table: str, header: List[str]) -> Optional[int]:
    """从表头获取同比列索引"""
    hs = [(i, normalize_key(h)) for i, h in enumerate(header)]
    hints = HEADER_ROLE_HINTS.get(table, {}).get("yoy", []) + GENERIC_HEADER_HINTS.get("yoy", [])
    if table == "core_performance_indicators_sheet":
        for i, h in hs:
            if "年初至报告期末比上年同期增减" in h or "年初至报告期末比上年同期变动幅度" in h:
                return i
    for i, h in hs:
        if match_header_hint(h, hints):
            return i
    return None


def select_col_for_field(cols: List[HeaderColumn], field_type: str) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """根据字段类型选择当期列、上期列、同比列"""
    yoy_cols = [c for c in cols if c.col_type == "yoy"]
    cum_cols = [c for c in cols if c.col_type == "cumulative"]
    pe_cols = [c for c in cols if c.col_type == "period_end"]
    p_cols = [c for c in cols if c.col_type == "period"]
    prev_cols = [c for c in cols if c.col_type in ("previous", "year")]

    yoy_idx = yoy_cols[0].index if yoy_cols else None

    if field_type == "period_end":
        if pe_cols:
            cur = pe_cols[0].index
            pre = next((c.index for c in prev_cols if "上年度" in c.norm), None)
            return cur, pre, yoy_idx
        if cum_cols:
            cur = cum_cols[0].index
            pre = prev_cols[0].index if prev_cols else None
            return cur, pre, yoy_idx
        if p_cols:
            return p_cols[0].index, prev_cols[0].index if prev_cols else None, yoy_idx

    elif field_type == "cumulative":
        if cum_cols:
            cur = cum_cols[0].index
            pre = next((c.index for c in prev_cols if "上年同期" in c.norm), None)
            return cur, pre, yoy_idx
        if p_cols:
            return p_cols[0].index, prev_cols[0].index if prev_cols else None, yoy_idx
        if pe_cols:
            return pe_cols[0].index, prev_cols[0].index if prev_cols else None, yoy_idx

    elif field_type == "ratio":
        if yoy_cols:
            return yoy_cols[0].index, None, yoy_idx
        for c in cols:
            if c.is_value and c.index > 0:
                return c.index, None, yoy_idx

    value_cols = [c for c in cols if c.is_value and c.index > 0]
    if value_cols:
        cur = value_cols[0].index
        pre = value_cols[1].index if len(value_cols) > 1 else None
        return cur, pre, yoy_idx

    return None, None, yoy_idx


# ============================================================
# 表格解析 (table_parser.py)
# ============================================================

def parse_table_rows(sec: str) -> List[List[str]]:
    """解析表格行（支持HTML和Markdown格式）"""
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


def locate_header_row(rows: List[List[str]]) -> int:
    """定位表头行"""
    header_hints = GENERIC_HEADER_HINTS.get("header_row", [])
    for i, cs in enumerate(rows[:12]):
        joined = " ".join(cs)
        if any(h in joined for h in header_hints):
            return i
    return 0


def parse_table_matrix(sec: str) -> Tuple[List[List[str]], List[str], List[List[str]]]:
    """解析表格矩阵"""
    rows = parse_table_rows(sec)
    if not rows:
        return [], [], []
    hi = locate_header_row(rows)
    header = rows[hi]
    body = rows[hi + 1:]
    return rows, header, body


def extract_quarterly_data(md: str) -> Dict[str, Dict[str, float]]:
    """从Markdown中提取分季度数据

    财报中通常有"九、XXXX年分季度主要财务数据"表格，包含Q1-Q4的单季度值。
    这个数据可以用来计算FY和Q1的环比(QoQ)指标。

    Returns:
        {
            "revenue": {"Q1": xxx, "Q2": xxx, "Q3": xxx, "Q4": xxx},
            "net_profit": {"Q1": xxx, "Q2": xxx, "Q3": xxx, "Q4": xxx}
        }
    """
    quarterly_data = {"revenue": {}, "net_profit": {}}

    # 查找分季度数据表格
    # 格式通常是："九、XXXX年分季度主要财务数据" 或类似标题
    import re
    patterns = [
        r'分季度.*?主要财务数据.*?<table>(.*?)</table>',
        r'季度.*?财务.*?<table>(.*?)</table>',
    ]

    for pattern in patterns:
        matches = re.findall(pattern, md, re.DOTALL | re.IGNORECASE)
        for match in matches:
            # 解析表格内容
            table_content = match
            # 提取行数据
            rows = []
            for row_match in re.findall(r'<tr[^>]*>(.*?)</tr>', table_content, re.DOTALL | re.I):
                cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', row_match, re.DOTALL | re.I)
                if cells:
                    cleaned_cells = []
                    for cell in cells:
                        # 清理HTML标签和空格
                        cell_text = re.sub(r'<[^>]+>', '', cell).strip()
                        cell_text = cell_text.replace('\xa0', ' ').replace(' ', '')
                        cleaned_cells.append(cell_text)
                    rows.append(cleaned_cells)

            # 查找季度列
            quarter_cols = {}
            header_row = None
            for i, row in enumerate(rows[:3]):
                for j, cell in enumerate(row):
                    cell_lower = cell.lower()
                    if '季度' in cell or '1-3' in cell or '4-6' in cell or '7-9' in cell or '10-12' in cell:
                        quarter_cols[i] = j
                        if header_row is None:
                            header_row = i

            # 查找营收和净利润行
            for row in rows[header_row + 1:]:
                if len(row) < 2:
                    continue
                label = row[0]
                label = re.sub(r'<[^>]+>', '', label).strip()

                # 清理标签中的空格
                label = label.replace('\xa0', ' ').replace(' ', '')

                # 匹配营收行
                if '营业' in label and '收入' in label and '总' not in label:
                    for q_name, q_idx in [('Q1', 1), ('Q2', 2), ('Q3', 3), ('Q4', 4)]:
                        if q_idx < len(row):
                            val_text = row[q_idx].replace('\xa0', ' ').replace(' ', '')
                            val = clean_num(val_text, 1.0)
                            if val is not None and val > 0:
                                quarterly_data["revenue"][q_name] = val

                # 匹配净利润行
                if '归属' in label and '净利润' in label and '非经常' not in label and '扣除' not in label:
                    for q_name, q_idx in [('Q1', 1), ('Q2', 2), ('Q3', 3), ('Q4', 4)]:
                        if q_idx < len(row):
                            val_text = row[q_idx].replace('\xa0', ' ').replace(' ', '')
                            val = clean_num(val_text, 1.0)
                            if val is not None and abs(val) > 0:
                                quarterly_data["net_profit"][q_name] = val

            # 如果成功提取到数据，返回
            if quarterly_data["revenue"] or quarterly_data["net_profit"]:
                return quarterly_data

    return quarterly_data


def standardize_table_section(sec: str) -> str:
    """标准化表格区块"""
    sec = sec.replace("\r\n", "\n").replace("\r", "\n")
    sec = sec.replace("−", "-").replace("—", "-").replace("–", "-")
    return normalize_text_cell(sec)


def build_alias(fs, table: str = "") -> dict:
    """构建别名映射"""
    m = {}
    for f in fs:
        c = f.cn_name
        m[normalize_key(c)] = f.field_name
        m[normalize_key(c.replace("-", ""))] = f.field_name
        m[normalize_key(
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
        )] = f.field_name
        for alias in TABLE_FIELD_ALIASES.get(table, {}).get(f.field_name, []):
            m[normalize_key(alias)] = f.field_name

    for k, v in SYNONYM.items():
        kk = normalize_key(k)
        if kk not in m:
            m[kk] = v
    return m


def first_number_values(cs: List[str], uf: float) -> List[float]:
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


def value_by_idx(cs: List[str], idx: int, uf: float) -> float:
    """按索引获取值"""
    if idx is None or idx < 0 or idx >= len(cs):
        return None
    if cs and cs[0].strip() == "":
        return None
    return clean_num(cs[idx], uf)


def extract_core_row_metrics(cs: List[str], uf: float) -> Tuple[float, float, float]:
    """提取核心行指标"""
    nums = []
    for idx, c in enumerate(cs[1:], start=1):
        text = normalize_text_cell(c)
        if not text or text in {"不适用", "N/A"}:
            continue
        is_pct = "%" in text
        val = clean_num(text, 1.0 if is_pct else uf)
        if val is None:
            continue
        nums.append((idx, val, is_pct))

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


def extract_adjusted_core_metrics(cs: List[str], header: List[str], uf: float) -> Tuple[float, float, float]:
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


def derive_values_from_row(cs: List[str], uf: float, table: str, tf: str, curr_is_none: bool = False) -> List[float]:
    """从行推导值"""
    vals = first_number_values(cs, uf)
    if not vals:
        return []

    if table == "core_performance_indicators_sheet":
        if tf in {"eps", "roe", "roe_weighted_excl_non_recurring", "net_asset_per_share", "operating_cf_per_share"}:
            return vals[:2]
        if len(vals) >= 3:
            return [vals[0], vals[1], vals[2]]
        return vals

    if table == "balance_sheet" and tf in BALANCE_SHEET_ZERO_FIELDS:
        if curr_is_none:
            return []
        return [vals[0]] if vals else []

    if table in {"income_sheet", "cash_flow_sheet", "balance_sheet"}:
        return vals[:2]

    return vals


# ============================================================
# 区块分段 (sectioning.py)
# ============================================================

def slice_by_boundaries(md: str, table: str) -> List[str]:
    """按边界切分"""
    parts = []
    for start_kw, end_kw in SECTION_BOUNDARIES.get(table, []):
        start = md.find(start_kw)
        if start < 0:
            continue
        end = md.find(end_kw, start + len(start_kw)) if end_kw else -1
        parts.append(md[start: (end if end > start else len(md))])
    return parts


def keyword_sections(md: str, table: str) -> List[str]:
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


def score_section(sec: str, table: str) -> int:
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


def sections(md: str, table: str) -> List[str]:
    """获取最优区块"""
    candidates = slice_by_boundaries(md, table) or keyword_sections(md, table)
    if not candidates:
        return []
    ranked = sorted(
        ((score_section(sec, table), idx, sec) for idx, sec in enumerate(candidates)),
        key=lambda x: (-x[0], x[1]),
    )
    return [sec for _, _, sec in ranked[:3]]


# ============================================================
# 跨表指标计算 (cross_table.py)
# ============================================================

def calc_growth_pct(curr: float, prev: float, field_name: str = None) -> Optional[float]:
    """计算增长率（带合理性检查）"""
    if curr is None or prev in (None, 0):
        return None
    growth = round((curr - prev) / abs(prev) * 100, 4)
    return growth


def split_report_period(full_report_period: str) -> tuple:
    """拆分完整报告期字符串为年份和期号

    Args:
        full_report_period: 完整报告期字符串，如 "2023Q1"、"2023FY"、"Q1"、"FY"

    Returns:
        (year, period) 元组，year为整数年份或None，period为报告期字符串或None
    """
    if not full_report_period:
        return None, None

    # 如果是纯期号（不含年份），返回 (None, period)
    if full_report_period in ("Q1", "Q2", "Q3", "Q4", "HY", "FY"):
        return None, full_report_period

    # 尝试从字符串中提取年份和期号
    # 支持格式: "2023Q1", "2023HY", "2023FY"
    import re
    match = re.match(r'^(\d{4})(Q[1234]|HY|FY)$', full_report_period)
    if match:
        return int(match.group(1)), match.group(2)

    # 如果无法解析，返回 (None, original)
    return None, full_report_period


def build_full_report_period(report_year: int, report_period: str) -> str:
    """构建完整报告期字符串

    Args:
        report_year: 报告年份，如 2023
        report_period: 报告期，如 "Q1"、"HY"、"FY"

    Returns:
        完整报告期字符串，如 "2023Q1"
    """
    if report_year and report_period:
        return f"{report_year}{report_period}"
    return ""


def previous_period_in_year(report_period: str) -> Optional[str]:
    """获取年内上一期

    Args:
        report_period: 完整报告期字符串，如 "2023HY"

    Returns:
        年内上一期的完整报告期字符串，如 "2023HY" 的上一期是 "2023Q1"
    """
    year, period = split_report_period(report_period)
    if not year or not period:
        return None
    mapping = {"HY": "Q1", "Q3": "HY", "FY": "Q3"}
    prev_period = mapping.get(period)
    if prev_period:
        return f"{year}{prev_period}"
    return None


def previous_quarter_period(report_period: str) -> Optional[str]:
    """获取上一季度

    Args:
        report_period: 完整报告期字符串，如 "2023Q1"、"2023HY"

    Returns:
        上一季度的完整报告期字符串，如 "2023Q1" 的上一季度是 "2022FY"
    """
    year, period = split_report_period(report_period)
    if not year or not period:
        return None
    mapping = {
        "Q1": f"{year - 1}FY",
        "HY": f"{year}Q1",
        "Q3": f"{year}HY",
        "FY": f"{year}Q3"
    }
    return mapping.get(period)


def get_table_row(history: Dict, table_name: str, report_period: str) -> Dict:
    """获取表行数据"""
    return history.get(table_name, {}).get(report_period, {})


def cumulative_value_for_period(history: Dict, parsed: Dict, table_name: str, field_name: str, report_period: str) -> Optional[float]:
    """获取期间累计值"""
    current_period = parsed.get("__current_report_period__")
    if report_period == current_period:
        row = parsed.get(table_name, ({}, []))[0]
    else:
        row = get_table_row(history, table_name, report_period)
    return row.get(field_name)


def quarter_value(history: Dict, parsed: Dict, table_name: str, field_name: str, report_period: str) -> Optional[float]:
    """获取季度值（从累计值计算单季度值）

    Args:
        history: 历史数据查询表
        parsed: 当前解析结果
        table_name: 表名
        field_name: 字段名
        report_period: 完整报告期字符串，如 "2023HY"

    Returns:
        单季度值（如有）
    """
    year, period = split_report_period(report_period)
    if not year or not period:
        return None

    cumulative = cumulative_value_for_period(history, parsed, table_name, field_name, report_period)
    if cumulative is None:
        return None

    if period == "Q1":
        return cumulative

    if period == "FY":
        return None

    prev_in_year = previous_period_in_year(report_period)
    if prev_in_year:
        prev_cumulative = cumulative_value_for_period(history, parsed, table_name, field_name, prev_in_year)
        if prev_cumulative is not None:
            quarter_val = round(cumulative - prev_cumulative, 4)
            if isinstance(quarter_val, (int, float)):
                return quarter_val
            return None

    return None


def fallback_per_share(equity_total_equity: float) -> Optional[float]:
    """计算每股净资产后备值"""
    if equity_total_equity is None:
        return None
    return round(equity_total_equity * 10000 / DEFAULT_SHARES_OUTSTANDING, 4)


def build_history_lookup(parsed_history: Dict) -> Dict:
    """构建历史查询表

    Args:
        parsed_history: 键为 (report_year, report_period) 元组或完整报告期字符串

    Returns:
        键为完整报告期字符串（如 "2023Q1"）的历史查询表
    """
    history = {}
    for key, parsed in parsed_history.items():
        # 解析键：支持元组 (year, period) 或字符串 "2023Q1"
        if isinstance(key, tuple) and len(key) == 2:
            report_year, report_period = key
            full_period = f"{report_year}{report_period}" if report_year else report_period
        elif isinstance(key, str):
            full_period = key
        else:
            full_period = str(key)

        for table, (row, _) in parsed.items():
            history.setdefault(table, {})[full_period] = row
    return history


def cross_table_quality(parsed: Dict, pdf_name: str, report_period: str, q: Dict, evidence_by_table: Dict = None) -> None:
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
                {"left_value": lv, "right_value": rv, "relative_diff": round(rel, 4)},
                {
                    "left": (evidence_by_table or {}).get(lt, {}).get(lf, {}),
                    "right": (evidence_by_table or {}).get(rt, {}).get(rf, {}),
                },
            )

    cf = parsed.get("cash_flow_sheet", ({}, []))[0]
    if cf.get("net_cash_flow") not in (None, 0):
        ratio_sum = sum(
            cf.get(k, 0)
            for k in ["operating_cf_ratio_of_net_cf", "investing_cf_ratio_of_net_cf", "financing_cf_ratio_of_net_cf"]
            if isinstance(cf.get(k), (int, float))
        )
        if ratio_sum and abs(ratio_sum - 100) > 50:
            append_quality_issue(
                q, pdf_name, report_period, "cash_flow_sheet",
                "cashflow_ratio_sum", "现金流占比合计明显异常",
                {"ratio_sum": round(ratio_sum, 4)},
            )


def enrich_cross_table_metrics(
    parsed: Dict,
    schema: Dict,
    parsed_history: Dict = None,
    current_report_year: int = None,
    current_report_period: str = "",
    quarterly_data: Dict = None
) -> Dict:
    """丰富跨表指标

    Args:
        parsed: 解析结果字典
        schema: 表结构定义
        parsed_history: 历史解析数据（键为完整报告期字符串如"2023Q1"）
        current_report_year: 当前报告年份
        current_report_period: 当前报告期（FY/HY/Q1/Q3，不带年份）
        quarterly_data: 分季度数据 {"revenue": {"Q1": xxx, ...}, "net_profit": {...}}
    """
    core = parsed.get("core_performance_indicators_sheet", ({}, []))[0]
    bal = parsed.get("balance_sheet", ({}, []))[0]
    inc = parsed.get("income_sheet", ({}, []))[0]
    cf = parsed.get("cash_flow_sheet", ({}, []))[0]

    # 获取报告年份和期号
    report_year = current_report_year or core.get("report_year") or inc.get("report_year")
    report_period = current_report_period or core.get("report_period") or inc.get("report_period")

    # 构建完整报告期字符串
    if report_year and report_period:
        full_current_period = f"{report_year}{report_period}"
    else:
        full_current_period = ""

    # 跨表补充基础指标
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
        margin = round(
            (inc["total_operating_revenue"] - inc["operating_expense_cost_of_sales"])
            / inc["total_operating_revenue"] * 100, 4,
        )
        core["gross_profit_margin"] = margin

    if (
        core.get("net_profit_margin") is None
        and core.get("net_profit_10k_yuan") is not None
        and core.get("total_operating_revenue") not in (None, 0)
    ):
        margin = round(core["net_profit_10k_yuan"] / core["total_operating_revenue"] * 100, 4)
        core["net_profit_margin"] = margin

    fallback_navps = None
    if bal.get("equity_total_equity"):
        fallback_navps = round(bal["equity_total_equity"] * 10000 / DEFAULT_SHARES_OUTSTANDING, 4)
    if core.get("net_asset_per_share") is None and fallback_navps is not None:
        core["net_asset_per_share"] = fallback_navps
    elif fallback_navps is not None and not is_reasonable_per_share(core.get("net_asset_per_share")):
        core["net_asset_per_share"] = fallback_navps

    if core.get("operating_cf_per_share") is None and cf.get("operating_cf_net_amount") is not None:
        ocf_per_share = round(cf["operating_cf_net_amount"] * 10000 / DEFAULT_SHARES_OUTSTANDING, 4)
        core["operating_cf_per_share"] = ocf_per_share

    if core.get("roe") is None and bal.get("equity_total_equity") not in (None, 0):
        base_profit = core.get("net_profit_10k_yuan")
        if base_profit is not None:
            calculated_roe = round(base_profit / bal["equity_total_equity"] * 100, 4)
            core["roe"] = calculated_roe

    if (
        core.get("roe_weighted_excl_non_recurring") is None
        and core.get("net_profit_excl_non_recurring") is not None
        and bal.get("equity_total_equity") not in (None, 0)
    ):
        calculated_roe_weighted = round(
            core["net_profit_excl_non_recurring"] / bal["equity_total_equity"] * 100, 4
        )
        core["roe_weighted_excl_non_recurring"] = calculated_roe_weighted

    if (
        bal.get("asset_liability_ratio") is None
        and bal.get("liability_total_liabilities") is not None
        and bal.get("asset_total_assets") not in (None, 0)
    ):
        ratio = round(bal["liability_total_liabilities"] / bal["asset_total_assets"] * 100, 4)
        bal["asset_liability_ratio"] = ratio

    if (
        cf.get("operating_cf_ratio_of_net_cf") is None
        and cf.get("operating_cf_net_amount") is not None
        and cf.get("net_cash_flow") not in (None, 0)
    ):
        ratio = round(cf["operating_cf_net_amount"] / cf["net_cash_flow"] * 100, 4)
        cf["operating_cf_ratio_of_net_cf"] = ratio

    if (
        cf.get("investing_cf_ratio_of_net_cf") is None
        and cf.get("investing_cf_net_amount") is not None
        and cf.get("net_cash_flow") not in (None, 0)
    ):
        ratio = round(cf["investing_cf_net_amount"] / cf["net_cash_flow"] * 100, 4)
        cf["investing_cf_ratio_of_net_cf"] = ratio

    if (
        cf.get("financing_cf_ratio_of_net_cf") is None
        and cf.get("financing_cf_net_amount") is not None
        and cf.get("net_cash_flow") not in (None, 0)
    ):
        ratio = round(cf["financing_cf_net_amount"] / cf["net_cash_flow"] * 100, 4)
        cf["financing_cf_ratio_of_net_cf"] = ratio

    # 设置当前报告期用于QoQ计算
    if full_current_period:
        parsed["__current_report_period__"] = full_current_period

    # 使用历史数据计算YoY和QoQ
    if full_current_period and parsed_history:
        history = build_history_lookup(parsed_history)

        # 计算同比期（上一年同报告期）
        prev_yoy_period = None
        if report_year and report_period:
            prev_yoy_period = f"{report_year - 1}{report_period}"

        # 计算环比期（上一季度）
        prev_qoq_period = previous_quarter_period(full_current_period)

        # ========== YoY计算（同比：同报告期，上一年）==========

        # 核心指标YoY
        if core.get("operating_revenue_yoy_growth") is None and prev_yoy_period:
            curr_val = core.get("total_operating_revenue") or inc.get("total_operating_revenue")
            prev_row = get_table_row(history, "core_performance_indicators_sheet", prev_yoy_period)
            prev_val = prev_row.get("total_operating_revenue")
            if prev_val is None:
                prev_row = get_table_row(history, "income_sheet", prev_yoy_period)
                prev_val = prev_row.get("total_operating_revenue")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    core["operating_revenue_yoy_growth"] = yoy

        if core.get("net_profit_yoy_growth") is None and prev_yoy_period:
            curr_val = core.get("net_profit_10k_yuan") or inc.get("net_profit")
            prev_row = get_table_row(history, "core_performance_indicators_sheet", prev_yoy_period)
            prev_val = prev_row.get("net_profit_10k_yuan")
            if prev_val is None:
                prev_row = get_table_row(history, "income_sheet", prev_yoy_period)
                prev_val = prev_row.get("net_profit")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    core["net_profit_yoy_growth"] = yoy

        # 资产负债表YoY
        if bal.get("asset_total_assets_yoy_growth") is None and prev_yoy_period:
            curr_val = bal.get("asset_total_assets")
            prev_row = get_table_row(history, "balance_sheet", prev_yoy_period)
            prev_val = prev_row.get("asset_total_assets")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    bal["asset_total_assets_yoy_growth"] = yoy

        if bal.get("liability_total_liabilities_yoy_growth") is None and prev_yoy_period:
            curr_val = bal.get("liability_total_liabilities")
            prev_row = get_table_row(history, "balance_sheet", prev_yoy_period)
            prev_val = prev_row.get("liability_total_liabilities")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    bal["liability_total_liabilities_yoy_growth"] = yoy

        # 现金流量表YoY
        if cf.get("net_cash_flow_yoy_growth") is None and prev_yoy_period:
            curr_val = cf.get("net_cash_flow")
            prev_row = get_table_row(history, "cash_flow_sheet", prev_yoy_period)
            prev_val = prev_row.get("net_cash_flow")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    cf["net_cash_flow_yoy_growth"] = yoy

        # 利润表YoY
        if inc.get("net_profit_yoy_growth") is None and prev_yoy_period:
            curr_val = inc.get("net_profit")
            prev_row = get_table_row(history, "income_sheet", prev_yoy_period)
            prev_val = prev_row.get("net_profit")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    inc["net_profit_yoy_growth"] = yoy

        if inc.get("operating_revenue_yoy_growth") is None and prev_yoy_period:
            curr_val = inc.get("total_operating_revenue")
            prev_row = get_table_row(history, "income_sheet", prev_yoy_period)
            prev_val = prev_row.get("total_operating_revenue")
            if curr_val is not None and prev_val is not None and prev_val != 0:
                yoy = round((curr_val - prev_val) / abs(prev_val) * 100, 4)
                if abs(yoy) <= MAX_REASONABLE_GROWTH_ABS:
                    inc["operating_revenue_yoy_growth"] = yoy

        # ========== QoQ计算（环比：上一季度）==========

        if prev_qoq_period:
            if core.get("operating_revenue_qoq_growth") is None:
                curr_rev = quarter_value(history, parsed, "core_performance_indicators_sheet", "total_operating_revenue", full_current_period)
                prev_rev = quarter_value(history, parsed, "core_performance_indicators_sheet", "total_operating_revenue", prev_qoq_period)
                if curr_rev is None:
                    curr_rev = quarter_value(history, parsed, "income_sheet", "total_operating_revenue", full_current_period)
                if prev_rev is None:
                    prev_rev = quarter_value(history, parsed, "income_sheet", "total_operating_revenue", prev_qoq_period)
                qoq = calc_growth_pct(curr_rev, prev_rev, "operating_revenue_qoq_growth")
                if qoq is not None:
                    core["operating_revenue_qoq_growth"] = qoq
            if core.get("net_profit_qoq_growth") is None:
                curr_np = quarter_value(history, parsed, "core_performance_indicators_sheet", "net_profit_10k_yuan", full_current_period)
                prev_np = quarter_value(history, parsed, "core_performance_indicators_sheet", "net_profit_10k_yuan", prev_qoq_period)
                if curr_np is None:
                    curr_np = quarter_value(history, parsed, "income_sheet", "net_profit", full_current_period)
                if prev_np is None:
                    prev_np = quarter_value(history, parsed, "income_sheet", "net_profit", prev_qoq_period)
                qoq = calc_growth_pct(curr_np, prev_np, "net_profit_qoq_growth")
                if qoq is not None:
                    core["net_profit_qoq_growth"] = qoq

        # ========== 从分季度数据计算FY/Q1的QoQ ==========
        # 当 quarter_value 返回 None 时（FY和Q1无法直接计算），尝试从分季度数据推算
        if quarterly_data and report_period in ("FY", "Q1"):
            # FY 的 QoQ: Q4环比Q3
            if core.get("operating_revenue_qoq_growth") is None and report_period == "FY":
                curr_rev = core.get("total_operating_revenue") or inc.get("total_operating_revenue")
                # 从分季度数据计算 Q4 单季度值
                if curr_rev is not None and quarterly_data.get("revenue"):
                    q1 = quarterly_data["revenue"].get("Q1")
                    q2 = quarterly_data["revenue"].get("Q2")
                    q3 = quarterly_data["revenue"].get("Q3")
                    if q1 and q2 and q3:
                        q4 = curr_rev - q1 - q2 - q3
                        if q4 > 0 and q3 > 0:
                            qoq = round((q4 / q3 - 1) * 100, 4)
                            if abs(qoq) <= MAX_REASONABLE_GROWTH_ABS:
                                core["operating_revenue_qoq_growth"] = qoq

            # FY 的净利润 QoQ
            if core.get("net_profit_qoq_growth") is None and report_period == "FY":
                curr_np = core.get("net_profit_10k_yuan") or inc.get("net_profit")
                if curr_np is not None and quarterly_data.get("net_profit"):
                    q1 = quarterly_data["net_profit"].get("Q1")
                    q2 = quarterly_data["net_profit"].get("Q2")
                    q3 = quarterly_data["net_profit"].get("Q3")
                    if q1 is not None and q2 is not None and q3 is not None:
                        q4 = curr_np - q1 - q2 - q3
                        if q3 != 0:
                            qoq = round((q4 / q3 - 1) * 100, 4)
                            if abs(qoq) <= MAX_REASONABLE_GROWTH_ABS:
                                core["net_profit_qoq_growth"] = qoq

            # Q1 的 QoQ: Q1环比Q4（上年度）
            if core.get("operating_revenue_qoq_growth") is None and report_period == "Q1":
                curr_rev = core.get("total_operating_revenue") or inc.get("total_operating_revenue")
                q1 = quarterly_data.get("revenue", {}).get("Q1")
                if curr_rev is not None and q1 is not None:
                    # Q1 = curr_rev，说明Q1是单季度值
                    # 需要从历史数据获取上年度Q4
                    prev_year = report_year - 1 if report_year else None
                    if prev_year:
                        prev_q4_period = f"{prev_year}Q3"  # Q3是Q4之前的最后一个累计期
                        prev_q3 = quarter_value(history, parsed, "core_performance_indicators_sheet", "total_operating_revenue", prev_q4_period)
                        if prev_q3 is None:
                            prev_q3 = quarter_value(history, parsed, "income_sheet", "total_operating_revenue", prev_q4_period)
                        prev_fy_period = f"{prev_year}FY"
                        prev_fy = quarter_value(history, parsed, "core_performance_indicators_sheet", "total_operating_revenue", prev_fy_period)
                        if prev_fy is None:
                            prev_fy = quarter_value(history, parsed, "income_sheet", "total_operating_revenue", prev_fy_period)
                        if prev_fy is not None and prev_q3 is not None:
                            prev_q4 = prev_fy - prev_q3
                            if prev_q4 > 0 and q1 > 0:
                                qoq = round((q1 / prev_q4 - 1) * 100, 4)
                                if abs(qoq) <= MAX_REASONABLE_GROWTH_ABS:
                                    core["operating_revenue_qoq_growth"] = qoq

            # Q1 的净利润 QoQ
            if core.get("net_profit_qoq_growth") is None and report_period == "Q1":
                curr_np = core.get("net_profit_10k_yuan") or inc.get("net_profit")
                q1 = quarterly_data.get("net_profit", {}).get("Q1")
                if curr_np is not None and q1 is not None:
                    prev_year = report_year - 1 if report_year else None
                    if prev_year:
                        prev_q4_period = f"{prev_year}Q3"
                        prev_q3 = quarter_value(history, parsed, "core_performance_indicators_sheet", "net_profit_10k_yuan", prev_q4_period)
                        if prev_q3 is None:
                            prev_q3 = quarter_value(history, parsed, "income_sheet", "net_profit", prev_q4_period)
                        prev_fy_period = f"{prev_year}FY"
                        prev_fy = quarter_value(history, parsed, "core_performance_indicators_sheet", "net_profit_10k_yuan", prev_fy_period)
                        if prev_fy is None:
                            prev_fy = quarter_value(history, parsed, "income_sheet", "net_profit", prev_fy_period)
                        if prev_fy is not None and prev_q3 is not None:
                            prev_q4 = prev_fy - prev_q3
                            if prev_q4 != 0 and q1 != 0:
                                qoq = round((q1 / prev_q4 - 1) * 100, 4)
                                if abs(qoq) <= MAX_REASONABLE_GROWTH_ABS:
                                    core["net_profit_qoq_growth"] = qoq

    parsed.pop("__current_report_period__", None)
    for t, fs in schema.items():
        ext, sn = parsed[t]
        parsed[t] = (filter_to_schema(ext, fs), sn)

    return parsed


# ============================================================
# 主解析器 (md_parser.py)
# ============================================================

def score_core_candidate(row: Dict, seen: Dict, sec: str) -> int:
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


def enrich_table_fields(table: str, row: Dict, seen: Dict) -> Dict:
    """丰富表格字段"""
    if table == "core_performance_indicators_sheet":
        if "total_operating_revenue" in seen and "operating_revenue_yoy_growth" not in row:
            vs = seen["total_operating_revenue"]
            yoy = vs[2] if len(vs) >= 3 and is_plausible_growth(vs[2]) else calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None, "operating_revenue_yoy_growth"
            )
            if is_plausible_growth(yoy):
                row["operating_revenue_yoy_growth"] = yoy
        if "net_profit_10k_yuan" in seen and "net_profit_yoy_growth" not in row:
            vs = seen["net_profit_10k_yuan"]
            yoy = vs[2] if len(vs) >= 3 and is_plausible_growth(vs[2]) else calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None, "net_profit_yoy_growth"
            )
            if is_plausible_growth(yoy):
                row["net_profit_yoy_growth"] = yoy
        if "net_profit_excl_non_recurring" in seen and "net_profit_excl_non_recurring_yoy" not in row:
            vs = seen["net_profit_excl_non_recurring"]
            yoy = vs[2] if len(vs) >= 3 and is_plausible_growth(vs[2]) else calc_growth(
                vs[0], vs[1] if len(vs) >= 2 else None, "net_profit_excl_non_recurring_yoy"
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
            growth = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None, "asset_total_assets_yrowth")
            if growth is not None:
                row["asset_total_assets_yoy_growth"] = growth
        if "liability_total_liabilities" in seen and "liability_total_liabilities_yoy_growth" not in row:
            vs = seen["liability_total_liabilities"]
            growth = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None, "liability_total_liabilities_yoy_growth")
            if growth is not None:
                row["liability_total_liabilities_yoy_growth"] = growth
        if (
            row.get("liability_total_liabilities") is not None
            and row.get("asset_total_assets") not in (None, 0)
            and "asset_liability_ratio" not in row
        ):
            row["asset_liability_ratio"] = round(row["liability_total_liabilities"] / row["asset_total_assets"] * 100, 4)

    if table == "income_sheet":
        if "net_profit" in seen and "net_profit_yoy_growth" not in row:
            vs = seen["net_profit"]
            yoy = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None, "net_profit_yoy_growth")
            if is_plausible_growth(yoy):
                row["net_profit_yoy_growth"] = yoy
        if "total_operating_revenue" in seen and "operating_revenue_yoy_growth" not in row:
            vs = seen["total_operating_revenue"]
            yoy = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None, "operating_revenue_yoy_growth")
            if is_plausible_growth(yoy):
                row["operating_revenue_yoy_growth"] = yoy

    if table == "cash_flow_sheet":
        if "net_cash_flow" in seen and "net_cash_flow_yoy_growth" not in row:
            vs = seen["net_cash_flow"]
            growth = calc_growth(vs[0], vs[1] if len(vs) >= 2 else None, "net_cash_flow_yoy_growth")
            if growth is not None:
                row["net_cash_flow_yoy_growth"] = growth
        for a, b in [
            ("operating_cf_net_amount", "operating_cf_ratio_of_net_cf"),
            ("investing_cf_net_amount", "investing_cf_ratio_of_net_cf"),
            ("financing_cf_net_amount", "financing_cf_ratio_of_net_cf"),
        ]:
            if row.get(a) is not None and row.get("net_cash_flow") not in (None, 0) and b not in row:
                ratio = round(row[a] / row["net_cash_flow"] * 100, 4)
                row[b] = ratio

    return {k: v for k, v in row.items() if v is not None}


def parse_section_rows(sec: str, table: str, fs, am: dict) -> Tuple[Dict, Dict, Dict]:
    """解析区块行"""
    row = {}
    seen = {}
    field_evidence = {}

    uf = detect_unit(sec)
    rows, header, body = parse_table_matrix(sec)
    if not rows:
        return row, seen, field_evidence

    header_cols = analyze_header_columns(header)
    has_adjusted_layout = any(
        "调整后" in normalize_text_cell(cell)
        for row_cells in rows[:3]
        for cell in row_cells
    )

    yoy_idx = next((c.index for c in header_cols if c.col_type == "yoy"), None)

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

        if table == "core_performance_indicators_sheet":
            if tf in PERIOD_END_FIELDS:
                field_type = "period_end"
            elif tf in CUMULATIVE_FIELDS:
                field_type = "cumulative"
            elif tf in RATIO_FIELDS:
                field_type = "ratio"
            else:
                field_type = "other"
        else:
            field_type = "other"

        cur_idx, pre_idx, _ = select_col_for_field(header_cols, field_type)
        curr = value_by_idx(cs, cur_idx, uf) if cur_idx is not None else None
        prev = value_by_idx(cs, pre_idx, uf) if pre_idx is not None else None
        yoy = None

        if table == "core_performance_indicators_sheet":
            smart_curr, smart_prev, smart_yoy = extract_core_row_metrics(cs, uf)
            adjusted_curr, adjusted_prev, adjusted_yoy = extract_adjusted_core_metrics(cs, header, uf)
            if not has_adjusted_layout:
                adjusted_curr = adjusted_prev = adjusted_yoy = None

            if field_type == "cumulative":
                cum_col = next((c for c in header_cols if c.col_type == "cumulative"), None)
                if cum_col:
                    cum_val = value_by_idx(cs, cum_col.index, uf)
                    if cum_val is not None:
                        curr = cum_val
                    prev_col = next((c for c in header_cols if c.col_type == "previous" and "上年" in c.norm), None)
                    if prev_col:
                        prev = value_by_idx(cs, prev_col.index, uf)

            elif field_type == "period_end":
                pe_col = next((c for c in header_cols if c.col_type == "period_end"), None)
                if pe_col:
                    pe_val = value_by_idx(cs, pe_col.index, uf)
                    if pe_val is not None:
                        curr = pe_val
                    prev_col = next((c for c in header_cols if c.col_type == "previous" and "上年度" in c.norm), None)
                    if prev_col:
                        prev = value_by_idx(cs, prev_col.index, uf)

            elif field_type == "ratio":
                if smart_curr is not None:
                    curr = smart_curr
                if smart_prev is not None:
                    prev = smart_prev
                if smart_yoy is not None:
                    yoy = smart_yoy

            if curr is None and smart_curr is not None:
                curr = smart_curr
            if prev is None and smart_prev is not None:
                prev = smart_prev

            raw_yoy = normalize_text_cell(cs[yoy_idx]) if yoy_idx is not None and yoy_idx < len(cs) else ""
            if adjusted_yoy is not None and has_adjusted_layout:
                yoy = adjusted_yoy
            elif raw_yoy not in {"", "不适用", "N/A"} and yoy_idx is not None and yoy_idx != cur_idx:
                yoy = clean_num(raw_yoy, 1.0)
            elif adjusted_yoy is not None:
                yoy = adjusted_yoy
            elif smart_yoy is not None:
                yoy = smart_yoy
            elif curr is not None and prev is not None:
                yoy = calc_growth(curr, prev)

            if yoy is not None and is_plausible_growth(yoy):
                if tf == "total_operating_revenue" and "operating_revenue_yoy_growth" not in row:
                    row["operating_revenue_yoy_growth"] = round(yoy, 4)
                if tf == "net_profit_10k_yuan" and "net_profit_yoy_growth" not in row:
                    row["net_profit_yoy_growth"] = round(yoy, 4)
                if tf == "net_profit_excl_non_recurring" and "net_profit_excl_non_recurring_yoy" not in row:
                    row["net_profit_excl_non_recurring_yoy"] = round(yoy, 4)

        vals = [curr] if curr is not None else []
        if prev is not None:
            vals.append(prev)
        if yoy is not None and is_plausible_growth(yoy):
            vals.append(yoy)

        if not vals:
            vals = derive_values_from_row(cs, uf, table, tf, curr_is_none=(curr is None))

        vals = [v for v in vals if v is not None]
        if not vals:
            if table == "balance_sheet" and tf in BALANCE_SHEET_ZERO_FIELDS:
                row.setdefault(tf, 0)
            continue

        # 保存YoY字段到seen字典，供enrich_table_fields使用
        if yoy is not None and is_plausible_growth(yoy):
            if tf == "total_operating_revenue":
                seen.setdefault("operating_revenue_yoy_growth", []).append(yoy)
            elif tf == "net_profit_10k_yuan" or tf == "net_profit":
                seen.setdefault("net_profit_yoy_growth", []).append(yoy)
            elif tf == "net_profit_excl_non_recurring":
                seen.setdefault("net_profit_excl_non_recurring_yoy", []).append(yoy)
            elif tf == "asset_total_assets":
                seen.setdefault("asset_total_assets_yoy_growth", []).append(yoy)
            elif tf == "liability_total_liabilities":
                seen.setdefault("liability_total_liabilities_yoy_growth", []).append(yoy)
            elif tf == "net_cash_flow":
                seen.setdefault("net_cash_flow_yoy_growth", []).append(yoy)

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


def parse_md(md: str, table: str, fs) -> Tuple[Dict, List, Dict]:
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
# Markdown解析入库 (md_to_db.py)
# ============================================================

def parse_schema(x: Path) -> Dict[str, List[Field]]:
    """解析Schema Excel"""
    wb = openpyxl.load_workbook(x, data_only=True)
    r = {}
    for cn, t in TABLE_NAME_MAP.items():
        fs = []
        for row in wb[cn].iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                fs.append(
                    Field(
                        field_name=str(row[0]).strip() if row[0] else '',
                        cn_name=str(row[1]).strip() if len(row) > 1 and row[1] else str(row[0]).strip() if row[0] else '',
                        data_type=str(row[2]).strip() if len(row) > 2 and row[2] else 'TEXT'
                    )
                )
        r[t] = fs
    return r


def load_company_map(x: Path) -> Dict[str, str]:
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
            abbr = ""
            if ai is not None and ai < len(row) and row[ai]:
                abbr = str(row[ai]).strip()
            if abbr.isdigit() and len(abbr) >= 6:
                if ni is not None and ni < len(row) and row[ni]:
                    full_name = str(row[ni]).strip()
                    for suffix in ["股份有限公司", "有限公司", "集团"]:
                        if suffix in full_name:
                            abbr = full_name.split(suffix)[0]
                            if "（" in full_name:
                                abbr = full_name.split("（")[0]
                            break
                    if not abbr or abbr.isdigit():
                        abbr = ""
            if not abbr and ni is not None and ni < len(row) and row[ni]:
                full_name = str(row[ni]).strip()
                for suffix in ["股份有限公司", "有限公司"]:
                    if suffix in full_name:
                        abbr = full_name.split(suffix)[0]
                        break
            m[c] = abbr
    return m


def init_db(output_root: Path, schema: Dict) -> sqlite3.Connection:
    """初始化数据库"""
    conn = sqlite3.connect(output_root / "financial.db")
    ensure_tables(conn, schema)
    return conn


def parse_md_to_tables(
    md: str,
    schema: Dict,
    q: Dict,
    pdf_name: str = "",
    report_year: int = None,
    report_period: str = "",
    parsed_history: Dict = None
) -> Dict[str, Any]:
    """解析Markdown到表格数据

    Args:
        md: Markdown内容
        schema: 表结构定义
        q: 质量统计对象
        pdf_name: PDF文件名
        report_year: 报告年份
        report_period: 报告期（FY/HY/Q1/Q3，不带年份）
        parsed_history: 历史解析数据
    """
    parsed = {}
    evidence_by_table = {}

    for table_name, fields in schema.items():
        ext, snippets, evidence = parse_md(md, table_name, fields)
        ext = filter_to_schema(ext, fields)

        cov, hit, total = check_table_quality(ext, fields, table_name, pdf_name, report_period, q, evidence)
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

    # 提取分季度数据，用于计算FY/Q1的QoQ
    quarterly_data = extract_quarterly_data(md)

    parsed = enrich_cross_table_metrics(
        parsed,
        schema,
        parsed_history=parsed_history,
        current_report_year=report_year,
        current_report_period=report_period,
        quarterly_data=quarterly_data,
    )
    check_cross_table_quality(parsed, pdf_name, report_period, q, evidence_by_table)
    return parsed


def write_tables_to_db(conn: sqlite3.Connection, meta: Meta, parsed: Dict) -> None:
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
        for snippet in snippets[:3]:
            cur = conn.execute(
                "SELECT id FROM raw_table_snippets WHERE stock_code=? AND report_year=? AND report_period=? AND target_table=? AND source_section=? AND snippet=?",
                (meta.stock_code, meta.report_year, meta.report_period, table_name, table_name, snippet)
            )
            if cur.fetchone() is None:
                conn.execute(
                    "INSERT INTO raw_table_snippets(stock_code,report_year,report_period,target_table,source_section,snippet,created_at) VALUES (?,?,?,?,?,?,?)",
                    (
                        meta.stock_code, meta.report_year, meta.report_period,
                        table_name, table_name, snippet, dt.datetime.now().isoformat(),
                    ),
                )


def rebuild_stock_qoq_metrics(conn: sqlite3.Connection, schema: Dict, stock_code: str, stock_period_history: Dict) -> None:
    """重建股票的环比指标

    Args:
        conn: 数据库连接
        schema: 表结构定义
        stock_code: 股票代码
        stock_period_history: 历史数据（键为完整报告期字符串如"2023Q1"）
    """
    def sort_key(key):
        """排序键函数，支持完整报告期字符串"""
        if isinstance(key, tuple) and len(key) == 2:
            year, period = key
            return (year or 0, QUARTER_SEQUENCE.get(period, 0))
        # 支持字符串格式 "2023Q1" 或 "2023FY"
        import re
        match = re.match(r'^(\d{4})(Q[1234]|HY|FY)$', key) if isinstance(key, str) else None
        if match:
            year = int(match.group(1))
            period = match.group(2)
            return (year, QUARTER_SEQUENCE.get(period, 0))
        return (0, 0)

    ordered_periods = sorted(stock_period_history.keys(), key=sort_key)
    rolling_history = {}

    for history_key in ordered_periods:
        # 解析完整报告期
        if isinstance(history_key, tuple) and len(history_key) == 2:
            report_year, report_period = history_key
        else:
            # 支持字符串格式
            import re
            match = re.match(r'^(\d{4})(Q[1234]|HY|FY)$', history_key) if isinstance(history_key, str) else None
            if match:
                report_year = int(match.group(1))
                report_period = match.group(2)
            else:
                report_year, report_period = None, history_key

        parsed = {
            table_name: (dict(ext), list(snippets))
            for table_name, (ext, snippets) in stock_period_history[history_key].items()
        }

        enriched = enrich_cross_table_metrics(
            parsed,
            schema,
            parsed_history=rolling_history,
            current_report_year=report_year,
            current_report_period=report_period,
        )

        qoq_fields = ['operating_revenue_qoq_growth', 'net_profit_qoq_growth']
        yoy_fields = ['operating_revenue_yoy_growth', 'net_profit_yoy_growth']

        for table_name, (ext, _) in enriched.items():
            if report_year is None or report_period is None:
                continue
            cur = conn.execute(
                f'SELECT * FROM "{table_name}" WHERE stock_code=? AND report_year=? AND report_period=?',
                (stock_code, report_year, report_period)
            ).fetchone()

            if cur:
                existing = dict(zip([d[0] for d in conn.execute(f'SELECT * FROM "{table_name}" WHERE 1=0').description], cur))
                update_fields = {}
                for field in qoq_fields:
                    if field in ext and ext[field] is not None:
                        update_fields[field] = ext[field]

                # 从income_sheet补充YoY字段
                if table_name == "core_performance_indicators_sheet":
                    inc_ext = enriched.get("income_sheet", ({}, []))[0]
                    for field in yoy_fields:
                        if field in inc_ext and inc_ext[field] is not None:
                            if existing.get(field) is None:
                                update_fields[field] = inc_ext[field]

                if update_fields:
                    set_clause = ", ".join([f'"{k}"=?' for k in update_fields.keys()])
                    conn.execute(
                        f'UPDATE "{table_name}" SET {set_clause} WHERE stock_code=? AND report_year=? AND report_period=?',
                        list(update_fields.values()) + [stock_code, report_year, report_period]
                    )

        # 将元组键转换为字符串键，兼容 build_history_lookup
        if isinstance(history_key, tuple) and len(history_key) == 2:
            full_key = f"{history_key[0]}{history_key[1]}"
        else:
            full_key = history_key
        rolling_history[full_key] = {
            table_name: (dict(ext), list(snippets))
            for table_name, (ext, snippets) in enriched.items()
        }


def process_pdf_to_db(
    pdf: Path,
    md_cache_root: Path,
    conn: sqlite3.Connection,
    schema: Dict,
    company_map: Dict[str, str],
    quality_stats: Dict,
    parsed_history: Dict = None,
) -> tuple:
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
    print(f"    [阶段2/4] 识别元数据：stock_code={meta.stock_code or '-'}, report_period={meta.report_period or '-'}")

    if not meta.report_period or not meta.stock_code:
        log(conn, pdf.name, "failed", "元数据识别失败")
        quality_stats["failed"] += 1
        return meta, None

    print("    [阶段3/4] 解析Markdown：开始")
    stock_history = parsed_history.get(meta.stock_code, {}) if parsed_history else {}
    parsed = parse_md_to_tables(
        md, schema, quality_stats, pdf.name, meta.report_year, meta.report_period, parsed_history=stock_history,
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


def dump_quality(out: Path, q: Dict) -> None:
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

    (out / "quality_report.json").write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / "quality_issues.json").write_text(json.dumps(q.get("issues", []), ensure_ascii=False, indent=2), encoding="utf-8")


# ============================================================
# 入口程序 (run_md_to_db.py)
# ============================================================

def is_summary_file(name: str) -> bool:
    """判断是否为摘要文件名"""
    return any(
        k in name
        for k in [
            "报告摘要", "年度报告摘要", "半年度报告摘要", "季度报告摘要",
            "一季度报告摘要", "三季度报告摘要", "摘要版",
        ]
    )


def build_company_map_from_dirs(md_dirs) -> dict:
    """从Markdown目录构建公司映射"""
    company_map = {}
    for md_dir in md_dirs:
        name = md_dir.name if isinstance(md_dir, Path) else str(md_dir)
        m = re.search(r"(?<!\d)(\d{6})(?!\d)", name)
        if m:
            stock_code = m.group(1)
            left = name.split("：")[0].strip() if "：" in name else name.split("_")[0].strip()
            if left and left not in company_map.values():
                company_map[stock_code] = left.split("（")[0].split("(")[0].strip()
    return company_map


def find_md_dirs(base_dir: Path) -> list:
    """递归查找所有包含report.md的目录"""
    md_dirs = []
    for item in base_dir.iterdir():
        if item.is_dir():
            if (item / "report.md").exists():
                md_dirs.append(item)
            else:
                md_dirs.extend(find_md_dirs(item))
    return md_dirs


def process_md_dir(
    md_dir: Path,
    conn: sqlite3.Connection,
    schema: dict,
    company_map: dict,
    quality_stats: dict,
    parsed_history: dict = None,
):
    """处理Markdown目录到数据库"""
    md_name = md_dir.name

    if is_summary_file(md_name):
        print("    [skip] summary detected by filename")
        log(conn, md_name, "skipped", "文件名命中摘要规则，跳过")
        quality_stats.setdefault("skipped", 0)
        quality_stats["skipped"] += 1
        return None, None

    md_path = md_dir / "report.md"
    if not md_path.exists():
        log(conn, md_name, "failed", "未找到 report.md")
        quality_stats["failed"] += 1
        return None, None

    md = md_path.read_text(encoding="utf-8", errors="ignore")
    if detect_summary_from_md(md):
        print("    [skip] summary detected by markdown content")
        log(conn, md_name, "skipped", "Markdown内容判定为摘要，跳过")
        quality_stats.setdefault("skipped", 0)
        quality_stats["skipped"] += 1
        return None, None

    meta = detect_meta(md_dir, company_map, md)
    if meta.stock_code and meta.stock_code in company_map:
        meta = Meta(
            stock_code=meta.stock_code,
            stock_abbr=company_map[meta.stock_code],
            report_year=meta.report_year,
            report_type=meta.report_type,
            report_period=meta.report_period,
            source_file_name=meta.source_file_name or md_name
        )
    print(f"    [阶段2/4] 识别元数据：stock_code={meta.stock_code or '-'}, report_period={meta.report_period or '-'}")

    if not meta.report_period or not meta.stock_code:
        log(conn, md_name, "failed", "元数据识别失败")
        quality_stats["failed"] += 1
        return meta, None

    print("    [阶段3/4] 解析Markdown：开始")
    stock_history = parsed_history.get(meta.stock_code, {}) if parsed_history else {}
    parsed = parse_md_to_tables(
        md, schema, quality_stats, md_name, meta.report_year, meta.report_period, parsed_history=stock_history,
    )
    hit_total = sum(len(ext) for ext, _ in parsed.values())
    print(f"    [阶段3/4] 解析Markdown：完成, 命中字段={hit_total}")

    print("    [阶段4/4] 写入数据库：开始")
    write_tables_to_db(conn, meta, parsed)
    print("    [阶段4/4] 写入数据库：完成")

    if hit_total:
        log(conn, md_name, "ok", "处理完成")
    else:
        log(conn, md_name, "warning", "仅写入元数据，未命中业务字段")
    quality_stats["ok"] += 1
    return meta, parsed


def main():
    """主函数"""
    MD_DIR = Path(r"d:\14 Tedicup\FS\data\md\financial report")
    DB_OUTPUT = Path(r"d:\14 Tedicup\FS\data\financial.db")
    SCHEMA_XLSX = Path(r"d:\14 Tedicup\全部数据\正式数据\附件3：数据库-表名及字段说明.xlsx")
    COMPANY_XLSX = Path(r"d:\14 Tedicup\全部数据\正式数据\附件1：中药上市公司基本信息（截至到2025年12月22日）.xlsx")

    print("=" * 60)
    print("财务报表 Markdown 解析入库工具 (整合版)")
    print("=" * 60)
    print(f"Markdown源目录: {MD_DIR}")
    print(f"数据库输出路径: {DB_OUTPUT}")
    print(f"Schema来源: {SCHEMA_XLSX}")

    if not MD_DIR.exists():
        print(f"[错误] Markdown目录不存在: {MD_DIR}")
        return

    schema = parse_schema(SCHEMA_XLSX)
    print(f"\nSchema包含 {sum(len(fields) for fields in schema.values())} 个字段")

    all_md_dirs = find_md_dirs(MD_DIR)

    if COMPANY_XLSX.exists():
        company_map = load_company_map(COMPANY_XLSX)
        print(f"从Excel加载 {len(company_map)} 个公司映射")
    else:
        company_map = build_company_map_from_dirs(all_md_dirs)
        print(f"从目录名构建 {len(company_map)} 个公司映射")

    DB_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    if DB_OUTPUT.exists():
        try:
            DB_OUTPUT.unlink()
            print("已删除旧数据库")
        except Exception:
            pass

    conn = sqlite3.connect(str(DB_OUTPUT))
    ensure_tables(conn, schema)
    print(f"数据库已初始化: {DB_OUTPUT}")

    md_dirs = find_md_dirs(MD_DIR)
    print(f"\n找到 {len(md_dirs)} 个Markdown目录")

    quality_stats = {"ok": 0, "failed": 0, "skipped": 0, "blocked": 0, "warnings": 0, "errors": 0, "tables": {}, "issues": []}
    parsed_history = {}

    total = len(md_dirs)
    for index, md_dir in enumerate(sorted(md_dirs), 1):
        print(f"\n[{index}/{total}] 处理: {md_dir.name}")
        try:
            meta, parsed = process_md_dir(
                md_dir, conn, schema, company_map, quality_stats, parsed_history=parsed_history,
            )
            if meta and parsed and meta.stock_code:
                history_key = (meta.report_year, meta.report_period)
                parsed_history.setdefault(meta.stock_code, {})[history_key] = {
                    table_name: (dict(ext), list(snippets))
                    for table_name, (ext, snippets) in parsed.items()
                }
            conn.commit()
        except Exception as exc:
            import traceback
            log(conn, md_dir.name, "failed", f"异常:{exc}")
            quality_stats["failed"] += 1
            conn.commit()
            print(f"  -> 异常: {exc}")
            traceback.print_exc()

    print("\n[后处理] 重建环比指标...")
    for stock_code, stock_period_history in parsed_history.items():
        rebuild_stock_qoq_metrics(conn, schema, stock_code, stock_period_history)
    conn.commit()
    print("[后处理] 环比指标重建完成")

    conn.close()
    dump_quality(DB_OUTPUT.parent, quality_stats)

    print("\n" + "=" * 60)
    print("完成!")
    print(f"  成功: {quality_stats['ok']}")
    print(f"  失败: {quality_stats['failed']}")
    print(f"  跳过: {quality_stats.get('skipped', 0)}")
    print(f"  阻断: {quality_stats.get('blocked', 0)}")
    print(f"  警告: {quality_stats.get('warnings', 0)}")
    print(f"  错误: {quality_stats.get('errors', 0)}")
    print(f"数据库: {DB_OUTPUT}")
    print("=" * 60)


if __name__ == "__main__":
    main()
