# -*- coding: utf-8 -*-
"""数据可视化工具 — 基于 matplotlib，封装为 StructuredTool。

参考 B黎_1/agent/tools/visualizer.py，适配 LangGraph 状态流转。
"""
from __future__ import annotations

import json
import re
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

matplotlib.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "Arial Unicode MS"]
matplotlib.rcParams["axes.unicode_minus"] = False

from langchain_core.tools import StructuredTool
from pydantic import BaseModel, Field

# 图片输出目录
ASSET_DIR = Path(__file__).resolve().parent.parent / "result" / "img"
ASSET_DIR.mkdir(parents=True, exist_ok=True)


# ──────────────────────────────────────────────
# 绘图核心
# ──────────────────────────────────────────────

def _plot_line(data: Dict[str, Any]) -> None:
    labels = data.get("labels", [])
    if "datasets" in data:
        for ds in data["datasets"]:
            plt.plot(labels, ds.get("values", []), marker="o", label=ds.get("label", ""))
        plt.legend()
    else:
        plt.plot(labels, data.get("values", []), marker="o", color="b", label=data.get("label", ""))
        if data.get("label"):
            plt.legend()


def _plot_bar(data: Dict[str, Any]) -> None:
    import numpy as np

    labels = data.get("labels", [])
    if "datasets" in data:
        datasets = data["datasets"]
        x = np.arange(len(labels))
        width = 0.8 / max(len(datasets), 1)
        for i, ds in enumerate(datasets):
            plt.bar(x + (i - len(datasets) / 2 + 0.5) * width, ds.get("values", []), width, label=ds.get("label", ""))
        plt.xticks(x, labels)
        plt.legend()
    else:
        plt.bar(labels, data.get("values", []), color="skyblue")


def _plot_pie(data: Dict[str, Any]) -> None:
    values = data.get("values", [])
    labels = data.get("labels", [])
    if not values or len(values) != len(labels):
        raise ValueError("饼图数据 values 与 labels 长度不匹配")
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=140, shadow=False)
    plt.axis("equal")


_PLOT_FUNCS = {
    "line": _plot_line,
    "bar": _plot_bar,
    "pie": _plot_pie,
}


def _generate_chart(chart_type: str, title: str, data: Dict[str, Any], y_label: str = "") -> Optional[str]:
    """生成图表并保存为 PNG，返回相对路径；失败返回 None。"""
    if not data or (not data.get("values") and not data.get("datasets")):
        return None

    plot_fn = _PLOT_FUNCS.get(chart_type)
    if plot_fn is None:
        return None

    plt.figure(figsize=(10, 6))
    try:
        plot_fn(data)
        plt.title(title, fontsize=14)
        if y_label and chart_type != "pie":
            plt.ylabel(y_label)
        if chart_type != "pie":
            plt.grid(True, linestyle="--", alpha=0.6)

        filename = f"chart_{uuid.uuid4().hex[:8]}.png"
        filepath = ASSET_DIR / filename
        plt.savefig(filepath, bbox_inches="tight", dpi=150)
        plt.close()
        return str(filepath)
    except Exception:
        plt.close()
        return None


# ──────────────────────────────────────────────
# StructuredTool 定义
# ──────────────────────────────────────────────

class VisualizerInput(BaseModel):
    chart_type: str = Field(description="图表类型：line（折线图，趋势）、bar（柱状图，对比）、pie（饼图，占比）")
    title: str = Field(description="图表标题（应包含单位信息）")
    y_label: str = Field(default="", description="Y轴标签/单位，如 '亿元'、'万元'、'%'。饼图可留空。")
    data_json: str = Field(
        description=(
            "图表数据，JSON 字符串格式。"
            '单系列: {"labels":["A","B"],"values":[1,2]}；'
            '多系列: {"labels":["A","B"],"datasets":[{"label":"系列1","values":[1,2]},{"label":"系列2","values":[3,4]}]}'
        )
    )


def _visualizer_run(chart_type: str, title: str, data_json: str, y_label: str = "") -> str:
    """解析 data_json 并生成图表。"""
    try:
        data = json.loads(data_json)
    except json.JSONDecodeError as e:
        return f"数据解析失败: {e}"

    path = _generate_chart(chart_type, title, data, y_label)
    if path:
        return f"图表已保存: {path}"
    return "图表生成失败，请检查 chart_type 和 data_json 格式。"


data_visualizer_tool = StructuredTool.from_function(
    func=_visualizer_run,
    name="data_visualizer",
    description="根据结构化数据生成图表（折线图、柱状图、饼图），保存为 PNG 图片并返回路径。",
    args_schema=VisualizerInput,
)
