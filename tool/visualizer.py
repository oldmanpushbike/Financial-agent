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

_next_chart_name: str | None = None
_chart_prefix: str = ""
_chart_counter: int = 0


def set_chart_prefix(prefix: str):
    global _chart_prefix, _chart_counter
    _chart_prefix = prefix
    _chart_counter = 0


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


def _plot_radar(data: Dict[str, Any]) -> None:
    import numpy as np

    labels = data.get("labels", [])
    n = len(labels)
    if n < 3:
        raise ValueError("雷达图至少需要 3 个维度")
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]

    ax = plt.gca()
    ax.remove()
    ax = plt.gcf().add_subplot(111, polar=True)

    if "datasets" in data:
        for ds in data["datasets"]:
            vals = ds.get("values", [])[:n]
            vals += vals[:1]
            ax.plot(angles, vals, "o-", label=ds.get("label", ""))
            ax.fill(angles, vals, alpha=0.15)
        ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1))
    else:
        vals = data.get("values", [])[:n]
        vals += vals[:1]
        ax.plot(angles, vals, "o-", color="b")
        ax.fill(angles, vals, alpha=0.25, color="b")

    ax.set_thetagrids(np.degrees(angles[:-1]), labels)


def _plot_histogram(data: Dict[str, Any]) -> None:
    values = data.get("values", [])
    if not values:
        raise ValueError("直方图缺少 values 数据")
    bins = data.get("bins", 10)
    plt.hist(values, bins=bins, color="skyblue", edgecolor="black", alpha=0.7)
    plt.xlabel(data.get("x_label", ""))


def _plot_double_bar(data: Dict[str, Any]) -> None:
    import numpy as np

    labels = data.get("labels", [])
    datasets = data.get("datasets", [])
    if not datasets:
        values = data.get("values", [])
        if values:
            labels = labels[:len(values)]
            plt.barh(range(len(labels)), values)
            plt.yticks(range(len(labels)), labels)
            return
        raise ValueError("双条形图需要 datasets 或 values")
    max_len = max(len(ds.get("values", [])) for ds in datasets)
    labels = labels[:max_len]
    if len(datasets) == 1:
        plt.barh(range(len(labels)), datasets[0].get("values", []), label=datasets[0].get("label", ""))
        plt.yticks(range(len(labels)), labels)
        plt.legend()
        return
    x = np.arange(len(labels))
    width = 0.35
    plt.barh(x + width / 2, datasets[0].get("values", []), width, label=datasets[0].get("label", ""))
    plt.barh(x - width / 2, datasets[1].get("values", []), width, label=datasets[1].get("label", ""))
    plt.yticks(x, labels)
    plt.legend()


def _plot_scatter(data: Dict[str, Any]) -> None:
    x_vals = data.get("x_values", [])
    y_vals = data.get("y_values", [])
    if not x_vals or not y_vals:
        raise ValueError("散点图需要 x_values 和 y_values")
    point_labels = data.get("labels", [])
    plt.scatter(x_vals, y_vals, c="steelblue", alpha=0.7, edgecolors="black", s=60)
    plt.xlabel(data.get("x_label", ""))
    plt.ylabel(data.get("y_label", ""))
    for i, lbl in enumerate(point_labels):
        if i < len(x_vals):
            plt.annotate(lbl, (x_vals[i], y_vals[i]), textcoords="offset points",
                         xytext=(5, 5), fontsize=8)


def _plot_table(data: Dict[str, Any]) -> None:
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    if not headers or not rows:
        raise ValueError("表格需要 headers 和 rows")
    ax = plt.gca()
    ax.axis("off")
    table = ax.table(cellText=rows, colLabels=headers, loc="center", cellLoc="center")
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.5)


def _plot_boxplot(data: Dict[str, Any]) -> None:
    datasets = data.get("datasets", [])
    labels = data.get("labels", [])
    if not datasets:
        raise ValueError("箱线图缺少 datasets")
    box_data = [ds.get("values", []) for ds in datasets]
    bp = plt.boxplot(box_data, patch_artist=True)
    colors = ["#4C72B0", "#55A868", "#C44E52", "#8172B2", "#CCB974", "#64B5CD"]
    for i, patch in enumerate(bp["boxes"]):
        patch.set_facecolor(colors[i % len(colors)])
    if labels:
        plt.xticks(range(1, len(labels) + 1), labels)


_PLOT_FUNCS = {
    "line": _plot_line,
    "bar": _plot_bar,
    "pie": _plot_pie,
    "radar": _plot_radar,
    "histogram": _plot_histogram,
    "double_bar": _plot_double_bar,
    "scatter": _plot_scatter,
    "table": _plot_table,
    "boxplot": _plot_boxplot,
}


def _generate_chart(chart_type: str, title: str, data: Dict[str, Any], y_label: str = "") -> Optional[str]:
    """生成图表并保存为 PNG，返回相对路径；失败返回 None。"""
    has_data = (
        data.get("values") or data.get("datasets")
        or data.get("rows") or data.get("x_values")
    )
    if not data or not has_data:
        return None

    plot_fn = _PLOT_FUNCS.get(chart_type)
    if plot_fn is None:
        return None

    plt.figure(figsize=(10, 6))
    try:
        plot_fn(data)
        plt.title(title, fontsize=14)
        if y_label and chart_type not in ("pie", "table", "radar"):
            plt.ylabel(y_label)
        if chart_type not in ("pie", "table", "radar"):
            plt.grid(True, linestyle="--", alpha=0.6)

        global _chart_prefix, _chart_counter
        if _chart_prefix:
            _chart_counter += 1
            filename = f"{_chart_prefix}_{_chart_counter}.png"
        else:
            filename = f"chart_{uuid.uuid4().hex[:8]}.png"
        filepath = ASSET_DIR / filename
        plt.savefig(filepath, bbox_inches="tight", dpi=150)
        plt.close()
        return f"result/img/{filename}"
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
