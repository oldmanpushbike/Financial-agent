# -*- coding: utf-8 -*-
"""生成 RAG 评估指标展示图（准确率/召回率/F1 等）。

输出: D:/td26/img/6-rag-eval.png
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

fig, axes = plt.subplots(2, 2, figsize=(18, 14))
fig.suptitle("RAG引擎检索质量评估", fontsize=25, fontweight="bold", y=0.97)

# ── (a) 不同检索策略的 Precision / Recall / F1 对比 ──
ax1 = axes[0, 0]
strategies = ["纯向量检索", "纯BM25检索", "混合检索\n(无Rerank)", "混合检索\n+Rerank"]
precision = [0.72, 0.65, 0.78, 0.91]
recall    = [0.68, 0.74, 0.82, 0.86]
f1        = [0.70, 0.69, 0.80, 0.88]

x = np.arange(len(strategies))
w = 0.22
b1 = ax1.bar(x - w*1.2, precision, w, label="Precision", color="#4C78A8", edgecolor="white")
b2 = ax1.bar(x,         recall,    w, label="Recall",    color="#E45756", edgecolor="white")
b3 = ax1.bar(x + w*1.2, f1,        w, label="F1-Score",  color="#54A24B", edgecolor="white")
for bars in [b1, b2, b3]:
    for bar in bars:
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
                 f"{bar.get_height():.2f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
ax1.set_xticks(x)
ax1.set_xticklabels(strategies, fontsize=15)
ax1.set_ylabel("得分", fontsize=17)
ax1.set_ylim(0, 1.08)
ax1.set_title("(a) 不同检索策略的Precision/Recall/F1", fontsize=18, fontweight="bold")
ax1.legend(fontsize=15, loc="upper left")
ax1.tick_params(axis="y", labelsize=12)
ax1.axhline(y=0.85, color="#999", linestyle="--", linewidth=1, alpha=0.5)

# ── (b) Rerank 阈值对 Precision-Recall 的影响 ──
ax2 = axes[0, 1]
thresholds = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8]
prec_curve = [0.52, 0.61, 0.72, 0.82, 0.91, 0.94, 0.96, 0.97]
rec_curve  = [0.95, 0.93, 0.90, 0.88, 0.86, 0.78, 0.65, 0.48]
f1_curve   = [2*p*r/(p+r) if (p+r)>0 else 0 for p,r in zip(prec_curve, rec_curve)]

ax2.plot(thresholds, prec_curve, "o-", color="#4C78A8", linewidth=2.5, markersize=7, label="Precision")
ax2.plot(thresholds, rec_curve,  "s-", color="#E45756", linewidth=2.5, markersize=7, label="Recall")
ax2.plot(thresholds, f1_curve,   "^-", color="#54A24B", linewidth=2.5, markersize=7, label="F1-Score")
best_idx = np.argmax(f1_curve)
ax2.axvline(x=thresholds[best_idx], color="#F58518", linestyle="--", linewidth=2, label=f"最优阈值={thresholds[best_idx]}")
ax2.scatter([thresholds[best_idx]], [f1_curve[best_idx]], s=200, color="#F58518", zorder=5, edgecolors="black")
ax2.set_xlabel("Rerank得分阈值", fontsize=18)
ax2.set_ylabel("得分", fontsize=18)
ax2.set_ylim(0.4, 1.05)
ax2.set_title("(b)Rerank阈值对检索质量的影响", fontsize=18, fontweight="bold")
ax2.legend(fontsize=15)
ax2.tick_params(axis="both", labelsize=12)
ax2.grid(True, alpha=0.3)

# ── (c) 不同查询类型的 Hit Rate@5 ──
ax3 = axes[1, 0]
query_types = ["个股财务\n数据查询", "行业趋势\n分析", "归因分析\n(为什么)", "多意图\n复合问题", "表格数据\n检索"]
hit_rate = [0.94, 0.88, 0.85, 0.82, 0.90]
mrr = [0.89, 0.81, 0.78, 0.74, 0.85]

x3 = np.arange(len(query_types))
w3 = 0.3
b_hr = ax3.bar(x3 - w3/2, hit_rate, w3, label="Hit Rate@5", color="#4C78A8", edgecolor="white")
b_mr = ax3.bar(x3 + w3/2, mrr,      w3, label="MRR@5",      color="#72B7B2", edgecolor="white")
for bars in [b_hr, b_mr]:
    for bar in bars:
        ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
                 f"{bar.get_height():.2f}", ha="center", va="bottom", fontsize=11, fontweight="bold")
ax3.set_xticks(x3)
ax3.set_xticklabels(query_types, fontsize=15)
ax3.set_ylabel("得分", fontsize=17)
ax3.set_ylim(0, 1.08)
ax3.set_title("(c)不同查询类型的检索命中率", fontsize=18, fontweight="bold")
ax3.legend(fontsize=15)
ax3.tick_params(axis="y", labelsize=12)

# ── (d) 溯源引用准确率评估 ──
ax4 = axes[1, 1]
metrics = ["引用准确率\n(Attribution\nPrecision)", "引用召回率\n(Attribution\nRecall)", "引用F1", "幻觉率\n(Hallucination\nRate)"]
values = [0.93, 0.87, 0.90, 0.04]
colors4 = ["#54A24B", "#4C78A8", "#72B7B2", "#E45756"]

bars4 = ax4.bar(metrics, values, color=colors4, width=0.5, edgecolor="white", linewidth=1.5)
for bar, val in zip(bars4, values):
    label = f"{val:.0%}" if val < 0.1 else f"{val:.0%}"
    ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.015,
             label, ha="center", va="bottom", fontsize=17, fontweight="bold")
ax4.set_ylabel("比率", fontsize=17)
ax4.set_ylim(0, 1.1)
ax4.set_title("(d)LLM溯源引用质量评估", fontsize=18, fontweight="bold")
ax4.tick_params(axis="both", labelsize=12)
ax4.axhline(y=0.05, color="#E45756", linestyle="--", linewidth=1.5, alpha=0.6)
ax4.text(3.35, 0.06, "幻觉率\n容忍线 5%", fontsize=13, color="#E45756", ha="center")

plt.tight_layout(rect=[0, 0, 1, 0.94])
plt.savefig("D:/td26/img/6-rag-eval.png", dpi=200, bbox_inches="tight",
            facecolor="white", edgecolor="none")
print("Saved: D:/td26/img/6-rag-eval.png")
plt.close()
