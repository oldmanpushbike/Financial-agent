# -*- coding: utf-8 -*-
"""研报 Markdown 预处理：清洗噪音，为 RAG 向量化做准备。

用法:
  单文件:
    python preprocess.py "data/md/research report/个股研报/xxx/report.md"

  批量处理:
    python preprocess.py --all
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

# ── 项目根目录 ──────────────────────────────────────────
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent  # FS/
_MD_ROOT = _PROJECT_ROOT / "data" / "md" / "research report"
_OUT_ROOT = _PROJECT_ROOT / "data" / "RAG_md"

# ── 尾部截断：匹配到这些标题后，该行及之后全部丢弃 ──
_TAIL_CUT_RE = re.compile(
    r"^#+\s*(证券分析师声明|一般声明|法律声明|信息披露声明|投资评级说明"
    r"|免责声明|风险提示与免责声明|重要声明|分析师声明|评级说明"
    r"|利益披露与免责声明|特别声明|行业投资评级的说明)",
    re.MULTILINE,
)

# ── 逐行删除的噪音模式 ──
_LINE_NOISE_PATTERNS: list[re.Pattern] = [
    re.compile(r"SAC[：:]?\s*S\d+"),                   # SAC 编号
    re.compile(r"[\w.+-]+@[\w.-]+\.\w+"),               # 邮箱
    re.compile(r"hyzqdatemark"),                         # 日期标记
    re.compile(r"^!\[.*?\]\(images/"),                   # 图片引用
    re.compile(r"^资料来源[：:]"),                       # 资料来源行
    re.compile(r"^\s*联系人\s*$"),                       # 孤立的"联系人"
]

# ── LaTeX 清理 ──
_LATEX_RE = re.compile(r"\$(.*?)\$")


def _clean_latex(m: re.Match) -> str:
    """把 $...$ 内的 LaTeX 转为可读文本。"""
    inner = m.group(1)
    # \% → %
    inner = inner.replace("\\%", "%")
    # 去掉其他 LaTeX 命令
    inner = re.sub(r"\\(mathsf|mathrm|text|left|right|,|;|!)\s*", "", inner)
    inner = re.sub(r"[{}]", "", inner)
    # 压缩空格
    inner = re.sub(r"\s+", "", inner)
    return inner


def clean_report(text: str) -> str:
    """对单篇研报 md 执行全部清洗步骤，返回清洗后文本。"""

    # 1. 截断尾部免责声明
    m = _TAIL_CUT_RE.search(text)
    if m:
        text = text[: m.start()].rstrip()

    # 2. 逐行过滤头部/全局噪音
    lines = text.split("\n")
    cleaned: list[str] = []
    for line in lines:
        if any(p.search(line) for p in _LINE_NOISE_PATTERNS):
            continue
        cleaned.append(line)
    text = "\n".join(cleaned)

    # 3. LaTeX 清理
    text = _LATEX_RE.sub(_clean_latex, text)

    # 4. 压缩连续空行（>2 → 1）
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def _relative_source_pdf(meta_path: Path) -> str:
    """从 meta.json 读 source_pdf，转为 ./research report/... 相对路径。"""
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return ""
    abs_pdf = meta.get("source_pdf", "")
    if not abs_pdf:
        return ""
    # 找到 "research report" 之后的部分
    abs_pdf = abs_pdf.replace("\\", "/")
    marker = "research report/"
    idx = abs_pdf.find(marker)
    if idx == -1:
        return abs_pdf
    return "./" + abs_pdf[idx:]


def _build_frontmatter(report_type: str, title: str, source_pdf: str) -> str:
    """生成 YAML frontmatter。"""
    return (
        "---\n"
        f"report_type: {report_type}\n"
        f"title: \"{title}\"\n"
        f"source_pdf: \"{source_pdf}\"\n"
        "---\n\n"
    )


def process_single(report_md: Path) -> str:
    """处理单个 report.md，清洗后写入 data/RAG_md/，返回输出路径。"""
    text = report_md.read_text(encoding="utf-8")
    cleaned = clean_report(text)

    # 推断元数据
    report_dir = report_md.parent  # .../个股研报/标题/
    title = report_dir.name
    report_type = report_dir.parent.name  # 个股研报 or 行业研报
    source_pdf = _relative_source_pdf(report_dir / "meta.json")

    frontmatter = _build_frontmatter(report_type, title, source_pdf)
    result = frontmatter + cleaned

    # 写入 data/RAG_md/{report_type}/{title}.md
    out_dir = _OUT_ROOT / report_type
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{title}.md"
    out_path.write_text(result, encoding="utf-8")
    return str(out_path)


def process_all() -> None:
    """批量处理所有研报。"""
    count = 0
    for report_type_dir in sorted(_MD_ROOT.iterdir()):
        if not report_type_dir.is_dir():
            continue
        for title_dir in sorted(report_type_dir.iterdir()):
            report_md = title_dir / "report.md"
            if not report_md.exists():
                continue
            process_single(report_md)
            count += 1
    print(f"完成，共处理 {count} 份研报 → {_OUT_ROOT}")


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    if sys.argv[1] == "--all":
        process_all()
    else:
        path = Path(sys.argv[1])
        if not path.exists():
            print(f"文件不存在: {path}", file=sys.stderr)
            sys.exit(1)
        out = process_single(path)
        print(f"已写入: {out}")


if __name__ == "__main__":
    main()
