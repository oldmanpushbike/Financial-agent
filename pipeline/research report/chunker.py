# -*- coding: utf-8 -*-
"""研报 Markdown 语义切分：按 # 标题层级 + <table> 边界切分 chunk。

用法:
  单文件测试:
    python chunker.py "data/RAG_md/个股研报/xxx.md"

  批量切分:
    python chunker.py --all

输出: data/chunks/{个股研报|行业研报}/{报告名}.jsonl
"""
from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent  # FS/
_RAG_MD_ROOT = _PROJECT_ROOT / "data" / "RAG_md"
_CHUNKS_DIR = _PROJECT_ROOT / "data" / "chunks"

# ── frontmatter 解析 ────────────────────────────────────
_FM_RE = re.compile(r"^---\s*\n(.*?)\n---\s*\n", re.DOTALL)
_FM_KV_RE = re.compile(r'^(\w+):\s*"?(.+?)"?\s*$', re.MULTILINE)

# ── 图表标题模式 ────────────────────────────────────────
_CHART_TITLE_RE = re.compile(r"^(图表\d+[：:].+)$", re.MULTILINE)

# ── heading 模式 ────────────────────────────────────────
_HEADING_RE = re.compile(r"^(#{1,4})\s+(.+)$", re.MULTILINE)


# ── PLACEHOLDER_FOR_APPEND ──


@dataclass
class Chunk:
    text: str
    metadata: dict = field(default_factory=dict)


def _parse_frontmatter(content: str) -> tuple[dict, str]:
    """解析 YAML frontmatter，返回 (meta_dict, body)。"""
    m = _FM_RE.match(content)
    if not m:
        return {}, content
    fm_block = m.group(1)
    meta = {k: v for k, v in _FM_KV_RE.findall(fm_block)}
    body = content[m.end():]
    return meta, body


def _find_table_title(lines: list[str], table_line_idx: int) -> str:
    """向上查找 table 对应的标题。

    优先匹配 "图表N：xxx" 格式；
    否则取紧邻 table 上方的非空文本行（如 "资产负债表（百万元）"）。
    """
    for i in range(table_line_idx - 1, max(table_line_idx - 4, -1), -1):
        if i < 0:
            break
        line = lines[i].strip()
        if not line:
            continue
        if _CHART_TITLE_RE.match(line):
            return line
        # 非空非标题行，当作 table 的描述行
        if not line.startswith("#") and "<table>" not in line:
            return line
        break
    return ""


def chunk_file(filepath: Path) -> list[Chunk]:
    """对单个 RAG_md 文件执行语义切分。

    切分策略：
    1. 按 # heading 拆分为 section
    2. section 内遇到 <table> 行，将 table（含上方标题行）独立为 table chunk
    3. table 前后的文本各自成为 text chunk
    4. 过短的 text chunk（<30 字）合并到相邻 chunk
    """
    content = filepath.read_text(encoding="utf-8")
    file_meta, body = _parse_frontmatter(content)
    source_pdf = file_meta.get("source_pdf", "")
    report_type = file_meta.get("report_type", "")
    title = file_meta.get("title", "")

    lines = body.split("\n")
    chunks: list[Chunk] = []
    current_heading = ""
    buf: list[str] = []  # 当前文本缓冲

    def _flush_text_buf():
        """将文本缓冲刷出为一个 text chunk。"""
        text = "\n".join(buf).strip()
        buf.clear()
        if len(text) < 30:
            # 过短，尝试合并到上一个 chunk
            if chunks and chunks[-1].metadata.get("chunk_type") == "text":
                chunks[-1].text += "\n\n" + text
                return
            if not text:
                return
        chunks.append(Chunk(
            text=text,
            metadata={
                "source_pdf": source_pdf,
                "report_type": report_type,
                "title": title,
                "chunk_type": "text",
                "heading": current_heading,
            },
        ))

    i = 0
    while i < len(lines):
        line = lines[i]

        # ── 检测 heading ──
        hm = _HEADING_RE.match(line)
        if hm:
            # 先刷出之前的缓冲
            _flush_text_buf()
            current_heading = hm.group(2).strip()
            # heading 只存 metadata，不放进 chunk text
            i += 1
            continue

        # ── 检测 <table> ──
        if "<table>" in line:
            # 向上回收 table 标题行（从 buf 末尾取）
            table_title = ""
            title_line = ""
            if buf:
                candidate = buf[-1].strip()
                if candidate and not candidate.startswith("#"):
                    table_title = candidate
                    title_line = buf.pop()

            # 先刷出 table 前的文本
            _flush_text_buf()

            # 如果 buf 里没找到标题，尝试从上一个 text chunk 末尾回收
            if not table_title and chunks:
                prev = chunks[-1]
                if prev.metadata.get("chunk_type") == "text":
                    prev_lines = prev.text.rstrip().split("\n")
                    candidate = prev_lines[-1].strip()
                    if candidate and not candidate.startswith("#"):
                        table_title = candidate
                        title_line = prev_lines.pop()
                        prev.text = "\n".join(prev_lines).strip()
                        # 如果回收后上一个 chunk 变空/过短，删掉它
                        if len(prev.text) < 30:
                            if len(chunks) >= 2 and chunks[-2].metadata.get("chunk_type") == "text":
                                chunks[-2].text += "\n\n" + prev.text
                            chunks.pop()

            # 收集完整 table（可能跨多行，直到 </table>）
            table_lines = []
            while i < len(lines):
                table_lines.append(lines[i])
                if "</table>" in lines[i]:
                    i += 1
                    break
                i += 1

            table_text = "\n".join(table_lines).strip()
            chunks.append(Chunk(
                text=table_text,
                metadata={
                    "source_pdf": source_pdf,
                    "report_type": report_type,
                    "title": title,
                    "chunk_type": "table",
                    "heading": current_heading,
                    "table_title": table_title,
                },
            ))
            continue

        # ── 普通文本行 ──
        buf.append(line)
        i += 1

    # 刷出最后的缓冲
    _flush_text_buf()

    return chunks


def _write_chunks(chunks: list[Chunk], md_path: Path) -> Path:
    """将 chunks 写入 data/chunks/{report_type}/{stem}.jsonl，返回输出路径。"""
    # 从第一个 chunk 的 metadata 取 report_type，兜底用文件父目录名
    report_type = ""
    if chunks and chunks[0].metadata.get("report_type"):
        report_type = chunks[0].metadata["report_type"]
    else:
        report_type = md_path.parent.name

    out_dir = _CHUNKS_DIR / report_type
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{md_path.stem}.jsonl"
    with open(out_path, "w", encoding="utf-8") as f:
        for c in chunks:
            f.write(json.dumps(asdict(c), ensure_ascii=False) + "\n")
    return out_path


def chunk_all() -> int:
    """批量切分所有 RAG_md 文件，按文件分别存储。"""
    total_chunks = 0
    total_files = 0
    for md_file in sorted(_RAG_MD_ROOT.rglob("*.md")):
        chunks = chunk_file(md_file)
        _write_chunks(chunks, md_file)
        total_chunks += len(chunks)
        total_files += 1
    print(f"完成，共 {total_files} 份研报，{total_chunks} 个 chunk → {_CHUNKS_DIR}")
    return total_chunks


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    if sys.argv[1] == "--all":
        chunk_all()
    else:
        path = Path(sys.argv[1])
        if not path.exists():
            print(f"文件不存在: {path}", file=sys.stderr)
            sys.exit(1)
        chunks = chunk_file(path)
        out = _write_chunks(chunks, path)
        print(f"共 {len(chunks)} 个 chunk → {out}")


if __name__ == "__main__":
    main()

