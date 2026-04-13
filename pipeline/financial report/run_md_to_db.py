# -*- coding: utf-8 -*-
"""
Markdown解析入库工具 (适配版本)
直接处理已存在的Markdown文件，不依赖PDF转换
"""
import sys
import sqlite3
from pathlib import Path
from collections import defaultdict

# 导入md_to_db中的所有函数
from md_to_db import (
    parse_schema, init_db, rebuild_stock_qoq_metrics,
    dump_quality, load_company_map, detect_meta,
    detect_summary_from_md, parse_md_to_tables, write_tables_to_db,
    log, ensure_tables, M
)


def is_summary_file(name):
    """判断是否为摘要文件名"""
    return any(
        k in name
        for k in [
            "报告摘要", "年度报告摘要", "半年度报告摘要", "季度报告摘要",
            "一季度报告摘要", "三季度报告摘要", "摘要版",
        ]
    )


def build_company_map_from_dirs(md_dirs):
    """从Markdown目录构建公司映射"""
    company_map = {}
    import re
    
    for md_dir in md_dirs:
        name = md_dir.name if isinstance(md_dir, Path) else str(md_dir)
        # 尝试提取股票代码（6位数字，前后不是数字）
        m = re.search(r"(?<!\d)(\d{6})(?!\d)", name)
        if m:
            stock_code = m.group(1)
            # 从公司名称中提取简称
            left = name.split("：")[0].strip() if "：" in name else name.split("_")[0].strip()
            # 如果没有已知的简称，使用目录名的一部分
            if left and left not in company_map.values():
                company_map[stock_code] = left.split("（")[0].split("(")[0].strip()
    
    return company_map


def find_md_dirs(base_dir):
    """递归查找所有包含report.md的目录"""
    md_dirs = []
    for item in base_dir.iterdir():
        if item.is_dir():
            if (item / "report.md").exists():
                md_dirs.append(item)
            else:
                # 继续递归查找
                md_dirs.extend(find_md_dirs(item))
    return md_dirs


def process_md_dir(
    md_dir,
    conn,
    schema,
    company_map,
    quality_stats,
    parsed_history=None,
):
    """处理Markdown目录到数据库（适配版本）"""
    # 确保 md_dir 是 Path 对象
    md_dir = Path(md_dir) if not isinstance(md_dir, Path) else md_dir
    md_name = md_dir.name
    
    # 检查是否为摘要文件
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

    # 使用 detect_meta 函数 - 它需要 Path 对象
    # 强制使用company_map中的官方简称
    meta = detect_meta(md_dir, company_map, md)
    # 确保使用官方简称
    if meta.stock_code and meta.stock_code in company_map:
        meta = M(
            stock_code=meta.stock_code,
            stock_abbr=company_map[meta.stock_code],
            report_year=meta.report_year,
            report_type=meta.report_type,
            report_period=meta.report_period,
            source_file_name=meta.source_file_name or md_name
        )
    print(
        f"    [阶段2/4] 识别元数据：stock_code={meta.stock_code or '-'}, report_period={meta.report_period or '-'}"
    )

    if not meta.report_period or not meta.stock_code:
        log(conn, md_name, "failed", "元数据识别失败")
        quality_stats["failed"] += 1
        return meta, None

    print("    [阶段3/4] 解析Markdown：开始")
    stock_history = parsed_history.get(meta.stock_code, {}) if parsed_history else {}
    parsed = parse_md_to_tables(
        md,
        schema,
        quality_stats,
        md_name,
        meta.report_period,
        parsed_history=stock_history,
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
    # 路径配置
    MD_DIR = Path(r"d:\14 Tedicup\FS\data\md\financial report")
    DB_OUTPUT = Path(r"d:\14 Tedicup\FS\data\financial.db")
    SCHEMA_XLSX = Path(r"d:\14 Tedicup\FS\data\pdf\附件3：数据库-表名及字段说明.xlsx")
    # 优先使用附件1的公司信息（更完整）
    COMPANY_XLSX = Path(r"d:\14 Tedicup\FS\data\pdf\附件1：中药上市公司基本信息（截至到2025年12月22日）.xlsx")
    
    print("=" * 60)
    print("财务报表 Markdown 解析入库工具 (适配版)")
    print("=" * 60)
    print(f"Markdown源目录: {MD_DIR}")
    print(f"数据库输出路径: {DB_OUTPUT}")
    print(f"Schema来源: {SCHEMA_XLSX}")
    
    # 检查目录
    if not MD_DIR.exists():
        print(f"[错误] Markdown目录不存在: {MD_DIR}")
        return
    
    # 加载Schema
    schema = parse_schema(SCHEMA_XLSX)
    print(f"\nSchema包含 {sum(len(fields) for fields in schema.values())} 个字段")
    
    # 加载公司映射（优先从Excel加载）
    # 先收集所有MD目录（用于后续处理）
    all_md_dirs = find_md_dirs(MD_DIR)
    
    if COMPANY_XLSX.exists():
        company_map = load_company_map(COMPANY_XLSX)
        print(f"从Excel加载 {len(company_map)} 个公司映射")
    else:
        # 使用所有找到的MD目录来构建映射
        company_map = build_company_map_from_dirs(all_md_dirs)
        print(f"从目录名构建 {len(company_map)} 个公司映射")
    
    # 初始化数据库
    DB_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    
    # 如果数据库已存在，先删除
    if DB_OUTPUT.exists():
        try:
            DB_OUTPUT.unlink()
            print("已删除旧数据库")
        except:
            pass
    
    conn = sqlite3.connect(str(DB_OUTPUT))
    ensure_tables(conn, schema)
    print(f"数据库已初始化: {DB_OUTPUT}")
    
    # 获取所有Markdown目录（遍历所有子目录）
    md_dirs = find_md_dirs(MD_DIR)
    print(f"\n找到 {len(md_dirs)} 个Markdown目录")
    
    quality_stats = {"ok": 0, "failed": 0, "skipped": 0, "tables": {}, "issues": []}
    parsed_history = {}
    
    # 处理每个目录
    total = len(md_dirs)
    for index, md_dir in enumerate(sorted(md_dirs), 1):
        print(f"\n[{index}/{total}] 处理: {md_dir.name}")
        try:
            meta, parsed = process_md_dir(
                md_dir,
                conn,
                schema,
                company_map,
                quality_stats,
                parsed_history=parsed_history,
            )
            if meta and parsed and meta.stock_code:
                parsed_history.setdefault(meta.stock_code, {})[meta.report_period] = {
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
    
    # 重建环比指标
    print("\n[后处理] 重建环比指标...")
    for stock_code, stock_period_history in parsed_history.items():
        rebuild_stock_qoq_metrics(conn, schema, stock_code, stock_period_history)
    conn.commit()
    print("[后处理] 环比指标重建完成")
    
    conn.close()
    
    # 输出质量报告
    dump_quality(DB_OUTPUT.parent, quality_stats)
    
    print("\n" + "=" * 60)
    print("完成!")
    print(f"  成功: {quality_stats['ok']}")
    print(f"  失败: {quality_stats['failed']}")
    print(f"  跳过: {quality_stats.get('skipped', 0)}")
    print(f"数据库: {DB_OUTPUT}")
    print("=" * 60)


if __name__ == "__main__":
    main()
