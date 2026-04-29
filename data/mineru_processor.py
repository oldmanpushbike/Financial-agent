# -*- coding: utf-8 -*-
"""
MinerU PDF转Markdown处理脚本
将 financial report 和 research report PDF 文件转换为 Markdown

输入目录结构:
    FS/data/pdf/financial report/      - 财务报表PDF
    FS/data/pdf/research report/       - 研究报告PDF

输出目录结构:
    FS/data/md/financial report/      - 财务报表Markdown缓存
    FS/data/md/research report/        - 研究报告Markdown缓存

每个PDF生成:
    {pdf_stem}/
        report.md     - Markdown内容
        meta.json     - 元信息

环境变量:
    MINERU_TOKEN  - MinerU API Token (必需)
"""

import os
import json
import time
import re
import hashlib
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from datetime import datetime

# 加载 .env 文件（支持 MINERU_TOKEN 等配置）
def load_env_file():
    """从脚本同级目录的 .env 文件加载环境变量"""
    script_dir = Path(__file__).parent.resolve()
    env_file = script_dir / ".env"
    if env_file.exists():
        with open(env_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    if key and value and key not in os.environ:
                        os.environ[key] = value

load_env_file()

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# =============================================================================
# API 配置
# =============================================================================
MINERU_BASE_URL = "https://mineru.net"
DEFAULT_TIMEOUT = 0  # 超时时间（秒），0表示不限制
POLL_INTERVAL = 10      # 轮询间隔（秒）
BATCH_SIZE = 10         # 减小批次避免429限流


# =============================================================================
# 路径配置
# =============================================================================
# 获取当前脚本所在目录（pipeline目录）
SCRIPT_DIR = Path(__file__).parent.resolve()
FS_ROOT = SCRIPT_DIR.parent

PDF_DIR = FS_ROOT / "data" / "pdf"
MD_DIR = FS_ROOT / "data" / "md"

# PDF子目录
FINANCIAL_PDF_DIR = PDF_DIR / "financial report"
RESEARCH_PDF_DIR = PDF_DIR / "research report"

# 财务报表子目录
FINANCIAL_PDF_SUBDIRS = {
    "reports_上交所": FINANCIAL_PDF_DIR / "reports-上交所",
    "reports_深交所": FINANCIAL_PDF_DIR / "reports-深交所",
}

# 研究报告子目录
RESEARCH_PDF_SUBDIRS = {
    "个股研报": RESEARCH_PDF_DIR / "个股研报",
    "行业研报": RESEARCH_PDF_DIR / "行业研报",
}

# MD子目录
FINANCIAL_MD_DIR = MD_DIR / "financial report"
RESEARCH_MD_DIR = MD_DIR / "research report"

# MD输出子目录（与PDF子目录对应）
FINANCIAL_MD_SUBDIRS = {
    "reports_上交所": FINANCIAL_MD_DIR / "reports-上交所",
    "reports_深交所": FINANCIAL_MD_DIR / "reports-深交所",
}

RESEARCH_MD_SUBDIRS = {
    "个股研报": RESEARCH_MD_DIR / "个股研报",
    "行业研报": RESEARCH_MD_DIR / "行业研报",
}


# =============================================================================
# MinerU API 客户端类
# =============================================================================
class MinerUV4Client:
    """MinerU v4 API 客户端封装"""

    def __init__(self, token: str, model_version: str = "vlm", base_url: str = MINERU_BASE_URL, timeout: int = 30):
        self.token = token
        self.model_version = model_version
        self.base_url = base_url.rstrip('/')
        self.timeout = timeout

        # 配置HTTP会话及重试策略
        self.session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.token}"
        }

    def upload_and_submit(self, file_paths: List[Path]) -> str:
        """上传文件并提交批量解析任务
        
        Args:
            file_paths: PDF文件路径列表
            
        Returns:
            batch_id: 批次ID
        """
        url = f"{self.base_url}/api/v4/file-urls/batch"

        files_data = []
        for fp in file_paths:
            data_id = _stem_to_data_id(fp.stem)
            files_data.append({"name": fp.name, "data_id": data_id})

        data = {
            "files": files_data,
            "model_version": self.model_version
        }

        response = self.session.post(url, headers=self.headers, json=data, timeout=self.timeout)
        response.raise_for_status()

        result = response.json()
        if result.get("code") != 0:
            raise Exception(f"申请上传URL失败: {result.get('msg', '未知错误')}")

        batch_id = result["data"]["batch_id"]
        file_urls = result["data"]["file_urls"]

        # 上传所有文件到预签名URL
        for i, fp in enumerate(file_paths):
            with open(fp, 'rb') as f:
                upload_response = self.session.put(file_urls[i], data=f, timeout=self.timeout)
                if upload_response.status_code != 200:
                    raise Exception(f"文件上传失败: {fp.name}, 状态码: {upload_response.status_code}")

        return batch_id

    def _poll_results(self, batch_id: str, file_count: int, timeout: int = DEFAULT_TIMEOUT) -> Dict[str, Any]:
        """轮询获取解析结果
        
        Args:
            batch_id: 批次ID
            file_count: 文件数量
            timeout: 超时时间（秒）
            
        Returns:
            解析结果字典
        """
        result_url = f"{self.base_url}/api/v4/extract-results/batch/{batch_id}"

        start_time = time.time()
        poll_count = 0

        print(f"    [API] 开始轮询，超时时间: {timeout}秒", end="", flush=True)

        while True:
            elapsed = time.time() - start_time
            if timeout > 0 and elapsed > timeout:
                print(f"\n    [API] 轮询超时({elapsed:.0f}秒)，任务可能在后台继续运行")
                print(f"    [API] Batch ID: {batch_id} (可用于后续查询)")
                return {
                    "status": "timeout",
                    "batch_id": batch_id,
                    "file_count": file_count,
                    "elapsed": elapsed,
                    "extract_result": []
                }

            response = self.session.get(result_url, headers=self.headers, timeout=self.timeout)
            response.raise_for_status()

            result = response.json()
            if result.get("code") != 0:
                raise Exception(f"获取结果失败: {result.get('msg', '未知错误')}")

            data = result.get("data", {})
            extract_results = data.get("extract_result", [])

            # 统计各状态数量
            done_count = sum(1 for r in extract_results if r.get("state") == "done")
            running_count = sum(1 for r in extract_results if r.get("state") == "running")
            failed_count = sum(1 for r in extract_results if r.get("state") == "failed")

            # 全部完成
            if extract_results and done_count + failed_count == len(extract_results):
                print(f"\n    [API] 批量解析完成！")
                data["status"] = "done"
                return data

            # 部分失败
            if extract_results and failed_count > 0 and running_count == 0:
                print(f"\n    [API] 部分任务失败")
                data["status"] = "partial"
                return data

            # 定期打印进度
            if poll_count > 0 and poll_count % 20 == 0:
                print(f" ({elapsed:.0f}秒)", end="", flush=True)

            poll_count += 1
            time.sleep(POLL_INTERVAL)


# =============================================================================
# PDF解析函数
# =============================================================================
def _stem_to_data_id(stem: str) -> str:
    return hashlib.md5(stem.encode('utf-8')).hexdigest()[:16]


def parse_files_batch(pdf_paths: List[Path], token: str, model_version: str = "vlm", timeout: int = 0) -> Dict[str, Dict[str, Any]]:
    """批量解析PDF文件，返回 {pdf_stem: result} 字典。"""
    import zipfile
    import io

    if not pdf_paths:
        return {}

    hash_to_stem = {_stem_to_data_id(fp.stem): fp.stem for fp in pdf_paths}

    client = MinerUV4Client(token=token, model_version=model_version)

    print(f"\n{'='*50}")
    print(f"开始批量解析 {len(pdf_paths)} 个文件...")
    print(f"{'='*50}")

    results = {}
    try:
        batch_id = client.upload_and_submit(pdf_paths)
        print(f"    [API] Batch ID: {batch_id}, 文件数: {len(pdf_paths)}")

        batch_result = client._poll_results(batch_id, len(pdf_paths), timeout=timeout)

        if batch_result.get("status") == "timeout":
            batch_id = batch_result.get("batch_id")
            print(f"\n    [API] 轮询超时，任务可能仍在处理中")
            for p in pdf_paths:
                results[p.stem] = {
                    "status": "timeout",
                    "content": None,
                    "error": f"轮询超时({batch_result.get('elapsed', 0):.0f}秒)",
                    "batch_id": batch_id
                }
            return results

        extract_results = batch_result.get("extract_result", [])

        for item in extract_results:
            data_id = item.get("data_id", "")
            stem = hash_to_stem.get(data_id, data_id)
            file_state = item.get("state", "")
            err_msg = item.get("err_msg", "")
            zip_url = item.get("full_zip_url")

            if file_state == "done":
                content = None
                if zip_url:
                    try:
                        headers = {"Authorization": f"Bearer {token}"}
                        zip_response = requests.get(zip_url, headers=headers, timeout=120)
                        zip_response.raise_for_status()

                        with zipfile.ZipFile(io.BytesIO(zip_response.content)) as zf:
                            md_files = [f for f in zf.namelist() if f.endswith('.md')]
                            if md_files:
                                for md_file in md_files:
                                    if 'image' not in md_file.lower() and 'img' not in md_file.lower():
                                        content = zf.read(md_file).decode('utf-8')
                                        break
                                if not content and md_files:
                                    content = zf.read(md_files[0]).decode('utf-8')
                    except Exception as e:
                        err_msg = f"下载失败: {e}"

                results[stem] = {
                    "status": "success" if content else "failed",
                    "content": content,
                    "error": None if content else err_msg,
                    "batch_id": batch_result.get("batch_id")
                }
            else:
                results[stem] = {
                    "status": "failed",
                    "content": None,
                    "error": err_msg or "解析失败",
                    "batch_id": batch_result.get("batch_id")
                }

    except Exception as e:
        print(f"\n    [API] 批量解析异常: {e}")
        for p in pdf_paths:
            results[p.stem] = {
                "status": "failed",
                "content": None,
                "error": str(e),
                "batch_id": None
            }

    return results


# =============================================================================
# 缓存处理函数
# =============================================================================
def batch_ensure_md_for_pdf(
    pdfs: List[Path],
    md_output_dir: Path,
    token: str = None,
    batch_size: int = BATCH_SIZE,
    model_version: str = "vlm",
    max_timeout_retries: int = 2
) -> Dict[str, Tuple[Path, Optional[Path]]]:
    """批量处理PDF文件，确保每个都有Markdown缓存
    
    Args:
        pdfs: PDF文件路径列表
        md_output_dir: Markdown输出根目录
        token: API Token（从环境变量MINERU_TOKEN读取）
        batch_size: 每批最大文件数
        model_version: 模型版本
        max_timeout_retries: 超时后继续等待的次数
        
    Returns:
        {pdf_stem: (output_subdir, md_path)} 字典
    """
    if token is None:
        token = os.environ.get("MINERU_TOKEN")
    if token is None:
        raise ValueError("请设置 MINERU_TOKEN 环境变量")

    # 创建输出目录
    md_output_dir.mkdir(parents=True, exist_ok=True)

    # 分类：已有缓存 vs 需要解析
    to_process = []
    cached_results = {}

    for pdf in pdfs:
        # 每个PDF输出到独立的子目录
        output_subdir = md_output_dir / pdf.stem
        output_subdir.mkdir(parents=True, exist_ok=True)
        cached_md = output_subdir / "report.md"

        if cached_md.exists():
            cached_results[pdf.stem] = (output_subdir, cached_md)
        else:
            to_process.append(pdf)

    print(f"\n总计 {len(pdfs)} 个 PDF")
    print(f"  - 已有缓存: {len(cached_results)} 个")
    print(f"  - 需要解析: {len(to_process)} 个")
    print(f"  - 输出目录: {md_output_dir}")
    print(f"  - 每批大小: {batch_size}")

    # 分批处理
    for i in range(0, len(to_process), batch_size):
        batch = to_process[i:i + batch_size]
        batch_num = i // batch_size + 1
        total_batches = (len(to_process) + batch_size - 1) // batch_size

        print(f"\n{'='*50}")
        print(f"处理批次 {batch_num}/{total_batches}，文件数: {len(batch)}")
        print(f"{'='*50}")

        # 批量解析
        results = parse_files_batch(batch, token, model_version)

        # 保存结果
        for pdf in batch:
            stem = pdf.stem
            output_subdir = md_output_dir / stem
            cached_md = output_subdir / "report.md"

            if stem in results and results[stem]["status"] == "success":
                content = results[stem]["content"]
                if content:
                    # 保存markdown内容
                    cached_md.write_text(content, encoding="utf-8")
                    
                    # 生成meta.json
                    meta = {
                        "source_pdf": str(pdf.resolve()),
                        "generated_at": datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f")
                    }
                    meta_file = output_subdir / "meta.json"
                    meta_file.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
                    
                    print(f"    [保存] {pdf.name} -> {stem}/")
                    cached_results[stem] = (output_subdir, cached_md)
                else:
                    print(f"    [失败] {pdf.name} - 内容为空")
                    cached_results[stem] = (output_subdir, None)
            elif stem in results and results[stem].get("status") == "timeout":
                print(f"    [超时] {pdf.name} - 任务可能仍在处理中")
                cached_results[stem] = (output_subdir, None)
            else:
                error = results.get(stem, {}).get("error", "未知错误")
                print(f"    [失败] {pdf.name} - {error}")
                cached_results[stem] = (output_subdir, None)

        if i + batch_size < len(to_process):
            print(f"    批次间等待30秒避免限流...")
            time.sleep(30)

    return cached_results


# =============================================================================
# 主处理函数
# =============================================================================
def process_all_reports(
    token: str = None,
    batch_size: int = BATCH_SIZE,
    model_version: str = "vlm"
) -> Dict[str, Dict[str, Any]]:
    """处理所有类型的报告PDF
    
    Args:
        token: API Token
        batch_size: 每批最大文件数
        model_version: 模型版本
        
    Returns:
        处理结果统计
    """
    if token is None:
        token = os.environ.get("MINERU_TOKEN")
    if token is None:
        raise ValueError("请设置 MINERU_TOKEN 环境变量")

    results = {}

    # 处理财务报表（上交所和深交所）
    results["financial"] = {}
    for subdir_name, pdf_subdir in FINANCIAL_PDF_SUBDIRS.items():
        md_subdir = FINANCIAL_MD_SUBDIRS[subdir_name]
        if pdf_subdir.exists():
            pdfs = list(pdf_subdir.glob("*.pdf"))
            if pdfs:
                print(f"\n{'#'*60}")
                print(f"# 处理财务报表 [{subdir_name}]: {len(pdfs)} 个文件")
                print(f"# 输入: {pdf_subdir}")
                print(f"# 输出: {md_subdir}")
                print(f"{'#'*60}")
                
                md_subdir.mkdir(parents=True, exist_ok=True)
                sub_result = batch_ensure_md_for_pdf(
                    pdfs,
                    md_subdir,
                    token,
                    batch_size,
                    model_version
                )
                results["financial"][subdir_name] = sub_result
            else:
                print(f"\n[跳过] 目录为空: {pdf_subdir}")
        else:
            print(f"\n[跳过] 目录不存在: {pdf_subdir}")

    # 处理研究报告（个股研报和行业研报）
    results["research"] = {}
    for subdir_name, pdf_subdir in RESEARCH_PDF_SUBDIRS.items():
        md_subdir = RESEARCH_MD_SUBDIRS[subdir_name]
        if pdf_subdir.exists():
            pdfs = list(pdf_subdir.glob("*.pdf"))
            if pdfs:
                print(f"\n{'#'*60}")
                print(f"# 处理研究报告 [{subdir_name}]: {len(pdfs)} 个文件")
                print(f"# 输入: {pdf_subdir}")
                print(f"# 输出: {md_subdir}")
                print(f"{'#'*60}")
                
                md_subdir.mkdir(parents=True, exist_ok=True)
                sub_result = batch_ensure_md_for_pdf(
                    pdfs,
                    md_subdir,
                    token,
                    batch_size,
                    model_version
                )
                results["research"][subdir_name] = sub_result
            else:
                print(f"\n[跳过] 目录为空: {pdf_subdir}")
        else:
            print(f"\n[跳过] 目录不存在: {pdf_subdir}")
        results["research"] = {}

    # 打印汇总
    print(f"\n{'='*60}")
    print(f"处理完成！汇总:")
    print(f"{'='*60}")
    
    for report_type, result_dict in results.items():
        total = len(result_dict)
        success = sum(1 for _, md_path in result_dict.values() if md_path is not None and md_path.exists())
        print(f"  {report_type}: {success}/{total} 成功")

    return results


# =============================================================================
# 测试函数
# =============================================================================
def test_api_connection(token: str = None) -> bool:
    """测试API连接"""
    try:
        if token is None:
            token = os.environ.get("MINERU_TOKEN")
        if token is None:
            print("[FAIL] 未设置 MINERU_TOKEN")
            return False

        client = MinerUV4Client(token=token)
        print(f"[OK] API 客户端创建成功")
        print(f"     Token: {token[:20]}...")
        print(f"     Base URL: {MINERU_BASE_URL}")
        return True
    except Exception as e:
        print(f"[FAIL] API 连接失败: {e}")
        return False


# =============================================================================
# 入口点
# =============================================================================
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="MinerU PDF转Markdown处理工具")
    parser.add_argument("--token", "-t", help="MinerU API Token（可选，默认从环境变量读取）")
    parser.add_argument("--batch-size", "-b", type=int, default=BATCH_SIZE, help=f"每批最大文件数（默认: {BATCH_SIZE}）")
    parser.add_argument("--model-version", "-m", default="vlm", help="模型版本（默认: vlm）")
    parser.add_argument("--test", action="store_true", help="仅测试API连接")
    parser.add_argument("--list", action="store_true", help="仅列出待处理文件")
    
    args = parser.parse_args()
    
    token = args.token or os.environ.get("MINERU_TOKEN")
    
    if args.test:
        print("测试API连接...")
        test_api_connection(token)
    elif args.list:
        print("待处理文件:")
        print(f"\n财务报表:")
        for subdir_name, pdf_subdir in FINANCIAL_PDF_SUBDIRS.items():
            if pdf_subdir.exists():
                pdfs = list(pdf_subdir.glob("*.pdf"))
                print(f"  [{subdir_name}] {pdf_subdir}: {len(pdfs)} 个文件")
                for pdf in pdfs:
                    print(f"    - {pdf.name}")
            else:
                print(f"  [{subdir_name}] {pdf_subdir}: 目录不存在")
        
        print(f"\n研究报告:")
        for subdir_name, pdf_subdir in RESEARCH_PDF_SUBDIRS.items():
            if pdf_subdir.exists():
                pdfs = list(pdf_subdir.glob("*.pdf"))
                print(f"  [{subdir_name}] {pdf_subdir}: {len(pdfs)} 个文件")
                for pdf in pdfs:
                    print(f"    - {pdf.name}")
            else:
                print(f"  [{subdir_name}] {pdf_subdir}: 目录不存在")
    else:
        print("MinerU PDF转Markdown处理工具")
        print(f"输入目录: {PDF_DIR}")
        print(f"输出目录: {MD_DIR}")
        print()
        
        process_all_reports(
            token=token,
            batch_size=args.batch_size,
            model_version=args.model_version
        )
