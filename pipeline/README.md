# MinerU PDF转Markdown处理工具

## 概述

本模块负责将财务报表（financial report）和研究报告（research report）的PDF文件转换为Markdown格式，并生成相应的元信息文件。

## 目录结构

```
FS/
├── data/
│   ├── pdf/
│   │   ├── financial report/      # 财务报表PDF源文件
│   │   └── research report/       # 研究报告PDF源文件
│   └── md/
│       ├── financial report/      # 财务报表Markdown缓存
│       │   ├── {pdf_stem}/
│       │   │   ├── report.md      # Markdown内容
│       │   │   └── meta.json      # 元信息
│       │   └── ...
│       └── research report/       # 研究报告Markdown缓存
│           └── ...
└── pipeline/
    └── mineru_processor.py        # 主处理脚本
```

## 输出文件格式

### Markdown文件 (`report.md`)
- 提取的PDF文本内容
- 保留原有的层级结构
- 不保存图片等中间文件

### 元信息文件 (`meta.json`)
```json
{
  "source_pdf": "D:\\...\\financial report\\xxx.pdf",
  "generated_at": "2026-04-10T17:30:00.123456"
}
```

## 使用方法

### 1. 环境准备

设置MinerU API Token环境变量：
```bash
# Windows
set MINERU_TOKEN=你的Token

# Linux/Mac
export MINERU_TOKEN=你的Token
```

### 2. 基本用法

处理所有PDF文件：
```bash
python mineru_processor.py
```

### 3. 命令行参数

| 参数 | 说明 |
|------|------|
| `-t, --token` | MinerU API Token（可选，默认从环境变量读取） |
| `-b, --batch-size` | 每批最大文件数（默认: 200） |
| `-m, --model-version` | 模型版本（默认: vlm） |
| `--test` | 仅测试API连接 |
| `--list` | 仅列出待处理文件 |

### 4. 使用示例

```bash
# 测试API连接
python mineru_processor.py --test

# 查看待处理文件
python mineru_processor.py --list

# 指定Token和处理批次大小
python mineru_processor.py --token YOUR_TOKEN --batch-size 50

# 使用不同的模型版本
python mineru_processor.py --model-version ocr
```

## API配置

- **Base URL**: https://mineru.net
- **默认超时**: 3600秒（1小时）
- **轮询间隔**: 10秒
- **批量大小**: 200个文件/批

## 缓存机制

- 脚本会跳过已存在`report.md`的PDF文件
- 重新处理时删除旧文件即可

## 错误处理

- 超时任务可后续通过batch_id重试
- 失败文件会在日志中显示原因
