# FS Agent — 上市公司智能问数系统

基于 LangGraph + 智谱 GLM-4 的财务数据分析 Agent，支持自然语言查询财务报表、检索研究报告、生成可视化图表。

## 功能特性

- **自然语言查数**：Text-to-SQL 自动生成并执行 SQL，支持两阶段策略（直接生成 + 分解回退）
- **研报检索**：四阶段混合检索（向量检索 + BM25 + 千帆 Rerank + Parent-Child 块映射）
- **图表生成**：支持折线图、柱状图、饼图、雷达图、散点图、直方图、箱线图、双条形图、表格共 9 种图表类型
- **智能编排**：LangGraph 统一任务列表架构，自动规划 → 执行 → 综合 → 格式化

## 项目结构

```
FS/
├── agent/                  # 核心 Agent 模块
│   ├── main.py             # CLI 入口
│   ├── config.py           # 配置加载
│   ├── state.py            # 状态定义
│   ├── planner.py          # 意图解析
│   ├── graph.py            # LangGraph 编排
│   ├── conditions.py       # 路由条件
│   └── llm.py              # LLM 工厂
├── tool/                   # 工具模块
│   ├── sql_query.py        # SQL 生成与执行
│   ├── rag_search.py       # 混合检索
│   ├── visualizer.py       # 图表生成
│   └── json_formatter.py   # 输出格式化
├── pipeline/               # 离线数据处理管道
│   ├── financial report/   # 财务报表解析
│   └── research report/    # 研报预处理、分块、向量库构建
├── data/                   # 数据目录
│   ├── .env                # API 密钥（需自行创建）
│   ├── financial.db        # SQLite 财务数据库（需自行构建）
│   ├── chroma_db/          # ChromaDB 向量库（需自行构建）
│   └── RAG_md/             # 研报 Markdown
├── scripts/                # 批量测试脚本
├── result/                 # 输出结果
├── docs/                   # 文档
├── pyproject.toml          # 项目配置
└── .env.example            # 环境变量模板
```

## 快速开始

### 环境要求

- Python >= 3.11
- [uv](https://docs.astral.sh/uv/) 包管理器

### 安装

```bash
# 克隆仓库
git clone <repo-url>
cd FS

# 安装依赖
uv sync

# 配置环境变量
cp .env.example data/.env
# 编辑 data/.env，填入你的 API 密钥
```

### 运行

```bash
uv run python -m agent.main
```

交互式 CLI 对话：

```
FS Agent 已启动 (thread=xxx)
输入问题开始对话，exit/quit 退出。
知识库加载中... 就绪

You> 凯莱英2024年营收和净利润是多少？
Agent> {"Q": "...", "A": {"content": "...", "image": null, "references": []}}
```

## 架构

```
用户输入
  │
  ▼
Planner ──► Executor ──► Synthesis ──► Formatter ──► END
（意图解析）  （工具执行）  （结果汇总）  （JSON格式化）
                │
    ┌───────────┼───────────┐
    ▼           ▼           ▼
   SQL        RAG        Visualize
```

- **Planner**：使用 GLM-4-Flash 解析用户意图，输出统一任务列表 `[{id, tool, query, depends_on}]`
- **Executor**：按依赖顺序执行子任务，支持 sql / rag / visualize / clarify / chat 五种工具
- **Synthesis**：汇总所有证据，生成最终自然语言回答
- **Formatter**：标准化 JSON 输出 `{"Q": "...", "A": {"content", "image", "references"}}`

### SQL 执行策略

两阶段执行保证鲁棒性：
1. **Phase 1**：直接生成 SQL 并执行，含 3 次重试（自动修正语法/字段错误）
2. **Phase 2**：失败时自动分解为多条单表查询，逐条执行后 LLM 综合结果

### RAG 检索流程

四阶段混合检索：
1. 向量检索（ChromaDB + BAAI/bge-large-zh-v1.5）
2. BM25 关键词检索
3. 百度千帆 bce-reranker-base 重排序
4. Parent-Child 块映射，返回完整上下文

## 数据准备

需要自行准备以下数据（详见 `pipeline/` 目录下的脚本）：

1. **财务报表**：将 PDF 财报通过 MinerU API 解析为 Markdown，运行 `pipeline/financial report/merged_financial_parser.py` 导入 SQLite
2. **研究报告**：将 PDF 研报解析为 Markdown，依次运行 `preprocess.py` → `chunker.py` → `build_vectordb.py` 构建向量库

## 依赖

核心依赖参见 [pyproject.toml](pyproject.toml)，主要依赖：

- **LangChain + LangGraph**：Agent 框架与编排
- **智谱 AI glm-4-flash**：LLM（通过 OpenAI 兼容接口）
- **百度千帆 Embedding + Rerank**：向量检索与重排序
- **ChromaDB**：向量数据库
- **Matplotlib**：图表生成
- **sentence-transformers + modelscope**：本地 Embedding 模型

## 适用数据范围（可替换）

- **财务报表**：资产负债表、利润表、现金流量表、核心绩效指标表
- **研究报告**：个股研报 + 行业研报

