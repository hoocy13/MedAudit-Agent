# MedAudit Agent

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-blue.svg" alt="Python 3.10+">
  <img src="https://img.shields.io/badge/FastAPI-0.115+-green.svg" alt="FastAPI">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License: MIT">
</p>

> 医学病例合规自动审核系统 — 上传 PPT 病例文件，通过 LLM + 药品规则库进行多维度合规审查，返回结构化审核意见与专业 HTML 报告。

---

## 功能特性

- **REST API 服务** — FastAPI 驱动，支持文件上传、异步审核、报告下载
- **智能审核** — LLM 多维度合规审查：基础质控 + 合规专项审核
- **批量处理** — 支持单文件 / 多文件批量上传，后台异步执行
- **专业报告** — HTML 审核报告，分组展示，支持 PDF 导出
- **规则库驱动** — 按药品×适应症结构化管理，支持热更新

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置 API

```bash
cp .env.example .env
```

编辑 `.env`，填写 LLM API 配置：

```env
OPENAI_BASE_URL=https://your-llm-api.com/v1
OPENAI_API_KEY=your_api_key
OPENAI_MODEL_ID=your-model-id
```

### 3. 启动服务

```bash
python main.py
```

服务默认运行在 `http://localhost:8000`，访问 `http://localhost:8000/docs` 查看交互式 API 文档。

---

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| `POST` | `/api/audit/upload` | 上传单个 PPT，返回 `task_id` |
| `POST` | `/api/audit/batch` | 上传多个 PPT，批量审核 |
| `GET` | `/api/audit/tasks` | 列出所有任务 |
| `GET` | `/api/audit/{task_id}` | 查询任务状态与审核结果 |
| `GET` | `/api/audit/{task_id}/report` | 下载 HTML 审核报告 |

### 使用示例

```bash
# 上传单个 PPT
curl -X POST http://localhost:8000/api/audit/upload \
  -F "file=@病例分享.pptx"

# 返回: {"task_id": "a1b2c3d4", "status": "pending", ...}

# 查询结果
curl http://localhost:8000/api/audit/a1b2c3d4

# 下载报告
curl -O http://localhost:8000/api/audit/a1b2c3d4/report
```

---

## 审核维度

### 基础质控检查（每份报告必含）

| 检查项 | 说明 |
|--------|------|
| 文本规范性 | 错别字、语法错误、标点符号误用 |
| 通用名规范 | 药品是否使用通用名（INN），非商品名 |
| 医学逻辑 | 药物与适应症匹配性、剂量合理性 |
| 单位规范 | 剂量/频率/检验指标单位正确性 |

### 合规专项审核（按规则库匹配）

| 维度 | 说明 |
|------|------|
| 适应症合规 | 药品与适应症匹配、年龄/EOS/既往治疗要求 |
| 隐私保护 | 患者个人信息泄露检测 |
| Storyline | 就诊/注射/随访时间逻辑通顺性 |
| 安全性风险 | 急性加重期用药、注射间隔合规性 |
| 科学幻灯 | 引用文献、数据一致性 |

---

## 药品规则库

| 药品 | 适应症 | onlabel | 病例 | 幻灯 |
|------|--------|---------|------|------|
| 美泊利珠单抗 | SEA / COPD / CRSwNP / EGPA | 2-5 | 6-8 | 1-7 |
| 氟替美维 | COPD / 哮喘 | 3 | 10 | 7 |
| 西甲硅油 | default | 3 | 3 | 0 |

规则配置文件：`config/rules_config.json`（结构化 JSON） / `config/rules_checklist.csv`（Excel 可编辑）

---

## 项目结构

```
MedAudit-Agent/
├── app/
│   ├── main.py              # FastAPI 应用入口
│   ├── config.py             # 配置管理 (pydantic-settings)
│   ├── models.py             # Pydantic 数据模型
│   ├── api/
│   │   └── audit.py          # 审核 API 路由
│   ├── core/
│   │   ├── extractor.py      # PPT 文本/图片提取
│   │   ├── structurer.py     # LLM 结构化患者数据
│   │   ├── auditor.py        # 合规审核
│   │   ├── reporter.py       # HTML 报告生成
│   │   └── pipeline.py       # 流程编排
│   ├── services/
│   │   ├── llm_client.py     # LLM 客户端封装
│   │   └── rules_loader.py   # 规则加载与匹配
│   └── templates/            # Jinja2 HTML 模板
├── config/                   # 药品规则库
├── main.py                   # uvicorn 启动入口
├── requirements.txt
└── .env.example
```

---

## 配置说明

所有配置通过 `.env` 环境变量管理：

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `OPENAI_BASE_URL` | LLM API 地址 | `https://token-plan-cn.xiaomimimo.com/v1` |
| `OPENAI_API_KEY` | API Key | — |
| `OPENAI_MODEL_ID` | 模型 ID | `mimo-v2.5-pro` |
| `OPENAI_MULTIMODAL_MODEL` | 视觉模型（可选） | 空 = 禁用 |
| `MAX_WORKERS` | 并行线程数 | `3` |

---

## 技术栈

- **FastAPI** — Web 框架
- **python-pptx** — PPT 解析
- **OpenAI SDK** — LLM 调用（兼容任意 OpenAI API）
- **Jinja2** — HTML 报告模板
- **pydantic-settings** — 配置管理

---

## License

MIT
