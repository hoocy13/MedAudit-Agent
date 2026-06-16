# MedAudit Agent — 开发者文档

## 1. 系统架构

FastAPI 服务，接收 PPT 文件上传，经四阶段 Pipeline 审核后返回结构化结果与 HTML 报告。

```
[客户端] ──POST /api/audit/upload──▶ [FastAPI]
                                        │
                                  BackgroundTasks
                                        │
                                        ▼
                              ┌─────────────────────┐
                              │  pipeline.run_audit  │
                              │                     │
                              │  1. extractor       │  python-pptx 提取文本+图片
                              │  2. structurer      │  LLM → 结构化患者 JSON
                              │  3. auditor         │  规则匹配 + LLM 合规审核
                              │  4. reporter        │  Jinja2 → HTML 报告
                              └─────────────────────┘
                                        │
                                        ▼
                              [GET /api/audit/{task_id}]
                              [GET /api/audit/{task_id}/report]
```

### 核心设计决策

- **异步任务模型**：上传即返回 `task_id`，后台 `BackgroundTasks` 执行审核流水线，前端轮询结果
- **内存任务存储**：任务状态存在内存字典（`_tasks`），服务重启后丢失，后续可扩展 Redis/SQLite
- **规则库 + LLM 混合审核**：规则从 `rules_config.json` 模糊匹配后注入 LLM prompt，由 LLM 执行判断

---

## 2. 模块说明

### `app/config.py` — 配置管理

使用 `pydantic-settings` 从 `.env` 加载配置，全局单例 `settings`。

### `app/models.py` — 数据模型

Pydantic v2 模型定义请求/响应结构：`TaskResult`、`AuditResult`、`AuditItem`、`PatientData` 等。

### `app/api/audit.py` — API 路由

5 个端点，任务存储在模块级 `_tasks` 字典。`_process_task()` 作为 `BackgroundTasks` 回调执行完整流水线。

### `app/core/extractor.py` — PPT 提取

- `extract_text_from_pptx()` — 文本框 + 表格
- `extract_images_from_pptx()` — 嵌入图片，跳过不支持格式
- `describe_images_with_llm()` — 多模态视觉识别（需配置 `OPENAI_MULTIMODAL_MODEL`）

### `app/core/structurer.py` — LLM 结构化

`transform_raw_text_to_structured()` — 原始文本 → 12 字段患者 JSON，含药品/适应症约束和后处理。

### `app/core/auditor.py` — 合规审核

`audit_patient_data()` — 患者数据 + 规则 → 审核意见数组。4 项基础质控（必选）+ 规则库专项审核。

### `app/core/reporter.py` — 报告生成

从 `app/templates/` 加载 Jinja2 模板，支持渲染为文件或 HTML 字符串。

### `app/core/pipeline.py` — 流程编排

`run_audit_pipeline()` — 串联 4 个 Agent 的完整处理流程。

### `app/services/llm_client.py` — LLM 客户端

封装 OpenAI 同步/异步客户端，含 `safe_json_parse()` JSON 容错解析与 LLM 重试。

### `app/services/rules_loader.py` — 规则服务

启动时加载 `rules_config.json` 并缓存，提供 `match_rules(drug, indication)` 模糊匹配。

---

## 3. 审核维度

### 基础质控检查（每份报告必含）

| 检查项 | 说明 |
|--------|------|
| 文本规范性 | 错别字、语法错误、标点符号误用、医学术语拼写 |
| 通用名规范 | 药品是否使用通用名（INN），非商品名/品牌名 |
| 医学逻辑 | 药物与适应症匹配、剂量合理性、疗程合规性、联合用药风险 |
| 单位规范 | 剂量/频率/检验指标/体重身高等单位正确性 |

### 合规专项审核（按规则库匹配）

| 维度 | 说明 |
|------|------|
| 适应症合规 | 药品与适应症匹配、年龄/EOS/既往治疗要求 |
| 隐私保护 | 患者个人信息泄露检测 |
| Storyline | 就诊/注射/随访时间逻辑通顺性 |
| 安全性风险 | 急性加重期用药、注射间隔合规性 |
| 科学幻灯 | 引用文献、数据一致性 |
| 信息完整性 | 随访数据、既往治疗史完整性 |

---

## 4. 药品规则库

| 药品 | 适应症 | onlabel | 病例 | 幻灯 |
|------|--------|---------|------|------|
| 美泊利珠单抗 | SEA | 5 | 8 | 4 |
| 美泊利珠单抗 | COPD | 5 | 6 | 7 |
| 美泊利珠单抗 | CRSwNP | 4 | 7 | 7 |
| 美泊利珠单抗 | EGPA | 2 | 6 | 1 |
| 氟替美维 | COPD | 3 | 10 | 7 |
| 氟替美维 | 哮喘 | 3 | 10 | 7 |
| 西甲硅油 | default | 3 | 3 | 0 |

规则文件：`config/rules_config.json`（程序读取） / `config/rules_checklist.csv`（Excel 编辑）

结构：`药品 → 适应症 → { onlabel_rules, case_rules, slide_rules }`

---

## 5. 配置

`.env` 环境变量：

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `OPENAI_BASE_URL` | LLM API 地址 | `https://token-plan-cn.xiaomimimo.com/v1` |
| `OPENAI_API_KEY` | API Key | — |
| `OPENAI_MODEL_ID` | 模型 ID | `mimo-v2.5-pro` |
| `OPENAI_MULTIMODAL_MODEL` | 视觉模型（空=禁用） | 空 |
| `UPLOAD_DIR` | 上传目录 | `data/uploads` |
| `OUTPUT_DIR` | 输出目录 | `data/output_reports` |
| `RULES_CONFIG_PATH` | 规则库路径 | `config/rules_config.json` |
| `MAX_WORKERS` | 并行线程数 | `3` |

---

## 6. 扩展

### 新增药品规则

1. 编辑 `config/rules_checklist.csv` 添加规则
2. 同步更新 `config/rules_config.json`
3. 调用 `POST /api/audit/upload` 验证（或重启服务后规则自动重载）

### 启用视觉模型

设置 `OPENAI_MULTIMODAL_MODEL=your-vision-model`，extractor 会自动调用多模态 LLM 识别 PPT 图片。

### 切换 LLM API

修改 `.env` 中 `OPENAI_BASE_URL`、`OPENAI_API_KEY`、`OPENAI_MODEL_ID`，无需改代码。

---

## 7. 待办

- [ ] 前端页面（Gradio/Streamlit）
- [ ] 持久化存储（SQLite/Redis）
- [ ] WebSocket 实时推送审核进度
- [ ] 用户认证
- [ ] Docker 部署
- [ ] 单元测试

---

## 8. 更新日志

### 2026-06-15

- FastAPI 服务化改造，拆分为模块化 `app/` 包结构
- REST API：文件上传、异步审核、任务查询、报告下载
- 配置管理迁移至 `pydantic-settings`

### 2026-06-12

- 基础质控检查（4 项固定检查）
- PDF 导出、并行处理、审核报告分组展示
- 药品×适应症规则库（3 药品 7 适应症 56 条规则）
