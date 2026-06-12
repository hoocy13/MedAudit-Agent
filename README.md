# MedAudit Agent - 医学病例合规自动审核系统

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-blue.svg" alt="Python 3.10+">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License: MIT">
  <img src="https://img.shields.io/badge/Status-Active-brightgreen.svg" alt="Status: Active">
</p>

> 🏥 自动化医学病例合规审核 Agent，通过大模型与药品规则库进行多维度合规审查，生成专业审核报告。

---

## ✨ 功能特性

- **🤖 智能审核**：基于 LLM 的多维度合规审查，支持基础质控 + 合规专项审核
- **📦 批量处理**：自动扫描目录下所有 PPT 文件，支持并行处理（最多 3 线程）
- **📋 基础质控检查**：文本规范性、通用名、医学逻辑、单位规范 4 项固定检查
- **🔍 合规专项审核**：适应症合规、隐私保护、Storyline 故事线、安全性风险等
- **📄 专业报告**：HTML 审核报告，分组展示，支持 PDF 导出
- **⚡ 高性能**：并行处理 + JSON 容错重试，快速完成批量审核

---

## 📸 报告预览

```
┌─────────────────────────────────────────────────────────┐
│  📄 导出PDF 按钮（右上角）                                │
├─────────────────────────────────────────────────────────┤
│  病例医学合规自动审核意见书                               │
├─────────────────────────────────────────────────────────┤
│  患者信息：编号、性别/年龄、诊断、治疗药物、适应症         │
├─────────────────────────────────────────────────────────┤
│  统计摘要：✓ Pass | ⚠ Warning | ✗ Fail                  │
├─────────────────────────────────────────────────────────┤
│  🔍 基础质控检查（蓝色背景）                              │
│    - 文本规范性检查 | 通用名规范检查                      │
│    - 医学逻辑审核 | 单位规范检查                          │
├─────────────────────────────────────────────────────────┤
│  📋 合规专项审核                                         │
│    - 适应症合规 | 隐私保护 | 安全性风险                   │
│    - Storyline故事线 | 科学幻灯 | 信息完整性             │
└─────────────────────────────────────────────────────────┘
```

---

## 🚀 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/yourusername/MedAudit-Agent.git
cd MedAudit-Agent
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置 API

复制 `.env.example` 为 `.env`，填写 API 配置：

```bash
cp .env.example .env
```

编辑 `.env` 文件：

```env
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL_ID=gpt-4
OPENAI_MULTIMODAL_MODEL=          # 视觉模型，为空则跳过图片识别
```

### 4. 准备 PPT 文件

将待审核的病例 PPT 文件放入 `data/input_cases/` 目录。

### 5. 运行审核

```bash
python main.py
```

审核报告将生成在 `data/output_reports/` 目录。

---

## 📁 项目结构

```
MedAudit-Agent/
│
├── data/
│   ├── input_cases/             # 存放待审核的 PPT 文件
│   └── output_reports/          # 生成的 HTML 报告 + 图片
│       └── images/              # 提取的 PPT 图片（按文件名分子目录）
│
├── config/
│   ├── rules_config.json        # 药品规则库（按产品×适应症结构化）
│   └── rules_checklist.csv      # 规则源文件（便于 Excel 编辑）
│
├── logs/
│   └── medaudit.log             # 运行日志（DEBUG 级别）
│
├── .env.example                 # API 配置模板
├── .gitignore
├── requirements.txt             # 依赖环境
├── main.py                      # 核心运行脚本
├── agent.md                     # 开发说明文档
└── README.md                    # 本文件
```

---

## 📊 审核维度

### 基础质控检查（必选）

| 检查项 | 说明 |
|--------|------|
| **文本规范性检查** | 检查错别字、语法错误、标点符号误用、医学术语拼写错误 |
| **通用名规范检查** | 检查药品是否使用通用名（INN），而非商品名/品牌名 |
| **医学逻辑审核** | 检查药物与适应症匹配性、剂量合理性、疗程合规性 |
| **单位规范检查** | 检查剂量单位、给药频率、检验指标单位等 |

### 合规专项审核（按规则库）

| 维度 | 说明 |
|------|------|
| **适应症合规** | 药品与适应症匹配、年龄/EOS/既往治疗是否符合要求 |
| **隐私保护** | 患者个人信息是否泄露 |
| **Storyline 故事线** | 就诊时间、注射时间、随访时间逻辑通顺性 |
| **安全性风险** | 急性加重期用药、注射间隔合规性 |
| **科学幻灯** | 引用文献、数据一致性 |
| **信息完整性** | 随访数据、既往治疗史完整性 |

---

## 📋 已录入的药品规则

| 药品 | 适应症 | onlabel 规则 | 病例规则 | 科学幻灯规则 |
|------|--------|-------------|---------|-------------|
| 美泊利珠单抗 | SEA | 5 | 8 | 4 |
| 美泊利珠单抗 | COPD | 5 | 6 | 7 |
| 美泊利珠单抗 | CRSwNP | 4 | 7 | 7 |
| 美泊利珠单抗 | EGPA | 2 | 6 | 1 |
| 氟替美维 | COPD | 3 | 10 | 7 |
| 氟替美维 | 哮喘 | 3 | 10 | 7 |
| 西甲硅油 | default | 3 | 3 | 0 |

---

## ⚙️ 核心配置

### rules_config.json

按「产品 → 适应症 → 规则类别」三层结构组织：

```json
{
  "美泊利珠单抗": {
    "SEA": {
      "onlabel_rules": ["成人≥18岁及12岁以上青少年", "入院时EOS≥150", ...],
      "case_rules": ["首页就诊时间需要为本次就诊时间...", ...],
      "slide_rules": ["对应内容补充相关参考文献", ...]
    },
    "COPD": { ... }
  }
}
```

### rules_checklist.csv

CSV 格式便于非技术人员在 Excel 中直接编辑维护。

---

## 🔧 扩展指南

### 新增药品规则

1. 编辑 `config/rules_checklist.csv`，添加新规则行
2. 同步更新 `config/rules_config.json`
3. 重新运行 `python main.py`

### 启用视觉模型

在 `.env` 中设置：
```env
OPENAI_MULTIMODAL_MODEL=your-vision-model-name
```

### 切换 API

修改 `.env` 中的 `OPENAI_BASE_URL`、`OPENAI_API_KEY`、`OPENAI_MODEL_ID` 即可，无需改动代码。

---

## 📈 更新日志

### 2026-06-12

**新功能：**
- ✅ 基础质控检查：4 项固定检查（文本规范性、通用名、医学逻辑、单位规范）
- ✅ PDF 导出：右上角"导出PDF"按钮
- ✅ 并行处理：ThreadPoolExecutor 并行，最多 3 线程
- ✅ 审核报告分组展示：基础质控检查 + 合规专项审核

**优化：**
- ✅ 图片提取错误处理：自动跳过不支持的格式（MPO等）
- ✅ 日志系统：树形结构日志，线程锁防止交错
- ✅ 表格列宽优化：审核维度/状态列紧凑，详情/建议列更宽

---

## 🛠️ 技术栈

- **Python 3.10+**
- **python-pptx**：PPT 文件解析
- **OpenAI API**：大模型调用（支持任意 OpenAI 兼容 API）
- **Jinja2**：HTML 报告模板渲染
- **python-dotenv**：环境变量管理

---

## 📄 License

MIT License

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

## 📧 联系方式

如有问题，请提交 Issue 或联系项目维护者。

---

<p align="center">
  Made with ❤️ by MedAudit Agent Team
</p>
