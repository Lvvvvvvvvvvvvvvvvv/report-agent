# 数据分析报告智能体

> **G 端安防行业 · 公安监狱 · 下级向上级汇报的数据分析报告自动化生成工具**

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

---

## 项目简介

本项目是一款面向公安监狱等 G 端安防行业的**数据分析报告撰写智能体**。

核心场景：监区/中队的管理人员将**罪犯训练任务数据**与**异常行为数据**导入工具，系统自动完成统计分析、图表生成，并输出符合机关公文规范的 Word 格式报告（.docx）。

接入小米 **Mimo 大模型**后，可在保留所有数字与事实不变的前提下，对报告正文进行专业化润色，使文风更接近真实公文。

---

## 核心功能

| 功能模块 | 说明 |
|---|---|
| 📊 数据分析 | 自动统计训练达标率、出勤率、成绩分布；识别重点关注人员 |
| ⚠️ 异常行为分析 | 自动汇总异常类型、风险等级、中队分布；生成高风险个案清单 |
| 📈 可视化图表 | 自动生成 4 张 matplotlib 图表（中队对比、类别对比、异常类型、风险分布）|
| 📄 Word 报告 | 一键输出含封面、目录、正文、图表的规范公文 .docx |
| 🤖 大模型润色 | 接入 Mimo（小米）大模型，保留数字事实，自动优化行文风格 |
| ⚙️ 灵活配置 | 支持自定义单位名称、汇报周期、密级、达标阈值、章节开关等 |

---

## 智能体架构

```
输入层              大脑层                        输出层
──────────────────────────────────────────────────────────
Excel 数据    →   统计计算（pandas）          →
配置参数      →   可视化（matplotlib）        →   Word 报告 (.docx)
                  报告起草（python-docx）     →
                  语言润色（Mimo 大模型）     →
```

**核心设计原则：代码计算事实，大模型只润色文字。**  
所有数字、人名、日期均由 Python 代码精确计算，LLM 仅负责改善语言表达，从根源上杜绝数据幻觉。

---

## 快速开始（本地运行）

### 1. 克隆项目

```bash
git clone https://github.com/YOUR_USERNAME/report-agent.git
cd report-agent
```

### 2. 创建虚拟环境并安装依赖

```bash
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. （可选）配置大模型 API Key

如需启用 Mimo 大模型润色功能，在项目根目录新建 `mimo_key.txt`，把你的 Mimo Token Plan Key 粘贴进去（仅一行），保存。

```
tp-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

> ⚠️ `mimo_key.txt` 已在 `.gitignore` 中排除，**不会被上传到 GitHub**。请勿将 Key 分享给他人。

### 4. 生成示例数据

```bash
python generate_data.py
```

这会在 `data/` 目录生成两张示例 Excel 文件（60 名罪犯、3 个中队的脱敏虚构数据）。

### 5. 启动 Web 界面

```bash
streamlit run app.py
```

浏览器自动打开 `http://localhost:8501`。

macOS 用户也可直接**双击** `启动报告工具.command` 启动。

---

## 文件说明

```
report-agent/
├── app.py                    # Streamlit Web 前端（主入口）
├── report_agent.py           # 核心引擎：数据分析 + 图表 + Word 生成 + LLM 润色
├── generate_data.py          # 生成脱敏示例数据
├── requirements.txt          # Python 依赖
├── .gitignore                # Git 忽略规则（含 mimo_key.txt）
├── 启动报告工具.command       # macOS 双击启动脚本
├── 智能体全景图.html          # 智能体架构可视化（面试/汇报用）
├── data/
│   ├── 训练任务数据.xlsx      # 示例训练数据（虚构脱敏）
│   └── 异常行为数据.xlsx      # 示例异常数据（虚构脱敏）
├── output/                   # 生成的报告输出目录（本地）
└── .streamlit/
    └── config.toml           # Streamlit 主题配置
```

---

## 报告配置项说明

| 配置项 | 说明 | 默认值 |
|---|---|---|
| 报送单位 | 撰写报告的单位名称 | 三监区 |
| 呈报单位 | 上级接收单位 | 监狱教育改造科 |
| 编制/审核/批准 | 签字人信息 | 空（可选填）|
| 汇报周期 | 日/周/月度/季度/年度 | 月度 |
| 密级标注 | 公开/内部资料/保密/机密 | 内部资料 |
| 成绩达标线 | 低于此分数记为未达标 | 60 分 |
| 出勤率达标线 | 低于此比例记为未达标 | 80% |
| 重点关注成绩线 | 低于此分数列入重点关注人员 | 70 分 |
| 统计范围 | 选择纳入统计的中队 | 全部 |
| 章节开关 | 精细控制报告包含哪些章节 | 全部开启 |

---

## 部署到 Streamlit Community Cloud

1. 将代码 Push 到 GitHub 公开仓库（参见下方 GitHub 发布步骤）
2. 访问 [share.streamlit.io](https://share.streamlit.io)，用 GitHub 账号登录
3. 点击 **New app** → 选择本仓库 → Main file: `app.py` → Deploy
4. 部署完成后，进入 App Settings → **Secrets**，添加：
   ```toml
   MIMO_API_KEY = "tp-你的Key"
   ```
5. 保存后 App 自动重启，大模型功能即可在公网使用

---

## 技术栈

| 组件 | 技术 |
|---|---|
| Web 前端 | Streamlit |
| 数据处理 | pandas · numpy |
| 可视化 | matplotlib |
| Word 生成 | python-docx |
| 大模型 | 小米 Mimo（Anthropic 兼容 API）|
| 运行环境 | Python 3.9+ |

---

## 数据安全说明

- 本项目所有示例数据均为**虚构脱敏数据**，不含任何真实人员信息
- 公安监狱数据属高度敏感信息，生产环境请务必在内网/离线环境部署
- API Key 通过本地文件或 Streamlit Secrets 管理，不进入代码仓库
- 输出报告如含真实数据，请按单位保密规定妥善保管

---

## License

MIT License — 仅供学习与演示使用。生产环境接入真实涉密数据前，请进行安全评估。
