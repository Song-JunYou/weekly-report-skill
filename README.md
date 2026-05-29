# 周报助手 (Weekly Report Skill) for OpenCode

> 告别手动填 Excel！和 AI 聊几句，周报自动生成。

一个 [OpenCode](https://opencode.ai) Skill，通过自然语言对话，自动将你的工作描述整理并填充到公司标准的 Excel 周报模板中。

---

## 🚀 快速开始

### 1. 安装（30 秒）

```bash
# Linux/macOS
git clone https://github.com/Song-JunYou/weekly-report-skill.git \
  ~/.config/opencode/superpowers/skills/weekly-report

# Windows (PowerShell)
git clone https://github.com/Song-JunYou/weekly-report-skill.git \
  $env:USERPROFILE\.config\opencode\superpowers\skills\weekly-report
```

> 重启 OpenCode，系统会自动加载新 Skill。

### 2. 使用

对 OpenCode 说任意一句：

| 中文 | English |
|------|---------|
| "帮我写周报" | "weekly report" |
| "写周报" | "fill my weekly report" |
| "更新周报" | "update weekly report" |

AI 会进入向导模式，按流程引导你：

```
1. 确认基本信息（周期、姓名、部门）
2. 收集上周工作（内容、状态、成果、问题）
3. 收集本周工作（内容、优先级、协作人、时间）
4. 收集下周计划（同上）
5. AI 自动评估饱和度（≥85%）+ 询问协调事项
6. 完整预览，等你确认 "生成吧"
7. 自动生成 Excel 文件到桌面
```

---

## 功能特性

- **对话式交互** —— 不需要记模板格式，告诉 AI "我这周做了什么" 就行
- **智能填充** —— 自动将工作归类到「上周完成 / 本周重点 / 下周计划」三大板块
- **自动评估饱和度** —— AI 根据工作量自动计算工作饱和度，最低保证 **85%**
- **保留模板样式** —— 使用 `openpyxl` 在原模板上填充，**不破坏任何格式和公式**
- **文件名智能命名** —— 自动使用「姓名 + 完整周期范围」，如 `2026-05-25至2026-05-29`
- **协作人追踪** —— 主动追问每条工作的协作人，避免遗漏
- **确认后生成** —— 收集完所有信息并给你完整预览，确认无误再生成文件

---

## 一图看懂

### 🔄 交互流程

<img src="https://raw.githubusercontent.com/Song-JunYou/weekly-report-skill/main/docs/assets/workflow.png" alt="交互流程图" width="800"/>

> 和 AI 聊 8 步，周报自动生成。不用记格式，不用手动填 Excel。

### 🧠 饱和度怎么算的？

<img src="https://raw.githubusercontent.com/Song-JunYou/weekly-report-skill/main/docs/assets/saturation.png" alt="饱和度计算逻辑" width="600"/>

> AI 根据你的工作量自动评估，**最低保证 85%**，无需手动填写。

### 🏗️ 项目架构

<img src="https://raw.githubusercontent.com/Song-JunYou/weekly-report-skill/main/docs/assets/architecture.png" alt="项目架构图" width="700"/>

> 用户 → AI 解析 → Python 脚本 → 填充模板 → 生成文件。全程保留 Excel 格式。

---

## 项目结构

```
weekly-report-skill/
├── SKILL.md                    # Skill 定义（交互规范、规则、工作流）
├── README.md                   # 本文件
├── .gitignore
├── evals/
│   └── evals.json              # 测试用例
├── scripts/
│   └── generate_report.py      # 核心生成脚本（openpyxl）
└── templates/
    └── 赛仕科工作周报模板.xlsx   # 内置 Excel 模板（随 Skill 自带）
```

---

## 模板说明

`赛仕科工作周报模板.xlsx` 包含以下板块：

| 板块 | 字段 |
|------|------|
| **表头** | 周期、姓名、部门 |
| **上周工作完成情况** | 工作内容、时间、状态、成果、问题、解决方案 |
| **本周重要工作项** | 工作内容、优先级、时间、责任人、**协作人** |
| **下周工作计划** | 同上 |
| **工作饱和度评估** | 饱和度（%）、负荷说明、协调资源、其他建议 |

你也可以使用自己的模板：

```bash
export WKR_TEMPLATE_PATH="/path/to/your/周报模板.xlsx"
```

---

## 配置选项

通过环境变量自定义行为：

| 环境变量 | 说明 | 默认值 |
|----------|------|--------|
| `WKR_TEMPLATE_PATH` | 自定义 Excel 模板路径 | Skill 自带模板 |
| `WKR_OUTPUT_DIR` | 周报输出目录 | `~/Desktop/周报/` |

---

## 依赖

- Python 3.8+
- openpyxl

```bash
pip install openpyxl
```

---

## 使用示例

**场景一：自由描述**

```
用户：我这周搞定了数据迁移，还在搞 API 优化，下周要处理 bug。

AI：好的！我记录一下。请再确认几个细节：
  1. 数据迁移的完成状态？
  2. API 优化预计什么时候完成？协作人是谁？
  3. 下周处理 bug 的优先级？
  ...（追问完毕，展示预览）
  以上信息确认无误吗？确认后生成周报。

用户：确认

AI：周报已生成！📄 赛仕科工作周报 - OLIVER SONG 宋俊佑 2026-05-25至2026-05-29.xlsx
```

**场景二：向导模式**

```
用户：写周报

AI：好的，我来帮你写周报！首先确认本周周期是 2026/5/25-5/29 吗？
     ...（逐项引导，不跳过任何板块）
```

---

## 核心设计

### 为什么不用 pandas？

因为周报模板通常带有：
- 合并单元格（表头、标题）
- 条件格式（优先级高亮）
- 公式（进度条自动计算）
- 公司 Logo 和样式

`pandas` 生成的新文件会丢失这些格式。我们使用 `openpyxl` 直接**在原模板上修改**，保留一切样式。

### 饱和度怎么算？

AI 根据上周、本周、下周的工作量自动评估：

| 工作量 | 饱和度 |
|--------|--------|
| 1-2 项 | 85% ~ 90% |
| 3-4 项 | 90% ~ 95% |
| 5 项以上 / 跨多项目 | 95% ~ 100% |

**最低不低于 85%**，无需你手动填写。

---

## 更新日志

| 版本 | 更新内容 |
|------|----------|
| v1.1 | 饱和度自动评估、文件名完整周期格式、强制询问所有板块 |
| v1.0 | 初始版本，支持向导式周报生成 |

---

## 贡献

欢迎 Issue 和 PR！

- 发现 Bug → 提 [Issue](https://github.com/Song-JunYou/weekly-report-skill/issues)
- 想支持其他模板格式 → 提 PR
- 有功能建议 → 先开 Issue 讨论

---

## 鸣谢

- 基于 [OpenCode](https://opencode.ai) 的 Superpowers Skill 框架开发
- 模板样式感谢赛仕科（SATSHK）的标准化周报设计

---

<div align="center">

Made with ☕ by [Song-JunYou](https://github.com/Song-JunYou)

</div>
