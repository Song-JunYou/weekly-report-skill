# Weekly Report Skill (周报助手)

## 目录结构

```
weekly-report/
├── SKILL.md                      # Skill 定义和交互规范
├── README.md                     # 本文件
├── evals/
│   └── evals.json               # 测试用例
├── scripts/
│   └── generate_report.py       # Python 生成脚本
└── templates/
    └── 赛仕科工作周报模板.xlsx   # Excel 周报模板
```

## 安装方法

### 方式一：复制到 OpenCode Skills 目录

1. 将本 `weekly-report` 文件夹复制到 OpenCode 的 skills 目录：

```bash
# Linux/macOS
cp -r weekly-report ~/.config/opencode/superpowers/skills/

# Windows
xcopy /E /I weekly-report %USERPROFILE%\.config\opencode\superpowers\skills\
```

2. **重启 OpenCode**（关闭并重新打开），系统会自动扫描并加载新 Skill

### 方式二：通过 OpenCode CLI 安装

```bash
# 如果 OpenCode 支持 skill 安装命令
opencode skill install weekly-report
```

## 使用方法

安装并重启后，直接对 OpenCode 说：

- **"帮我写周报"**
- **"写周报"**
- **"weekly report"**
- **"更新周报"**

AI 会自动进入向导模式，按以下流程引导你：

1. **确认基本信息**：周期、姓名、部门
2. **收集上周工作**：工作内容、完成状态、成果、问题
3. **收集本周工作**：工作内容、优先级、时间、协作人（必须问）
4. **收集下周计划**：工作内容、优先级、时间、协作人（必须问）
5. **饱和度评估**：AI 自动根据工作量评估（最低 85%）
6. **询问协调事项**：需协调的资源、其他说明
7. **预览确认**：展示所有信息，等待你确认
8. **生成文件**：确认后自动生成 Excel 周报

## 高级配置

可通过环境变量自定义路径：

```bash
# 使用自定义模板（可选，默认使用 Skill 自带模板）
export WKR_TEMPLATE_PATH="/path/to/your/template.xlsx"

# 自定义输出目录（可选，默认桌面/周报/）
export WKR_OUTPUT_DIR="/path/to/output/dir/"
```

## 模板说明

本 Skill 自带标准周报模板 `赛仕科工作周报模板.xlsx`，包含以下板块：

- **表头**：周期、姓名、部门
- **上周工作完成情况**：序号、工作内容、时间、状态、成果、问题
- **本周重要工作项**：序号、内容、优先级、时间、责任人、协作人
- **下周工作计划**：同上
- **工作饱和度评估**：饱和度（%）、负荷说明、协调资源、其他建议

## 依赖要求

- Python 3.8+
- openpyxl (`pip install openpyxl`)

## 测试

```bash
# 测试脚本
python3 scripts/generate_report.py --test
```

## 注意事项

- **必须先重启 OpenCode** 才能识别新安装的 Skill
- Skill 采用向导式交互，必须等所有信息收集完成并确认后才生成文件
- 饱和度由 AI 自动评估，最低不低于 85%
- 文件名使用完整周期范围（如：`2026-05-25至2026-05-29`）

## 更新日志

- **v1.1**: 修复饱和度自动评估、文件名格式、强制询问所有板块
- **v1.0**: 初始版本，支持向导式周报生成
