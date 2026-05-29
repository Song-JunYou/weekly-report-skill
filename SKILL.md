---
name: weekly-report
description: Use when the user mentions writing, updating, creating, or filling out their weekly work report, work summary, or 周报. Triggers on phrases like "写周报", "更新周报", "weekly report", "工作总结", or when the user describes work progress that should be formatted into the company's Excel weekly report template.
---

# 周报助手 (Weekly Report Assistant)

## Overview
通过对话式交互，自动将用户的工作描述整理并填充到公司标准的 Excel 周报模板中。

**核心原则：在收集完所有必填信息并经过用户明确确认前，绝不可以生成文件。**

## When to Use
- 用户说"写周报"、"更新周报"、"weekly report"
- 用户描述一周工作内容，需要整理成正式周报格式
- 用户需要基于模板生成新的周报文件
- **When NOT to use**: 用户只是询问模板结构或查看历史周报，不需要生成新文件

## CRITICAL RULES (绝不违反)

1. **绝不提前生成文件**：在收集完所有 4 大板块（上周、本周、下周、饱和度）信息前，绝不可以调用生成脚本。
2. **必须用户确认**：生成文件前，必须向用户完整展示所有已收集信息，获得用户明确同意（如"确认生成"、"可以了"、"生成吧"）。
3. **文件名必须包含完整周期**：`赛仕科工作周报 - {姓名} {周期起始日期}-{周期结束日期}.xlsx`，例如 `赛仕科工作周报 - OLIVER SONG 宋俊佑 2026-05-25至2026-05-29.xlsx`
4. **饱和度使用百分号**：如 "95%"，绝不使用小数 "0.95"。

## Required Information Checklist (必填信息清单)

生成周报前，必须确认以下信息已收集完整：

### ☐ 1. 基本信息
- [ ] 周报周期（如：2026/5/25-2026/5/29）
- [ ] 姓名（默认 OLIVER SONG 宋俊佑）
- [ ] 部门（默认 IT信息技术部）

### ☐ 2. 上周工作完成情况（至少1条，建议2-4条）
每条必须包含：
- [ ] 工作内容
- [ ] 完成状态（已完成/进行中/未开始/延期）
- [ ] 成果说明
- [ ] 遇到的问题（可选，但建议询问）
- [ ] 解决方案（可选）

### ☐ 3. 本周重点工作（至少1条，建议2-4条）
每条必须包含：
- [ ] 工作内容
- [ ] 优先级（高/中/低）
- [ ] 计划开始时间
- [ ] 计划完成时间
- [ ] 责任人（默认本人）
- [ ] 协作人（必须询问！不可遗漏）
- [ ] 遇到的问题（可选）

### ☐ 4. 下周工作计划（至少1条，建议2-4条）
每条必须包含：
- [ ] 工作内容
- [ ] 优先级（高/中/低）
- [ ] 计划开始时间
- [ ] 计划完成时间
- [ ] 责任人（默认本人）
- [ ] 协作人（必须询问！不可遗漏）
- [ ] 遇到的问题（可选）

### ☐ 5. 工作饱和度评估与其他说明（AI自动评估）
- [x] 整体工作饱和度 — **由AI根据工作内容自动评估，默认85%以上**
- [x] 工作负荷说明 — **由AI根据工作内容自动生成**
- [ ] 需协调支持的资源/事项 — 必须询问用户
- [ ] 其他工作说明/建议 — 必须询问用户

**饱和度规则**：AI根据上周、本周、下周的工作量自动计算，最低不低于85%。
- 1-2项工作：85%-90%
- 3-4项工作：90%-95%
- 5项以上或跨多项目：95%-100%

**以上 5 大板块全部确认完成后，才允许生成周报文件。**

## Core Workflow

### 向导式交互流程 (强制使用)

```
1. 【开场】用户说"写周报"
   → AI: "好的，我来帮你写周报！首先确认一下，本周的周报周期是 XX 吗？"
   → 确认周期、姓名、部门

2. 【信息收集阶段】严格按照清单逐项询问，不可跳过任何板块：
   
   2.1 上周工作完成情况
   → "请告诉我上周完成了哪些工作？可以一条一条说，我会记录。"
   → 对每条工作追问：完成状态？成果说明？遇到问题了吗？怎么解决的？
   
   2.2 本周重点工作
   → "本周有哪些重点工作？"
   → 对每条工作追问：优先级？计划时间？协作人是谁？（必须问协作人！）
   
   2.3 下周工作计划  
   → "下周计划做哪些工作？"
   → 对每条工作追问：优先级？计划时间？协作人是谁？（必须问协作人！）
   
   2.4 工作饱和度评估（AI自动计算，告知用户）
   → AI根据已收集的上周、本周、下周工作内容，自动评估饱和度（最低85%）
   → 向用户说明："根据您本周的工作内容，我评估工作饱和度为XX%，主要因为..."
   → 询问："有需要协调的资源或支持吗？"
   → 询问："还有其他需要说明的事项吗？"

3. 【预览确认阶段】所有信息收集完毕后：
   → 向用户完整展示所有已收集内容（按板块列出）
   → "以上内容确认无误吗？确认后我将生成周报文件。"
   → 等待用户明确回复"确认"、"可以"、"生成吧"等

4. 【生成阶段】获得用户明确确认后：
   → 调用 generate_report.py 生成 Excel 文件
   → 向用户报告文件路径和生成结果
```

### 自由式（用户一次性描述）

```
1. 用户一次性描述所有工作内容
2. AI 解析并归类到对应板块
3. AI 主动识别缺失项，向用户补充询问（特别是协作人、协调事项）
4. AI 根据所有工作内容自动评估饱和度（最低85%），告知用户
5. 向用户展示完整预览
6. 获得用户明确确认后生成
```

## Implementation

### 文件位置约定
- **模板文件**: 默认随 Skill 自带，位于 `templates/赛仕科工作周报模板.xlsx`
  - 支持通过环境变量 `WKR_TEMPLATE_PATH` 覆盖
  - 若用户已有自定义模板，优先使用用户配置的路径
- **输出目录**: 默认输出到用户桌面周报文件夹（可通过 `WKR_OUTPUT_DIR` 覆盖）
- **输出文件名**: `赛仕科工作周报 - {姓名} {周期起始日期}至{周期结束日期}.xlsx`
  - 例如：`赛仕科工作周报 - OLIVER SONG 宋俊佑 2026-05-25至2026-05-29.xlsx`
  - 绝不使用单日期：`赛仕科工作周报 - OLIVER SONG 宋俊佑 2026-05-29.xlsx` ❌

### 数据结构映射

#### 上周工作完成情况
| 列 | 字段 | 说明 |
|---|---|------|
| B | 序号 | 自动填充 1,2,3... |
| C | 工作内容 | 用户描述 |
| D | 预计完成时间 | YYYY-MM-DD |
| E | 实际完成时间 | YYYY-MM-DD |
| F | 完成状态 | 下拉：已完成/进行中/未开始/延期 |
| G | 完成进度 | 保留模板公式，不覆盖 |
| H | 成果说明 | 多行文本，wrap_text |
| I | 遇到的问题 | 多行文本 |
| J | 解决方案 | 多行文本 |
| K | 备注 | 多行文本 |

#### 本周重要工作项
| 列 | 字段 | 说明 |
|---|---|------|
| B | 序号 | 自动填充 |
| C | 工作内容 | 用户描述 |
| D | 优先级 | 下拉：高/中/低 |
| E | 计划开始时间 | YYYY-MM-DD |
| F | 计划完成时间 | YYYY-MM-DD |
| G | 责任人 | 默认"OLIVER SONG 宋俊佑" |
| H | 协作人 | 必须询问，不可遗漏 |
| I | 遇到的问题 | 多行文本 |
| J | 预计完成时间 | YYYY-MM-DD |
| K | 备注 | 多行文本 |

#### 下周工作计划
结构同"本周重要工作项"

#### 工作饱和度评估
| 位置 | 字段 | 说明 |
|------|------|------|
| C28 | 整体工作饱和度 | **百分号格式**，如 "95%"、"80%" |
| F28 | 工作负荷说明 | 多行文本 |
| C29 | 需协调支持的资源 | 多行文本 |
| F29 | 其他工作说明/建议 | 多行文本 |

### Python 实现代码

```python
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import datetime
import os

def find_section_rows(sheet):
    """动态查找各板块标题行和数据起始行"""
    sections = {}
    for row in range(1, sheet.max_row + 1):
        cell_val = sheet.cell(row=row, column=1).value
        if cell_val and isinstance(cell_val, str):
            if '上周工作完成情况' in cell_val:
                sections['last_week_data_start'] = row + 2
            elif '本周重要工作项' in cell_val:
                sections['this_week_data_start'] = row + 2
            elif '下周工作计划' in cell_val:
                sections['next_week_data_start'] = row + 2
            elif '工作饱和度评估' in cell_val:
                sections['saturation_data'] = row + 1
    return sections

def set_cell_value(sheet, row, col, value):
    cell = sheet.cell(row=row, column=col, value=value)
    if isinstance(value, str) and '\n' in value:
        cell.alignment = Alignment(wrap_text=True, vertical='top')
    return cell

def generate_weekly_report(report_data):
    template_path = os.environ.get('WKR_TEMPLATE_PATH', 
        "/mnt/c/Users/OLIVER_SONG.AADDS/Desktop/周报/模板/赛仕科工作周报模板.xlsx")
    wb = load_workbook(template_path)
    sheet = wb['员工工作周报']
    sections = find_section_rows(sheet)
    
    # 表头 (Row 2)
    sheet['B2'] = report_data.get('period', '')
    sheet['E2'] = report_data.get('name', 'OLIVER SONG 宋俊佑')
    sheet['H2'] = report_data.get('department', 'IT信息技术部')
    
    # 上周完成情况
    last_week_cols = {
        'content': 2, 'planned_date': 3, 'actual_date': 4,
        'status': 5, 'results': 7,
        'problems': 8, 'solutions': 9, 'notes': 10
    }
    for i, item in enumerate(report_data.get('last_week', [])):
        row = sections['last_week_data_start'] + i
        for field, col in last_week_cols.items():
            if field in item and item[field] is not None:
                set_cell_value(sheet, row, col, item[field])
    
    # 本周工作项
    this_week_cols = {
        'content': 2, 'priority': 3, 'start_date': 4, 'end_date': 5,
        'owner': 6, 'collaborators': 7, 'problems': 8, 'eta': 9, 'notes': 10
    }
    for i, item in enumerate(report_data.get('this_week', [])):
        row = sections['this_week_data_start'] + i
        for field, col in this_week_cols.items():
            if field in item and item[field] is not None:
                set_cell_value(sheet, row, col, item[field])
    
    # 下周计划
    for i, item in enumerate(report_data.get('next_week', [])):
        row = sections['next_week_data_start'] + i
        for field, col in this_week_cols.items():
            if field in item and item[field] is not None:
                set_cell_value(sheet, row, col, item[field])
    
    # 饱和度评估 - 使用百分号格式
    sat_row = sections['saturation_data']
    saturation = report_data.get('saturation', '')
    # 确保饱和度是百分号格式
    if saturation and not str(saturation).endswith('%'):
        saturation = f"{saturation}%"
    sheet.cell(row=sat_row, column=2, value=saturation)
    sheet.cell(row=sat_row, column=5, value=report_data.get('workload_desc', ''))
    sheet.cell(row=sat_row + 1, column=2, value=report_data.get('resources_needed', ''))
    sheet.cell(row=sat_row + 1, column=5, value=report_data.get('other_notes', ''))
    
    # 保存：文件名使用完整周期范围
    period = report_data.get('period', '')
    if '-' in period:
        parts = period.split('-')
        if len(parts) == 2:
            start_date = parts[0].strip().replace('/', '-')
            end_date = parts[1].strip().replace('/', '-')
            # 统一格式为 YYYY-MM-DD
            try:
                start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                end_dt = datetime.strptime(end_date, '%Y-%m-%d')
                date_suffix = f"{start_dt.strftime('%Y-%m-%d')}至{end_dt.strftime('%Y-%m-%d')}"
            except ValueError:
                date_suffix = f"{start_date}至{end_date}"
        else:
            date_suffix = datetime.now().strftime('%Y-%m-%d')
    else:
        date_suffix = datetime.now().strftime('%Y-%m-%d')
    
    filename = f"赛仕科工作周报 - {report_data.get('name', 'OLIVER SONG 宋俊佑')} {date_suffix}.xlsx"
    output_path = os.path.join(
        os.environ.get('WKR_OUTPUT_DIR', "/mnt/c/Users/OLIVER_SONG.AADDS/Desktop/周报/"),
        filename
    )
    wb.save(output_path)
    return output_path
```

### 自动迁移功能 (可选)

如果用户选择"基于上周周报生成":
1. 找到最新的周报文件
2. 读取"本周重要工作项" → 作为新的"上周工作完成情况"
3. 读取"下周工作计划" → 作为新的"本周重要工作项"
4. 清空"下周工作计划"，等待用户补充
5. AI根据迁移后的工作内容自动评估饱和度（最低85%），并告知用户
6. 询问协调事项、其他说明等信息

## Quick Reference

| 用户场景 | 操作 |
|---------|------|
| "写周报" | 启动向导模式，严格按照清单逐项引导，绝不提前生成 |
| "更新周报" | 询问基于上周迁移还是全新填写，仍需完整收集所有信息 |
| "这周我做了..." | 解析内容，自动归类，识别缺失项并补充询问，预览后确认生成 |
| "帮我看看上周周报" | 读取并展示上周内容，不生成新文件 |

## Common Mistakes

1. **绝不在信息收集完整前生成文件** - 必须等4大板块全部确认
2. **绝不遗漏协作人** - 本周和下周的每条工作都必须询问协作人
3. **饱和度由AI自动评估** - 根据工作内容自动计算，最低85%，无需询问用户但需告知用户评估结果
4. **文件名必须包含完整周期** - 使用 "2026-05-25至2026-05-29" 格式
5. **饱和度使用百分号** - "95%" 而非 "0.95"
6. **不要猜测信息** - 用户未提及的字段留空或追问，不要编造
7. **保留模板格式** - 使用 openpyxl 加载模板后修改，不要新建空白文件
8. **日期格式统一** - 始终使用 YYYY-MM-DD 格式
9. **多行文本处理** - 设置 wrap_text=True，确保换行显示正常
10. **完成状态验证** - 只接受：已完成/进行中/未开始/延期
11. **优先级验证** - 只接受：高/中/低

## Red Flags

| 错误思维 | 正确做法 |
|---------|------|
| "用户说了一些我就生成" | 收集完所有4大板块并预览确认后再生成 |
| "先随便填一个" | 确认后再写入，不要猜测 |
| "覆盖上周文件" | 总是生成新文件，保留历史 |
| "用 pandas 生成" | 用 openpyxl 保留模板格式和样式 |
| "用户说完了就保存" | 生成前给用户完整预览，等待明确确认 |
| "协作人用户没说就算了" | 主动追问每条工作的协作人 |
| "饱和度用户没提就跳过" | AI自动根据工作内容评估饱和度（最低85%），并告知用户 |
| "文件名用单日期就行" | 必须用完整周期范围作为文件名 |
