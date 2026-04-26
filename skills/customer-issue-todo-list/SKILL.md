---
name: customer-issue-todo-list
description: Use when the user imports a meeting note, customer feedback document, issue collection record, DOCX, or Markdown file and asks to generate, regenerate, extract, or output a “客户问题收集待办列表”, “问题收集待办”, “重点待办台账”, “问题清单详情”, “客户问题模板”, or similar Chinese follow-up document.
---

# Customer Issue Todo List

## Overview

Extract an imported document into the same template style as the reference document `2026-04-24下午 张顾问跟进优邦PLM系统录入问题.docx`. Treat every structural element in that reference as template requirements; replace only the concrete meeting content with the new document's facts.

## Required Output Shape

Use `assets/customer-issue-todo-template.md` as the required skeleton. Preserve this order:

1. Title line: `{日期/时间段} {跟进人或会议名称}{客户/系统}{问题主题}`
2. Metadata line: `会议时间：{时间}参会人员：{人员列表}`
3. `一、会议结论简述`
4. `二、重点待办台账`
5. `三、问题清单详情描述如下`
6. Priority detail groups:
   - `高优先级待办`
   - `中优先级待办`
   - `低优先级待办（体验优化）`

Do not add unrelated summary sections such as “按责任人汇总”“建议推进顺序” unless the user explicitly asks. The reference document's structure is the contract.

## Extraction Rules

- Keep the document in Chinese.
- Use Markdown by default. If the user asks for Word, create a DOCX using the same content structure.
- Preserve source-specific names, systems, modules, account names, material codes, examples, and dates.
- Do not invent people, statuses, deadlines, or impact ranges. If missing, write `待确认`.
- Merge duplicate issues only when they clearly describe the same problem and share the same owner/priority.
- Split an issue when the source implies different owners, priorities, or follow-up actions.

## Section Rules

### Title And Metadata

- Derive the title from the source filename or document title.
- Keep a single metadata line after the title:
  `会议时间： {会议时间}参会人员： {参会人员}`
- If either value is missing, use `待确认`.

### 一、会议结论简述

Write 1 short introductory sentence, then 2-4 numbered category sentences in this style:

- `一是{问题类别}，{影响说明}。`
- `二是{问题类别}，{影响说明}。`
- `三是{问题类别}，{影响说明}。`

The categories should be abstracted from the source, not copied from the reference meeting unless they apply.

### 二、重点待办台账

Always include the explanation lines:

- `优先级分为 P1（高）、P2（中）、P3（低）`
- `状态分为 未开始、进行中、已完成`

Use this exact table schema:

| 序号 | 待办事项 | 责任人 | 优先级 | 状态 | 备注 |

Todo item wording should start with a concrete action, such as `完成`, `修复`, `排查`, `补充`, `系统性核查`, `优化`, `增加`, `评估`, `讨论`, `调整`, `支持`, `获取`, `沟通确认`, or `持续`.

### 三、问题清单详情描述如下

Group detailed items by priority. Number items continuously across all priority groups, following the reference style.

Each detail item should use this pattern when information exists:

```markdown
{序号}. {问题名称}
问题描述： {现象/背景/规则/用户反馈}
影响范围： {受影响对象/场景/用户/流程}
待办事项：
- {行动一}
- {行动二}
```

If `影响范围` is absent in the source, omit that line instead of inventing it. If a future design is discussed but not confirmed, use `待讨论设计功能：{内容}`.

## Priority And Status Rules

Priority:

- `P1`: data accuracy, master data completeness, blocking use, release commitment, business continuity, customer-visible critical display, or urgent repair.
- `P2`: workflow flexibility, approval/signoff, version traceability, collaboration rules, or ongoing collection/validation.
- `P3`: usability, default values, search convenience, popup/table/rich-text experience, documentation clarification, or low-risk enhancement.

Status:

- `已完成`: source states completed, released, imported, fixed, synchronized, or closed.
- `进行中`: source states ongoing, continuing, under investigation, being used, or serving as a feedback/validation input.
- `未开始`: no progress is stated.

Owner:

- Keep named owners exactly when present.
- For vendor or implementation tasks, keep the source's team name if present; otherwise use `开发团队`.
- For customer-side confirmation, use the named person/role from source; otherwise use `客户方负责人`.
- For ongoing feedback collection by users, use the named team from source; otherwise use `使用团队`.

## Quality Bar

- The output should read like a usable internal follow-up document, not a generic summary.
- Preserve concrete evidence in `备注`, such as examples, codes, affected accounts, or planned release dates.
- Keep table remarks short; put longer explanation in the detail section.
- Avoid adding advice, interpretation, or new fields that are not in the template.

## DOCX Requests

If the user asks for a Word file, use the documents skill and visually verify the rendered DOCX before delivery. Keep the same order, headings, table fields, and detail-item pattern.
