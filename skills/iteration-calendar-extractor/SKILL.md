---
name: "iteration-calendar-extractor"
description: "当用户需要从 Excel 迭代日历中抽取结构化迭代数据时使用。默认读取第一个工作表，从第 9 行开始，输出格式为“迭代I42，2025/12/31 是迭代启动日期，2026/1/14 是迭代发布日期”。"
---

# 迭代日历抽取器

此 skill 用于从 Excel 文件第一个工作表中抽取迭代日历数据。默认从第 9 行开始读取，并将每个迭代周期格式化为易读的 Markdown 条目。

## 适用场景

- 从 Excel 文件中抽取迭代日历数据
- 将 Excel 中的 sprint/迭代计划转换为 Markdown
- 处理包含迭代编号、启动日期和发布日期的项目时间表
- 生成标准格式的迭代周期清单

## 使用方法

1. 使用 Excel 文件路径运行 PowerShell 脚本。
2. 脚本只读取**第一个工作表**。
3. 从**第 9 行**开始处理，默认跳过表头。
4. 抽取迭代编号、启动日期和发布日期。
5. 输出格式为：`迭代I42，2025/12/31 是迭代启动日期，2026/1/14 是迭代发布日期`。
6. 生成 Markdown 文档。

## 预期 Excel 结构

Excel 第一个工作表从第 9 行开始应包含以下列：

- A 列：迭代编号，例如 I42、I43
- B 列：启动日期，例如 2025/12/31
- C 列：结束/发布日期，例如 2026/1/14

## 输出格式

```markdown
# 迭代日历

**来源：** 2026年迭代日历V2.xlsx

**生成时间：** 2026-04-21 13:14:00

---

## 迭代周期列表

- 迭代I42，2025/12/31 是迭代启动日期，2026/1/14 是迭代发布日期
- 迭代I43，2026/1/15 是迭代启动日期，2026/1/28 是迭代发布日期
- 迭代I44，2026/1/30 是迭代启动日期，2026/2/26 是迭代发布日期
...
```

## 实现方式

使用同目录下的 PowerShell 脚本：

```powershell
powershell -ExecutionPolicy Bypass -File Extract-IterationCalendar.ps1 -ExcelFilePath "c:\code\2026年迭代日历V2.xlsx"
```

脚本逻辑：

1. 校验 Excel 文件是否存在。
2. 默认输出到 Excel 同目录的同名 `.md` 文件。
3. 通过本机 Microsoft Excel COM 对象打开工作簿。
4. 读取第一个工作表，从第 9 行开始逐行扫描。
5. 读取 A/B/C 三列，分别作为迭代编号、启动日期、发布日期。
6. 当 A 列为空时停止读取。
7. 生成 UTF-8 编码的 Markdown 文件。

## 命令行用法

```powershell
powershell -ExecutionPolicy Bypass -File Extract-IterationCalendar.ps1 -ExcelFilePath "c:\code\2026年迭代日历V2.xlsx"
```

如需指定输出路径：

```powershell
powershell -ExecutionPolicy Bypass -File Extract-IterationCalendar.ps1 -ExcelFilePath "c:\code\2026年迭代日历V2.xlsx" -OutputPath "c:\code\迭代日历.md"
```

## 注意事项

- 运行环境需要安装 Microsoft Excel。
- 只处理第一个工作表。
- 默认从第 9 行开始读取；如表头位置不同，可修改脚本中的 `$row = 9`。
- 默认期望数据位于 A、B、C 三列。
- A 列为空时停止读取。
- 输出文件使用 UTF-8 编码。
