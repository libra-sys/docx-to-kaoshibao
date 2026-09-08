---
name: docx-to-kaoshibao
description: This skill should be used when the user wants to extract multiple-choice questions from a Word document (especially scanned/OCR documents) and convert them into an Excel question bank that can be imported into 考试宝 (kaoshibao.com). It handles text extraction, answer selection, Chinese translation, explanation writing, and Excel formatting.
agent_created: true
display_name: 题库转换助手（DOCX 转考试宝）
display_name_en: DOCX to Kaoshibao Question Bank
description_zh: 从 Word 文档（含扫描/OCR）中提取选择题，生成可一键导入考试宝的 Excel 题库，含题干校对、答案解析与中文翻译。
description_en: Extract multiple-choice questions from Word documents (including scanned/OCR) into an Excel question bank importable by Kaoshibao.
category: 教育服务
author: libra-eyes
version: 1.0.0
---

# DOCX 英语选择题转考试宝题库

## Overview

将 Word 文档中的英语选择题整理成 考试宝 可一键导入的 Excel 题库。该流程适用于扫描/OCR 生成的 docx 文件以及普通文档：先提取文本，再人工校对题干与选项，确定正确答案并补充翻译与解析，最终生成符合考试宝标准的 Excel 文件。

## When to Use

在以下场景触发此 skill：
- 用户提供了 docx 文件，要求整理其中的选择题。
- 用户要求生成“考试宝可以一键导入的题库”。
- 源文档是扫描件转换成的 docx（OCR 文本可能包含错误、乱码或特殊符号）。
- 用户希望答案、翻译、解析一并包含在题库中。

## Workflow

### Step 1: Extract text from the DOCX

- 使用 Python 读取 `word/document.xml`（docx 本质是 zip 包），提取所有段落文本。
- 同时检查 `word/media/` 下是否包含图片，必要时提取图片以确认手写标注或 OCR 遗漏内容（例如第 19 题中的 “not” 标注图片）。
- 如果文本 OCR 质量差，先列出所有段落供用户确认，不要直接猜测。

### Step 2: Identify and clean multiple-choice questions

- 区分选择题、填空题、翻译题和语法说明文字。用户通常只要求整理“选择题”。
- 将 OCR 错误进行合理修复：
  - 去掉题干中的题号占位符，用 `______` 表示填空位置。
  - 修正错乱的选项字母和分隔符（如 `A. Was, had`）。
  - 若题干或选项明显残缺，标记出来并询问用户，不要强行编造。

### Step 3: Determine correct answers

- 根据英语语法规则逐一判断正确选项。
- 对于常见谚语、固定搭配或情态动词用法，给出简要依据。
- 若存在多个可能答案，选择最符合教材/考试惯例的一项，并在解析中说明。

### Step 4: Write Chinese translation and explanation

- 翻译：给出完整、通顺的中文句意。
- 解析：说明关键语法点（时态、语态、情态动词、固定搭配、主谓一致等）。
- 解析较长时，可在 Excel 中合并为“【句意】...\n【解析】...”格式。

### Step 5: Generate 考试宝 Excel

- 表头使用 考试宝 标准导入格式：
  - `题型` — 统一填 `单选题`
  - `题干` — 清理后的题目
  - `选项A` / `选项B` / `选项C` / `选项D`
  - `正确答案` — 大写字母 A/B/C/D
  - `解析` — 翻译 + 解析
  - `难度` — 默认 `中等`
- 使用 openpyxl 生成 `.xlsx`。
- 设置列宽、自动换行、表头样式，便于用户直接预览和导入。
- 同时生成一份 Markdown 版答案解析供用户核对。

## Output Files

1. `英语练习题库_考试宝导入.xlsx` — 考试宝一键导入文件。
2. `英语练习题库_答案解析.md` — 包含全部题目、答案、翻译和解析的文本版本，便于人工核对。

## 考试宝导入说明（可写入 Markdown 给用户）

- 打开 Excel 文件。
- 列头已按考试宝标准格式设置：题型、题干、选项A、选项B、选项C、选项D、正确答案、解析、难度。
- 在考试宝中进入 题库 → 导入题库 → 选择 Excel 导入，按提示上传即可。
- 若平台只识别固定列名，保留“题型/题干/选项A~D/正确答案/解析”即可。

## Common Pitfalls

- OCR 文档中选项可能被截断或合并，必须人工核对。
- 扫描件中的手写标注（如图片里的 “not”）会影响答案，需要单独提取图片确认。
- 有些题目从第 9 题开始编号，说明前面题目可能缺失，需在最终输出中说明。
- 考试宝对 Excel 表头敏感，确保使用“题型”而不是“题目类型”等变体。
