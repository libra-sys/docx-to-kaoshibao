# 题库转换助手（DOCX 转考试宝）

> WorkBuddy Skill · 作者 libra-eyes · MIT 协议

## 这是什么

从 Word 文档（含扫描/OCR）中提取选择题，生成可一键导入考试宝的 Excel 题库，含题干校对、答案解析与中文翻译。

> Extract multiple-choice questions from Word documents (including scanned/OCR) into an Excel question bank importable by Kaoshibao.

- **服务类目**：教育服务
- **标识**：`docx-to-kaoshibao`

## 触发词

- 见 SKILL.md 内触发说明

## 安装

把 `docx-to-kaoshibao/` 整个目录复制到 WorkBuddy 用户级技能目录：

- Windows：`C:\Users\<你的用户名>\.workbuddy\skills\`
- macOS / Linux：`~/.workbuddy/skills/`

然后在 WorkBuddy 中重新加载技能即可使用。

也可以打包上传到 [WorkBuddy 开放平台](https://open.workbuddy.cn) 的技能市场。

## 目录结构

```
docx-to-kaoshibao/
├── SKILL.md
├── scripts/build_kaoshibao_bank.py
```

## 许可

MIT License · Copyright (c) 2026 libra-eyes
