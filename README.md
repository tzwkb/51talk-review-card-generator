# A1 Review Card Generator

[中文](README_ZH.md) | English


## Overview

 Automated review-card generator for 51Talk Business English lessons, producing PDF, PNG, and HTML deliverables.

## Key Capabilities

- Turns lesson material into after-class review cards.
- Supports multiple output formats for delivery and preview.
- Fits batch production of course assets.

## Usage

 Prepare input files and run the generation commands described in the detailed README below.

## Status

 This repository is maintained or used according to the current README notes.

## Notes

 The original format, parameter, and output-directory notes are kept below.

## Command and Configuration Reference

The following code blocks are preserved from the primary README. Commands, paths, and configuration keys are not translated; adjust them for the actual environment.

```bash
pip install openai openpyxl pymupdf

export REPORT_API_KEY="your-api-key"
```

```
report_generator/
├── main.py              # 入口
├── cli.py               # 交互菜单
├── config.py            # API 配置 + AI Prompts
├── core/                # 核心引擎
│   ├── ai.py            # AI 调用
│   ├── render.py        # HTML 渲染 + 阿拉伯语 RTL
│   ├── export.py        # PDF/PNG 导出
│   ├── record.py        # 记录构建 + 验证
│   ├── pipeline.py      # 单课/单元流程
│   ├── batch.py         # 批量处理
│   └── utils.py         # 通用工具
├── template/
│   ├── lesson_template.html
│   └── unit_template.html
├── test/                # 测试脚本
└── KSA BE 20260324.xlsx # 数据源示例
```

```bash
python main.py
```

```bash
python main.py --excel "KSA BE 20260324.xlsx" --output-dir "./output"
```

```markdown
### Unit 1: Greetings & Introductions

## Lesson 1: Saying Hello

**Goal:** ...

**Vocabulary**
- Hello — A polite greeting.
- ...
```

```bash
cd test

# 快速渲染测试（不调用 AI）
python test_quick_render.py

# 端到端测试（调用 AI）
python test_real_lesson.py
python test_real_unit_excel.py
```

## Detailed Technical Notes

The primary README keeps the original technical details, history notes, full commands, and file layout. This file maintains the English version of the core documentation; consult the primary README code blocks and paths when exact commands are needed.
