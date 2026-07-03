# A1 Review Card Generator

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)

English | [中文](README_ZH.md)

## Overview

 Automated review-card generator for 51Talk Business English lessons, producing PDF, PNG, and HTML deliverables.

## Key Capabilities

- Turns lesson material into after-class review cards.
- Supports multiple output formats for delivery and preview.
- Fits batch production of course assets.

## Usage

 Prepare input files and run the generation commands described in the detailed README below.

## Notes

 The original format, parameter, and output-directory notes are kept below.

## Command and Configuration Reference

The following code blocks keep commands, paths, filenames, and configuration keys literal; explanatory comments are translated for the English README.

```bash
pip install openai openpyxl pymupdf

export REPORT_API_KEY="your-api-key"
```

```
report_generator/
├── main.py              # entry point
├── cli.py               # interactive menu
├── config.py            # API configuration + AI Prompts
├── core/                # core engine
│   ├── ai.py            # AI calls
│   ├── render.py        # HTML rendering + Arabic RTL
│   ├── export.py        # PDF/PNG export
│   ├── record.py        # record building + validation
│   ├── pipeline.py      # lesson/unit pipeline
│   ├── batch.py         # batch processing
│   └── utils.py         # common utilities
├── template/
│   ├── lesson_template.html
│   └── unit_template.html
├── test/                # test scripts
└── KSA BE 20260324.xlsx # sample data source
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

# quick render test (does not call AI)
python test_quick_render.py

# end-to-end tests (calls AI)
python test_real_lesson.py
python test_real_unit_excel.py
```

## Usage Modes

### Interactive Mode

Interactive mode is the recommended workflow for manually entering or reviewing class data before generating review cards.

### Excel CLI Mode

The command-line Excel mode reads structured class or student data from spreadsheets and generates review-card output without manual UI steps.

## Input and Output

Inputs can be Markdown or Excel depending on the chosen workflow. Outputs are generated as review-card artifacts for classroom or delivery use.

## Environment Variables

Use environment variables only for configuration that should not be hard-coded. Keep secrets out of committed files.

## Tests and Tech Stack

The project is Python-based. Run the local test or smoke-check commands documented in the repository before using generated cards in delivery workflows.
