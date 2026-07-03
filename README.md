# A1 Review Card Generator

English | [中文](README_ZH.md)


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

## Detailed Technical Notes

The primary README keeps the original technical details, history notes, full commands, and file layout. This file maintains the English version of the core documentation; consult the primary README code blocks and paths when exact commands are needed.
