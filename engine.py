from __future__ import annotations

import asyncio
import base64
import json
import os
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from config import DEFAULT_COURSE
from ai import tprint, extract_json_from_markdown, parse_sheet_selection

CONCURRENCY_LIMIT = 12

TEMPLATE_DIR         = Path(__file__).resolve().parent / "template"
LESSON_TEMPLATE_PATH = TEMPLATE_DIR / "lesson_template.html"
UNIT_TEMPLATE_PATH   = TEMPLATE_DIR / "unit_template.html"
LOGO_PATH = Path(__file__).resolve().parent / "51talklogo.png"
INPUT_FILE_RE = re.compile(r"^Unit(\d+)_Lesson(\d+)\.md$", re.I)


def logo_data_uri() -> str:
    if not LOGO_PATH.exists():
        return ""
    mime = "image/png"
    encoded = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def infer_course(source: Path | str | None) -> str:
    if not source:
        return DEFAULT_COURSE
    name = Path(source).name.upper()
    if "GE" in name or "GENERAL" in name:
        return "General English"
    if "BE" in name or "BUSINESS" in name:
        return "Business English"
    return DEFAULT_COURSE


def existing_lessons(output_dir: Path) -> set[str]:
    existing = set()
    if not output_dir.exists():
        return existing
    for level_dir in output_dir.iterdir():
        if not level_dir.is_dir() or not level_dir.name.startswith("L"):
            continue
        for unit_dir in level_dir.iterdir():
            if not unit_dir.is_dir() or not unit_dir.name.startswith("U"):
                continue
            for pdf in unit_dir.glob("*.pdf"):
                existing.add(pdf.stem)
    return existing


import html


_ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]")

_CARD_CLASSES = (
    "say-item", "keyword-item", "pattern-block", "review-chip",
    "unlock-item", "lesson-block", "unit-word-item", "master-syntax",
)


def _esc(value: str) -> str:
    return html.escape(value or "", quote=True)


def _esc_br(value: str) -> str:
    return html.escape(value or "", quote=True).replace("&lt;br&gt;", "<br>").replace("&lt;br/&gt;", "<br/>")


def rtl_arabic_in_html(html_text: str, compact: bool = False) -> str:
    """Wrap Arabic text in RTL block spans.
    If compact=False (non-card paragraphs like Tips), insert blank lines
    between bilingual pairs. If compact=True (cards), keep tight."""
    if not html_text or not _ARABIC_RE.search(html_text):
        return html_text
    parts = re.split(r"(<br\s*/?>)", html_text)

    wrapped = []
    for part in parts:
        if part.lower().startswith("<br"):
            wrapped.append(part)
            continue
        if not _ARABIC_RE.search(part):
            wrapped.append(part)
            continue
        segments = re.split(r"(<[^>]+>)", part)
        processed = []
        for seg in segments:
            if seg.startswith("<") and seg.endswith(">"):
                processed.append(seg)
            elif _ARABIC_RE.search(seg):
                processed.append(
                    f'<span dir="rtl" style="display:block;text-align:right;">{seg}</span>'
                )
            else:
                processed.append(seg)
        wrapped.append("".join(processed))

    if compact:
        return "".join(wrapped)

    result = []
    for i, part in enumerate(wrapped):
        result.append(part)
        if part.lower().startswith("<br"):
            prev_idx, next_idx = i - 1, i + 1
            if prev_idx >= 0 and next_idx < len(wrapped):
                prev, nxt = wrapped[prev_idx], wrapped[next_idx]
                if 'dir="rtl"' in prev and nxt and not nxt.lower().startswith("<br"):
                    nxt_text = re.sub(r"<[^>]+>", "", nxt).strip()
                    if nxt_text and not _ARABIC_RE.search(nxt_text):
                        result.append("<br>")
    return "".join(result)


def render_html(record: dict[str, Any], template: str) -> str:
    """Fill template with record data."""
    content_blocks = ""
    for sec in record.get("sections") or []:
        if isinstance(sec, dict) and sec.get("html"):
            html = sec["html"]
            has_card = any(cls in html for cls in _CARD_CLASSES)
            content_blocks += f'<div class="section">{rtl_arabic_in_html(html, compact=has_card)}</div>\n'

    return template.format(
        lesson_code=_esc(record.get("lesson_code", "")),
        meta_line=_esc(record.get("meta_line", "")),
        title=_esc(record.get("title", "")),
        goal=rtl_arabic_in_html(_esc_br(record.get("goal", ""))),
        content_blocks=content_blocks,
        logo_uri=logo_data_uri(),
        level_meta=_esc(record.get("level_meta", record.get("meta_line", ""))),
        goal_statement=rtl_arabic_in_html(_esc_br(record.get("goal_statement", record.get("goal", "")))),
        unit_code=_esc(record.get("unit_code", record.get("lesson_code", ""))),
        unit_number=_esc(str(record.get("unit_number", record.get("unit", "")))),
        unit_title=_esc(record.get("unit_title", record.get("title", ""))),
        word_count=_esc(str(record.get("word_count", ""))),
        topic_count=_esc(str(record.get("topic_count", ""))),
        grammar_count=_esc(str(record.get("grammar_count", ""))),
    )


import shutil
import subprocess

import fitz


def _find_browser() -> str:
    for candidate in [
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    ]:
        if candidate.exists():
            return str(candidate)
    for name in ("msedge.exe", "chrome.exe"):
        found = shutil.which(name)
        if found:
            return found
    return ""


def _flatten_pdf(pdf_path: Path) -> bool:
    try:
        doc = fitz.open(str(pdf_path))
        new_doc = fitz.open()
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            new_page = new_doc.new_page(width=pix.width, height=pix.height)
            new_page.insert_image(new_page.rect, pixmap=pix)
        doc.close()
        new_doc.save(str(pdf_path), deflate=True)
        new_doc.close()
        return True
    except Exception as exc:
        print(f"  [warn] PDF flatten failed: {exc}")
        return False


def generate_pdf(html_path: Path, pdf_path: Path) -> bool:
    browser = _find_browser()
    if not browser:
        print(f"  [warn] No browser found for PDF export — skipping {pdf_path.name}")
        return False
    try:
        subprocess.run(
            [
                browser,
                "--headless",
                "--disable-gpu",
                "--no-pdf-header-footer",
                "--run-all-compositor-stages-before-draw",
                "--virtual-time-budget=5000",
                f"--print-to-pdf={pdf_path.resolve()}",
                html_path.resolve().as_uri(),
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        _flatten_pdf(pdf_path)
        return True
    except Exception as exc:
        print(f"  [warn] PDF export failed: {exc}")
        return False


def generate_png(pdf_path: Path, png_path: Path) -> bool:
    try:
        doc = fitz.open(str(pdf_path))
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        base = png_path.with_suffix("")
        for i in range(len(doc)):
            page = doc[i]
            pix = page.get_pixmap(matrix=mat)
            page_png = Path(f"{base}_p{i+1}.png")
            pix.save(str(page_png))
        doc.close()
        # remove old single-page PNG if exists
        if png_path.exists():
            png_path.unlink()
        return True
    except Exception as exc:
        print(f"  [warn] PNG export failed: {exc}")
        return False





def _lesson_code_from_path(source_path: Path) -> tuple[str, str, str]:
    m = INPUT_FILE_RE.match(source_path.name)
    if m:
        return m.group(1), m.group(2), f"U{m.group(1)}L{m.group(2)}"
    return "?", "?", source_path.stem


def build_record(
    ai_data: dict[str, Any],
    lesson_code: str,
    level: str,
    unit: str,
    lesson: str,
    source_file: str,
    logo_path: str,
    course_name: str = DEFAULT_COURSE,
) -> dict[str, Any]:
    sections = ai_data.get("sections") or []
    if not isinstance(sections, list):
        sections = []

    cleaned_sections = []
    for sec in sections:
        if isinstance(sec, dict) and sec.get("html"):
            cleaned_sections.append({
                "heading": str(sec.get("heading", "")).strip(),
                "html": str(sec.get("html", "")).strip(),
            })

    return {
        "lesson_code": lesson_code,
        "level":       level,
        "unit":        unit,
        "lesson":      lesson,
        "meta_line":   " | ".join(filter(None, [course_name, f"Level {level} (المستوى {level})", f"Unit {unit} (الوحدة {unit})", f"Lesson {lesson} (الدرس {lesson})"])),
        "title":       (ai_data.get("title") or "").strip(),
        "goal":        (ai_data.get("goal") or "").strip(),
        "sections":    cleaned_sections,
        "source_file": source_file,
        "logo_path":   logo_path,
    }


def extract_unit_title(markdown_text: str) -> str:
    m = re.search(r"###\s*Unit\s*\d+[:\-]\s*(.+)", markdown_text)
    if m:
        title = m.group(1).strip()
        title = re.sub(r"[\U0001f300-\U0001f9ff]", "", title).strip()
        return title
    return ""


def _count_grammar_sections(sections: list[dict[str, Any]]) -> int:
    keywords = {"grammar", "pattern", "syntax", "structure"}
    count = 0
    for sec in sections:
        heading = str(sec.get("heading", "")).lower()
        if any(k in heading for k in keywords):
            count += 1
    return count


def build_unit_record(
    ai_data: dict[str, Any],
    unit_code: str,
    level: str,
    unit: str,
    source_files: list[str],
    logo_path: str,
    unit_title: str = "",
    course_name: str = DEFAULT_COURSE,
) -> dict[str, Any]:
    sections = ai_data.get("sections") or []
    if not isinstance(sections, list):
        sections = []
    cleaned_sections = []
    for sec in sections:
        if isinstance(sec, dict) and sec.get("html"):
            cleaned_sections.append({
                "heading": str(sec.get("heading", "")).strip(),
                "html": str(sec.get("html", "")).strip(),
            })
    title = unit_title.strip() if unit_title.strip() else (ai_data.get("title") or "").strip()
    return {
        "lesson_code": unit_code,
        "level":       level,
        "unit":        unit,
        "lesson":      "Summary",
        "meta_line":   " | ".join(filter(None, [course_name, f"Level {level} (المستوى {level})", f"Unit {unit} (الوحدة {unit})", "Unit Summary (ملخص الوحدة)"])),
        "title":       title,
        "goal":        (ai_data.get("goal") or "").strip(),
        "sections":    cleaned_sections,
        "source_files": source_files,
        "logo_path":   logo_path,
        "word_count":  str(ai_data.get("word_count", "")),
        "topic_count": str(ai_data.get("topic_count", "")),
        "grammar_count": str(ai_data.get("grammar_count", "")),
    }


BANNED_TOKENS = [
    "Teacher:", "Student:", "_____", "Option A", "Option B",
    "Use the correct sentence pattern", "Use this lesson pattern",
]


def validate(record: dict[str, Any], html_text: str) -> list[str]:
    issues = []
    if not record.get("title"):
        issues.append("missing_title")
    if not record.get("goal"):
        issues.append("missing_goal")
    for token in BANNED_TOKENS:
        if token in html_text:
            issues.append(f"banned_token:{token[:20]}")
    return issues


import json

from config import DEFAULT_COURSE, SYSTEM_PROMPT, UNIT_SYSTEM_PROMPT
from ai import extract_json_from_markdown


async def process_file(
    source_path: Path,
    output_dir: Path,
    template: str,
    logo_path: str,
    save_json: bool,
    course_name: str = DEFAULT_COURSE,
) -> dict[str, Any]:
    markdown_text = source_path.read_text(encoding="utf-8")

    level = "?"
    for part in reversed(source_path.resolve().parts):
        m = re.match(r"^L(\d+)$", part, re.I)
        if m:
            level = m.group(1)
            break

    prompt = SYSTEM_PROMPT
    md = markdown_text
    if level in {"0", "1", "2", "3"}:
        md = (
            "IMPORTANT: This lesson is for beginner learners (Level " + level + "). "
            "Output ALL section content in BOTH English and Arabic. "
            "English text first, Arabic translation on a new line using <br>. "
            "Do not skip the Arabic text.\n\n"
            + markdown_text
        )

    await tprint(f"  [AI]  Extracting {source_path.name}...")
    try:
        ai_data = await extract_json_from_markdown(md, prompt)
    except Exception as exc:
        await tprint(f"  [ERROR] AI extraction failed for {source_path.name}: {exc}")
        return {"lesson_code": source_path.stem, "html": "", "pdf": "", "png": "", "json": "", "issues": [f"ai_error:{exc}"]}

    unit, lesson, code_suffix = _lesson_code_from_path(source_path)
    ai_level = str(ai_data.get("level", "?"))
    if ai_level != "?":
        level = ai_level
    lesson_code = f"L{level}{code_suffix}" if level != "?" else code_suffix
    record = build_record(ai_data, lesson_code, level, unit, lesson, str(source_path), logo_path, course_name)

    html_path = output_dir / f"{lesson_code}.html"
    pdf_path  = output_dir / f"{lesson_code}.pdf"
    png_path  = output_dir / f"{lesson_code}.png"
    json_path = output_dir / f"{lesson_code}.json"

    html_text = render_html(record, template)
    html_path.write_text(html_text, encoding="utf-8")

    await asyncio.to_thread(generate_pdf, html_path, pdf_path)
    await asyncio.to_thread(generate_png, pdf_path, png_path)

    if save_json:
        json_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")

    issues = validate(record, html_text)
    return {
        "lesson_code": lesson_code,
        "html": str(html_path),
        "pdf":  str(pdf_path),
        "png":  str(png_path),
        "json": str(json_path) if save_json else "",
        "issues": issues,
    }


async def process_unit(
    unit_code: str,
    combined_markdown: str,
    output_dir: Path,
    template: str,
    logo_path: str,
    level: str,
    unit: str,
    unit_title: str = "",
    course_name: str = DEFAULT_COURSE,
    next_unit_preview_md: str = "",
) -> dict[str, Any]:
    await tprint(f"  [AI]  Summarizing {unit_code}...")
    prompt = UNIT_SYSTEM_PROMPT
    markdown = combined_markdown
    if level in {"0", "1", "2", "3"}:
        markdown = (
            "IMPORTANT: This unit is for beginner learners (Level " + level + "). "
            "Output ALL section content in BOTH English and Arabic. "
            "English text first, Arabic translation on a new line using <br>. "
            "Do not skip the Arabic text.\n\n"
            + combined_markdown
        )
    if next_unit_preview_md:
        markdown += (
            "\n\n---\n\nNEXT UNIT PREVIEW (Unit " + str(int(unit) + 1) + " Lesson 1):\n\n"
            + next_unit_preview_md
        )
    try:
        ai_data = await extract_json_from_markdown(markdown, prompt)
    except Exception as exc:
        await tprint(f"  [ERROR] AI summarization failed for {unit_code}: {exc}")
        return {"lesson_code": unit_code, "html": "", "pdf": "", "png": "", "json": "", "issues": [f"ai_error:{exc}"]}

    record = build_unit_record(ai_data, unit_code, level, unit, [], logo_path, unit_title, course_name)

    unit_dir = output_dir / f"L{level}" / f"U{unit}"
    unit_dir.mkdir(parents=True, exist_ok=True)

    html_path = unit_dir / f"{unit_code}_summary.html"
    pdf_path  = unit_dir / f"{unit_code}_summary.pdf"
    png_path  = unit_dir / f"{unit_code}_summary.png"
    json_path = unit_dir / f"{unit_code}_summary.json"

    html_text = render_html(record, template)
    html_path.write_text(html_text, encoding="utf-8")

    await asyncio.to_thread(generate_pdf, html_path, pdf_path)
    await asyncio.to_thread(generate_png, pdf_path, png_path)

    json_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")

    issues = validate(record, html_text)
    await tprint(f"  [done] {unit_code}  {'⚠ ' + ', '.join(issues) if issues else '✓'}")
    return {
        "lesson_code": unit_code,
        "html": str(html_path),
        "pdf":  str(pdf_path),
        "png":  str(png_path),
        "json": str(json_path),
        "issues": issues,
    }





def write_report(output_dir: Path, reports: list[dict[str, Any]]) -> None:
    payload = {
        "status": "pass" if all(not r["issues"] for r in reports) else "warn",
        "lesson_count": len(reports),
        "issue_count":  sum(len(r["issues"]) for r in reports),
        "reports": reports,
    }
    (output_dir / "_batch_report.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    lines = [
        "# Batch Report", "",
        f"- Status: {payload['status']}",
        f"- Lessons: {payload['lesson_count']}",
        f"- Issues:  {payload['issue_count']}", "",
    ]
    for r in reports:
        tag = ", ".join(r["issues"]) if r["issues"] else "ok"
        lines.append(f"- `{r['lesson_code']}`: {tag}")
    (output_dir / "_batch_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


async def run_batch(
    sources: list[Path],
    output_dir: Path,
    logo_path: str,
    save_json: bool,
    course_name: str,
) -> list[dict[str, Any]]:
    template = LESSON_TEMPLATE_PATH.read_text(encoding="utf-8")
    output_dir.mkdir(parents=True, exist_ok=True)
    total = len(sources)

    if total == 1:
        print(f"  [1/1] {sources[0].name}")
        return [await process_file(sources[0], output_dir, template, logo_path, save_json, course_name)]

    print(f"  Processing {total} files with asyncio (concurrency limit {CONCURRENCY_LIMIT})...")
    sem = asyncio.Semaphore(CONCURRENCY_LIMIT)
    done = 0

    async def _process_one(src: Path) -> tuple[Path, dict[str, Any]]:
        nonlocal done
        async with sem:
            try:
                result = await process_file(src, output_dir, template, logo_path, save_json, course_name)
            except Exception as exc:
                await tprint(f"  [ERROR] {src.name}: {exc}")
                result = {"lesson_code": src.stem, "html": "", "pdf": "", "png": "", "json": "", "issues": [f"error:{exc}"]}
            done += 1
            await tprint(f"  [{done}/{total}] completed")
            return src, result

    tasks = [asyncio.create_task(_process_one(src)) for src in sources]
    results: dict[Path, dict[str, Any]] = {}
    for coro in asyncio.as_completed(tasks):
        src, result = await coro
        results[src] = result

    return [results[src] for src in sources]


async def process_excel_cell(
    markdown_text: str,
    level: str,
    unit: str,
    lesson: str,
    output_dir: Path,
    template: str,
    logo_path: str,
    save_json: bool,
    course_name: str,
) -> dict[str, Any]:
    lesson_code = f"L{level}U{unit}L{lesson}"

    await tprint(f"  [AI]  Extracting {lesson_code}...")
    try:
        prompt = SYSTEM_PROMPT
        md = markdown_text
        if level in {"0", "1", "2", "3"}:
            md = (
                "IMPORTANT: This lesson is for beginner learners (Level " + level + "). "
                "Output ALL section content in BOTH English and Arabic. "
                "English text first, Arabic translation on a new line using <br>. "
                "Do not skip the Arabic text.\n\n"
                + markdown_text
            )
        ai_data = await extract_json_from_markdown(md, prompt)
    except Exception as exc:
        await tprint(f"  [ERROR] AI extraction failed for {lesson_code}: {exc}")
        return {"lesson_code": lesson_code, "html": "", "pdf": "", "png": "", "json": "", "issues": [f"ai_error:{exc}"]}

    ai_data["level"] = level
    record = build_record(ai_data, lesson_code, level, unit, lesson, "excel", logo_path, course_name)

    lesson_dir = output_dir / f"L{level}" / f"U{unit}"
    lesson_dir.mkdir(parents=True, exist_ok=True)

    html_path = lesson_dir / f"{lesson_code}.html"
    pdf_path  = lesson_dir / f"{lesson_code}.pdf"
    png_path  = lesson_dir / f"{lesson_code}.png"
    json_path = lesson_dir / f"{lesson_code}.json"

    html_text = render_html(record, template)
    html_path.write_text(html_text, encoding="utf-8")

    await asyncio.to_thread(generate_pdf, html_path, pdf_path)
    await asyncio.to_thread(generate_png, pdf_path, png_path)

    if save_json:
        json_path.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")

    issues = validate(record, html_text)
    await tprint(f"  [done] {lesson_code}  {'⚠ ' + ', '.join(issues) if issues else '✓'}")
    return {
        "lesson_code": lesson_code,
        "html": str(html_path),
        "pdf":  str(pdf_path),
        "png":  str(png_path),
        "json": str(json_path) if save_json else "",
        "issues": issues,
    }
