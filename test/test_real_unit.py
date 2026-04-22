import asyncio
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from engine import process_unit, _extract_unit_title

ROOT = Path(__file__).parent.parent
MD_PATH = ROOT / "L1" / "Unit1_Lesson1.md"
TEMPLATE_PATH = ROOT / "template" / "unit_template.html"
LOGO_PATH = ROOT / "51talklogo.png"
OUTPUT_DIR = ROOT / "test_output_real"


def _infer_course(excel_path: Path) -> str:
    name = excel_path.name.upper()
    if "GE" in name or "GENERAL" in name:
        return "General English"
    if "BE" in name or "BUSINESS" in name:
        return "Business English"
    return "Business English"


async def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    markdown = MD_PATH.read_text(encoding="utf-8")
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    unit_title = _extract_unit_title(markdown)
    logo = str(LOGO_PATH) if LOGO_PATH.exists() else ""

    next_preview_path = ROOT / "L1" / "Unit2_Lesson1.md"
    next_preview = next_preview_path.read_text(encoding="utf-8") if next_preview_path.exists() else ""

    result = await process_unit(
        unit_code="L1U1",
        combined_markdown=markdown,
        output_dir=OUTPUT_DIR,
        template=template,
        logo_path=logo,
        level="1",
        unit="1",
        unit_title=unit_title,
        course_name="Business English",
        next_unit_preview_md=next_preview,
    )
    print("Result:", result)


if __name__ == "__main__":
    asyncio.run(main())
