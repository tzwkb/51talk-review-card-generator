import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from engine import process_file

ROOT = Path(__file__).parent.parent
MD_PATH = ROOT / "L1" / "Unit1_Lesson1.md"
TEMPLATE_PATH = ROOT / "template" / "lesson_template.html"
LOGO_PATH = ROOT / "51talklogo.png"
OUTPUT_DIR = ROOT / "test_output_real_lesson"


async def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    logo = str(LOGO_PATH) if LOGO_PATH.exists() else ""

    result = await process_file(
        source_path=MD_PATH,
        output_dir=OUTPUT_DIR,
        template=template,
        logo_path=logo,
        save_json=True,
        course_name="Business English",
    )
    print("Result:", result)


if __name__ == "__main__":
    asyncio.run(main())
