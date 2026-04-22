import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from engine import generate_pdf, generate_png

HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: #f5f3ff; padding: 32px 24px; }
.unit-card { width: 720px; margin: 0 auto; background: #ffffff; border-radius: 16px; overflow: hidden; border-top: 6px solid #6d28d9; }
.section { padding: 24px 32px; position: relative; }
.section-title { font-size: 18px; font-weight: 800; color: #4c1d95; margin-bottom: 16px; }
.unlock-list { display: flex; flex-wrap: wrap; gap: 12px; }
.unlock-item { display: flex; align-items: center; gap: 8px; background: #f5f3ff; border: 1px solid #ddd6fe; border-radius: 10px; padding: 10px 14px; font-size: 13px; color: #5b21b6; font-weight: 600; }
</style>
</head>
<body>
<div class="unit-card">
  <div class="section">
    <div class="section-title">Achievement Summary</div>
    <div class="unlock-list">
      <div class="unlock-item">&#127942; 15 Words Learned</div>
      <div class="unlock-item">&#128161; 3 Skills Unlocked</div>
      <div class="unlock-item">&#128640; 2 Scenarios Handled</div>
      <div class="unlock-item">&#10024; 2 Grammar Mastered</div>
    </div>
  </div>
</div>
</body>
</html>
"""


async def main():
    out = Path("test_output")
    out.mkdir(exist_ok=True)
    html_path = out / "test_unlock.html"
    pdf_path = out / "test_unlock.pdf"
    png_path = out / "test_unlock.png"
    html_path.write_text(HTML, encoding="utf-8")
    ok_pdf = await asyncio.to_thread(generate_pdf, html_path.resolve(), pdf_path.resolve())
    ok_png = await asyncio.to_thread(generate_png, pdf_path.resolve(), png_path.resolve())
    print(f"PDF: {'OK' if ok_pdf else 'FAIL'} -> {pdf_path}")
    print(f"PNG: {'OK' if ok_png else 'FAIL'} -> {png_path}")


if __name__ == "__main__":
    asyncio.run(main())
