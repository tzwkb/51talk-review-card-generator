import sys
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from engine import render_html, generate_pdf

ROOT = Path(__file__).parent.parent
TEMPLATE_PATH = ROOT / "template" / "unit_template.html"
OUTPUT_DIR = ROOT / "test_output_real_excel" / "L1" / "U1"
JSON_PATH = OUTPUT_DIR / "L1U1_summary.json"


def main():
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    record = json.loads(JSON_PATH.read_text(encoding="utf-8"))

    html_text = render_html(record, template)
    html_path = OUTPUT_DIR / "L1U1_summary_test.html"
    pdf_path = OUTPUT_DIR / "L1U1_summary_test.pdf"

    html_path.write_text(html_text, encoding="utf-8")
    print(f"HTML written: {html_path}")

    ok = generate_pdf(html_path, pdf_path)
    print(f"PDF generated: {pdf_path} ({'OK' if ok else 'FAIL'})")


if __name__ == "__main__":
    main()
