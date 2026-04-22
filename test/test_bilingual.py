import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


def inject_bilingual(level: str, combined_markdown: str) -> str:
    if level in {"1", "2", "3"}:
        return (
            "IMPORTANT: This unit is for beginner learners (Level " + level + "). "
            "Output ALL section content in BOTH English and Arabic. "
            "English text first, Arabic translation on a new line using <br>. "
            "Do not skip the Arabic text.\n\n"
            + combined_markdown
        )
    return combined_markdown


def test():
    base = "Original markdown text"
    for lvl in ("1", "2", "3"):
        out = inject_bilingual(lvl, base)
        assert "IMPORTANT" in out, f"Level {lvl}: missing IMPORTANT"
        assert f"Level {lvl}" in out, f"Level {lvl}: missing level number"
        assert "Arabic" in out, f"Level {lvl}: missing Arabic keyword"
        assert base in out, f"Level {lvl}: missing original text"
    for lvl in ("4", "5", "?", ""):
        out = inject_bilingual(lvl, base)
        assert out == base, f"Level {lvl}: should not mutate, got: {out[:60]}"
    print("Bilingual logic: OK")


if __name__ == "__main__":
    test()
