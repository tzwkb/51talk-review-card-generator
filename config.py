# ============================================================
# config.py — A1 Review Card Generator — API & Prompt Config
# ============================================================

import os

# --- AI API (OpenAI-compatible) ---
API_KEY    = os.environ.get("REPORT_API_KEY", "")
BASE_URL   = "https://api.vectorengine.ai/v1"
MODEL      = "gemini-3.1-pro-preview"
MAX_TOKENS = 8192
MAX_RETRIES = 3

# --- Default lesson metadata ---
DEFAULT_COURSE = "Business English"

# --- AI extraction prompt ---
SYSTEM_PROMPT = """\
You are generating a compact AFTER-CLASS REVIEW CARD for adult English learners.
The card must be scannable in 10–15 seconds. ONLY the most essential takeaways should appear.

Return ONLY valid JSON. No explanation. No markdown code fences. No extra keys.

Required JSON structure:
{
  "title": "Short lesson title — strip any leading 'Lesson N:' or similar prefix",
  "goal": "One tight sentence starting with a verb describing the communicative outcome.",
  "sections": [
    {
      "heading": "Section heading",
      "html": "Compact HTML snippet for that section"
    }
  ]
}

Section inclusion rules (critical):
1. INCLUDE only these core knowledge types:
   - Vocabulary / Key Words (with meaning + example)
   - Power Words / Useful Phrases (high-impact chunks)
   - Say This (1–3 must-practise sentences, if the lesson has explicit speaking targets)
   - Grammar Points (the core rule + 1 concise example, ignore common-mistake long explanations)
   - Quick Review (2–4 bullets under 8 words each)
   - Tips (1–2 short cultural or usage tips, if genuinely useful)
2. EXCLUDE all of the following completely:
   - Warm Up / Introduction
   - Use It Here / Real-life Situation / Scenario
   - Discussion / Analysis Questions / Thinking Time
   - Logic Quiz / Quick Check / Choose A or B / any quiz or test items
   - Role Play / Your Mission / Level Up Challenge / Action Plan / Expansion Drill / Self-Study tasks
   - Long context-setting paragraphs or redundant explanations
3. If a module doesn't fit the INCLUDE list, drop it. Do not put it in extra_sections or anywhere else.

Style rules for the card:
- ULTRA-CONCISE. Strip every sentence that doesn't add hard value.
- Use bullet points and short phrases over long prose.
- For vocabulary, keep to 4–6 most useful items per lesson (skip ultra-common words like "the", "is").
- For grammar, state the rule in one line and give ONE example.
- For Quick Review, each bullet must be a complete thought but under 8 words.

HTML snippet rules:
- Each section's "html" is the inner content of a <div class="section">. DO NOT wrap it in <div class="section">.
- Start every snippet with: <p class="section-label">Heading Name</p>
- MANDATORY class usage by section type:
  * Unit Vocabulary / Key Words → MUST use <ul class="keyword-grid"><li class="keyword-item">...</li></ul> (2-column cards with word + meaning + example)
  * Say This → MUST use <ul class="say-list"><li class="say-item">...</li></ul> (blue highlight blocks)
  * Grammar Points → MUST use <div class="pattern-block"><p class="pattern-text">...</p></div> (formula box)
  * Quick Review / Tips / Common Mistakes → MUST use <ul class="review-list"><li class="review-chip">...</li></ul> (green rounded chips)
  * Real-life Scenario / Power Words → use <ul> / <li>, <p>, <b> as needed
- Use these CSS classes only when they fit:
  * <ul class="keyword-grid">... — vocabulary (2-column cards: word + meaning + example)
  * <ul class="say-list"><li class="say-item">...</li></ul> — practice sentences (blue highlight)
  * <ul class="review-list"><li class="review-chip">...</li></ul> — quick review / tips (green chips)
  * <div class="pattern-block"><p class="pattern-text">...</p></div> — grammar rule
  * <ul class="unlock-list"><li class="unlock-item">...</li></ul> — achievement summary (purple gradient cards)
  * <ul> / <li>, <p>, <b> — for anything else
- No inline styles. No <script>. No images.
- Clean English only. No blanks (___), no ellipsis (...), no speaker labels.

Return ONLY the JSON object. No surrounding text.
"""

# --- AI unit summarization prompt ---
UNIT_SYSTEM_PROMPT = """\
You are generating a comprehensive UNIT REVIEW CARD for adult English learners.
This card synthesizes the MOST ESSENTIAL takeaways across ALL lessons in one unit.
The output must be rich enough to fill 2–3 pages when rendered as a PDF (720px wide).

Return ONLY valid JSON. No explanation. No markdown code fences. No extra keys.

Required JSON structure:
{
  "title": "Short unit title — strip any leading 'Unit N:' or similar prefix",
  "goal": "One tight sentence starting with a verb describing the unit's overall communicative outcome.",
  "word_count": 10,
  "topic_count": 4,
  "grammar_count": 2,
  "sections": [
    {
      "heading": "Section heading",
      "html": "Compact HTML snippet for that section"
    }
  ]
}

Section inclusion rules (critical):
1. INCLUDE these core knowledge types (synthesized across the whole unit):
   - Unit Vocabulary & Power Words (ONE combined section).
     Include ONLY words from the "Vocabulary" and "Power Words" modules of each lesson.
     Do NOT pull words from Warm Up, Grammar, Role Play, Tips, Summary, Quiz, or other sections.
     Target: approximately 15–20 items per unit (depending on how many lessons the unit has).
     Format: word/phrase + meaning + example.
     Use <ul class="keyword-grid"><li class="keyword-item"> for this section.
   - Say This (4–6 must-practise sentences drawn from the most important speaking targets in the unit)
   - Grammar Points (2–3 consolidated rules across lessons; each: one rule + one concise example)
   - Quick Review (6–8 bullets under 10 words each, covering the whole unit)
   - Common Mistakes (1–2 frequent errors learners make in this unit + correction)
   - Tips (2–3 short cultural or usage tips, if genuinely useful)
   - Real-life Scenario (1 short scenario paragraph showing the unit language in action, 2–3 sentences)
   - Achievement Summary (1 section celebrating learner progress).
     Frame every achievement as a PREMIUM MEMBER milestone.
     Use high-impact, emotional language that makes the learner feel they have unlocked EXCLUSIVE, high-value skills.
     You MAY use numbers, but they must feel BIG (e.g., "15+", "20+", "a wide range of", "dozens of").
     NEVER use small counts like "3 words" or "5 phrases" — if the actual count is low, use broader language.
     Let the content vary naturally: skills unlocked, scenarios handled, confidence gained, professional growth, etc.
     Use <ul class="unlock-list"><li class="unlock-item"> for this section.
   - Next Unit Preview (ONLY if the input contains a "NEXT UNIT PREVIEW" section after the separator).
     Generate 1 teaser section titled "Coming Up Next" that excites the learner about the next unit.
     Highlight 2–3 exciting things they will learn next. Use enthusiastic, high-value language.
2. EXCLUDE all of the following completely:
   - Lesson-by-lesson breakdowns or references to individual lesson numbers
   - Warm Up / Introduction / Scenario texts that do not demonstrate language use
   - Discussion / Analysis Questions / Thinking Time
   - Logic Quiz / Quick Check / Choose A or B / any quiz or test items
   - Role Play / Your Mission / Level Up Challenge / Action Plan / Expansion Drill / Self-Study tasks
   - Long context-setting paragraphs or redundant explanations
3. Deduplicate: if the same word or grammar point appears in multiple lessons, list it ONCE.
4. If a module doesn't fit the INCLUDE list, drop it.

Style rules for the card:
- PACKED WITH VALUE. The card should feel DENSE with knowledge — like a premium cheat-sheet.
- Use bullet points and short phrases over long prose.
- For vocabulary, include EVERY useful item; do not artificially limit the list.
- For grammar, state the rule in one line and give ONE example.
- For Quick Review, each bullet must be a complete thought but under 10 words.
- The total content should fill approximately 2–4 pages when printed at 720px width.

HTML snippet rules:
- Each section's "html" is the inner content of a <div class="section">. DO NOT wrap it in <div class="section">.
- Start every snippet with: <p class="section-label">Heading Name</p>
- Use these CSS classes only when they fit:
  * <ul class="keyword-grid">... — vocabulary (2-column cards: word + meaning + example)
  * <ul class="say-list"><li class="say-item">...</li></ul> — practice sentences (blue highlight)
  * <ul class="review-list"><li class="review-chip">...</li></ul> — quick review / tips (green chips)
  * <div class="pattern-block"><p class="pattern-text">...</p></div> — grammar rule
  * <ul> / <li>, <p>, <b> — for anything else
- No inline styles. No <script>. No images.
- Clean English only. No blanks (___), no ellipsis (...), no speaker labels.

Return ONLY the JSON object. No surrounding text.
"""
