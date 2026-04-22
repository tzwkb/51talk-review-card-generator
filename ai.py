from __future__ import annotations

import asyncio
import json
import re
import sys
from typing import Any

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

from openai import AsyncOpenAI

from config import API_KEY, BASE_URL, MODEL, MAX_TOKENS, MAX_RETRIES, SYSTEM_PROMPT

_AI_CLIENT = AsyncOpenAI(api_key=API_KEY, base_url=BASE_URL)
_PRINT_LOCK = asyncio.Lock()


async def tprint(*args, **kwargs) -> None:
    async with _PRINT_LOCK:
        print(*args, **kwargs)


async def extract_json_from_markdown(markdown_text: str, system_prompt: str | None = None) -> dict[str, Any]:
    prompt = system_prompt if system_prompt is not None else SYSTEM_PROMPT
    last_error: Exception | None = None
    for attempt in range(MAX_RETRIES):
        try:
            response = await _AI_CLIENT.chat.completions.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user",   "content": markdown_text},
                ],
            )
            raw = response.choices[0].message.content or ""
            return _parse_json_response(raw)
        except (ValueError, json.JSONDecodeError) as exc:
            last_error = exc
            if attempt < MAX_RETRIES - 1:
                wait = [5, 15, 30][attempt]
                await tprint(f"  [retry {attempt+1}/{MAX_RETRIES}] JSON parse error: {exc}. Waiting {wait}s...")
                await asyncio.sleep(wait)
        except Exception as exc:
            last_error = exc
            if attempt < MAX_RETRIES - 1:
                wait = [10, 30, 60][attempt]
                await tprint(f"  [retry {attempt+1}/{MAX_RETRIES}] API error: {exc}. Waiting {wait}s...")
                await asyncio.sleep(wait)

    raise last_error  # type: ignore[misc]


def _parse_json_response(raw: str) -> dict[str, Any]:
    cleaned = raw.strip()
    fence = re.search(r"```(?:json)?\s*([\s\S]+?)\s*```", cleaned)
    if fence:
        cleaned = fence.group(1)
    return json.loads(cleaned)


def parse_sheet_selection(raw: str, total: int) -> list[int]:
    if not raw or raw.lower() == "all":
        return list(range(total))
    indices: set[int] = set()
    for part in raw.split(","):
        part = part.strip()
        if "-" in part:
            try:
                start, end = part.split("-", 1)
                for i in range(int(start.strip()) - 1, int(end.strip())):
                    if 0 <= i < total:
                        indices.add(i)
            except ValueError:
                continue
        else:
            try:
                i = int(part) - 1
                if 0 <= i < total:
                    indices.add(i)
            except ValueError:
                continue
    return sorted(indices)
