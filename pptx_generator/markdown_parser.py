"""Markdown → slides data parser.

Converts a simple Markdown file into the same structure as slides.json,
so users don't need to hand-craft JSON.

Supported Markdown conventions:
    - YAML front matter (``---`` fenced) → presentation_metadata
    - ``# H1`` → title_slide (first) or section_slide (subsequent)
    - ``## H2`` → new slide (bullet_points / code_demo depending on content)
    - Bullet lists (``-`` / ``*`` / ``1.``) → points[]
    - Fenced code blocks (``` ``` ```) → code_demo
    - ``> blockquote`` first line → speaker notes
    - ``---`` horizontal rule between sections → section_slide marker
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any


def _parse_front_matter(text: str) -> tuple[dict[str, str], str]:
    """Extract YAML front matter if present. Returns (metadata, remaining_text)."""
    if not text.startswith("---"):
        return {}, text
    end = text.find("---", 3)
    if end == -1:
        return {}, text
    fm_block = text[3:end].strip()
    remaining = text[end + 3 :].strip()

    meta: dict[str, str] = {}
    for line in fm_block.splitlines():
        if ":" in line:
            key, _, val = line.partition(":")
            meta[key.strip()] = val.strip().strip("\"'")
    return meta, remaining


def _detect_code_block(lines: list[str], start: int) -> tuple[str, str, int]:
    """Detect a fenced code block starting at *start*. Returns (code, language, end_index)."""
    opener = lines[start]
    lang = opener.strip().lstrip("`").strip()
    code_lines: list[str] = []
    i = start + 1
    while i < len(lines):
        if lines[i].strip().startswith("```"):
            return "\n".join(code_lines), lang, i
        code_lines.append(lines[i])
        i += 1
    # Unclosed block — treat rest as code
    return "\n".join(code_lines), lang, i - 1


def parse_markdown(text: str) -> dict[str, Any]:
    """Parse Markdown text into slides.json-compatible dict.

    Returns:
        ``{"presentation_metadata": {...}, "slides": [...]}``
    """
    meta, body = _parse_front_matter(text)
    presentation_metadata = {
        "title": meta.get("title", "Untitled"),
        "version": meta.get("version", meta.get("date", "")),
    }

    slides: list[dict[str, Any]] = []
    current_slide: dict[str, Any] | None = None
    seen_h1 = False
    lines = body.splitlines()
    i = 0

    def _flush() -> None:
        nonlocal current_slide
        if current_slide is not None:
            slides.append(current_slide)
            current_slide = None

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # --- Horizontal rule → section marker ---
        if re.match(r"^-{3,}$", stripped) or re.match(r"^\*{3,}$", stripped):
            _flush()
            i += 1
            continue

        # --- H1 heading ---
        if stripped.startswith("# ") and not stripped.startswith("## "):
            _flush()
            title = stripped.lstrip("# ").strip()
            if not seen_h1:
                seen_h1 = True
                # Collect subtitle from next non-empty lines
                sub_lines: list[str] = []
                j = i + 1
                while (
                    j < len(lines)
                    and lines[j].strip()
                    and not lines[j].strip().startswith("#")
                ):
                    sub_lines.append(lines[j].strip())
                    j += 1
                current_slide = {
                    "type": "title_slide",
                    "content": {
                        "title": title,
                        "sub_title": "\n".join(sub_lines) if sub_lines else "",
                    },
                }
                i = j
                _flush()
                continue
            else:
                current_slide = {
                    "type": "section_slide",
                    "content": {"title": title, "sub_title": ""},
                }
                i += 1
                _flush()
                continue

        # --- H2 heading → new slide ---
        if stripped.startswith("## ") and not stripped.startswith("### "):
            _flush()
            title = stripped.lstrip("# ").strip()
            current_slide = {
                "type": "bullet_points",
                "content": {"title": title, "points": []},
            }
            i += 1
            continue

        # --- Fenced code block ---
        if stripped.startswith("```"):
            code, lang, end_i = _detect_code_block(lines, i)
            if current_slide is None:
                current_slide = {
                    "type": "code_demo",
                    "content": {
                        "title": "Code",
                        "code": code,
                        "language": lang or "text",
                    },
                }
                _flush()
            elif (
                current_slide["type"] == "bullet_points"
                and not current_slide["content"]["points"]
            ):
                # Convert to code_demo if no bullets yet
                current_slide["type"] = "code_demo"
                current_slide["content"] = {
                    "title": current_slide["content"]["title"],
                    "code": code,
                    "language": lang or "text",
                }
                _flush()
            else:
                # Flush current, create code slide
                _flush()
                current_slide = {
                    "type": "code_demo",
                    "content": {
                        "title": "Code",
                        "code": code,
                        "language": lang or "text",
                    },
                }
                _flush()
            i = end_i + 1
            continue

        # --- Blockquote → speaker notes ---
        if stripped.startswith("> ") and current_slide is not None:
            note = stripped.lstrip("> ").strip()
            current_slide["content"]["notes"] = note
            i += 1
            continue

        # --- Bullet / numbered list ---
        if re.match(r"^(\s*)([-*•]|\d+[.)])\s+", line) and current_slide is not None:
            if current_slide["type"] == "bullet_points":
                # Preserve indentation for sub-bullets
                current_slide["content"]["points"].append(line.rstrip())
            i += 1
            continue

        # --- Plain text under a slide → append as bullet ---
        if (
            stripped
            and current_slide is not None
            and current_slide["type"] == "bullet_points"
        ):
            current_slide["content"]["points"].append(stripped)
            i += 1
            continue

        i += 1

    _flush()

    # Assign IDs and add outline slide if enough slides
    for idx, s in enumerate(slides, start=1):
        s["id"] = idx

    # If title is from front matter and first slide is title_slide, update it
    if slides and slides[0]["type"] == "title_slide" and meta.get("title"):
        presentation_metadata["title"] = slides[0]["content"]["title"]

    return {"presentation_metadata": presentation_metadata, "slides": slides}


def parse_markdown_file(path: Path) -> dict[str, Any]:
    """Read a Markdown file and parse it into slides data."""
    text = path.read_text(encoding="utf-8")
    return parse_markdown(text)
