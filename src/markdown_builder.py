"""Build AI-optimized Obsidian Markdown with frontmatter and per-slide embeds."""

import re


def split_markdown_by_slides(full_md: str, slide_count: int) -> list[str]:
    """
    Split markitdown output into one chunk per slide.
    Uses ## or # at line-start as section boundaries; if we get fewer chunks than
    slide_count, we pad with empty strings or merge the last chunk.
    """
    if not full_md or slide_count <= 0:
        return [""] * max(1, slide_count) if slide_count else []

    stripped = full_md.strip()
    if not stripped:
        return [""] * slide_count

    # Split by lines that start with ## or # (at start or after newline)
    parts = re.split(r"\n(?=#{1,6}\s)", stripped)
    # If first line isn't a header, the first "part" is the intro; keep it
    if parts and not re.match(r"^#{1,6}\s", parts[0].strip()):
        # First part is body before first header
        if len(parts) <= slide_count:
            # One slide gets intro + first header content; rest get subsequent
            result: list[str] = []
            for i in range(slide_count):
                if i == 0:
                    result.append(parts[0] if parts else "")
                elif i < len(parts):
                    result.append(parts[i].strip())
                else:
                    result.append("")
            return result
        else:
            parts = [parts[0]] + parts[1:]  # no change to split logic
    else:
        # All parts start with a header
        pass

    # Now we have parts that are header-led sections
    if len(parts) >= slide_count:
        return [p.strip() for p in parts[: slide_count]]
    # Too few parts: assign to slides, last slide gets remainder
    result = [p.strip() for p in parts]
    while len(result) < slide_count:
        result.append("")
    return result


def build_markdown(
    *,
    pdf_basename: str,
    slide_count: int,
    speaker_notes: list[str],
    body_sections: list[str],
    title: str | None = None,
) -> str:
    """
    Build the Obsidian-ready Markdown:
    - YAML frontmatter with speaker_notes (and optional title)
    - For each slide: ![[Lecture.pdf#page=N]] then that slide's text
    """
    pdf_ref = f"{pdf_basename}.pdf"
    lines: list[str] = []

    # Frontmatter: speaker notes as list
    notes_for_yaml = [n for n in speaker_notes if n]
    frontmatter: dict[str, list[str] | str] = {}
    if title:
        frontmatter["title"] = title
    if notes_for_yaml:
        frontmatter["speaker_notes"] = notes_for_yaml
    if frontmatter:
        lines.append("---")
        for k, v in frontmatter.items():
            if k == "speaker_notes" and isinstance(v, list):
                lines.append("speaker_notes:")
                for n in v:
                    # Escape for YAML: backslash and quotes
                    n_esc = n.replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ")
                    lines.append(f'  - "{n_esc}"')
            elif isinstance(v, str):
                v_esc = v.replace("\\", "\\\\").replace('"', '\\"')
                lines.append(f'{k}: "{v_esc}"')
            else:
                lines.append(f"{k}: {v}")
        lines.append("---")
        lines.append("")

    # Ensure we have one section per slide
    sections = body_sections if len(body_sections) >= slide_count else body_sections + [""] * (slide_count - len(body_sections))
    sections = sections[: slide_count]

    for n, section_text in enumerate(sections, start=1):
        lines.append(f"![[{pdf_ref}#page={n}]]")
        lines.append("")
        if section_text:
            lines.append(section_text.strip())
            lines.append("")
        lines.append("")

    return "\n".join(lines).strip() + "\n"
