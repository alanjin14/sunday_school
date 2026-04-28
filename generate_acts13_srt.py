"""Generate SRT subtitle file for Acts 13 sermon from docx manuscript.

YouTube auto-subs were blocked by bot detection. This script reconstructs an SRT
file from the user-provided sermon manuscript (Acts 13 2nd Talk v4.docx) by
distributing the manuscript text across the user-specified sermon time span
(28:00 - 1:16:00 = 48 minutes).

Each subtitle line is roughly equally distributed in time. The text is the
actual manuscript prose so the content is accurate even though timestamps are
synthetic (manuscript-derived, not audio-aligned).
"""

from docx import Document
import re
import os


SERMON_START_SEC = 28 * 60          # 28:00
SERMON_END_SEC = 1 * 3600 + 16 * 60  # 1:16:00
TOTAL_SEC = SERMON_END_SEC - SERMON_START_SEC

DOCX_PATH = r"C:\sunday_school\Acts 13 2nd Talk v4.docx"
SRT_PATH = r"C:\sunday_school\ECC Redmond English Sunday Service 4.26.2026 - Your Mission in the Great Commission.srt"


def fmt_ts(total_seconds: float) -> str:
    """Format seconds as SRT timestamp HH:MM:SS,mmm."""
    ms = int(round((total_seconds - int(total_seconds)) * 1000))
    s = int(total_seconds)
    h = s // 3600
    m = (s % 3600) // 60
    sec = s % 60
    return f"{h:02d}:{m:02d}:{sec:02d},{ms:03d}"


def split_into_chunks(text: str, target_chars: int = 90):
    """Split text into subtitle-sized chunks, breaking at sentence boundaries
    when possible, falling back to word-boundary splits for long sentences."""
    sentences = re.split(r"(?<=[\.\!\?])\s+", text.strip())
    chunks = []
    for sent in sentences:
        sent = sent.strip()
        if not sent:
            continue
        if len(sent) <= target_chars:
            chunks.append(sent)
        else:
            words = sent.split()
            cur = []
            cur_len = 0
            for w in words:
                if cur_len + len(w) + 1 > target_chars and cur:
                    chunks.append(" ".join(cur))
                    cur = [w]
                    cur_len = len(w)
                else:
                    cur.append(w)
                    cur_len += len(w) + 1
            if cur:
                chunks.append(" ".join(cur))
    return chunks


def main():
    doc = Document(DOCX_PATH)

    # Skip the outline/intro paragraphs at the top of the doc; the actual
    # spoken sermon manuscript begins after "Start of Talk".
    paragraphs = [p.text for p in doc.paragraphs]
    start_idx = 0
    for i, t in enumerate(paragraphs):
        if t.strip().lower().startswith("start of talk"):
            start_idx = i + 1
            break

    body = []
    for t in paragraphs[start_idx:]:
        t = t.strip()
        if not t:
            continue
        # Skip outline restatements / verse blocks that are pure headings
        if t.lower() in ("outline:",):
            continue
        body.append(t)

    full_text = " ".join(body)
    # Normalize unicode quotes for clean SRT
    full_text = (
        full_text
        .replace("’", "'")
        .replace("‘", "'")
        .replace("“", '"')
        .replace("”", '"')
        .replace("—", "--")
        .replace("–", "-")
        .replace("…", "...")
    )

    chunks = split_into_chunks(full_text, target_chars=90)
    n = len(chunks)
    if n == 0:
        raise RuntimeError("No subtitle chunks produced from manuscript.")

    per_chunk = TOTAL_SEC / n
    # Add a tiny gap between cues so they don't overlap
    gap = 0.05

    lines = []
    for i, chunk in enumerate(chunks):
        start_sec = SERMON_START_SEC + i * per_chunk
        end_sec = SERMON_START_SEC + (i + 1) * per_chunk - gap
        if end_sec <= start_sec:
            end_sec = start_sec + max(per_chunk - gap, 0.5)
        lines.append(str(i + 1))
        lines.append(f"{fmt_ts(start_sec)} --> {fmt_ts(end_sec)}")
        lines.append(chunk)
        lines.append("")

    with open(SRT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"Saved -> {SRT_PATH}")
    print(f"Cues: {n}")
    print(f"Span: 28:00 - 1:16:00 ({TOTAL_SEC}s, ~{per_chunk:.2f}s per cue)")


if __name__ == "__main__":
    main()
