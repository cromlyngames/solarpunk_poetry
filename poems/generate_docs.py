"""Generate .docx, .odt, .html, and .rtf copies of the awkward-formatting test poems."""
import os
from pathlib import Path

# python-docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# odfpy
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties, ParagraphProperties
from odf.text import P, Span

POEMS_DIR = Path(r"c:\Users\PatrickBarry\Desktop\poems")

TEST_FILES = [
    "test_slash_markers.txt",
    "test_em_dashes.txt",
    "test_curly_brackets.txt",
    "test_whitespace_spaces.txt",
    "test_whitespace_tabs.txt",
    "test_ellipsis.txt",
]


def read_poem(filename: str) -> str:
    return (POEMS_DIR / filename).read_text(encoding="utf-8")


# ── .docx ────────────────────────────────────────────────────────────────────

def write_docx(stem: str, text: str) -> None:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Courier New"
    style.font.size = Pt(12)
    for line in text.splitlines():
        para = doc.add_paragraph(line)
        para.paragraph_format.space_after = Pt(0)
    doc.save(POEMS_DIR / f"{stem}.docx")


# ── .odt ─────────────────────────────────────────────────────────────────────

def write_odt(stem: str, text: str) -> None:
    doc = OpenDocumentText()

    # paragraph style with no extra spacing
    para_style = Style(name="Poem", family="paragraph")
    para_style.addElement(ParagraphProperties(marginbottom="0cm"))
    para_style.addElement(TextProperties(fontfamily="Courier New", fontsize="12pt"))
    doc.styles.addElement(para_style)

    for line in text.splitlines():
        p = P(stylename=para_style, text=line if line else " ")
        doc.text.addElement(p)

    doc.save(str(POEMS_DIR / f"{stem}.odt"))


# ── .html (importable into Google Docs) ──────────────────────────────────────

def write_html(stem: str, text: str) -> None:
    lines_html = "\n".join(
        f"    <p>{line if line.strip() else '&nbsp;'}</p>"
        for line in text.splitlines()
    )
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>{stem}</title>
  <style>
    body {{ font-family: "Courier New", monospace; font-size: 12pt; margin: 2cm; }}
    p    {{ margin: 0; white-space: pre-wrap; }}
  </style>
</head>
<body>
{lines_html}
</body>
</html>
"""
    (POEMS_DIR / f"{stem}.html").write_text(html, encoding="utf-8")


# ── .rtf (opens natively in macOS TextEdit / Pages) ──────────────────────────

def _rtf_escape(text: str) -> str:
    """Escape non-ASCII characters and RTF special chars."""
    out = []
    for ch in text:
        if ch == "\\":
            out.append("\\\\")
        elif ch == "{":
            out.append("\\{")
        elif ch == "}":
            out.append("\\}")
        elif ord(ch) > 127:
            # encode as \uN? for Unicode
            out.append(f"\\u{ord(ch)}?")
        else:
            out.append(ch)
    return "".join(out)


def write_rtf(stem: str, text: str) -> None:
    lines = text.splitlines()
    body = "\\par\n".join(_rtf_escape(ln) for ln in lines)
    rtf = (
        "{\\rtf1\\ansi\\deff0\n"
        "{\\fonttbl{\\f0\\fmodern\\fcharset0 Courier New;}}\n"
        "{\\colortbl;\\red0\\green0\\blue0;}\n"
        "\\f0\\fs24\\cf1\n"
        f"{body}\n"
        "}"
    )
    (POEMS_DIR / f"{stem}.rtf").write_text(rtf, encoding="ascii", errors="replace")


# ── main ──────────────────────────────────────────────────────────────────────

for filename in TEST_FILES:
    stem = Path(filename).stem
    text = read_poem(filename)
    write_docx(stem, text)
    write_odt(stem, text)
    write_html(stem, text)
    write_rtf(stem, text)
    print(f"  {stem}: .docx  .odt  .html  .rtf")

print("Done.")
