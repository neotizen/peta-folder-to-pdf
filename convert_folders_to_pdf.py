#!/usr/bin/env python3
from __future__ import annotations

import argparse
import contextlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast
from urllib.parse import parse_qs, urlparse
from xml.sax.saxutils import escape

from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Preformatted, SimpleDocTemplate, Spacer, Table, TableStyle
from split_pdf_by_size import DEFAULT_SPLIT_SIZE_MB, parse_split_size_mb, split_pdf_by_size

if TYPE_CHECKING:
    from google.auth.transport.requests import AuthorizedSession, Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

try:
    from google.auth.transport.requests import AuthorizedSession as _AuthorizedSession, Request as _Request
    from google.oauth2.credentials import Credentials as _Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow as _InstalledAppFlow
except ImportError:  # pragma: no cover - optional at runtime
    _AuthorizedSession = None
    _Request = None
    _Credentials = None
    _InstalledAppFlow = None

AuthorizedSession = cast(Any, _AuthorizedSession)
Request = cast(Any, _Request)
Credentials = cast(Any, _Credentials)
InstalledAppFlow = cast(Any, _InstalledAppFlow)


PDF_EXTS = {".pdf"}
TEXT_EXTS = {".txt"}
MARKDOWN_EXTS = {".md", ".markdown"}
OFFICE_EXTS = {".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".xlx", ".hwp", ".hwpx"}
GOOGLE_EXTS = {
    ".gdoc": "document",
    ".gsheet": "spreadsheet",
    ".gslides": "presentation",
}
SUPPORTED_EXTS = PDF_EXTS | TEXT_EXTS | MARKDOWN_EXTS | OFFICE_EXTS | set(GOOGLE_EXTS)

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
EXPORT_MIME_TYPE = "application/pdf"

BASE_DIR = Path(__file__).resolve().parent
FONT_PATH = BASE_DIR / "fonts" / "MALGUN.TTF"
FONT_NAME = "Helvetica"
MARKDOWN_KO_FONT_NAME = "Helvetica"
MARKDOWN_EN_FONT_NAME = "Times-Roman"
MARKDOWN_EN_FONT_BOLD_NAME = "Times-Bold"
MARKDOWN_EN_FONT_ITALIC_NAME = "Times-Italic"
MARKDOWN_EN_FONT_BOLDITALIC_NAME = "Times-BoldItalic"
MARKDOWN_KO_FONT_CANDIDATES = [
    Path("/System/Library/Fonts/Supplemental/Batang.ttc"),
    Path("/System/Library/Fonts/Supplemental/AppleMyungjo.ttf"),
]
MARKDOWN_EN_FONT_CANDIDATES = {
    "normal": [Path("/System/Library/Fonts/Supplemental/Times New Roman.ttf")],
    "bold": [Path("/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf")],
    "italic": [Path("/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf")],
    "bold_italic": [Path("/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf")],
}

PAGE_W, PAGE_H = A4
MARGIN_LR = 40
MARGIN_TOP = 40
MARGIN_BOTTOM = 40
FONT_SIZE = 10
LINE_H = 14
MAX_TEXT_WIDTH = PAGE_W - (MARGIN_LR * 2)
HEADER_FONT_SIZE = 8
HEADER_GAP = 10
HEADER_COLOR = colors.HexColor("#808080")
HEADER_MAX_WIDTH = PAGE_W - (MARGIN_LR * 2)
CONTENT_TOP_Y = PAGE_H - MARGIN_TOP - HEADER_FONT_SIZE - HEADER_GAP
MARKDOWN_MARGIN = 2.54 * cm

GOOGLE_DOC_URL_RE = re.compile(r"https://docs\.google\.com/document/d/([A-Za-z0-9_-]+)")
GOOGLE_SHEET_URL_RE = re.compile(r"https://docs\.google\.com/spreadsheets/d/([A-Za-z0-9_-]+)")
GOOGLE_SLIDES_URL_RE = re.compile(r"https://docs\.google\.com/presentation/d/([A-Za-z0-9_-]+)")
RESOURCE_ID_RE = re.compile(r"(document|spreadsheet|presentation):([A-Za-z0-9_-]+)")
FILE_ID_RE = re.compile(r'"(?:fileId|doc_id|sheet_id|id)"\s*:\s*"([A-Za-z0-9_-]+)"')
KOREAN_CHAR_RE = re.compile(r"[\u1100-\u11FF\u3130-\u318F\uA960-\uA97F\uAC00-\uD7AF\uD7B0-\uD7FF\u4E00-\u9FFF]")


@dataclass
class GoogleFileRef:
    file_id: str
    kind: str
    resource_key: str | None = None
    source_url: str | None = None


@dataclass
class FolderResult:
    subfolder: Path
    output_pdf: Path
    source_count: int
    success_count: int
    failure_count: int
    split_pdfs: list[Path] = field(default_factory=list)


class TeeStream:
    def __init__(self, primary: Any, log_file: Any):
        self.primary = primary
        self.log_file = log_file

    def write(self, data: str) -> int:
        self.primary.write(data)
        self.log_file.write(data)
        return len(data)

    def flush(self) -> None:
        self.primary.flush()
        self.log_file.flush()

    def isatty(self) -> bool:
        return bool(getattr(self.primary, "isatty", lambda: False)())

    def __getattr__(self, name: str) -> Any:
        return getattr(self.primary, name)


def init_font() -> None:
    global FONT_NAME, MARKDOWN_KO_FONT_NAME, MARKDOWN_EN_FONT_NAME
    global MARKDOWN_EN_FONT_BOLD_NAME, MARKDOWN_EN_FONT_ITALIC_NAME, MARKDOWN_EN_FONT_BOLDITALIC_NAME
    if FONT_PATH.exists():
        pdfmetrics.registerFont(TTFont("MalgunGothic", str(FONT_PATH)))
        FONT_NAME = "MalgunGothic"

    for candidate in MARKDOWN_KO_FONT_CANDIDATES:
        if candidate.exists():
            pdfmetrics.registerFont(TTFont("MarkdownKorean", str(candidate)))
            MARKDOWN_KO_FONT_NAME = "MarkdownKorean"
            break

    english_aliases = {
        "normal": ("MarkdownEnglish", "MARKDOWN_EN_FONT_NAME"),
        "bold": ("MarkdownEnglishBold", "MARKDOWN_EN_FONT_BOLD_NAME"),
        "italic": ("MarkdownEnglishItalic", "MARKDOWN_EN_FONT_ITALIC_NAME"),
        "bold_italic": ("MarkdownEnglishBoldItalic", "MARKDOWN_EN_FONT_BOLDITALIC_NAME"),
    }
    for key, (alias, _) in english_aliases.items():
        for candidate in MARKDOWN_EN_FONT_CANDIDATES[key]:
            if candidate.exists():
                pdfmetrics.registerFont(TTFont(alias, str(candidate)))
                if key == "normal":
                    MARKDOWN_EN_FONT_NAME = alias
                elif key == "bold":
                    MARKDOWN_EN_FONT_BOLD_NAME = alias
                elif key == "italic":
                    MARKDOWN_EN_FONT_ITALIC_NAME = alias
                else:
                    MARKDOWN_EN_FONT_BOLDITALIC_NAME = alias
                break

    pdfmetrics.registerFontFamily(
        "MarkdownEnglishFamily",
        normal=MARKDOWN_EN_FONT_NAME,
        bold=MARKDOWN_EN_FONT_BOLD_NAME,
        italic=MARKDOWN_EN_FONT_ITALIC_NAME,
        boldItalic=MARKDOWN_EN_FONT_BOLDITALIC_NAME,
    )


def nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text or "")


def sanitize_filename(name: str) -> str:
    clean = nfc(name).strip()
    clean = re.sub(r'[\\/:*?"<>|]+', "_", clean)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean or "output"


def is_excluded_name(name: str) -> bool:
    return nfc(name).startswith("제외")


def build_log_path(output_dir: Path, date_str: str, root_name: str) -> Path:
    return output_dir / f"{date_str}-{sanitize_filename(root_name)}-log.txt"


@contextlib.contextmanager
def tee_output(log_path: Path):
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as log_file:
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        sys.stdout = TeeStream(original_stdout, log_file)
        sys.stderr = TeeStream(original_stderr, log_file)
        try:
            yield
        finally:
            sys.stdout.flush()
            sys.stderr.flush()
            sys.stdout = original_stdout
            sys.stderr = original_stderr


def wrap_line_by_width(line: str, max_width: float) -> list[str]:
    line = nfc(line)
    if line == "":
        return [""]

    out: list[str] = []
    current = line
    while stringWidth(current, FONT_NAME, FONT_SIZE) > max_width:
        cut = len(current)
        while cut > 0 and stringWidth(current[:cut], FONT_NAME, FONT_SIZE) > max_width:
            cut -= 1
        if cut <= 0:
            break
        space_pos = current.rfind(" ", 0, cut)
        if space_pos > 0:
            out.append(current[:space_pos])
            current = current[space_pos + 1 :]
        else:
            out.append(current[:cut])
            current = current[cut:]
    out.append(current)
    return out


def wrap_text(text: str) -> list[str]:
    lines: list[str] = []
    normalized = nfc(text.replace("\r\n", "\n").replace("\r", "\n"))
    for raw in normalized.split("\n"):
        lines.extend(wrap_line_by_width(raw, MAX_TEXT_WIDTH))
    return lines


def fit_text_to_width(text: str, font_name: str, font_size: int, max_width: float) -> str:
    fitted = nfc(text)
    ellipsis = "..."
    if stringWidth(fitted, font_name, font_size) <= max_width:
        return fitted

    while fitted:
        candidate = fitted.rstrip()
        if stringWidth(candidate + ellipsis, font_name, font_size) <= max_width:
            return candidate + ellipsis
        fitted = fitted[:-1]
    return ellipsis


def make_header_text(source_label: str, page_width: float) -> str:
    return fit_text_to_width(source_label, FONT_NAME, HEADER_FONT_SIZE, page_width - (MARGIN_LR * 2))


def make_page_counter_text(page_number: int, total_pages: int) -> str:
    return f"{page_number} / {total_pages}"


def make_header_overlay_pdf(
    source_label: str,
    page_width: float,
    page_height: float,
    page_number: int,
    total_pages: int,
) -> PdfReader:
    bio = io.BytesIO()
    pdf = canvas.Canvas(bio, pagesize=(page_width, page_height))
    pdf.setFont(FONT_NAME, HEADER_FONT_SIZE)
    pdf.setFillColor(HEADER_COLOR)
    pdf.drawString(MARGIN_LR, page_height - MARGIN_TOP, make_header_text(source_label, page_width))
    pdf.drawString(
        MARGIN_LR,
        MARGIN_BOTTOM,
        make_page_counter_text(page_number, total_pages),
    )
    pdf.save()
    bio.seek(0)
    return PdfReader(bio)


def stamp_page_with_header(page: Any, source_label: str, page_number: int, total_pages: int) -> Any:
    overlay_reader = make_header_overlay_pdf(
        source_label,
        float(page.mediabox.width),
        float(page.mediabox.height),
        page_number,
        total_pages,
    )
    page.merge_page(overlay_reader.pages[0])
    return page


def render_text_pdf(text: str, output_pdf: Path, source_label: str | None = None) -> Path:
    lines = wrap_text(text)
    if not lines:
        lines = [""]

    bio = io.BytesIO()
    pdf = canvas.Canvas(bio, pagesize=A4)
    pdf.setFont(FONT_NAME, FONT_SIZE)

    x = MARGIN_LR
    y = CONTENT_TOP_Y

    for line in lines:
        pdf.drawString(x, y, nfc(line))
        y -= LINE_H
        if y < MARGIN_BOTTOM:
            pdf.showPage()
            pdf.setFont(FONT_NAME, FONT_SIZE)
            y = CONTENT_TOP_Y

    pdf.save()
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    output_pdf.write_bytes(bio.getvalue())
    return output_pdf


def render_notice_pdf(output_pdf: Path, title: str, details: str, source_label: str | None = None) -> Path:
    body = f"{title}\n\n{details}"
    return render_text_pdf(body, output_pdf, source_label=source_label)


def normalize_markdown_inline(text: str) -> str:
    value = nfc(text)
    value = re.sub(
        r"!\[([^\]]*)\]\(([^)]+)\)",
        lambda m: f"[이미지: {m.group(1).strip() or 'image'}] ({m.group(2).strip()})",
        value,
    )
    value = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        lambda m: f"{m.group(1).strip()} ({m.group(2).strip()})",
        value,
    )
    value = re.sub(r"`([^`]+)`", lambda m: f"「{m.group(1)}」", value)
    value = re.sub(r"\*\*(.+?)\*\*", r"\1", value)
    value = re.sub(r"__(.+?)__", r"\1", value)
    value = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"\1", value)
    value = re.sub(r"(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)", r"\1", value)
    value = re.sub(r"~~(.+?)~~", r"\1", value)
    return re.sub(r"\s+", " ", value).strip()


def is_korean_text(text: str) -> bool:
    return bool(KOREAN_CHAR_RE.search(text))


def markdown_font_name_for_text(text: str, bold: bool = False, italic: bool = False) -> str:
    if is_korean_text(text):
        return MARKDOWN_KO_FONT_NAME
    if bold and italic:
        return MARKDOWN_EN_FONT_BOLDITALIC_NAME
    if bold:
        return MARKDOWN_EN_FONT_BOLD_NAME
    if italic:
        return MARKDOWN_EN_FONT_ITALIC_NAME
    return MARKDOWN_EN_FONT_NAME


def markdown_font_markup(text: str, bold: bool = False, italic: bool = False) -> str:
    value = nfc(text)
    if not value:
        return ""

    chunks: list[tuple[str, str]] = []
    current_font = markdown_font_name_for_text(value[0], bold=bold, italic=italic)
    current_chars = [value[0]]

    for char in value[1:]:
        font_name = markdown_font_name_for_text(char, bold=bold, italic=italic)
        if font_name == current_font:
            current_chars.append(char)
            continue
        chunks.append((current_font, "".join(current_chars)))
        current_font = font_name
        current_chars = [char]
    chunks.append((current_font, "".join(current_chars)))

    parts: list[str] = []
    for font_name, chunk in chunks:
        parts.append(f'<font name="{font_name}">{escape(chunk)}</font>')
    return "".join(parts)


def markdown_paragraph(text: str, style: ParagraphStyle, bold: bool = False, italic: bool = False) -> Paragraph:
    return Paragraph(markdown_font_markup(text or " ", bold=bold, italic=italic), style)


def build_markdown_styles() -> dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "MarkdownBody",
        parent=styles["BodyText"],
        fontName=MARKDOWN_KO_FONT_NAME,
        fontSize=11,
        leading=16,
        spaceBefore=0,
        spaceAfter=8,
    )
    return {
        "body": body,
        "h1": ParagraphStyle("MarkdownH1", parent=body, fontSize=20, leading=26, spaceBefore=10, spaceAfter=12),
        "h2": ParagraphStyle("MarkdownH2", parent=body, fontSize=17, leading=23, spaceBefore=10, spaceAfter=10),
        "h3": ParagraphStyle("MarkdownH3", parent=body, fontSize=15, leading=20, spaceBefore=8, spaceAfter=8),
        "h4": ParagraphStyle("MarkdownH4", parent=body, fontSize=13, leading=18, spaceBefore=6, spaceAfter=6),
        "h5": ParagraphStyle("MarkdownH5", parent=body, fontSize=12, leading=17, spaceBefore=6, spaceAfter=6),
        "h6": ParagraphStyle("MarkdownH6", parent=body, fontSize=11, leading=16, spaceBefore=6, spaceAfter=6),
        "bullet": ParagraphStyle(
            "MarkdownBullet",
            parent=body,
            leftIndent=16,
            firstLineIndent=-12,
            spaceAfter=4,
        ),
        "quote": ParagraphStyle(
            "MarkdownQuote",
            parent=body,
            leftIndent=18,
            rightIndent=12,
            textColor=colors.HexColor("#555555"),
            spaceBefore=4,
            spaceAfter=8,
        ),
        "code": ParagraphStyle(
            "MarkdownCode",
            parent=body,
            fontName=MARKDOWN_EN_FONT_NAME,
            fontSize=9,
            leading=13,
            leftIndent=12,
            rightIndent=12,
            backColor=colors.HexColor("#F5F5F5"),
            borderPadding=8,
            spaceBefore=4,
            spaceAfter=8,
        ),
        "table_header": ParagraphStyle(
            "MarkdownTableHeader",
            parent=body,
            fontSize=10,
            leading=14,
            spaceAfter=0,
        ),
        "table_cell": ParagraphStyle(
            "MarkdownTableCell",
            parent=body,
            fontSize=10,
            leading=14,
            spaceAfter=0,
        ),
    }


def append_markdown_paragraph(story: list[Any], lines: list[str], style: ParagraphStyle) -> None:
    text = normalize_markdown_inline(" ".join(line.strip() for line in lines if line.strip()))
    if text:
        story.append(markdown_paragraph(text, style))
    lines.clear()


def split_markdown_table_row(line: str) -> list[str]:
    text = line.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|"):
        text = text[:-1]
    cells = re.split(r"(?<!\\)\|", text)
    return [cell.replace(r"\|", "|").strip() for cell in cells]


def is_markdown_table_separator_row(cells: list[str]) -> bool:
    if not cells:
        return False
    return all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells)


def markdown_table_alignments(separator_cells: list[str], column_count: int) -> list[str]:
    alignments: list[str] = []
    for index in range(column_count):
        cell = separator_cells[index].replace(" ", "") if index < len(separator_cells) else "---"
        if cell.startswith(":") and cell.endswith(":"):
            alignments.append("CENTER")
        elif cell.endswith(":"):
            alignments.append("RIGHT")
        else:
            alignments.append("LEFT")
    return alignments


def build_markdown_table(table_lines: list[str], styles: dict[str, ParagraphStyle]) -> Table | None:
    rows = [split_markdown_table_row(line) for line in table_lines if line.strip()]
    if not rows:
        return None

    header_cells = rows[0]
    data_rows = rows[1:]
    alignments = ["LEFT"] * len(header_cells)

    if data_rows and is_markdown_table_separator_row(data_rows[0]):
        alignments = markdown_table_alignments(data_rows[0], len(header_cells))
        data_rows = data_rows[1:]

    column_count = max([len(header_cells)] + [len(row) for row in data_rows] or [len(header_cells)])
    if column_count == 0:
        return None

    def padded(row: list[str]) -> list[str]:
        return row + ([""] * (column_count - len(row)))

    available_width = PAGE_W - (MARKDOWN_MARGIN * 2)
    col_widths = [available_width / column_count] * column_count

    table_data: list[list[Any]] = []
    header_row = []
    for cell in padded(header_cells):
        cell_text = normalize_markdown_inline(cell) or " "
        header_row.append(markdown_paragraph(cell_text, styles["table_header"], bold=True))
    table_data.append(header_row)

    for row in data_rows:
        body_row = []
        for cell in padded(row):
            cell_text = normalize_markdown_inline(cell) or " "
            body_row.append(markdown_paragraph(cell_text, styles["table_cell"]))
        table_data.append(body_row)

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table_style_commands: list[tuple[Any, ...]] = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3F4F6")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8BCC2")),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#9AA0A6")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for index, alignment in enumerate(alignments):
        table_style_commands.append(("ALIGN", (index, 0), (index, -1), alignment))
    table.setStyle(TableStyle(table_style_commands))
    return table


def build_markdown_story(markdown_text: str) -> list[Any]:
    styles = build_markdown_styles()
    story: list[Any] = []
    paragraph_lines: list[str] = []
    lines = nfc(markdown_text.replace("\r\n", "\n").replace("\r", "\n")).split("\n")
    index = 0

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if re.match(r"^(```+|~~~+)", stripped):
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            fence = stripped[:3]
            index += 1
            code_lines: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith(fence):
                code_lines.append(lines[index])
                index += 1
            story.append(Preformatted("\n".join(code_lines) or " ", styles["code"]))
            index += 1
            continue

        if stripped == "":
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            index += 1
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if heading_match:
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            level = len(heading_match.group(1))
            heading_text = normalize_markdown_inline(heading_match.group(2))
            story.append(markdown_paragraph(heading_text or " ", styles[f"h{level}"], bold=True))
            index += 1
            continue

        if re.fullmatch(r"(\*\s*){3,}|(-\s*){3,}|(_\s*){3,}", stripped):
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            story.append(Spacer(1, 12))
            index += 1
            continue

        if stripped.startswith(">"):
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            quote_lines: list[str] = []
            while index < len(lines) and lines[index].strip().startswith(">"):
                quote_lines.append(lines[index].strip()[1:].strip())
                index += 1
            quote_text = normalize_markdown_inline(" ".join(quote_lines))
            if quote_text:
                story.append(markdown_paragraph(quote_text, styles["quote"]))
            continue

        list_match = re.match(r"^([-*+])\s+(.*)$", stripped)
        ordered_match = re.match(r"^(\d+)[\.\)]\s+(.*)$", stripped)
        if list_match or ordered_match:
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            while index < len(lines):
                current = lines[index].strip()
                bullet_match = re.match(r"^([-*+])\s+(.*)$", current)
                number_match = re.match(r"^(\d+)[\.\)]\s+(.*)$", current)
                if not bullet_match and not number_match:
                    break
                if bullet_match:
                    item_text = f"• {normalize_markdown_inline(bullet_match.group(2))}"
                else:
                    item_text = f"{number_match.group(1)}. {normalize_markdown_inline(number_match.group(2))}"
                story.append(markdown_paragraph(item_text, styles["bullet"]))
                index += 1
            continue

        if stripped.startswith("|") and "|" in stripped[1:]:
            append_markdown_paragraph(story, paragraph_lines, styles["body"])
            table_lines: list[str] = []
            while index < len(lines):
                current = lines[index].rstrip()
                if not current.strip().startswith("|"):
                    break
                table_lines.append(current)
                index += 1
            table = build_markdown_table(table_lines, styles)
            if table is not None:
                story.append(table)
                story.append(Spacer(1, 8))
            else:
                story.append(Preformatted("\n".join(table_lines), styles["code"]))
            continue

        paragraph_lines.append(line)
        index += 1

    append_markdown_paragraph(story, paragraph_lines, styles["body"])
    if not story:
        story.append(markdown_paragraph("(빈 마크다운 파일)", styles["body"]))
    return story


def convert_markdown_to_pdf(input_path: Path, output_pdf: Path, source_label: str | None = None) -> Path:
    markdown_text = read_text_file(input_path).strip("\ufeff")
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=MARKDOWN_MARGIN,
        rightMargin=MARKDOWN_MARGIN,
        topMargin=MARKDOWN_MARGIN,
        bottomMargin=MARKDOWN_MARGIN,
    )
    document.build(build_markdown_story(markdown_text))
    return output_pdf


def normalize_prompt_path(raw: str) -> str:
    cleaned = raw.strip()
    for prefix in ("source ", "cd ", "open "):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix) :].strip()
            break

    if cleaned[:1] in {"'", '"'}:
        cleaned = cleaned[1:]
    if cleaned[-1:] in {"'", '"'}:
        cleaned = cleaned[:-1]
    return cleaned.strip()


def prompt_path(
    prompt: str,
    default: Path | None = None,
    must_exist: bool = False,
    must_be_dir: bool = False,
    bare_name_base: Path | None = None,
) -> Path:
    while True:
        suffix = f" [{default}]" if default else ""
        raw = input(f"{prompt}{suffix}: ").strip()
        if raw == "":
            if default is None:
                print("경로를 입력해 주세요.")
                continue
            candidate = default
        else:
            normalized = normalize_prompt_path(raw)
            if (
                bare_name_base is not None
                and normalized
                and "/" not in normalized
                and "\\" not in normalized
                and not normalized.startswith("~")
            ):
                candidate = bare_name_base / normalized
            else:
                candidate = Path(normalized).expanduser()
        candidate = candidate.resolve()
        if must_exist and not candidate.exists():
            print(f"존재하지 않는 경로입니다: {candidate}")
            continue
        if must_be_dir and candidate.exists() and not candidate.is_dir():
            print(f"폴더 경로를 입력해 주세요: {candidate}")
            continue
        return candidate


def prompt_yes_no(prompt: str, default: bool = True) -> bool:
    default_hint = "y" if default else "n"
    while True:
        raw = input(f"{prompt} [Enter={default_hint}, y/n]: ").strip()
        normalized = normalize_prompt_path(raw).lower()
        if normalized == "":
            return default
        if normalized in {"y", "yes"}:
            return True
        if normalized in {"n", "no"}:
            return False
        print("y 또는 n 으로 입력해 주세요.")


def prompt_split_size_mb(default_mb: float = DEFAULT_SPLIT_SIZE_MB) -> float | None:
    while True:
        raw = input(
            f"PDF 스플릿 크기(MB) [Enter={int(default_mb)}, n=스플릿 안 함]: "
        ).strip()
        normalized = normalize_prompt_path(raw)
        if normalized == "":
            return default_mb
        if normalized.lower() == "n":
            return None
        try:
            return parse_split_size_mb(normalized, default_mb=default_mb)
        except ValueError as exc:
            print(str(exc))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="하위 폴더 단위 또는 단일 파일을 PDF로 변환합니다."
    )
    parser.add_argument("--root", type=Path, default=None, help="상위 폴더 또는 단일 파일 경로")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="PDF 저장 폴더. 파일 입력 시 해당 파일 폴더, 폴더 입력 시 ~/Downloads",
    )
    parser.add_argument(
        "--credentials",
        type=Path,
        default=None,
        help="Google OAuth 클라이언트 JSON 경로",
    )
    parser.add_argument(
        "--token",
        type=Path,
        default=None,
        help="Google OAuth 토큰 저장 경로. 기본값은 스크립트 폴더의 google_token.json",
    )
    parser.add_argument(
        "--date",
        default=datetime.now().strftime("%y%m%d"),
        help="출력 파일 날짜 문자열(yymmdd). 기본값은 오늘",
    )
    parser.add_argument(
        "--include-subfolders",
        default=None,
        help="폴더 입력 시 하위폴더 포함 여부. y/yes 또는 n/no",
    )
    parser.add_argument(
        "--split-size-mb",
        default=None,
        help="PDF 스플릿 기준 크기(MB). 비우면 199, n이면 스플릿 안 함",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="필수 값 누락 시 프롬프트를 띄우지 않고 실패",
    )
    return parser.parse_args()


def ensure_valid_date(date_str: str) -> str:
    if len(date_str) != 6 or not date_str.isdigit():
        raise ValueError(f"잘못된 날짜 형식입니다: {date_str} (yymmdd 필요)")
    return date_str


def discover_credentials_file(explicit: Path | None) -> Path | None:
    if explicit is not None:
        candidate = explicit.expanduser().resolve()
        return candidate if candidate.exists() else candidate

    env_path = os.getenv("GOOGLE_OAUTH_CLIENT_SECRET")
    candidates: list[Path] = []
    if env_path:
        candidates.append(Path(env_path).expanduser())

    candidates.extend(
        [
            BASE_DIR / "credentials.json",
            BASE_DIR / "google_credentials.json",
            BASE_DIR / "client_secret.json",
        ]
    )

    for candidate in candidates:
        if candidate.exists():
            return candidate.resolve()
    return None


def default_token_path() -> Path:
    env_path = os.getenv("GOOGLE_OAUTH_TOKEN")
    if env_path:
        return Path(env_path).expanduser().resolve()
    return (BASE_DIR / "google_token.json").resolve()


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def collect_immediate_subfolders(root: Path) -> list[Path]:
    return sorted(
        [path for path in root.iterdir() if path.is_dir() and not is_excluded_name(path.name)],
        key=lambda item: item.name.lower(),
    )


def collect_supported_files(subfolder: Path, output_dir: Path, recursive: bool = True) -> list[Path]:
    files: list[Path] = []
    if not recursive:
        for path in sorted(subfolder.iterdir(), key=lambda item: item.name.lower()):
            if path.is_dir():
                continue
            if is_excluded_name(path.name):
                continue
            if is_relative_to(path, output_dir):
                continue
            if path.suffix.lower() in SUPPORTED_EXTS:
                files.append(path)
        return files

    for current_root, dirnames, filenames in os.walk(subfolder):
        current_root_path = Path(current_root)
        dirnames[:] = sorted(
            [
                dirname
                for dirname in dirnames
                if not is_excluded_name(dirname)
                and not is_relative_to(current_root_path / dirname, output_dir)
            ],
            key=str.lower,
        )
        for filename in sorted(filenames, key=str.lower):
            if is_excluded_name(filename):
                continue
            path = current_root_path / filename
            if is_relative_to(path, output_dir):
                continue
            if path.suffix.lower() in SUPPORTED_EXTS:
                files.append(path)
    return files


def iter_soffice_candidates() -> list[str]:
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        shutil.which("soffice"),
        "soffice",
    ]
    resolved: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        if not candidate:
            continue
        if candidate in seen:
            continue
        if shutil.which(candidate) or Path(candidate).exists():
            resolved.append(candidate)
            seen.add(candidate)
    return resolved


def convert_with_soffice(input_path: Path, out_dir: Path) -> Path:
    soffice_candidates = iter_soffice_candidates()
    if not soffice_candidates:
        raise RuntimeError("LibreOffice(soffice)를 찾을 수 없습니다.")

    out_dir.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []

    for soffice in soffice_candidates:
        for existing_pdf in out_dir.glob("*.pdf"):
            existing_pdf.unlink()

        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(input_path)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
        except subprocess.CalledProcessError as exc:
            message = exc.stderr.strip() or exc.stdout.strip() or "변환 실패"
            errors.append(f"{soffice}: {message}")
            continue

        output_pdfs = sorted(out_dir.glob("*.pdf"))
        if output_pdfs:
            return output_pdfs[0]
        errors.append(f"{soffice}: LibreOffice 변환 후 PDF가 생성되지 않았습니다.")

    joined = " | ".join(errors) if errors else "변환 실패"
    raise RuntimeError(joined)


def read_text_file(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp949", "euc-kr", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def convert_text_to_pdf(input_path: Path, output_pdf: Path, source_label: str | None = None) -> Path:
    text = read_text_file(input_path).strip("\ufeff")
    if text.strip() == "":
        text = "(빈 텍스트 파일)"
    return render_text_pdf(text, output_pdf, source_label=source_label)


def parse_google_file_ref(path: Path) -> GoogleFileRef:
    ext = path.suffix.lower()
    default_kind = GOOGLE_EXTS[ext]
    raw = path.read_bytes()

    text_candidates: list[str] = []
    for encoding in ("utf-8-sig", "utf-16", "utf-16le", "utf-16be", "cp949", "euc-kr", "latin-1"):
        try:
            decoded = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        if decoded not in text_candidates:
            text_candidates.append(decoded)

    file_id: str | None = None
    kind = default_kind
    resource_key: str | None = None
    source_url: str | None = None

    def inspect_url(value: str) -> None:
        nonlocal file_id, kind, resource_key, source_url
        doc_match = GOOGLE_DOC_URL_RE.search(value)
        sheet_match = GOOGLE_SHEET_URL_RE.search(value)
        slides_match = GOOGLE_SLIDES_URL_RE.search(value)
        full_url_match = re.search(r"https://docs\.google\.com/[^\s\"']+", value)
        if doc_match:
            file_id = doc_match.group(1)
            kind = "document"
            source_url = full_url_match.group(0) if full_url_match else doc_match.group(0)
        elif sheet_match:
            file_id = sheet_match.group(1)
            kind = "spreadsheet"
            source_url = full_url_match.group(0) if full_url_match else sheet_match.group(0)
        elif slides_match:
            file_id = slides_match.group(1)
            kind = "presentation"
            source_url = full_url_match.group(0) if full_url_match else slides_match.group(0)

        if not source_url:
            return

        parsed = urlparse(source_url)
        query = parse_qs(parsed.query)
        resource_key = query.get("resourcekey", query.get("resourceKey", [None]))[0]

    def inspect_text(value: str) -> None:
        nonlocal file_id, kind, resource_key
        inspect_url(value)
        if file_id and resource_key is not None:
            return

        resource_match = RESOURCE_ID_RE.search(value)
        if resource_match:
            kind = resource_match.group(1)
            file_id = resource_match.group(2)

        if resource_key is None:
            key_match = re.search(r'"(?:resourceKey|resource_key)"\s*:\s*"([^"]+)"', value)
            if key_match:
                resource_key = key_match.group(1)

        if file_id is None:
            id_match = FILE_ID_RE.search(value)
            if id_match:
                file_id = id_match.group(1)

        if file_id is None:
            url_line = re.search(r"URL=(https://[^\s]+)", value)
            if url_line:
                inspect_url(url_line.group(1))

    for candidate in text_candidates:
        stripped = candidate.strip()
        if stripped.startswith("{") and stripped.endswith("}"):
            try:
                payload = json.loads(stripped)
            except json.JSONDecodeError:
                payload = None
            if payload is not None:
                stack: list[Any] = [payload]
                while stack:
                    current = stack.pop()
                    if isinstance(current, dict):
                        stack.extend(current.values())
                    elif isinstance(current, list):
                        stack.extend(current)
                    elif isinstance(current, str):
                        inspect_text(current)
                    if file_id:
                        break
        inspect_text(candidate)
        if file_id:
            break

    if file_id is None:
        raise RuntimeError(f"Google 바로가기 파일에서 문서 ID를 찾지 못했습니다: {path}")

    return GoogleFileRef(file_id=file_id, kind=kind, resource_key=resource_key, source_url=source_url)


class GoogleExporter:
    def __init__(self, credentials_path: Path | None, token_path: Path, interactive: bool):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.interactive = interactive
        self._session: AuthorizedSession | None = None

    def _ensure_google_packages(self) -> None:
        if AuthorizedSession is None or Request is None or Credentials is None or InstalledAppFlow is None:
            raise RuntimeError(
                "Google 변환용 패키지가 설치되어 있지 않습니다. "
                "./venv/bin/python -m pip install google-auth google-auth-oauthlib 를 실행해 주세요."
            )

    def _build_session(self) -> AuthorizedSession:
        self._ensure_google_packages()

        creds: Credentials | None = None
        if self.token_path.exists():
            creds = Credentials.from_authorized_user_file(str(self.token_path), SCOPES)

        if creds and creds.valid:
            return AuthorizedSession(creds)

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if self.credentials_path is None or not self.credentials_path.exists():
                raise RuntimeError(
                    "Google OAuth 클라이언트 JSON을 찾지 못했습니다. "
                    "--credentials 옵션 또는 GOOGLE_OAUTH_CLIENT_SECRET 환경변수를 사용해 주세요."
                )
            if not self.interactive:
                raise RuntimeError("Google 인증이 필요하지만 비대화형 실행이라 진행할 수 없습니다.")

            flow = InstalledAppFlow.from_client_secrets_file(str(self.credentials_path), SCOPES)
            try:
                creds = flow.run_local_server(port=0, open_browser=True)
            except Exception:
                creds = flow.run_console()

        self.token_path.parent.mkdir(parents=True, exist_ok=True)
        self.token_path.write_text(creds.to_json(), encoding="utf-8")
        return AuthorizedSession(creds)

    @property
    def session(self) -> AuthorizedSession:
        if self._session is None:
            self._session = self._build_session()
        return self._session

    def export_to_pdf(self, ref: GoogleFileRef, output_pdf: Path) -> Path:
        headers: dict[str, str] = {}
        if ref.resource_key:
            headers["X-Goog-Drive-Resource-Keys"] = f"{ref.file_id}/{ref.resource_key}"

        response = self.session.get(
            f"https://www.googleapis.com/drive/v3/files/{ref.file_id}/export",
            params={"mimeType": EXPORT_MIME_TYPE},
            headers=headers,
            timeout=300,
        )
        if not response.ok:
            message = response.text.strip()
            raise RuntimeError(f"Google export 실패 ({response.status_code}): {message}")

        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        output_pdf.write_bytes(response.content)
        if output_pdf.stat().st_size == 0:
            raise RuntimeError("Google export 결과가 비어 있습니다.")
        return output_pdf


def append_pdf_to_writer(writer: PdfWriter, pdf_path: Path, source_label: str | None = None) -> None:
    reader = PdfReader(str(pdf_path), strict=False)
    if reader.is_encrypted:
        reader.decrypt("")
    total_pages = len(reader.pages)
    for index, page in enumerate(reader.pages, start=1):
        if source_label:
            stamp_page_with_header(page, source_label, index, total_pages)
        writer.add_page(page)


def save_writer(writer: PdfWriter, output_pdf: Path) -> None:
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with output_pdf.open("wb") as file_obj:
        writer.write(file_obj)


def cleanup_split_artifacts(output_pdf: Path) -> None:
    for part_path in sorted(output_pdf.parent.glob(f"{output_pdf.stem}-part*.pdf")):
        if part_path.is_file():
            part_path.unlink()


def build_output_name(date_str: str, root_name: str, sub_name: str) -> str:
    return f"{date_str}-{sanitize_filename(root_name)}-{sanitize_filename(sub_name)}.pdf"


def build_single_file_output_path(source_path: Path, output_dir: Path) -> Path:
    output_pdf = output_dir / f"{sanitize_filename(source_path.stem)}.pdf"
    if output_pdf.resolve() == source_path.resolve():
        output_pdf = output_dir / f"{sanitize_filename(source_path.stem)}-converted.pdf"
    return output_pdf


def build_temp_pdf_path(work_dir: Path, source_path: Path) -> Path:
    return work_dir / f"{sanitize_filename(source_path.stem)}.pdf"


def convert_source_to_pdf(
    source_path: Path,
    work_dir: Path,
    google_exporter: GoogleExporter | None,
    source_label: str | None = None,
) -> Path:
    ext = source_path.suffix.lower()
    if ext in PDF_EXTS:
        return source_path
    if ext in MARKDOWN_EXTS:
        return convert_markdown_to_pdf(
            source_path,
            build_temp_pdf_path(work_dir, source_path),
            source_label=source_label,
        )
    if ext in TEXT_EXTS:
        return convert_text_to_pdf(
            source_path,
            build_temp_pdf_path(work_dir, source_path),
            source_label=source_label,
        )
    if ext in OFFICE_EXTS:
        return convert_with_soffice(source_path, work_dir)
    if ext in GOOGLE_EXTS:
        if google_exporter is None:
            raise RuntimeError("Google 바로가기 변환기가 초기화되지 않았습니다.")
        ref = parse_google_file_ref(source_path)
        return google_exporter.export_to_pdf(ref, build_temp_pdf_path(work_dir, source_path))
    raise RuntimeError(f"지원되지 않는 확장자입니다: {source_path.suffix}")


def build_folder_pdf_result(
    display_name: str,
    source_label_prefix: str,
    source_folder: Path,
    output_pdf: Path,
    output_dir: Path,
    recursive: bool,
    google_exporter: GoogleExporter | None,
) -> FolderResult:
    source_files = collect_supported_files(source_folder, output_dir, recursive=recursive)
    success_count = 0
    failure_count = 0
    writer = PdfWriter()

    with tempfile.TemporaryDirectory(prefix="folder-to-pdf-") as temp_root:
        temp_root_path = Path(temp_root)

        if not source_files:
            notice_pdf = temp_root_path / "empty.pdf"
            source_label = nfc(source_label_prefix)
            render_notice_pdf(
                notice_pdf,
                f"{display_name}",
                "지원되는 파일이 없어 안내 페이지만 생성했습니다.",
                source_label=source_label,
            )
            append_pdf_to_writer(writer, notice_pdf, source_label=source_label)
        else:
            for index, source_file in enumerate(source_files, start=1):
                relative_path = source_file.relative_to(source_folder)
                source_label = nfc(f"{source_label_prefix}/{relative_path.as_posix()}")
                work_dir = temp_root_path / f"{index:04d}"
                work_dir.mkdir(parents=True, exist_ok=True)
                try:
                    converted_pdf = convert_source_to_pdf(
                        source_file,
                        work_dir,
                        google_exporter,
                        source_label=source_label,
                    )
                    append_pdf_to_writer(
                        writer,
                        converted_pdf,
                        source_label=source_label,
                    )
                    success_count += 1
                    print(f"OK\t{display_name}\t{relative_path}")
                except Exception as exc:  # pragma: no cover - runtime/file dependent
                    failure_count += 1
                    notice_pdf = work_dir / "failed.pdf"
                    render_notice_pdf(
                        notice_pdf,
                        f"변환 실패: {source_file.name}",
                        f"{source_file}\n\n{type(exc).__name__}: {exc}",
                        source_label=source_label,
                    )
                    append_pdf_to_writer(writer, notice_pdf, source_label=source_label)
                    print(f"FAIL\t{display_name}\t{relative_path}\t{type(exc).__name__}: {exc}")

        cleanup_split_artifacts(output_pdf)
        save_writer(writer, output_pdf)

    return FolderResult(
        subfolder=source_folder,
        output_pdf=output_pdf,
        source_count=len(source_files),
        success_count=success_count,
        failure_count=failure_count,
    )


def build_subfolder_pdf(
    root: Path,
    subfolder: Path,
    output_dir: Path,
    date_str: str,
    google_exporter: GoogleExporter | None,
) -> FolderResult:
    output_pdf = output_dir / build_output_name(date_str, root.name, subfolder.name)
    return build_folder_pdf_result(
        display_name=subfolder.name,
        source_label_prefix=f"{root.name}/{subfolder.name}",
        source_folder=subfolder,
        output_pdf=output_pdf,
        output_dir=output_dir,
        recursive=True,
        google_exporter=google_exporter,
    )


def build_aggregate_pdf(
    results: list[FolderResult],
    output_pdf: Path,
) -> Path:
    writer = PdfWriter()
    for result in results:
        append_pdf_to_writer(writer, result.output_pdf)
    cleanup_split_artifacts(output_pdf)
    save_writer(writer, output_pdf)
    return output_pdf


def build_single_file_pdf(
    source_path: Path,
    output_dir: Path,
    google_exporter: GoogleExporter | None,
) -> Path:
    if source_path.suffix.lower() not in SUPPORTED_EXTS:
        raise RuntimeError(f"지원되지 않는 확장자입니다: {source_path.suffix}")

    output_pdf = build_single_file_output_path(source_path, output_dir)
    source_label = nfc(source_path.name)
    writer = PdfWriter()

    with tempfile.TemporaryDirectory(prefix="file-to-pdf-") as temp_root:
        temp_root_path = Path(temp_root)
        converted_pdf = convert_source_to_pdf(
            source_path,
            temp_root_path,
            google_exporter,
            source_label=source_label,
        )
        append_pdf_to_writer(writer, converted_pdf, source_label=source_label)
        cleanup_split_artifacts(output_pdf)
        save_writer(writer, output_pdf)

    return output_pdf


def has_google_files(files: list[Path]) -> bool:
    return any(source_file.suffix.lower() in GOOGLE_EXTS for source_file in files)


def has_google_shortcuts(subfolders: list[Path], output_dir: Path, recursive: bool = True) -> bool:
    for subfolder in subfolders:
        for source_file in collect_supported_files(subfolder, output_dir, recursive=recursive):
            if source_file.suffix.lower() in GOOGLE_EXTS:
                return True
    return False


def resolve_paths(args: argparse.Namespace) -> tuple[Path, Path, Path | None, Path]:
    interactive = not args.non_interactive

    if args.root is not None:
        root = args.root.expanduser().resolve()
    elif interactive:
        root = prompt_path("PDF로 변환할 대상 파일 또는 폴더 경로", must_exist=True)
    else:
        raise RuntimeError("--root 가 필요합니다.")

    if not root.exists():
        raise RuntimeError(f"입력 경로를 찾을 수 없습니다: {root}")
    if not root.is_dir() and not root.is_file():
        raise RuntimeError(f"폴더 또는 파일 경로를 입력해 주세요: {root}")

    default_output = root.parent if root.is_file() else (Path.home() / "Downloads")
    if args.output_dir is not None:
        output_dir = args.output_dir.expanduser().resolve()
    elif interactive:
        output_dir = prompt_path(
            "PDF 저장 타겟 폴더",
            default=default_output,
            must_exist=False,
            bare_name_base=default_output,
        )
    else:
        output_dir = default_output.resolve()
    if output_dir.exists() and not output_dir.is_dir():
        raise RuntimeError(f"출력 경로가 폴더가 아닙니다: {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    credentials_path = discover_credentials_file(args.credentials)
    token_path = args.token.expanduser().resolve() if args.token else default_token_path()
    return root, output_dir, credentials_path, token_path


def resolve_split_size_mb(args: argparse.Namespace, interactive: bool) -> float | None:
    if args.split_size_mb is not None:
        normalized = normalize_prompt_path(str(args.split_size_mb))
        if normalized.lower() == "n":
            return None
        return parse_split_size_mb(normalized, default_mb=DEFAULT_SPLIT_SIZE_MB)
    if interactive:
        return prompt_split_size_mb()
    return DEFAULT_SPLIT_SIZE_MB


def resolve_include_subfolders(args: argparse.Namespace, root: Path, interactive: bool) -> bool:
    if not root.is_dir():
        return False
    if args.include_subfolders is not None:
        normalized = normalize_prompt_path(str(args.include_subfolders)).lower()
        if normalized in {"", "y", "yes"}:
            return True
        if normalized in {"n", "no"}:
            return False
        raise ValueError("--include-subfolders 는 y/yes 또는 n/no 이어야 합니다.")
    if interactive:
        return prompt_yes_no("하위폴더를 포함할까요?", default=True)
    return True


def maybe_prompt_credentials(
    credentials_path: Path | None,
    interactive: bool,
    token_path: Path,
) -> Path | None:
    if credentials_path is not None and credentials_path.exists():
        return credentials_path
    if token_path.exists():
        return credentials_path
    if not interactive:
        return credentials_path

    default_candidate = (BASE_DIR / "credentials.json").resolve()
    response = input(
        "Google OAuth client JSON 경로"
        f" [{default_candidate}]"
        " (Enter 시 기본 경로 사용): "
    ).strip()
    if response == "":
        return default_candidate if default_candidate.exists() else credentials_path
    return Path(response).expanduser().resolve()


def print_summary(results: list[FolderResult], aggregate_pdf: Path | None = None) -> None:
    print("\nSUMMARY")
    for result in results:
        split_suffix = ""
        if result.split_pdfs:
            split_suffix = f"\tsplit={', '.join(path.name for path in result.split_pdfs)}"
        print(
            f"{result.subfolder.name}\t"
            f"sources={result.source_count}\t"
            f"success={result.success_count}\t"
            f"failed={result.failure_count}\t"
            f"-> {result.output_pdf.name}"
            f"{split_suffix}"
        )
    if aggregate_pdf is not None:
        print(f"AGG\t{aggregate_pdf.name}")


def apply_optional_splits(results: list[FolderResult], aggregate_pdf: Path, split_size_mb: float | None) -> list[Path]:
    if split_size_mb is None:
        return []

    aggregate_parts = split_pdf_by_size(aggregate_pdf, max_size_mb=split_size_mb)
    for result in results:
        result.split_pdfs = split_pdf_by_size(result.output_pdf, max_size_mb=split_size_mb)
    return aggregate_parts


def main() -> int:
    init_font()
    args = parse_args()
    log_path: Path | None = None

    try:
        date_str = ensure_valid_date(args.date)
        root, output_dir, credentials_path, token_path = resolve_paths(args)
        log_path = build_log_path(output_dir, date_str, root.name)
        interactive = not args.non_interactive and sys.stdin.isatty()
        include_subfolders = resolve_include_subfolders(args, root, interactive)
        split_size_mb = resolve_split_size_mb(args, interactive)

        with tee_output(log_path):
            print(f"LOG\t{log_path.name}")
            print(f"START\t{datetime.now().isoformat(timespec='seconds')}")
            print(f"ROOT\t{root}")
            print(f"OUTPUT\t{output_dir}")
            print(
                "SPLIT\t"
                + ("disabled" if split_size_mb is None else f"{split_size_mb:g}MB")
            )

            if root.is_file():
                needs_google = root.suffix.lower() in GOOGLE_EXTS
                google_exporter: GoogleExporter | None = None
                if needs_google:
                    credentials_path = maybe_prompt_credentials(credentials_path, interactive, token_path)
                    google_exporter = GoogleExporter(
                        credentials_path=credentials_path,
                        token_path=token_path,
                        interactive=interactive,
                    )

                output_pdf = build_single_file_pdf(root, output_dir, google_exporter)
                split_pdfs = split_pdf_by_size(output_pdf, max_size_mb=split_size_mb) if split_size_mb is not None else []
                print(f"OK\t{root.name}\t-> {output_pdf.name}")
                if split_pdfs:
                    print(
                        f"SPLIT\t{output_pdf.name}\t"
                        f"{', '.join(path.name for path in split_pdfs)}"
                    )
                print("\nSUMMARY")
                print(f"FILE\t{root.name}\t-> {output_pdf.name}")
                return 0

            subfolders = collect_immediate_subfolders(root) if include_subfolders else []
            direct_root_files = collect_supported_files(root, output_dir, recursive=False)
            folder_mode = include_subfolders and bool(subfolders)
            needs_google = has_google_files(direct_root_files)
            if folder_mode:
                needs_google = needs_google or has_google_shortcuts(subfolders, output_dir, recursive=True)
            else:
                needs_google = needs_google or has_google_files(
                    collect_supported_files(root, output_dir, recursive=include_subfolders)
                )

            google_exporter: GoogleExporter | None = None
            if needs_google:
                credentials_path = maybe_prompt_credentials(credentials_path, interactive, token_path)
                google_exporter = GoogleExporter(
                    credentials_path=credentials_path,
                    token_path=token_path,
                    interactive=interactive,
                )

            results: list[FolderResult] = []
            if folder_mode and direct_root_files:
                root_output_pdf = output_dir / build_output_name(date_str, root.name, "root")
                results.append(
                    build_folder_pdf_result(
                        display_name="root",
                        source_label_prefix=root.name,
                        source_folder=root,
                        output_pdf=root_output_pdf,
                        output_dir=output_dir,
                        recursive=False,
                        google_exporter=google_exporter,
                    )
                )
            if folder_mode:
                for subfolder in subfolders:
                    result = build_subfolder_pdf(
                        root,
                        subfolder,
                        output_dir,
                        date_str,
                        google_exporter,
                    )
                    results.append(result)
            else:
                single_output_pdf = output_dir / build_output_name(date_str, root.name, "root")
                result = build_folder_pdf_result(
                    display_name=root.name,
                    source_label_prefix=root.name,
                    source_folder=root,
                    output_pdf=single_output_pdf,
                    output_dir=output_dir,
                    recursive=include_subfolders,
                    google_exporter=google_exporter,
                )
                results.append(result)

            if len(results) == 1:
                result = results[0]
                if split_size_mb is not None:
                    result.split_pdfs = split_pdf_by_size(result.output_pdf, max_size_mb=split_size_mb)
                print(f"OK\t{root.name}\t-> {result.output_pdf.name}")
                if result.split_pdfs:
                    print(
                        f"SPLIT\t{result.output_pdf.name}\t"
                        f"{', '.join(path.name for path in result.split_pdfs)}"
                    )
                print_summary(results)
                return 1 if result.failure_count else 0

            aggregate_pdf = output_dir / build_output_name(date_str, root.name, "agg")
            aggregate_pdf = build_aggregate_pdf(results, aggregate_pdf)
            aggregate_split_pdfs = apply_optional_splits(results, aggregate_pdf, split_size_mb)
            for result in results:
                if result.split_pdfs:
                    print(
                        f"SPLIT\t{result.output_pdf.name}\t"
                        f"{', '.join(path.name for path in result.split_pdfs)}"
                    )
            if aggregate_split_pdfs:
                print(
                    f"SPLIT\t{aggregate_pdf.name}\t"
                    f"{', '.join(path.name for path in aggregate_split_pdfs)}"
                )
            print_summary(results, aggregate_pdf)
            return 1 if any(result.failure_count for result in results) else 0
    except KeyboardInterrupt:  # pragma: no cover - interactive runtime
        if log_path is not None:
            with tee_output(log_path):
                print("\n사용자가 작업을 취소했습니다.")
        else:
            print("\n사용자가 작업을 취소했습니다.")
        return 130
    except Exception as exc:
        if log_path is not None:
            with tee_output(log_path):
                print(f"ERROR\t{type(exc).__name__}: {exc}")
        else:
            print(f"ERROR\t{type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
