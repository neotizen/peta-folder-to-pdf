import io
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
from pathlib import Path
from email import policy
from email.parser import BytesParser

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth

from pypdf import PdfReader, PdfWriter


# =========================
# 설정
# =========================
INPUT_DIR = Path("/Users/neotizen/Library/CloudStorage/GoogleDrive-neotizen@gmail.com/내 드라이브/Google - notShared/PJT-JK/김옥권/EMAIL")
OUTPUT_DIR = INPUT_DIR / "pdf"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_NAME = "combined_pdf_email"
MAX_BYTES = 100 * 1024 * 1024  # 100MB

# 첨부파일을 원본으로 보관하고 싶으면 True(원치 않으면 False)
SAVE_ATTACHMENTS_TO_DISK = True
ATTACH_ROOT = OUTPUT_DIR / "_attachments"
if SAVE_ATTACHMENTS_TO_DISK:
    ATTACH_ROOT.mkdir(parents=True, exist_ok=True)

# 폰트 (이미 가지고 있는 fonts/MALGUN.TTF)
BASE_DIR = Path(__file__).resolve().parent
FONT_PATH = BASE_DIR / "fonts" / "MALGUN.TTF"
FONT_NAME = "MalgunGothic"
pdfmetrics.registerFont(TTFont(FONT_NAME, str(FONT_PATH)))

# A4 페이지 설정
PAGE_W, PAGE_H = A4
MARGIN_LR = 40
MARGIN_TOP = 40
MARGIN_BOTTOM = 40

FONT_SIZE = 10
LINE_H = 14
MAX_TEXT_WIDTH = PAGE_W - (MARGIN_LR * 2)

# "출처" 헤더 영역(페이지 상단에 한 줄)
SOURCE_FONT_SIZE = 9
SOURCE_LINE_H = 12
SOURCE_BLOCK_H = SOURCE_LINE_H + 6  # 출처 줄 높이 + 여백
CONTENT_TOP_Y = PAGE_H - MARGIN_TOP - SOURCE_BLOCK_H  # 본문 시작 y

# 파일 타입 처리
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic"}
PDF_EXTS = {".pdf"}
DOC_EXTS = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".odp", ".rtf", ".txt", ".csv", ".hwp"}


# =========================
# 유틸
# =========================
def nfc(s: str) -> str:
    """한글 자모 분리(NFD) 등 깨짐 방지를 위해 NFC로 정규화."""
    return unicodedata.normalize("NFC", s or "")

def safe_str(s: str) -> str:
    return nfc((s or "").replace("\r\n", "\n").replace("\r", "\n"))

def sanitize_filename(name: str, fallback: str = "attachment"):
    name = nfc((name or fallback).strip())
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or fallback

def estimate_writer_size_bytes(writer: PdfWriter) -> int:
    """현재 writer를 메모리에 써서 바이트 크기 추정(정확)."""
    bio = io.BytesIO()
    writer.write(bio)
    return bio.tell()

def writer_to_file(writer: PdfWriter, out_path: Path):
    with open(out_path, "wb") as f:
        writer.write(f)

def wrap_line_by_width(line: str, font_name: str, font_size: int, max_width: float):
    line = nfc(line)
    if line == "":
        return [""]

    out = []
    s = line
    while stringWidth(s, font_name, font_size) > max_width:
        cut = len(s)
        while cut > 0 and stringWidth(s[:cut], font_name, font_size) > max_width:
            cut -= 1
        if cut <= 0:
            break

        # 공백이 있으면 단어 기준으로 끊기
        space_pos = s.rfind(" ", 0, cut)
        if space_pos > 0:
            out.append(s[:space_pos])
            s = s[space_pos + 1:]
        else:
            # 한글 연속 등 공백 없으면 문자 단위
            out.append(s[:cut])
            s = s[cut:]

    out.append(s)
    return out

def wrap_text(text: str, font_name: str, font_size: int, max_width: float):
    lines = []
    for raw in safe_str(text).split("\n"):
        lines.extend(wrap_line_by_width(raw, font_name, font_size, max_width))
    return lines

def extract_body(msg) -> str:
    """우선순위: text/plain -> 없으면 text/html(태그 포함 그대로)."""
    if msg.is_multipart():
        plain = None
        html = None
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", "")).lower()
            if "attachment" in disp:
                continue
            if ctype == "text/plain" and plain is None:
                plain = part.get_content()
            elif ctype == "text/html" and html is None:
                html = part.get_content()
        body = plain if plain else (html if html else "")
    else:
        body = msg.get_content()
    return safe_str(body)

def make_source_overlay_pdf(source_text: str) -> PdfReader:
    """A4 한 페이지짜리 overlay PDF(출처 텍스트 1줄) 생성."""
    source_text = nfc(source_text)
    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=A4)
    c.setFont(FONT_NAME, SOURCE_FONT_SIZE)

    x = MARGIN_LR
    y = PAGE_H - MARGIN_TOP - SOURCE_LINE_H  # 상단 여백 아래
    c.drawString(x, y, source_text)
    c.showPage()
    c.save()
    bio.seek(0)
    return PdfReader(bio)

def stamp_page_with_source(page, source_text: str):
    """기존 PDF 페이지 위에 [소스] 줄을 덮어씌움."""
    overlay_reader = make_source_overlay_pdf(source_text)
    overlay_page = overlay_reader.pages[0]
    page.merge_page(overlay_page)
    return page

def add_pdf_reader_to_writer_with_source(writer: PdfWriter, reader: PdfReader, source_text: str):
    """reader의 모든 페이지를 writer에 추가하면서 출처 줄을 스탬프."""
    for p in reader.pages:
        stamp_page_with_source(p, source_text)
        writer.add_page(p)

def make_email_text_pages_pdf_reader(source_text: str, text: str) -> PdfReader:
    """이메일 본문/헤더 텍스트를 A4 페이지로 만들어 PdfReader 반환(파일로 저장하지 않음)."""
    source_text = nfc(source_text)
    text = safe_str(text)
    lines = wrap_text(text, FONT_NAME, FONT_SIZE, MAX_TEXT_WIDTH)

    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=A4)
    c.setFont(FONT_NAME, FONT_SIZE)

    x = MARGIN_LR
    y = CONTENT_TOP_Y  # 출처 영역 아래에서 시작

    # 페이지마다 출처 표시
    def draw_source():
        c.setFont(FONT_NAME, SOURCE_FONT_SIZE)
        c.drawString(MARGIN_LR, PAGE_H - MARGIN_TOP - SOURCE_LINE_H, source_text)
        c.setFont(FONT_NAME, FONT_SIZE)

    draw_source()

    for line in lines:
        c.drawString(x, y, nfc(line))
        y -= LINE_H
        if y < MARGIN_BOTTOM:
            c.showPage()
            c.setFont(FONT_NAME, FONT_SIZE)
            y = CONTENT_TOP_Y
            draw_source()

    c.save()
    bio.seek(0)
    return PdfReader(bio)

def make_image_as_a4_pdf_reader(source_text: str, image_path: Path) -> PdfReader:
    """이미지를 A4에 맞춰 배치한 PDF(PdfReader) 생성."""
    source_text = nfc(source_text)
    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=A4)

    # 출처
    c.setFont(FONT_NAME, SOURCE_FONT_SIZE)
    c.drawString(MARGIN_LR, PAGE_H - MARGIN_TOP - SOURCE_LINE_H, source_text)

    # 이미지 영역(출처 아래 ~ 하단 여백 위)
    img_top = CONTENT_TOP_Y
    img_bottom = MARGIN_BOTTOM
    img_left = MARGIN_LR
    img_right = PAGE_W - MARGIN_LR
    box_w = img_right - img_left
    box_h = img_top - img_bottom

    img_reader = ImageReader(str(image_path))
    iw, ih = img_reader.getSize()

    # 비율 유지하면서 박스에 맞추기
    scale = min(box_w / iw, box_h / ih)
    draw_w = iw * scale
    draw_h = ih * scale
    x = img_left + (box_w - draw_w) / 2
    y = img_bottom + (box_h - draw_h) / 2

    c.drawImage(img_reader, x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask='auto')
    c.showPage()
    c.save()

    bio.seek(0)
    return PdfReader(bio)

def find_soffice():
    candidates = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    for c in candidates:
        if shutil.which(c) or Path(c).exists():
            return c
    return None

def convert_doc_to_pdf_with_soffice(input_path: Path, out_dir: Path) -> Path | None:
    """LibreOffice로 문서 -> PDF 변환. 성공 시 PDF 경로 반환."""
    soffice = find_soffice()
    if soffice is None:
        return None

    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(input_path)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        expected = out_dir / f"{input_path.stem}.pdf"
        return expected if expected.exists() else None
    except Exception:
        return None


# =========================
# 첨부 추출
# =========================
def extract_attachments_to_paths(msg, attach_dir: Path | None):
    """
    첨부파일을 (옵션) 디스크로 저장하고 경로 리스트 반환.
    SAVE_ATTACHMENTS_TO_DISK=False이면 임시로 메모리에만 두고 싶겠지만,
    대용량/다형식 처리 때문에 여기서는 디스크 저장 방식이 현실적으로 안정적입니다.
    """
    saved = []
    idx = 1

    if attach_dir is not None:
        attach_dir.mkdir(parents=True, exist_ok=True)

    for part in msg.walk():
        disp = str(part.get("Content-Disposition", "")).lower()
        filename = part.get_filename()

        # attachment 이거나 filename 있으면 첨부로 취급
        if "attachment" not in disp and not filename:
            continue

        filename = sanitize_filename(filename, fallback=f"attachment_{idx}")
        idx += 1

        payload = part.get_payload(decode=True)
        if payload is None:
            continue

        if attach_dir is None:
            # 디스크 저장 비활성화면 임시파일에 저장(코드 단순화 목적)
            tmpdir = Path(tempfile.gettempdir()) / "eml_attach_tmp"
            tmpdir.mkdir(parents=True, exist_ok=True)
            out_path = tmpdir / filename
        else:
            out_path = attach_dir / filename

        # 중복 방지
        if out_path.exists():
            stem, suf = out_path.stem, out_path.suffix
            n = 2
            while True:
                candidate = out_path.parent / f"{stem}_{n}{suf}"
                if not candidate.exists():
                    out_path = candidate
                    break
                n += 1

        out_path.write_bytes(payload)
        saved.append(out_path)

    return saved


# =========================
# 메일 1개를 writer에 append
# =========================
def append_one_eml(writer: PdfWriter, eml_path: Path):
    eml_name = nfc(eml_path.name)

    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    subject = safe_str(msg.get("subject", ""))
    sender = safe_str(msg.get("from", ""))
    to = safe_str(msg.get("to", ""))
    cc = safe_str(msg.get("cc", ""))
    date = safe_str(msg.get("date", ""))

    body = extract_body(msg)

    # 첨부 추출(원본 저장 폴더)
    attach_dir = None
    if SAVE_ATTACHMENTS_TO_DISK:
        attach_dir = ATTACH_ROOT / eml_path.stem

    attachments = extract_attachments_to_paths(msg, attach_dir)

    # 이메일 본문(출처 표기)
    email_source = f"[소스] {eml_name}"
    header = (
        f"Subject: {subject}\n"
        f"From: {sender}\n"
        f"To: {to}\n"
        f"CC: {cc}\n"
        f"Date: {date}\n"
        + ("-" * 90) + "\n"
        "첨부파일:\n"
    )
    if attachments:
        for a in attachments:
            header += f"- {nfc(a.name)}\n"
    else:
        header += "- (없음)\n"
    header += "\n"

    email_text = header + body
    email_reader = make_email_text_pages_pdf_reader(email_source, email_text)
    add_pdf_reader_to_writer_with_source(writer, email_reader, email_source)

    # 첨부를 페이지로 포함(출처: [소스] OO.eml의 첨부파일명)
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        conv_dir = td / "converted"
        conv_dir.mkdir(parents=True, exist_ok=True)

        for a in attachments:
            att_name = nfc(a.name)
            source_att = f"[소스] {eml_name}의 {att_name}"
            ext = a.suffix.lower()

            # PDF 첨부: 페이지 추가 + 출처 스탬프
            if ext in PDF_EXTS:
                try:
                    r = PdfReader(str(a))
                    add_pdf_reader_to_writer_with_source(writer, r, source_att)
                except Exception:
                    # 변환 실패 시 안내 페이지 1장 추가
                    notice = f"{source_att}\n\n(첨부 PDF를 읽지 못했습니다: {att_name})"
                    nr = make_email_text_pages_pdf_reader(source_att, notice)
                    add_pdf_reader_to_writer_with_source(writer, nr, source_att)
                continue

            # 이미지 첨부: A4에 배치해서 추가
            if ext in IMAGE_EXTS:
                try:
                    ir = make_image_as_a4_pdf_reader(source_att, a)
                    add_pdf_reader_to_writer_with_source(writer, ir, source_att)
                except Exception:
                    notice = f"{source_att}\n\n(이미지 변환 실패: {att_name})"
                    nr = make_email_text_pages_pdf_reader(source_att, notice)
                    add_pdf_reader_to_writer_with_source(writer, nr, source_att)
                continue

            # 문서 첨부: LibreOffice로 PDF 변환 후 추가
            if ext in DOC_EXTS:
                converted = convert_doc_to_pdf_with_soffice(a, conv_dir)
                if converted:
                    try:
                        r = PdfReader(str(converted))
                        add_pdf_reader_to_writer_with_source(writer, r, source_att)
                    except Exception:
                        notice = f"{source_att}\n\n(문서 PDF를 읽지 못했습니다: {att_name})"
                        nr = make_email_text_pages_pdf_reader(source_att, notice)
                        add_pdf_reader_to_writer_with_source(writer, nr, source_att)
                else:
                    # LibreOffice 없거나 변환 실패
                    notice = (
                        f"{source_att}\n\n"
                        f"(문서 첨부를 PDF로 변환하지 못했습니다: {att_name})\n"
                        f"- LibreOffice(soffice)가 설치되어 있는지 확인하세요.\n"
                        f"- 또는 해당 형식(HWP 등)이 변환 불가일 수 있습니다.\n"
                    )
                    nr = make_email_text_pages_pdf_reader(source_att, notice)
                    add_pdf_reader_to_writer_with_source(writer, nr, source_att)
                continue

            # 기타 파일: 안내 페이지 추가(원본은 _attachments에 보관될 수 있음)
            notice = f"{source_att}\n\n(이 첨부파일 형식은 PDF에 페이지로 포함하지 않습니다: {att_name})"
            nr = make_email_text_pages_pdf_reader(source_att, notice)
            add_pdf_reader_to_writer_with_source(writer, nr, source_att)


# =========================
# 100MB 기준 분할 저장
# =========================
def build_combined_pdfs():
    eml_files = sorted([p for p in INPUT_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".eml"])

    part_idx = 1
    writer = PdfWriter()

    def flush_current():
        nonlocal writer, part_idx
        if len(writer.pages) == 0:
            return

        out_path = OUTPUT_DIR / f"{BASE_NAME}_{part_idx:03d}.pdf"
        writer_to_file(writer, out_path)
        print(f"✅ SAVED: {out_path.name}")
        part_idx += 1
        writer = PdfWriter()

    for eml in eml_files:
        print("Appending:", nfc(eml.name))

        # 이 EML을 추가하기 전 크기
        before_size = estimate_writer_size_bytes(writer)

        # 일단 추가 시도
        pages_before = len(writer.pages)
        try:
            append_one_eml(writer, eml)
        except Exception as e:
            # 실패해도 “실패 안내 페이지”는 넣고 진행
            src = f"[소스] {nfc(eml.name)}"
            notice = f"{src}\n\n(이 EML 처리 중 오류 발생)\n{repr(e)}"
            nr = make_email_text_pages_pdf_reader(src, notice)
            add_pdf_reader_to_writer_with_source(writer, nr, src)

        # 추가 후 100MB 초과 검사
        after_size = estimate_writer_size_bytes(writer)

        if after_size > MAX_BYTES:
            # 방금 추가한 EML 때문에 초과 → 이전 writer를 저장하고, 새 writer에 방금 EML만 담기
            # (즉, 한 EML 단위로 분할)
            if pages_before > 0:
                # writer에서 방금 추가된 페이지를 떼어내는 기능이 pypdf에 깔끔히 없어서,
                # “이전 저장 + 새로 생성 후 해당 EML만 다시 추가” 방식으로 처리
                # 1) 이전까지 다시 구성하여 저장
                #    -> 현실적으로는: 초과 발생 시점에서 flush하고, 해당 eml을 새 writer에 재처리
                #    -> 이때 현재 writer는 이미 eml이 포함된 상태이므로, 저장 전략을 바꿈:
                #       - 저장은 "초과 이전" writer여야 하지만, rollback이 어려움.
                # 해결: "추가하기 전에" 크기를 봤으니,
                # - before_size가 0이면(처음부터 초과) 그대로 저장
                # - before_size > 0이면: 지금 writer를 그냥 저장하면 100MB 넘음 → 분할 요구 위반
                #
                # 그래서 가장 안전한 방식:
                # - EML 추가 전에 if before_size가 충분히 컸으면 flush 먼저.
                pass

    # 위 pass 문제를 해결하기 위해 “추가 전에 flush 판단” 방식으로 재작성
    # (아래에서 실제 동작 코드로 다시 구현)


def build_combined_pdfs_safe():
    eml_files = sorted([p for p in INPUT_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".eml"])

    part_idx = 1
    writer = PdfWriter()

    def save_part(w: PdfWriter, idx: int):
        out_path = OUTPUT_DIR / f"{BASE_NAME}_{idx:03d}.pdf"
        writer_to_file(w, out_path)
        print(f"✅ SAVED: {out_path.name} ({out_path.stat().st_size / (1024*1024):.2f} MB)")

    for eml in eml_files:
        # 1) “이 EML을 새 writer에 단독으로 넣었을 때” 크기를 먼저 측정
        test_writer = PdfWriter()
        try:
            append_one_eml(test_writer, eml)
        except Exception as e:
            src = f"[소스] {nfc(eml.name)}"
            notice = f"{src}\n\n(이 EML 처리 중 오류 발생)\n{repr(e)}"
            nr = make_email_text_pages_pdf_reader(src, notice)
            add_pdf_reader_to_writer_with_source(test_writer, nr, src)

        eml_size = estimate_writer_size_bytes(test_writer)

        # 2) 현재 writer에 합쳤을 때 100MB 넘어가면, 현재 part 저장 후 새 part 시작
        cur_size = estimate_writer_size_bytes(writer)
        if len(writer.pages) > 0 and (cur_size + eml_size) > MAX_BYTES:
            save_part(writer, part_idx)
            part_idx += 1
            writer = PdfWriter()

        # 3) 이제 안전하게 append(실제 추가)
        #    (test_writer pages를 writer로 옮김)
        for p in test_writer.pages:
            writer.add_page(p)

        print(f"Appended: {nfc(eml.name)} (EML block ~ {eml_size / (1024*1024):.2f} MB)")

        # 4) 만약 EML 단독이 이미 100MB를 초과하면, 그 EML만 있는 파일도 100MB 넘을 수 있음
        if eml_size > MAX_BYTES:
            print(f"⚠️ WARNING: 단일 EML({nfc(eml.name)})가 {MAX_BYTES/(1024*1024):.0f}MB를 초과합니다. 해당 파트는 100MB를 넘을 수 있습니다.")

            # 가능한 즉시 저장하고 새 part 시작
            save_part(writer, part_idx)
            part_idx += 1
            writer = PdfWriter()

    # 마지막 저장
    if len(writer.pages) > 0:
        save_part(writer, part_idx)

    print("DONE")


if __name__ == "__main__":
    build_combined_pdfs_safe()