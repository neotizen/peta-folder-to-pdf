#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata
import uuid
from dataclasses import dataclass
from html import escape
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

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


BASE_DIR = Path(__file__).resolve().parent
MARKDOWN_EXTS = {".md", ".markdown"}
SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/userinfo.email",
]
WARNING_TEXT = "경고: 이 파일을 수정하지 마세요. 변경사항이 저장되지 않습니다."


@dataclass
class UploadResult:
    source_path: Path
    shortcut_path: Path
    doc_id: str
    web_view_link: str


def nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text or "")


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


def prompt_path(prompt: str, must_exist: bool = False) -> Path:
    while True:
        raw = input(f"{prompt}: ").strip()
        normalized = normalize_prompt_path(raw)
        if not normalized:
            print("경로를 입력해 주세요.")
            continue
        candidate = Path(normalized).expanduser().resolve()
        if must_exist and not candidate.exists():
            print(f"존재하지 않는 경로입니다: {candidate}")
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


def read_text_file(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-16", "cp949", "euc-kr", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


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
    env_path = os.getenv("GOOGLE_MD_GDOC_OAUTH_TOKEN")
    if env_path:
        return Path(env_path).expanduser().resolve()
    return (BASE_DIR / "google_md_to_gdoc_token.json").resolve()


def prompt_credentials_path(credentials_path: Path | None, interactive: bool) -> Path | None:
    if credentials_path is not None and credentials_path.exists():
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Markdown 파일을 Google Docs(.gdoc)로 변환합니다.")
    parser.add_argument("--target", type=Path, default=None, help="Markdown 파일 또는 폴더 경로")
    parser.add_argument(
        "--include-subfolders",
        default=None,
        help="폴더 입력 시 하위폴더 포함 여부. y/yes 또는 n/no",
    )
    parser.add_argument("--credentials", type=Path, default=None, help="Google OAuth 클라이언트 JSON 경로")
    parser.add_argument(
        "--token",
        type=Path,
        default=None,
        help="Google OAuth 토큰 저장 경로. 기본값은 스크립트 폴더의 google_md_to_gdoc_token.json",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="필수 값 누락 시 프롬프트를 띄우지 않고 실패",
    )
    return parser.parse_args()


def resolve_include_subfolders(raw_value: Any, interactive: bool, target: Path) -> bool:
    if not target.is_dir():
        return False
    if raw_value is not None:
        normalized = normalize_prompt_path(str(raw_value)).lower()
        if normalized in {"", "y", "yes"}:
            return True
        if normalized in {"n", "no"}:
            return False
        raise ValueError("--include-subfolders 는 y/yes 또는 n/no 이어야 합니다.")
    if interactive:
        return prompt_yes_no("하위폴더를 포함할까요?", default=True)
    return True


def collect_markdown_files(target: Path, recursive: bool) -> list[Path]:
    if target.is_file():
        if target.suffix.lower() not in MARKDOWN_EXTS:
            raise RuntimeError(f"Markdown 파일만 지원합니다: {target.suffix}")
        return [target]

    iterator = target.rglob("*") if recursive else target.glob("*")
    files = [
        path
        for path in iterator
        if path.is_file() and path.suffix.lower() in MARKDOWN_EXTS
    ]
    return sorted(files, key=lambda item: str(item).lower())


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
            alignments.append("center")
        elif cell.endswith(":"):
            alignments.append("right")
        else:
            alignments.append("left")
    return alignments


def placeholder_token(index: int) -> str:
    return f"@@HTML_PLACEHOLDER_{index}@@"


def markdown_inline_to_html(text: str) -> str:
    placeholders: list[str] = []

    def keep(fragment: str) -> str:
        placeholders.append(fragment)
        return placeholder_token(len(placeholders) - 1)

    value = nfc(text)
    value = re.sub(
        r"!\[([^\]]*)\]\(([^)]+)\)",
        lambda m: keep(
            f'<a href="{escape(m.group(2).strip(), quote=True)}">[이미지: {escape(m.group(1).strip() or "image")}]</a>'
        ),
        value,
    )
    value = re.sub(
        r"\[([^\]]+)\]\(([^)]+)\)",
        lambda m: keep(
            f'<a href="{escape(m.group(2).strip(), quote=True)}">{escape(m.group(1).strip())}</a>'
        ),
        value,
    )
    value = re.sub(
        r"`([^`]+)`",
        lambda m: keep(f"<code>{escape(m.group(1))}</code>"),
        value,
    )
    value = escape(value)
    value = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", value)
    value = re.sub(r"__(.+?)__", r"<strong>\1</strong>", value)
    value = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"<em>\1</em>", value)
    value = re.sub(r"(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)", r"<em>\1</em>", value)
    value = re.sub(r"~~(.+?)~~", r"<del>\1</del>", value)

    for index, fragment in enumerate(placeholders):
        value = value.replace(placeholder_token(index), fragment)
    return value


def build_markdown_table_html(table_lines: list[str]) -> str:
    rows = [split_markdown_table_row(line) for line in table_lines if line.strip()]
    if not rows:
        return ""

    header_cells = rows[0]
    data_rows = rows[1:]
    alignments = ["left"] * len(header_cells)

    if data_rows and is_markdown_table_separator_row(data_rows[0]):
        alignments = markdown_table_alignments(data_rows[0], len(header_cells))
        data_rows = data_rows[1:]

    column_count = max([len(header_cells)] + [len(row) for row in data_rows] or [len(header_cells)])

    def padded(row: list[str]) -> list[str]:
        return row + ([""] * (column_count - len(row)))

    html_parts = [
        '<table style="border-collapse: collapse; width: 100%; margin: 8px 0 12px;">',
        "<thead><tr>",
    ]
    for index, cell in enumerate(padded(header_cells)):
        align = alignments[index] if index < len(alignments) else "left"
        html_parts.append(
            '<th style="border: 1px solid #b8bcc2; background: #f3f4f6; padding: 6px; text-align: '
            + align
            + ';">'
            + markdown_inline_to_html(cell)
            + "</th>"
        )
    html_parts.append("</tr></thead><tbody>")

    for row in data_rows:
        html_parts.append("<tr>")
        for index, cell in enumerate(padded(row)):
            align = alignments[index] if index < len(alignments) else "left"
            html_parts.append(
                '<td style="border: 1px solid #b8bcc2; padding: 6px; text-align: '
                + align
                + ';">'
                + markdown_inline_to_html(cell)
                + "</td>"
            )
        html_parts.append("</tr>")

    html_parts.append("</tbody></table>")
    return "".join(html_parts)


def markdown_to_html(markdown_text: str) -> str:
    lines = nfc(markdown_text.replace("\r\n", "\n").replace("\r", "\n")).split("\n")
    html_parts: list[str] = []
    paragraph_lines: list[str] = []
    list_type: str | None = None
    index = 0

    def flush_paragraph() -> None:
        if not paragraph_lines:
            return
        text = " ".join(line.strip() for line in paragraph_lines if line.strip())
        if text:
            html_parts.append(f"<p>{markdown_inline_to_html(text)}</p>")
        paragraph_lines.clear()

    def close_list() -> None:
        nonlocal list_type
        if list_type:
            html_parts.append(f"</{list_type}>")
            list_type = None

    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if re.match(r"^(```+|~~~+)", stripped):
            flush_paragraph()
            close_list()
            fence = stripped[:3]
            index += 1
            code_lines: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith(fence):
                code_lines.append(lines[index])
                index += 1
            html_parts.append(f"<pre><code>{escape('\n'.join(code_lines))}</code></pre>")
            index += 1
            continue

        if stripped == "":
            flush_paragraph()
            close_list()
            index += 1
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if heading_match:
            flush_paragraph()
            close_list()
            level = len(heading_match.group(1))
            html_parts.append(f"<h{level}>{markdown_inline_to_html(heading_match.group(2).strip())}</h{level}>")
            index += 1
            continue

        if re.fullmatch(r"(\*\s*){3,}|(-\s*){3,}|(_\s*){3,}", stripped):
            flush_paragraph()
            close_list()
            html_parts.append("<hr />")
            index += 1
            continue

        if stripped.startswith(">"):
            flush_paragraph()
            close_list()
            quote_lines: list[str] = []
            while index < len(lines) and lines[index].strip().startswith(">"):
                quote_lines.append(lines[index].strip()[1:].strip())
                index += 1
            quote_body = " ".join(line for line in quote_lines if line)
            html_parts.append(f"<blockquote><p>{markdown_inline_to_html(quote_body)}</p></blockquote>")
            continue

        list_match = re.match(r"^([-*+])\s+(.*)$", stripped)
        ordered_match = re.match(r"^(\d+)[\.\)]\s+(.*)$", stripped)
        if list_match or ordered_match:
            flush_paragraph()
            current_list_type = "ul" if list_match else "ol"
            if list_type != current_list_type:
                close_list()
                list_type = current_list_type
                html_parts.append(f"<{list_type}>")
            item_text = list_match.group(2) if list_match else ordered_match.group(2)
            html_parts.append(f"<li>{markdown_inline_to_html(item_text.strip())}</li>")
            index += 1
            continue

        if stripped.startswith("|") and "|" in stripped[1:]:
            flush_paragraph()
            close_list()
            table_lines: list[str] = []
            while index < len(lines):
                current = lines[index].rstrip()
                if not current.strip().startswith("|"):
                    break
                table_lines.append(current)
                index += 1
            table_html = build_markdown_table_html(table_lines)
            html_parts.append(table_html or f"<pre><code>{escape(chr(10).join(table_lines))}</code></pre>")
            continue

        paragraph_lines.append(line)
        index += 1

    flush_paragraph()
    close_list()
    if not html_parts:
        html_parts.append("<p>(빈 마크다운 파일)</p>")

    body = "\n".join(html_parts)
    return (
        "<!DOCTYPE html>"
        '<html><head><meta charset="utf-8">'
        "<style>"
        "body { font-family: 'Batang', 'AppleMyungjo', 'Times New Roman', serif; line-height: 1.6; }"
        "h1, h2, h3, h4, h5, h6 { margin: 1em 0 0.4em; }"
        "p, ul, ol, blockquote, pre, table { margin: 0 0 0.9em; }"
        "blockquote { border-left: 4px solid #d1d5db; padding-left: 12px; color: #555; }"
        "pre { background: #f5f5f5; padding: 10px; white-space: pre-wrap; }"
        "code { background: #f5f5f5; padding: 1px 3px; }"
        "</style></head><body>"
        + body
        + "</body></html>"
    )


class GoogleDocUploader:
    def __init__(self, credentials_path: Path | None, token_path: Path, interactive: bool):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.interactive = interactive
        self._session: AuthorizedSession | None = None
        self._user_email: str | None = None

    def _ensure_google_packages(self) -> None:
        if AuthorizedSession is None or Request is None or Credentials is None or InstalledAppFlow is None:
            raise RuntimeError(
                "Google 업로드용 패키지가 설치되어 있지 않습니다. "
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

    @property
    def user_email(self) -> str:
        if self._user_email is None:
            response = self.session.get("https://www.googleapis.com/oauth2/v2/userinfo", timeout=60)
            if response.ok:
                self._user_email = response.json().get("email", "")
            else:
                self._user_email = ""
        return self._user_email

    def create_google_doc(self, title: str, html_body: str) -> dict[str, Any]:
        boundary = f"===============_{uuid.uuid4().hex}"
        metadata = json.dumps(
            {
                "name": title,
                "mimeType": "application/vnd.google-apps.document",
            },
            ensure_ascii=False,
        )
        payload = (
            f"--{boundary}\r\n"
            "Content-Type: application/json; charset=UTF-8\r\n\r\n"
            f"{metadata}\r\n"
            f"--{boundary}\r\n"
            "Content-Type: text/html; charset=UTF-8\r\n\r\n"
            f"{html_body}\r\n"
            f"--{boundary}--\r\n"
        ).encode("utf-8")

        response = self.session.post(
            "https://www.googleapis.com/upload/drive/v3/files",
            params={
                "uploadType": "multipart",
                "fields": "id,name,webViewLink,resourceKey",
            },
            headers={"Content-Type": f"multipart/related; boundary={boundary}"},
            data=payload,
            timeout=300,
        )
        if not response.ok:
            message = response.text.strip()
            raise RuntimeError(f"Google Docs 생성 실패 ({response.status_code}): {message}")
        return response.json()


def write_gdoc_shortcut(shortcut_path: Path, doc_id: str, resource_key: str, email: str) -> Path:
    payload = {
        "": WARNING_TEXT,
        "doc_id": doc_id,
        "resource_key": resource_key,
        "email": email,
    }
    shortcut_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return shortcut_path


def convert_markdown_file(source_path: Path, uploader: GoogleDocUploader) -> UploadResult:
    markdown_text = read_text_file(source_path).strip("\ufeff")
    html_body = markdown_to_html(markdown_text)
    created = uploader.create_google_doc(source_path.stem, html_body)
    shortcut_path = source_path.with_suffix(".gdoc")
    write_gdoc_shortcut(
        shortcut_path=shortcut_path,
        doc_id=created["id"],
        resource_key=created.get("resourceKey", "") or "",
        email=uploader.user_email,
    )
    return UploadResult(
        source_path=source_path,
        shortcut_path=shortcut_path,
        doc_id=created["id"],
        web_view_link=created.get("webViewLink", f"https://docs.google.com/document/d/{created['id']}/edit"),
    )


def main() -> int:
    args = parse_args()
    interactive = not args.non_interactive and sys.stdin.isatty()

    try:
        if args.target is not None:
            target = args.target.expanduser().resolve()
        elif interactive:
            target = prompt_path("변환할 Markdown 파일 또는 폴더 경로", must_exist=True)
        else:
            raise RuntimeError("--target 이 필요합니다.")

        if not target.exists():
            raise RuntimeError(f"입력 경로를 찾을 수 없습니다: {target}")
        if not target.is_file() and not target.is_dir():
            raise RuntimeError(f"파일 또는 폴더 경로를 입력해 주세요: {target}")

        include_subfolders = resolve_include_subfolders(args.include_subfolders, interactive, target)
        credentials_path = prompt_credentials_path(discover_credentials_file(args.credentials), interactive)
        token_path = args.token.expanduser().resolve() if args.token else default_token_path()

        markdown_files = collect_markdown_files(target, recursive=include_subfolders)
        if not markdown_files:
            raise RuntimeError(f"Markdown 파일이 없습니다: {target}")

        uploader = GoogleDocUploader(
            credentials_path=credentials_path,
            token_path=token_path,
            interactive=interactive,
        )

        failures = 0
        print(f"TARGET\t{target}")
        print(f"RECURSIVE\t{'y' if include_subfolders else 'n'}")
        for source_path in markdown_files:
            try:
                result = convert_markdown_file(source_path, uploader)
                print(
                    f"OK\t{source_path}\t-> {result.shortcut_path.name}\t{result.web_view_link}"
                )
            except Exception as exc:  # pragma: no cover - runtime dependent
                failures += 1
                print(f"FAIL\t{source_path}\t{type(exc).__name__}: {exc}")

        return 1 if failures else 0
    except KeyboardInterrupt:  # pragma: no cover - interactive runtime
        print("\n사용자가 작업을 취소했습니다.")
        return 130
    except Exception as exc:
        print(f"ERROR\t{type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
