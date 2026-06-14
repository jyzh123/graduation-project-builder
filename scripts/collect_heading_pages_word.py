from __future__ import annotations

import argparse
import json
import re
import subprocess
import time
from pathlib import Path

try:
    import pythoncom  # type: ignore
    import pywintypes  # type: ignore
    import win32com.client  # type: ignore
except ModuleNotFoundError:
    pythoncom = None  # type: ignore
    pywintypes = None  # type: ignore
    win32com = None  # type: ignore


SCRIPT_DIR = Path(__file__).resolve().parent


def is_retryable_com_error(exc: pywintypes.com_error) -> bool:
    if not exc.args:
        return False
    try:
        hresult = int(exc.args[0])
    except Exception:
        hresult = 0
    if hresult == -2147418111:
        return True
    lowered = " ".join(str(part).lower() for part in exc.args if part is not None)
    return "call was rejected by callee" in lowered or "拒绝接收呼叫" in lowered


def com_call_with_retry(func, *, retries: int = 120, delay_s: float = 0.5):
    last_error: Exception | None = None
    for _ in range(retries):
        try:
            return func()
        except pywintypes.com_error as exc:
            if not is_retryable_com_error(exc):
                raise
            last_error = exc
            time.sleep(delay_s)
    if last_error:
        raise last_error
    raise RuntimeError("unexpected COM retry state")


def normalize_for_match(text: str) -> str:
    return re.sub(r"[\s\u25a1]+", "", text or "").lower()


def normalize_heading_dots(text: str) -> str:
    return re.sub(r"(?<=\d)[\s\u25a1]*[\.\uff0e][\s\u25a1]*(?=\d)", ".", text or "")


def collect_docx_heading_rows(input_path: Path) -> list[dict[str, object]]:
    from docx import Document  # type: ignore

    doc = Document(input_path)
    rows: list[dict[str, object]] = []
    for paragraph in doc.paragraphs:
        text = re.sub(r"[\r\a]+", "", paragraph.text or "").strip()
        if not text or "\t" in text:
            continue
        if normalize_for_match(text) in {normalize_for_match("摘    要"), normalize_for_match("Abstract"), normalize_for_match("目    录")}:
            continue
        style_name = getattr(getattr(paragraph, "style", None), "name", "") or ""
        if style_name.lower().startswith("toc"):
            continue
        level = heading_level_from_style(style_name) or heading_level_from_text(text)
        if level is None:
            continue
        rows.append({"text": text, "page": None, "level": level, "style": style_name})
    return rows


def export_pdf_for_fallback(input_path: Path, pdf_path: Path) -> bool:
    try:
        if pdf_path.exists():
            pdf_path.unlink()
    except OSError:
        pass
    proc = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(SCRIPT_DIR / "wps_export_pdf.ps1"),
            "-InputDocx",
            str(input_path),
            "-OutputPdf",
            str(pdf_path),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=900,
    )
    return proc.returncode == 0 and pdf_path.exists()


def assign_pages_from_pdf(rows: list[dict[str, object]], pdf_path: Path) -> list[dict[str, object]]:
    try:
        import fitz  # type: ignore
    except ModuleNotFoundError:
        return rows
    pdf = fitz.open(pdf_path)
    page_texts = []
    for page_index in range(len(pdf)):
        text = pdf.load_page(page_index).get_text("text")
        lines = [normalize_for_match(line) for line in text.splitlines() if normalize_for_match(line)]
        page_texts.append((page_index + 1, normalize_for_match(text), lines))
    last_page = 1
    for row in rows:
        needle = normalize_for_match(str(row.get("text") or ""))
        if not needle:
            continue
        matched_page = None
        for page_number, _page_text, page_lines in reversed(page_texts):
            if page_number < last_page:
                continue
            if needle in page_lines:
                matched_page = page_number
                break
        if matched_page is None:
            for page_number, _page_text, page_lines in reversed(page_texts):
                if page_number < last_page:
                    continue
                if needle in page_lines:
                    matched_page = page_number
                    break
        if matched_page is None:
            for page_number, page_text, _page_lines in reversed(page_texts):
                if page_number < last_page:
                    continue
                if needle in page_text:
                    matched_page = page_number
                    break
        if matched_page is None:
            for page_number, page_text, _page_lines in reversed(page_texts):
                if page_number < last_page:
                    continue
                if needle in page_text:
                    matched_page = page_number
                    break
        if matched_page is not None:
            row["page"] = matched_page
            last_page = matched_page
    return rows


def collect_heading_pages_fallback(input_path: Path, output_path: Path) -> list[dict[str, object]]:
    rows = collect_docx_heading_rows(input_path)
    pdf_path = output_path.with_suffix(".pdf")
    if export_pdf_for_fallback(input_path, pdf_path):
        rows = assign_pages_from_pdf(rows, pdf_path)
    return [row for row in rows if row.get("page") is not None]


def collect_heading_pages_from_rendered_pdf(input_path: Path, pdf_path: Path) -> list[dict[str, object]]:
    rows = collect_docx_heading_rows(input_path)
    if not pdf_path.exists():
        return []
    return [row for row in assign_pages_from_pdf(rows, pdf_path) if row.get("page") is not None]


def heading_level_from_style(style_name: str) -> int | None:
    if re.search(r"Heading 1|标题 1", style_name):
        return 1
    if re.search(r"Heading 2|标题 2", style_name):
        return 2
    if re.search(r"Heading 3|标题 3", style_name):
        return 3
    return None


def heading_level_from_text(text: str) -> int | None:
    stripped = re.sub(r"^[\s\u25a1]+", "", (text or "").strip())
    if "\u2026" in stripped:
        return None
    normalized = re.sub(r"[\s\u25a1]+", "", stripped).lower()
    if normalized in {"参考文献", "致谢", "references", "acknowledgements", "acknowledgments"}:
        return 1
    sep = r"[\s\u25a1]+"
    normalized_dots = normalize_heading_dots(stripped)
    if re.match(rf"^\d{{1,2}}{sep}\S", stripped) or re.match(r"^第[0-9一二三四五六七八九十]+章", stripped):
        return 1
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}{sep}\S", normalized_dots):
        return 2
    if re.match(rf"^\d{{1,2}}\.\d{{1,2}}\.\d{{1,2}}{sep}\S", normalized_dots):
        return 3
    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument(
        "--rendered-pdf",
        default="",
        help="Optional already-rendered PDF to use as the page authority instead of Word COM.",
    )
    parser.add_argument(
        "--force-fallback",
        action="store_true",
        help="Skip Word COM and collect heading pages from --rendered-pdf or the fallback renderer.",
    )
    args = parser.parse_args()

    input_path = str(Path(args.input).resolve())
    output_path = Path(args.output).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    rendered_pdf = Path(args.rendered_pdf).resolve() if args.rendered_pdf else None

    if args.force_fallback:
        rows = (
            collect_heading_pages_from_rendered_pdf(Path(input_path), rendered_pdf)
            if rendered_pdf is not None
            else collect_heading_pages_fallback(Path(input_path), output_path)
        )
        output_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        print(output_path)
        return 0 if rows else 1

    if pythoncom is None:
        rows = (
            collect_heading_pages_from_rendered_pdf(Path(input_path), rendered_pdf)
            if rendered_pdf is not None
            else collect_heading_pages_fallback(Path(input_path), output_path)
        )
        output_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        print(output_path)
        return 0 if rows else 1

    pythoncom.CoInitialize()
    word = None
    doc = None
    rows: list[dict[str, object]] = []
    try:
        word = com_call_with_retry(lambda: win32com.client.DispatchEx("Word.Application"))
        word.Visible = False
        word.DisplayAlerts = 0
        doc = com_call_with_retry(lambda: word.Documents.Open(input_path, False, True))

        para_count = int(com_call_with_retry(lambda: doc.Paragraphs.Count))
        for i in range(1, para_count + 1):
            try:
                para = com_call_with_retry(lambda idx=i: doc.Paragraphs.Item(idx))
            except Exception:
                continue
            try:
                para_range = para.Range
                text = re.sub(r"[\r\a]+", "", str(para_range.Text)).strip()
            except Exception:
                continue
            if not text:
                continue
            if normalize_for_match(text) in {normalize_for_match("摘    要"), normalize_for_match("Abstract"), normalize_for_match("目    录")}:
                continue
            style_name = ""
            try:
                style_name = str(para.Range.Style.NameLocal)
            except Exception:
                try:
                    style_name = str(para.Range.Style)
                except Exception:
                    style_name = ""
            if style_name.lower().startswith("toc") or "\t" in text:
                continue
            level = heading_level_from_style(style_name) or heading_level_from_text(text)
            if level is None:
                continue
            try:
                page = int(com_call_with_retry(lambda r=para_range: r.Information(3)))
            except Exception:
                continue
            rows.append(
                {
                    "text": text,
                    "page": page,
                    "level": level,
                    "style": style_name,
                }
            )
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not rows:
        rows = collect_heading_pages_fallback(Path(input_path), output_path)

    output_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output_path)
    return 0 if rows else 1


if __name__ == "__main__":
    raise SystemExit(main())
