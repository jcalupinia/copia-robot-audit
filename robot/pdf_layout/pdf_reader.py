from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

try:
    import fitz  # type: ignore
except Exception:
    fitz = None

try:
    from PIL import Image  # type: ignore
except Exception:
    Image = None

try:
    import pytesseract  # type: ignore
except Exception:
    pytesseract = None


@dataclass
class WordBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    page_index: int
    source: str = "layout"


@dataclass
class LineBox:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    page_index: int
    source: str = "layout"
    words: list[WordBox] = field(default_factory=list)
    region: str = "center"


@dataclass
class PageLayout:
    page_index: int
    width: float
    height: float
    lines: list[LineBox]
    used_ocr: bool = False


@dataclass
class DocumentLayout:
    pdf_path: Path
    pages: list[PageLayout]
    used_ocr: bool = False
    engine: str = "fitz"


def _assign_region(x0: float, y0: float, width: float, height: float) -> str:
    if y0 <= height * 0.34:
        return "top_left" if x0 <= width * 0.5 else "top_right"
    if y0 >= height * 0.72:
        return "bottom"
    return "center"


def _group_words_into_lines(words: list[WordBox], page_width: float, page_height: float) -> list[LineBox]:
    if not words:
        return []
    words = sorted(words, key=lambda w: (w.y0, w.x0))
    avg_height = sum(max(1.0, w.y1 - w.y0) for w in words) / max(1, len(words))
    tolerance = max(2.0, min(6.0, avg_height * 0.55))

    raw_groups: list[list[WordBox]] = []
    current: list[WordBox] = []
    current_y = None
    for word in words:
        if not current:
            current = [word]
            current_y = word.y0
            continue
        if current_y is not None and abs(word.y0 - current_y) <= tolerance:
            current.append(word)
            current_y = (current_y + word.y0) / 2
        else:
            raw_groups.append(sorted(current, key=lambda w: w.x0))
            current = [word]
            current_y = word.y0
    if current:
        raw_groups.append(sorted(current, key=lambda w: w.x0))

    lines: list[LineBox] = []
    for group in raw_groups:
        text = " ".join(w.text for w in group if w.text).strip()
        if not text:
            continue
        x0 = min(w.x0 for w in group)
        y0 = min(w.y0 for w in group)
        x1 = max(w.x1 for w in group)
        y1 = max(w.y1 for w in group)
        lines.append(
            LineBox(
                text=text,
                x0=x0,
                y0=y0,
                x1=x1,
                y1=y1,
                page_index=group[0].page_index,
                source=group[0].source,
                words=group,
                region=_assign_region(x0, y0, page_width, page_height),
            )
        )
    return lines


def _page_has_useful_text(page, min_chars: int = 24, min_words: int = 6) -> bool:
    if fitz is None:
        return False
    try:
        text = page.get_text("text") or ""
        words = page.get_text("words") or []
    except Exception:
        return False
    cleaned = text.strip()
    return len(cleaned) >= min_chars and len(words) >= min_words


def _extract_words_fitz(page, page_index: int) -> list[WordBox]:
    if fitz is None:
        return []
    try:
        raw_words = page.get_text("words") or []
    except Exception:
        return []
    words: list[WordBox] = []
    for item in raw_words:
        if len(item) < 5:
            continue
        x0, y0, x1, y1, text = item[:5]
        text = str(text or "").strip()
        if not text:
            continue
        words.append(WordBox(text=text, x0=float(x0), y0=float(y0), x1=float(x1), y1=float(y1), page_index=page_index))
    return words


def _extract_words_ocr(page, page_index: int) -> list[WordBox]:
    if fitz is None or pytesseract is None or Image is None:
        return []
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        image = Image.open(BytesIO(pix.tobytes("png")))
        ocr_data = pytesseract.image_to_data(
            image,
            output_type=pytesseract.Output.DICT,
            lang="spa+eng",
        )
    except Exception:
        return []

    scale_x = page.rect.width / max(1, image.width)
    scale_y = page.rect.height / max(1, image.height)
    words: list[WordBox] = []
    total = len(ocr_data.get("text", []))
    for idx in range(total):
        text = str(ocr_data["text"][idx] or "").strip()
        if not text:
            continue
        try:
            conf = float(ocr_data["conf"][idx])
        except Exception:
            conf = -1
        if conf < 20:
            continue
        x = float(ocr_data["left"][idx]) * scale_x
        y = float(ocr_data["top"][idx]) * scale_y
        w = float(ocr_data["width"][idx]) * scale_x
        h = float(ocr_data["height"][idx]) * scale_y
        words.append(
            WordBox(
                text=text,
                x0=x,
                y0=y,
                x1=x + w,
                y1=y + h,
                page_index=page_index,
                source="ocr",
            )
        )
    return words


def read_pdf_layout(pdf_path: str | Path) -> DocumentLayout:
    pdf_path = Path(pdf_path)
    if fitz is None:
        return DocumentLayout(pdf_path=pdf_path, pages=[], used_ocr=False, engine="unavailable")

    pages: list[PageLayout] = []
    used_ocr = False
    try:
        document = fitz.open(pdf_path)
    except Exception:
        return DocumentLayout(pdf_path=pdf_path, pages=[], used_ocr=False, engine="fitz")

    with document:
        for page_index in range(document.page_count):
            page = document.load_page(page_index)
            width = float(page.rect.width)
            height = float(page.rect.height)
            page_uses_ocr = False
            if _page_has_useful_text(page):
                words = _extract_words_fitz(page, page_index)
            else:
                words = _extract_words_ocr(page, page_index)
                page_uses_ocr = bool(words)
                used_ocr = used_ocr or page_uses_ocr
                if not words:
                    words = _extract_words_fitz(page, page_index)
            lines = _group_words_into_lines(words, width, height)
            pages.append(
                PageLayout(
                    page_index=page_index,
                    width=width,
                    height=height,
                    lines=lines,
                    used_ocr=page_uses_ocr,
                )
            )
    return DocumentLayout(pdf_path=pdf_path, pages=pages, used_ocr=used_ocr, engine="fitz")
