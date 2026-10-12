"""Convert uploaded documents / slide decks into video-ready material.

* PDF   -> one PNG per page (pdf2image/poppler) + extracted text
* PPT(X)-> PDF via LibreOffice then PNG per slide; if LibreOffice is missing,
           embedded pictures are extracted, and finally plain text slides are
           drawn with Pillow so that the user always gets something usable
* DOCX/TXT/MD -> extracted text only (optionally fed to the script)

Everything here is additive: the video pipeline just receives ordinary image
files through the existing "local" material source.
"""

import os
import shutil
import subprocess
import tempfile
import textwrap
from dataclasses import dataclass, field
from uuid import uuid4

from loguru import logger
from PIL import Image, ImageDraw, ImageFont

DOCUMENT_TEXT_EXTENSIONS = (".txt", ".md", ".docx")
DOCUMENT_PAGE_EXTENSIONS = (".pdf", ".ppt", ".pptx")
SUPPORTED_DOCUMENT_EXTENSIONS = DOCUMENT_TEXT_EXTENSIONS + DOCUMENT_PAGE_EXTENSIONS
MAX_PAGES = 30
MAX_TEXT_CHARS = 20000
SLIDE_SIZE = (1280, 720)


class DocumentConversionError(RuntimeError):
    pass


@dataclass
class DocumentResult:
    pages: list = field(default_factory=list)  # PNG paths
    text: str = ""
    notes: list = field(default_factory=list)  # user-facing hints/fallbacks


def _read_text_file(path):
    raw = open(path, "rb").read(5 * 1024 * 1024)
    for enc in ("utf-8-sig", "gb18030", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return ""


def extract_text(path, ext):
    ext = ext.lower()
    try:
        if ext in (".txt", ".md"):
            return _read_text_file(path)[:MAX_TEXT_CHARS]
        if ext == ".docx":
            import docx

            d = docx.Document(path)
            return "\n".join(p.text for p in d.paragraphs if p.text.strip())[
                :MAX_TEXT_CHARS
            ]
        if ext == ".pdf":
            from pypdf import PdfReader

            parts = []
            for page in PdfReader(path).pages[:MAX_PAGES]:
                parts.append((page.extract_text() or "").strip())
            return "\n".join(p for p in parts if p)[:MAX_TEXT_CHARS]
        if ext == ".pptx":
            from pptx import Presentation

            parts = []
            for slide in Presentation(path).slides:
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.text_frame.text.strip():
                        parts.append(shape.text_frame.text.strip())
            return "\n".join(parts)[:MAX_TEXT_CHARS]
    except Exception as exc:  # corrupt files must not crash the page
        logger.warning(f"extract_text failed: {path}: {exc}")
    return ""


def _soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def _pdf_to_images(pdf_path, out_dir, dpi=110):
    if not shutil.which("pdftoppm"):
        raise DocumentConversionError("系统缺少 poppler（pdftoppm），无法把 PDF 转为图片")
    from pdf2image import convert_from_path

    images = convert_from_path(pdf_path, dpi=dpi, last_page=MAX_PAGES)
    paths = []
    for image in images:
        p = os.path.join(out_dir, f"{uuid4().hex}.png")
        image.convert("RGB").save(p, "PNG")
        paths.append(p)
    return paths


def _office_to_pdf(path, work_dir):
    exe = _soffice()
    if not exe:
        return None
    try:
        subprocess.run(
            [exe, "--headless", "--convert-to", "pdf", "--outdir", work_dir, path],
            check=True,
            capture_output=True,
            timeout=180,
            env={**os.environ, "HOME": work_dir},
        )
    except Exception as exc:
        logger.warning(f"libreoffice conversion failed: {exc}")
        return None
    pdf = os.path.join(work_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    return pdf if os.path.isfile(pdf) else None


def _font(size):
    for candidate in (
        "resource/fonts/STHeitiMedium.ttc",
        "resource/fonts/MicrosoftYaHeiBold.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ):
        if os.path.isfile(candidate):
            try:
                return ImageFont.truetype(candidate, size)
            except Exception:
                continue
    return ImageFont.load_default()


def _text_slide(title, body, out_dir):
    img = Image.new("RGB", SLIDE_SIZE, (30, 41, 59))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 18, SLIDE_SIZE[1]], fill=(79, 70, 229))
    y = 80
    if title:
        d.text((80, y), title[:40], font=_font(56), fill=(255, 255, 255))
        y += 110
    for line in textwrap.wrap(body, width=34)[:9]:
        d.text((80, y), line, font=_font(36), fill=(226, 232, 240))
        y += 56
    p = os.path.join(out_dir, f"{uuid4().hex}.png")
    img.save(p, "PNG")
    return p


def _pptx_fallback(path, out_dir, result):
    """No LibreOffice: use embedded pictures, else draw text slides."""
    from pptx import Presentation

    prs = Presentation(path)
    for slide in list(prs.slides)[:MAX_PAGES]:
        picture = None
        for shape in slide.shapes:
            if shape.shape_type == 13:  # PICTURE
                picture = shape
                break
        title = body = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                if not title:
                    title = shape.text_frame.text.strip().split("\n")[0]
                else:
                    body += shape.text_frame.text.strip() + " "
        if picture is not None:
            try:
                import io

                im = Image.open(io.BytesIO(picture.image.blob)).convert("RGB")
                p = os.path.join(out_dir, f"{uuid4().hex}.png")
                im.save(p, "PNG")
                result.pages.append(p)
                continue
            except Exception:
                pass
        if title or body:
            result.pages.append(_text_slide(title, body, out_dir))


def convert_document(path, ext, out_dir):
    """Return DocumentResult. Raises DocumentConversionError when nothing usable."""
    ext = ext.lower()
    os.makedirs(out_dir, exist_ok=True)
    result = DocumentResult(text=extract_text(path, ext))
    if ext in DOCUMENT_TEXT_EXTENSIONS:
        if not result.text:
            raise DocumentConversionError("没有从文档中读取到文字")
        return result
    with tempfile.TemporaryDirectory() as work:
        if ext == ".pdf":
            result.pages = _pdf_to_images(path, out_dir)
        else:  # ppt / pptx
            pdf = _office_to_pdf(path, work)
            if pdf:
                result.pages = _pdf_to_images(pdf, out_dir)
            elif ext == ".pptx":
                result.notes.append(
                    "未检测到 LibreOffice，已改用简化方式：提取幻灯片内的图片或文字生成画面（版式会与原稿不同）"
                )
                _pptx_fallback(path, out_dir, result)
            else:
                raise DocumentConversionError(
                    "旧版 .ppt 需要安装 LibreOffice 才能转换，请先另存为 .pptx"
                )
    if not result.pages:
        raise DocumentConversionError("没有生成任何画面，文件可能为空或已损坏")
    return result
