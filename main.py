import io
import os
import json
import secrets
import tempfile
from datetime import datetime, timezone

from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont
import docx
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from pptx import Presentation

from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

MAX_FILE_SIZE = 5 * 1024 * 1024  # 5MB
ALLOWED_DOC_EXTENSIONS = {".docx", ".txt", ".pptx"}
API_KEYS_FILE = os.path.join(tempfile.gettempdir(), "api_keys.json")

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(status_code=500, content={"detail": "Something went wrong on our end. Please try again."})


@app.get("/", response_class=HTMLResponse)
def serve_homepage():
    with open("static/index.html", "r") as f:
        return f.read()


# ---------------------------------------------------------------------------
# Small shared helpers
# ---------------------------------------------------------------------------

def sanitize_text(text: str) -> str:
    """fpdf2's core fonts only support latin-1, so we replace anything else."""
    return (text or "").encode("latin-1", "replace").decode("latin-1")


def load_watermark_font(size: int):
    font_paths = ["arial.ttf", "Arial.ttf", "C:\\Windows\\Fonts\\arial.ttf", "DejaVuSans-Bold.ttf"]
    for path in font_paths:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def iter_block_items(document):
    """Walk a docx body in document order, yielding Paragraph/Table objects
    so tables show up in the right place instead of being handled separately."""
    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield DocxParagraph(child, document)
        elif child.tag.endswith("}tbl"):
            yield DocxTable(child, document)


def heading_level(style_name: str):
    name = style_name.lower().strip()
    if name.startswith("heading"):
        digits = "".join(ch for ch in name if ch.isdigit())
        return int(digits) if digits else 1
    if name in ("title",):
        return 1
    return 0


def heading_size(level: int) -> int:
    return {1: 18, 2: 16, 3: 14}.get(level, 12)


# ---------------------------------------------------------------------------
# Compress image (unchanged)
# ---------------------------------------------------------------------------

@app.post("/compress-image")
@limiter.limit("10/minute")
async def compress_image(request: Request, file: UploadFile = File(...), quality: int = 50):
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image.")
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large. Max size is 5MB.")
    try:
        image = Image.open(io.BytesIO(contents))
        image = image.convert("RGB")
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid or corrupted image file.")
    temp_path = os.path.join(tempfile.gettempdir(), "compressed_output.jpg")
    image.save(temp_path, format="JPEG", quality=quality, optimize=True)
    return FileResponse(temp_path, media_type="image/jpeg", filename="compressed.jpg")


# ---------------------------------------------------------------------------
# Document -> PDF, with real formatting (bold/italic/underline, headings,
# bullet lists, and basic tables for .docx; titles + indented bullets per
# slide for .pptx; plain paragraphs for .txt)
# ---------------------------------------------------------------------------

def render_docx_paragraph(pdf: FPDF, para: DocxParagraph):
    style_name = para.style.name if para.style else ""
    level = heading_level(style_name)
    is_list = "list" in style_name.lower() or "bullet" in style_name.lower()
    text = para.text.strip()
    if not text:
        return

    indent = 6 if is_list else 0
    pdf.set_x(pdf.l_margin + indent)
    if is_list:
        pdf.set_font("Helvetica", "", 12)
        pdf.write(7, "-  ")

    runs = para.runs if para.runs else None
    if not runs:
        size = heading_size(level) if level else 12
        style = "B" if level else ""
        pdf.set_font("Helvetica", style, size)
        pdf.write(9 if level else 7, sanitize_text(text))
    else:
        for run in runs:
            run_text = sanitize_text(run.text)
            if not run_text:
                continue
            style = ""
            if run.bold or level:
                style += "B"
            if run.italic:
                style += "I"
            if run.underline:
                style += "U"
            size = heading_size(level) if level else 12
            pdf.set_font("Helvetica", style, size)
            pdf.write(9 if level else 7, run_text)
    pdf.ln(10 if level else 7)


def render_docx_table(pdf: FPDF, table: DocxTable):
    col_count = max(len(table.columns), 1)
    usable_width = pdf.w - pdf.l_margin - pdf.r_margin
    col_width = usable_width / col_count
    row_height = 7

    pdf.ln(2)
    for row_index, row in enumerate(table.rows):
        x_start = pdf.get_x()
        y_start = pdf.get_y()
        max_y = y_start
        is_header = row_index == 0
        for i, cell in enumerate(row.cells):
            cell_text = sanitize_text(cell.text.strip())
            pdf.set_xy(x_start + i * col_width, y_start)
            pdf.set_font("Helvetica", "B" if is_header else "", 10)
            pdf.multi_cell(col_width, row_height, cell_text, border=1)
            if pdf.get_y() > max_y:
                max_y = pdf.get_y()
        pdf.set_xy(x_start, max_y)
    pdf.ln(4)


def build_pdf_from_docx(contents: bytes) -> FPDF:
    document = docx.Document(io.BytesIO(contents))
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    found_any = False
    for block in iter_block_items(document):
        if isinstance(block, DocxParagraph):
            if block.text.strip():
                found_any = True
            render_docx_paragraph(pdf, block)
        elif isinstance(block, DocxTable):
            if block.rows:
                found_any = True
            render_docx_table(pdf, block)

    if not found_any:
        raise ValueError("empty document")
    return pdf


def build_pdf_from_pptx(contents: bytes) -> FPDF:
    prs = Presentation(io.BytesIO(contents))
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    found_any = False
    for i, slide in enumerate(prs.slides, start=1):
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.multi_cell(0, 10, sanitize_text(f"Slide {i}"))
        pdf.ln(2)

        title_shape = slide.shapes.title
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            is_title = shape == title_shape
            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs).strip()
                if not text:
                    continue
                found_any = True
                indent = 0 if is_title else min(para.level, 4) * 6
                pdf.set_x(pdf.l_margin + indent)
                size = 15 if is_title else 12
                style = "B" if is_title else ""
                bullet = "" if is_title else "-  "
                pdf.set_font("Helvetica", style, size)
                pdf.multi_cell(0, 8, sanitize_text(bullet + text))
        pdf.ln(4)

    if not found_any:
        raise ValueError("empty document")
    return pdf


def build_pdf_from_txt(contents: bytes) -> FPDF:
    text_content = contents.decode("utf-8", errors="ignore")
    lines = [line.strip() for line in text_content.splitlines() if line.strip()]
    if not lines:
        raise ValueError("empty document")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "", 12)
    for line in lines:
        pdf.multi_cell(0, 8, sanitize_text(line))
        pdf.ln(1)
    return pdf


@app.post("/convert-document-to-pdf")
@limiter.limit("10/minute")
async def convert_document_to_pdf(request: Request, file: UploadFile = File(...)):
    filename = file.filename or ""
    ext = os.path.splitext(filename)[1].lower()

    if ext not in ALLOWED_DOC_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Unsupported file type. Please upload a .docx, .txt, or .pptx file.")

    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large. Max size is 5MB.")

    try:
        if ext == ".docx":
            pdf = build_pdf_from_docx(contents)
        elif ext == ".pptx":
            pdf = build_pdf_from_pptx(contents)
        else:
            pdf = build_pdf_from_txt(contents)
    except ValueError:
        raise HTTPException(status_code=400, detail="The document appears to be empty.")
    except Exception:
        raise HTTPException(status_code=400, detail="Could not read this document. It may be corrupted or invalid.")

    temp_path = os.path.join(tempfile.gettempdir(), "document_converted.pdf")
    pdf.output(temp_path)
    return FileResponse(temp_path, media_type="application/pdf", filename="converted_document.pdf")


# ---------------------------------------------------------------------------
# Image -> PDF (unchanged)
# ---------------------------------------------------------------------------

@app.post("/convert-image-to-pdf")
@limiter.limit("10/minute")
async def convert_image_to_pdf(request: Request, file: UploadFile = File(...)):
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image.")
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large. Max size is 5MB.")
    try:
        image = Image.open(io.BytesIO(contents))
        image = image.convert("RGB")
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid or corrupted image file.")
    temp_path = os.path.join(tempfile.gettempdir(), "converted.pdf")
    image.save(temp_path)
    return FileResponse(temp_path, media_type="application/pdf", filename="converted.pdf")


# ---------------------------------------------------------------------------
# Watermark image — single placement OR tiled/repeated pattern
# ---------------------------------------------------------------------------

@app.post("/watermark-image")
@limiter.limit("10/minute")
async def watermark_image(
    request: Request,
    file: UploadFile = File(...),
    text: str = "SAMPLE",
    mode: str = "single",          # "single" or "tiled"
    position: str = "bottom-right",  # used only when mode == "single"
    opacity: int = 180,
    angle: int = -30,              # used only when mode == "tiled"
):
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image.")
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large. Max size is 5MB.")
    try:
        image = Image.open(io.BytesIO(contents)).convert("RGBA")
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid or corrupted image file.")

    safe_opacity = max(0, min(255, opacity))
    overlay = Image.new("RGBA", image.size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(overlay)

    if mode == "tiled":
        tile_font_size = max(18, image.width // 20)
        tile_font = load_watermark_font(tile_font_size)

        bbox = draw.textbbox((0, 0), text, font=tile_font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        pad = 40
        tile = Image.new("RGBA", (tw + pad * 2, th + pad * 2), (255, 255, 255, 0))
        tile_draw = ImageDraw.Draw(tile)
        tile_draw.text((pad, pad), text, font=tile_font, fill=(255, 255, 255, safe_opacity))
        rotated_tile = tile.rotate(angle, expand=True)
        tile_w, tile_h = rotated_tile.size

        spacing_x = tile_w + 30
        spacing_y = tile_h + 30
        for y in range(-tile_h, image.height + tile_h, spacing_y):
            for x in range(-tile_w, image.width + tile_w, spacing_x):
                overlay.paste(rotated_tile, (x, y), rotated_tile)
    else:
        font_size = max(24, image.width // 15)
        font = load_watermark_font(font_size)
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        margin = 20

        positions = {
            "top-left": (margin, margin),
            "top-right": (image.width - text_width - margin, margin),
            "bottom-left": (margin, image.height - text_height - margin),
            "bottom-right": (image.width - text_width - margin, image.height - text_height - margin),
            "center": ((image.width - text_width) // 2, (image.height - text_height) // 2),
        }
        xy = positions.get(position, positions["bottom-right"])
        draw.text(xy, text, font=font, fill=(255, 255, 255, safe_opacity), stroke_width=3, stroke_fill=(0, 0, 0, safe_opacity))

    watermarked = Image.alpha_composite(image, overlay).convert("RGB")
    temp_path = os.path.join(tempfile.gettempdir(), "watermarked.jpg")
    watermarked.save(temp_path, format="JPEG")
    return FileResponse(temp_path, media_type="image/jpeg", filename="watermarked.jpg")


# ---------------------------------------------------------------------------
# Self-serve API key generation (informational only — not enforced on
# endpoints). Keys are logged to a JSON file so there's a record of who
# generated one, even though nothing currently checks for it.
# ---------------------------------------------------------------------------

def load_issued_keys():
    if not os.path.exists(API_KEYS_FILE):
        return []
    try:
        with open(API_KEYS_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return []


def save_issued_keys(keys):
    with open(API_KEYS_FILE, "w") as f:
        json.dump(keys, f, indent=2)


@app.post("/generate-key")
@limiter.limit("5/minute")
async def generate_key(request: Request, email: str = ""):
    clean_email = (email or "").strip()[:100]
    if "@" not in clean_email or "." not in clean_email.split("@")[-1] or len(clean_email) < 5:
        raise HTTPException(status_code=400, detail="Please enter a valid email address.")

    new_key = f"fpapi_{secrets.token_hex(16)}"

    keys = load_issued_keys()
    keys.append({
        "email": clean_email,
        "key": new_key,
        "date": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
    })
    save_issued_keys(keys)

    return {"email": clean_email, "api_key": new_key}