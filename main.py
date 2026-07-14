from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fpdf import FPDF
from PIL import Image, ImageDraw
import io
import tempfile
import os

from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

MAX_FILE_SIZE = 5 * 1024 * 1024  # 5MB in bytes

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"detail": "Something went wrong on our end. Please try again."},
    )


@app.get("/", response_class=HTMLResponse)
def serve_homepage():
    with open("static/index.html", "r") as f:
        return f.read()


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


@app.post("/generate-pdf")
@limiter.limit("10/minute")
def generate_pdf(request: Request, title: str, body: str):
    if not title.strip() or not body.strip():
        raise HTTPException(status_code=400, detail="Title and body cannot be empty.")

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=16)
    pdf.cell(0, 10, title, ln=True)
    pdf.set_font("Helvetica", size=12)
    pdf.multi_cell(0, 8, body)

    temp_path = os.path.join(tempfile.gettempdir(), "generated.pdf")
    pdf.output(temp_path)

    return FileResponse(temp_path, media_type="application/pdf", filename="generated.pdf")


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


@app.post("/watermark-image")
@limiter.limit("10/minute")
async def watermark_image(request: Request, file: UploadFile = File(...), text: str = "SAMPLE"):
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="File must be an image.")

    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(status_code=413, detail="File too large. Max size is 5MB.")

    try:
        image = Image.open(io.BytesIO(contents)).convert("RGB")
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid or corrupted image file.")

    draw = ImageDraw.Draw(image)
    draw.text((10, 10), text, fill=(255, 0, 0))

    temp_path = os.path.join(tempfile.gettempdir(), "watermarked.jpg")
    image.save(temp_path, format="JPEG")

    return FileResponse(temp_path, media_type="image/jpeg", filename="watermarked.jpg")