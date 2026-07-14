from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from fpdf import FPDF
from PIL import Image, ImageDraw
import io
import tempfile
import os

MAX_FILE_SIZE = 5 * 1024 * 1024  # 5MB in bytes

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
def serve_homepage():
    with open("static/index.html", "r") as f:
        return f.read()


@app.post("/compress-image")
async def compress_image(file: UploadFile = File(...), quality: int = 50):
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
def generate_pdf(title: str, body: str):
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
async def convert_image_to_pdf(file: UploadFile = File(...)):
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
async def watermark_image(file: UploadFile = File(...), text: str = "SAMPLE"):
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