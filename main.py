from fastapi import FastAPI, File, UploadFile
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from fpdf import FPDF
from PIL import Image, ImageDraw
import io
import tempfile
import os

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/", response_class=HTMLResponse)
def serve_homepage():
    with open("static/index.html", "r") as f:
        return f.read()


@app.post("/compress-image")
async def compress_image(file: UploadFile = File(...), quality: int = 50):
    image = Image.open(io.BytesIO(await file.read()))
    image = image.convert("RGB")

    # Create a temp file that works on any platform (local, Render, Vercel, etc.)
    temp_path = os.path.join(tempfile.gettempdir(), "compressed_output.jpg")
    image.save(temp_path, format="JPEG", quality=quality, optimize=True)

    return FileResponse(temp_path, media_type="image/jpeg", filename="compressed.jpg")


@app.post("/generate-pdf")
def generate_pdf(title: str, body: str):
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
    image = Image.open(io.BytesIO(await file.read()))
    image = image.convert("RGB")

    temp_path = os.path.join(tempfile.gettempdir(), "converted.pdf")
    image.save(temp_path)

    return FileResponse(temp_path, media_type="application/pdf", filename="converted.pdf")


@app.post("/watermark-image")
async def watermark_image(file: UploadFile = File(...), text: str = "SAMPLE"):
    image = Image.open(io.BytesIO(await file.read())).convert("RGB")
    draw = ImageDraw.Draw(image)
    draw.text((10, 10), text, fill=(255, 0, 0))

    temp_path = os.path.join(tempfile.gettempdir(), "watermarked.jpg")
    image.save(temp_path, format="JPEG")

    return FileResponse(temp_path, media_type="image/jpeg", filename="watermarked.jpg")