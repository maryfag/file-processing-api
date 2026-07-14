from fastapi import FastAPI, File, UploadFile
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from fpdf import FPDF
from PIL import Image, ImageDraw
import io

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
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=quality, optimize=True)
    buffer.seek(0)
    with open("compressed_output.jpg", "wb") as f:
        f.write(buffer.read())
    return FileResponse("compressed_output.jpg")


@app.post("/generate-pdf")
def generate_pdf(title: str, body: str):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=16)
    pdf.cell(0, 10, title, ln=True)
    pdf.set_font("Helvetica", size=12)
    pdf.multi_cell(0, 8, body)
    pdf.output("generated.pdf")
    return FileResponse("generated.pdf")


@app.post("/convert-image-to-pdf")
async def convert_image_to_pdf(file: UploadFile = File(...)):
    image = Image.open(io.BytesIO(await file.read()))
    image = image.convert("RGB")
    image.save("converted.pdf")
    return FileResponse("converted.pdf")


@app.post("/watermark-image")
async def watermark_image(file: UploadFile = File(...), text: str = "SAMPLE"):
    image = Image.open(io.BytesIO(await file.read())).convert("RGB")
    draw = ImageDraw.Draw(image)
    draw.text((10, 10), text, fill=(255, 0, 0))
    image.save("watermarked.jpg")
    return FileResponse("watermarked.jpg")