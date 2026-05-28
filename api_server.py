import os
import tempfile
import zipfile
from typing import List, Optional

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from PIL import Image

app = FastAPI(title="Pro File Tool API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_OUTPUTS = []


def save_upload(upload: UploadFile) -> str:
    ext = os.path.splitext(upload.filename or "upload.bin")[1]
    fd, path = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    with open(path, "wb") as f:
        f.write(upload.file.read())
    return path


def format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def mb_to_bytes(mb: float) -> int:
    return int(float(mb) * 1024 * 1024)


def track(path: Optional[str]) -> Optional[str]:
    if path:
        TEMP_OUTPUTS.append(path)
    return path


def make_zip(filepaths: List[str], base_name: str = "ai_parts") -> str:
    fd, zip_path = tempfile.mkstemp(suffix=".zip")
    os.close(fd)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for idx, path in enumerate(filepaths, start=1):
            zf.write(path, arcname=f"{base_name}_part_{idx:03d}.pdf")
    return track(zip_path)


def process_image_compress(path: str, quality: int = 65, max_width: int = 1800) -> str:
    img = Image.open(path)
    img_format = (img.format or "JPEG").upper()
    if img.mode in ("RGBA", "P") and img_format in ("JPEG", "JPG"):
        img = img.convert("RGB")

    width, height = img.size
    if width > max_width:
        new_height = int(height * (max_width / width))
        img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)

    ext = ".jpg" if img_format in ("JPEG", "JPG") else f".{img_format.lower()}"
    fd, out = tempfile.mkstemp(suffix=ext)
    os.close(fd)
    kwargs = {"optimize": True}
    if img_format in ("JPEG", "JPG", "WEBP"):
        kwargs["quality"] = quality
    if img_format == "PNG":
        kwargs["compress_level"] = 9
    img.save(out, format=img_format, **kwargs)
    return track(out)


def process_office_recompress(path: str) -> str:
    root, ext = os.path.splitext(path)
    out = f"{root}_compressed{ext}"
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
    return track(out)


def rasterize_pdf(path: str, dpi: int = 120, image_quality: int = 65, suffix: str = "compressed") -> str:
    import fitz

    root, _ = os.path.splitext(path)
    out = f"{root}_{suffix}.pdf"
    doc = fitz.open(path)
    new_doc = fitz.open()
    matrix = fitz.Matrix(dpi / 72, dpi / 72)

    for page in doc:
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        img_bytes = pix.tobytes("jpeg", jpg_quality=image_quality)
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(page.rect, stream=img_bytes)

    new_doc.save(out, garbage=4, deflate=True, clean=True)
    new_doc.close()
    doc.close()
    return track(out)


def compress_pdf_to_target(path: str, target_mb: Optional[float], mode: str) -> str:
    import fitz

    root, _ = os.path.splitext(path)
    target_bytes = mb_to_bytes(target_mb) if target_mb else None
    candidates = []

    doc = fitz.open(path)
    optimized = f"{root}_optimized.pdf"
    doc.save(optimized, garbage=4, deflate=True, clean=True)
    doc.close()
    candidates.append(track(optimized))

    if mode == "keep_text" and (not target_bytes or os.path.getsize(optimized) <= target_bytes):
        return optimized

    for dpi, quality in [(180, 80), (150, 70), (120, 60), (100, 50), (90, 40), (72, 35), (60, 30)]:
        out = rasterize_pdf(path, dpi=dpi, image_quality=quality, suffix=f"{dpi}dpi_q{quality}")
        candidates.append(out)
        if target_bytes and os.path.getsize(out) <= target_bytes:
            break

    best = min(candidates, key=lambda p: os.path.getsize(p))
    if target_bytes:
        under = [p for p in candidates if os.path.getsize(p) <= target_bytes]
        if under:
            best = max(under, key=lambda p: os.path.getsize(p))
    return best


def split_pdf_by_size(path: str, max_part_mb: float = 100) -> List[str]:
    import fitz

    max_bytes = mb_to_bytes(max_part_mb)
    doc = fitz.open(path)
    parts = []
    current = fitz.open()
    part_index = 1

    def save_part(part_doc, index: int) -> Optional[str]:
        if len(part_doc) == 0:
            return None
        out = os.path.join(tempfile.gettempdir(), f"ai_part_{index:03d}.pdf")
        part_doc.save(out, garbage=4, deflate=True, clean=True)
        return track(out)

    for page_index in range(len(doc)):
        test_doc = fitz.open()
        test_doc.insert_pdf(current)
        test_doc.insert_pdf(doc, from_page=page_index, to_page=page_index)
        fd, tmp = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        test_doc.save(tmp, garbage=4, deflate=True, clean=True)
        test_size = os.path.getsize(tmp)
        test_doc.close()
        os.remove(tmp)

        if len(current) > 0 and test_size > max_bytes:
            saved = save_part(current, part_index)
            if saved:
                parts.append(saved)
            current.close()
            current = fitz.open()
            part_index += 1

        current.insert_pdf(doc, from_page=page_index, to_page=page_index)

    saved = save_part(current, part_index)
    if saved:
        parts.append(saved)
    current.close()
    doc.close()
    return parts


@app.get("/health")
def health():
    return {"ok": True, "service": "pro-file-tool-api"}


@app.post("/process")
def process_file(
    file: UploadFile = File(...),
    target_mb: float = Form(100),
    max_part_mb: float = Form(100),
    split_for_ai: bool = Form(True),
    pdf_mode: str = Form("deep"),
):
    try:
        input_path = save_upload(file)
        input_size = os.path.getsize(input_path)
        ext = os.path.splitext(file.filename or "")[1].lower()
        base = os.path.splitext(file.filename or "result")[0]

        if ext in [".jpg", ".jpeg", ".png", ".webp"]:
            output = process_image_compress(input_path)
            download_name = f"{base}_compressed{ext}"
        elif ext == ".pdf":
            output = compress_pdf_to_target(input_path, target_mb, "keep_text" if pdf_mode == "keep_text" else "deep")
            if split_for_ai:
                parts = split_pdf_by_size(output, max_part_mb)
                output = make_zip(parts, base)
                download_name = f"{base}_ai_parts.zip"
            else:
                download_name = f"{base}_compressed.pdf"
        elif ext in [".docx", ".xlsx", ".pptx"]:
            output = process_office_recompress(input_path)
            download_name = f"{base}_compressed{ext}"
        else:
            return JSONResponse({"error": "Unsupported file type"}, status_code=400)

        output_size = os.path.getsize(output)
        headers = {
            "X-Input-Size": str(input_size),
            "X-Output-Size": str(output_size),
            "X-Input-Size-Text": format_file_size(input_size),
            "X-Output-Size-Text": format_file_size(output_size),
            "X-Saved-Percent": f"{max((input_size - output_size) / input_size * 100, 0):.1f}" if input_size else "0",
        }
        return FileResponse(output, filename=download_name, headers=headers)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
