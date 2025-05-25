from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
import shutil
import fitz  # PyMuPDF
import docx2txt
from docx import Document
from fastapi.responses import RedirectResponse


app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

UPLOAD_DIR = "uploads"
GENERATED_DIR = "generated_resumes"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)


def extract_text_from_file(upload_file: UploadFile) -> str:
    contents = ""
    file_path = os.path.join(UPLOAD_DIR, upload_file.filename)

    with open(file_path, "wb") as f:
        shutil.copyfileobj(upload_file.file, f)

    if upload_file.filename.endswith(".txt"):
        with open(file_path, "r", encoding="utf-8") as f:
            contents = f.read()
    elif upload_file.filename.endswith(".pdf"):
        doc = fitz.open(file_path)
        for page in doc:
            contents += page.get_text()
    elif upload_file.filename.endswith(".docx"):
        contents = docx2txt.process(file_path)

    return contents


def analyze_resume(resume_text: str, jd_text: str):
    resume_words = set(resume_text.lower().split())
    jd_words = set(jd_text.lower().split())
    match = resume_words.intersection(jd_words)
    score = round(len(match) / len(jd_words) * 100, 2) if jd_words else 0
    missing_keywords = jd_words - resume_words
    return score, ", ".join(missing_keywords)


def save_optimized_resume(filename: str, resume_text: str, suggestions: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    base_name = os.path.splitext(filename)[0]
    output_path = os.path.join(GENERATED_DIR, f"{base_name}{ext if ext == '.docx' else '.txt'}")

    if ext == ".docx":
        document = Document()
        document.add_heading('Optimized Resume Content', level=1)
        document.add_paragraph(resume_text)
        document.add_heading('Job Matching Suggestions', level=2)
        document.add_paragraph(suggestions)
        document.save(output_path)
    else:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("Optimized Resume Content:\n\n")
            f.write(resume_text)
            f.write("\n\nJob Matching Suggestions:\n")
            f.write(suggestions)

    return output_path


@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/upload", response_class=HTMLResponse)
async def upload_resume(request: Request,
                        resume: UploadFile = File(...),
                        jobdesc: UploadFile = File(...),
                        generate_download: str = Form("no")):
    resume_text = extract_text_from_file(resume)
    jobdesc_text = extract_text_from_file(jobdesc)

    score, suggestions = analyze_resume(resume_text, jobdesc_text)

    download_link = None
    if generate_download.lower() == "yes":
        output_path = save_optimized_resume(resume.filename, resume_text, suggestions)
        download_link = f"/download/{os.path.basename(output_path)}"

    return templates.TemplateResponse("index.html", {
        "request": request,
        "score": score,
        "suggestions": suggestions,
        "download_link": download_link
    })


@app.get("/download/{filename}")
async def download_optimized_resume(filename: str):
    file_path = os.path.join(GENERATED_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type='application/octet-stream', filename=filename)
    return {"detail": "File not found"}

@app.get("/", include_in_schema=False)
async def root():
    return RedirectResponse(url="/upload")
