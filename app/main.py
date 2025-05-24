from fastapi import FastAPI, File, UploadFile, Request, Form
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from PyPDF2 import PdfReader
import docx
import os

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

def extract_text(file: UploadFile):
    if file.filename.endswith(".txt"):
        return file.file.read().decode("utf-8")
    elif file.filename.endswith(".pdf"):
        reader = PdfReader(file.file)
        return " ".join([page.extract_text() or "" for page in reader.pages])
    elif file.filename.endswith(".docx"):
        doc = docx.Document(file.file)
        return "\n".join([para.text for para in doc.paragraphs])
    else:
        return ""

def calculate_score(resume_text, job_text):
    resume_words = set(resume_text.lower().split())
    job_words = set(job_text.lower().split())
    common_words = resume_words.intersection(job_words)
    score = round((len(common_words) / len(job_words)) * 100, 2) if job_words else 0
    missing_keywords = job_words - resume_words
    return score, missing_keywords

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload", response_class=HTMLResponse)
async def upload_files(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc: UploadFile = File(...)
):
    resume_text = extract_text(resume)
    job_text = extract_text(jobdesc)

    score, missing_keywords = calculate_score(resume_text, job_text)
    suggestion = (
        f"Try adding these keywords to improve your score: {', '.join(list(missing_keywords)[:10])}"
        if missing_keywords else "Great job! Your resume is well aligned with the job description."
    )

    return templates.TemplateResponse("index.html", {
        "request": request,
        "score": score,
        "suggestion": suggestion
    })
