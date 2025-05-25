from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docx import Document
import shutil, os, fitz, docx2txt, spacy
from collections import defaultdict

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

UPLOAD_DIR = "uploads"
GENERATED_DIR = "generated_resumes"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)

nlp = spacy.load("en_core_web_sm")

# Basic skill-type classification
HARD_KEYWORDS = {"python", "java", "cypress", "selenium", "api", "sql", "docker", "kubernetes", "aws", "jenkins", "jira", "linux", "bash"}
SOFT_KEYWORDS = {"communication", "leadership", "teamwork", "collaboration", "adaptability", "problem-solving", "critical thinking", "flexibility"}

def extract_text_from_file(upload_file: UploadFile) -> str:
    file_path = os.path.join(UPLOAD_DIR, upload_file.filename)
    with open(file_path, "wb") as f:
        shutil.copyfileobj(upload_file.file, f)

    if file_path.endswith(".txt"):
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    elif file_path.endswith(".pdf"):
        doc = fitz.open(file_path)
        return "".join(page.get_text() for page in doc)
    elif file_path.endswith(".docx"):
        return docx2txt.process(file_path)
    return ""

def categorize_keywords(jd_text: str):
    hard_skills, soft_skills = set(), set()
    doc = nlp(jd_text.lower())
    for token in doc:
        word = token.lemma_.strip()
        if word in HARD_KEYWORDS:
            hard_skills.add(word)
        elif word in SOFT_KEYWORDS:
            soft_skills.add(word)

    return hard_skills, soft_skills

def score_resume(resume_text: str, hard_skills, soft_skills):
    resume_words = set(resume_text.lower().split())
    matched_hard = hard_skills.intersection(resume_words)
    matched_soft = soft_skills.intersection(resume_words)

    hard_score = round(len(matched_hard) / len(hard_skills) * 100, 2) if hard_skills else 0
    soft_score = round(len(matched_soft) / len(soft_skills) * 100, 2) if soft_skills else 0
    search_score = 90 if len(resume_text) > 200 else 50  # crude readability heuristic

    final_score = round((0.5 * hard_score) + (0.3 * soft_score) + (0.2 * search_score), 2)
    missing_keywords = (hard_skills | soft_skills) - resume_words
    return final_score, hard_score, soft_score, search_score, ", ".join(missing_keywords)

def save_optimized_resume(filename: str, resume_text: str, suggestions: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    base_name = os.path.splitext(filename)[0]
    output_path = os.path.join(GENERATED_DIR, f"{base_name}{ext if ext == '.docx' else '.txt'}")

    if ext == ".docx":
        doc = Document()
        doc.add_heading("Optimized Resume Content", level=1)
        doc.add_paragraph(resume_text)
        doc.add_heading("Suggestions", level=2)
        doc.add_paragraph(suggestions)
        doc.save(output_path)
    else:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write("Optimized Resume Content:\n\n" + resume_text + "\n\nSuggestions:\n" + suggestions)

    return output_path

@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload", response_class=HTMLResponse)
async def upload_resume(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc: UploadFile = File(...),
    generate_download: str = Form("no")
):
    resume_text = extract_text_from_file(resume)
    jd_text = extract_text_from_file(jobdesc)

    hard_skills, soft_skills = categorize_keywords(jd_text)
    score, hard_score, soft_score, search_score, suggestions = score_resume(resume_text, hard_skills, soft_skills)

    download_link = None
    if generate_download.lower() == "yes":
        output_path = save_optimized_resume(resume.filename, resume_text, suggestions)
        download_link = f"/download/{os.path.basename(output_path)}"

    return templates.TemplateResponse("index.html", {
        "request": request,
        "score": score,
        "suggestions": suggestions,
        "download_link": download_link,
        "hard_score": hard_score,
        "soft_score": soft_score,
        "search_score": search_score
    })

@app.get("/download/{filename}")
async def download_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    if os.path.exists(path):
        return FileResponse(path, media_type="application/octet-stream", filename=filename)
    return {"detail": "File not found"}

@app.get("/", include_in_schema=False)
async def root():
    return HTMLResponse('<script>window.location.replace("/upload")</script>')
