# Import required modules from FastAPI and other libraries
from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docx import Document
import shutil, os, fitz, docx2txt, spacy, re
from app.utils import validate_resume_format, extract_keywords  # Reusable utility functions

# Initialize FastAPI app
app = FastAPI()

# Serve static files and HTML templates
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

# Create folders for storing uploaded and generated resumes
UPLOAD_DIR = "uploads"
GENERATED_DIR = "generated_resumes"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)

# Load spaCy model for NLP tasks
nlp = spacy.load("en_core_web_sm")

# Define sets of known hard and soft skills, and action verbs
HARD_KEYWORDS = {"python", "java", "cypress", "selenium", "api", "sql", "docker", "kubernetes", "aws", "jenkins", "jira", "linux", "bash"}
SOFT_KEYWORDS = {"communication", "leadership", "teamwork", "collaboration", "adaptability", "problem-solving", "critical thinking", "flexibility"}
ACTION_VERBS = {"developed", "led", "implemented", "designed", "created", "improved", "managed", "coordinated", "delivered", "built"}

# Function to extract plain text from uploaded resume file
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

# Extract hard and soft skills from job description using NLP
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

# Perform basic formatting checks on the resume text
def ats_formatting_warnings(resume_text: str, resume_filename: str):
    suggestions = []

    if not resume_filename.endswith((".doc", ".docx", ".txt")):
        suggestions.append("❌ Save your resume as .doc or .docx for better ATS compatibility.")
    if "Objective" not in resume_text and "Summary" not in resume_text:
        suggestions.append("❗ Include a clear 'Summary' or 'Objective' section at the top.")
    if re.search(r"[•▪▶❖➤]", resume_text):
        suggestions.append("❌ Use standard bullet points like '-' or '•' for better parsing.")
    if re.search(r"\d{1,2}/\d{2,4}", resume_text):
        suggestions.append("⚠️ Use consistent date format like 'MM/YYYY' or 'Month YYYY'.")
    if len(re.findall(r"([A-Z][A-Za-z\s]+):", resume_text)) < 3:
        suggestions.append("❗ Use standard section headings like Work Experience, Skills, Education.")
    if resume_text.lower().count("font-family") > 0:
        suggestions.append("❌ Avoid custom fonts/styles, use plain formatting (Arial, Calibri, etc.).")

    return suggestions

# Content-based ATS validations (keywords, action verbs, length, etc.)
def ats_content_warnings(resume_text: str, matched_hard: set, matched_soft: set):
    suggestions = []

    if len(matched_hard) < 3:
        suggestions.append("🔍 Add more hard skills relevant to the job.")
    if len(matched_soft) < 2:
        suggestions.append("💡 Mention soft skills like teamwork, communication, leadership.")
    if not any(verb in resume_text.lower() for verb in ACTION_VERBS):
        suggestions.append("🔧 Start bullet points with action verbs (e.g., 'Developed', 'Implemented').")
    if len(resume_text.split()) < 150:
        suggestions.append("✏️ Resume seems short. Expand with quantified achievements.")
    if resume_text.lower().count("lorem ipsum") > 0 or "dummy text" in resume_text.lower():
        suggestions.append("❌ Remove placeholder or dummy text.")

    return suggestions

# Scoring logic based on keyword match and readability
def score_resume(resume_text: str, hard_skills, soft_skills):
    resume_words = set(resume_text.lower().split())
    matched_hard = hard_skills.intersection(resume_words)
    matched_soft = soft_skills.intersection(resume_words)

    hard_score = round(len(matched_hard) / len(hard_skills) * 100, 2) if hard_skills else 0
    soft_score = round(len(matched_soft) / len(soft_skills) * 100, 2) if soft_skills else 0
    search_score = 90 if len(resume_text) > 300 and "-" in resume_text else 50

    final_score = round((0.5 * hard_score) + (0.3 * soft_score) + (0.2 * search_score), 2)
    missing_keywords = (hard_skills | soft_skills) - resume_words

    return final_score, hard_score, soft_score, search_score, ", ".join(sorted(missing_keywords)), matched_hard, matched_soft

# Save optimized resume with suggestions into a downloadable file
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

# Endpoint to serve upload form
@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Endpoint to handle resume and job description upload and analysis
@app.post("/upload", response_class=HTMLResponse)
async def upload_resume(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc: UploadFile = File(...),
    generate_download: str = Form("no")
):
    resume_text = extract_text_from_file(resume)
    jd_text = extract_text_from_file(jobdesc)

    if len(resume_text.strip()) < 30 or len(jd_text.strip()) < 30:
        return templates.TemplateResponse("index.html", {
            "request": request,
            "score": 0,
            "suggestions": "❗ Unable to extract enough text from Resume or JD file.",
            "download_link": None,
            "hard_score": 0,
            "soft_score": 0,
            "search_score": 0
        })

    # Perform ATS validations
    formatting_issues = ats_formatting_warnings(resume_text, resume.filename)
    content_issues = ats_content_warnings(resume_text, set(), set())

    # Extract job-specific keywords
    hard_skills, soft_skills = categorize_keywords(jd_text)

    # Score resume and get keyword matches
    score, hard_score, soft_score, search_score, missing_keywords, matched_hard, matched_soft = score_resume(resume_text, hard_skills, soft_skills)

    all_suggestions = formatting_issues + content_issues + [f"Missing Keywords: {missing_keywords}"]
    suggestions = " | ".join(all_suggestions)

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
        "search_score": search_score,
        "generate_download": generate_download,

    })

# Endpoint to serve the optimized resume file for download
@app.get("/download/{filename}")
async def download_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    if os.path.exists(path):
        return FileResponse(path, media_type="application/octet-stream", filename=filename)
    return {"detail": "File not found"}

# Root endpoint that redirects to the upload form
@app.get("/", include_in_schema=False)
async def root():
    return HTMLResponse('<script>window.location.replace("/upload")</script>')
