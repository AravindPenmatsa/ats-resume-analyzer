import spacy
import re
import logging
import subprocess
from app.core.config import HARD_KEYWORDS, SOFT_KEYWORDS, ACTION_VERBS

logger = logging.getLogger("app")

# Load spaCy model
try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    logger.info("spaCy model 'en_core_web_sm' not found. Downloading...")
    subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"], check=True)
    logger.info("spaCy model downloaded successfully.")
    nlp = spacy.load("en_core_web_sm")

def categorize_keywords(jd_text: str):
    """
    DEPRECATED: This function is kept for backward compatibility only.
    Skills should now be extracted using jd_skill_extractor.extract_skills_from_jd()
    
    This function now returns empty sets and logs a warning.
    """
    logger.warning("⚠️ categorize_keywords() is deprecated. Use jd_skill_extractor.extract_skills_from_jd() instead.")
    return set(), set()

def ats_formatting_warnings(resume_text: str, resume_filename: str):
    suggestions = []
    logger.info("Running ATS formatting checks...")

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

    logger.info(f"Found {len(suggestions)} formatting suggestions.")
    return suggestions

def ats_content_warnings(resume_text: str, matched_hard: set, matched_soft: set):
    suggestions = []
    logger.info("Running ATS content checks...")

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

    logger.info(f"Found {len(suggestions)} content suggestions.")
    return suggestions

def score_resume(resume_text: str, hard_skills, soft_skills):
    logger.info("Scoring resume against job description keywords...")
    resume_words = set(resume_text.lower().split())
    matched_hard = hard_skills.intersection(resume_words)
    matched_soft = soft_skills.intersection(resume_words)

    hard_score = round(len(matched_hard) / len(hard_skills) * 100, 2) if hard_skills else 0
    soft_score = round(len(matched_soft) / len(soft_skills) * 100, 2) if soft_skills else 0
    search_score = 90 if len(resume_text) > 300 and "-" in resume_text else 50

    final_score = round((0.5 * hard_score) + (0.3 * soft_score) + (0.2 * search_score), 2)
    missing_keywords = (hard_skills | soft_skills) - resume_words

    logger.info(f"Scoring complete. Final Score: {final_score}, Hard: {hard_score}, Soft: {soft_score}, Missing: {len(missing_keywords)} keywords.")
    return final_score, hard_score, soft_score, search_score, ", ".join(sorted(missing_keywords)), matched_hard, matched_soft
