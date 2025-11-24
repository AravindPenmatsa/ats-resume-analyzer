import logging
import os
import shutil
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, UploadFile, Form, Request, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates

from app.core.config import GENERATED_DIR, UPLOAD_DIR, templates
from app.services.resume_service import extract_text_from_file, add_header_section, enhance_resume_text
from app.services.scoring_service import score_resume, ats_formatting_warnings, ats_content_warnings
from app.services.pdf_service import generate_formatted_resume_pdf
from app.services.jd_skill_extractor import extract_skills_from_jd

router = APIRouter()
logger = logging.getLogger("app")

@router.get("/", include_in_schema=False)
async def root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@router.get("/upload", response_class=HTMLResponse)
async def upload_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@router.post("/upload", response_class=HTMLResponse)
async def upload_resume(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc_text: str = Form(...),
    generate_download: str = Form("no")
):
    logger.info(f"POST /upload - Received new request. Resume: '{resume.filename}', Generate Download: '{generate_download}'")
    try:
        resume_text = extract_text_from_file(resume)
        user_info = add_header_section(resume_text)

        resume_data = {
            "name": user_info['name'],
            "title": user_info['title'],
            "subtitle": user_info['subtitle'],
            "email": user_info['email'],
            "phone": user_info['phone'],
            "linkedin": user_info['linkedin'],
            "location": user_info['location'],
            "sections": []
        }
        jd_text = jobdesc_text

        if len(resume_text.strip()) < 30 or len(jd_text.strip()) < 30:
            logger.warning("Insufficient text from Resume or JD. Aborting analysis.")
            return templates.TemplateResponse("index.html", {
                "request": request,
                "score": 0,
                "suggestions": "❗ Unable to extract enough text from Resume or JD file.",
                "download_link": None,
                "hard_score": 0,
                "soft_score": 0,
                "search_score": 0
            })

        # Extract skills from job description using AI
        logger.info("🤖 Extracting skills from job description using AI...")
        hard_skills, soft_skills = extract_skills_from_jd(jd_text)
        logger.info(f"📊 Extracted {len(hard_skills)} hard skills and {len(soft_skills)} soft skills from JD")
        
        score, hard_score, soft_score, search_score, missing_keywords, matched_hard, matched_soft = score_resume(resume_text, hard_skills, soft_skills)

        filename = resume.filename or "unknown_file"
        formatting_issues = ats_formatting_warnings(resume_text, filename)
        content_issues = ats_content_warnings(resume_text, matched_hard, matched_soft)
        all_suggestions = formatting_issues + content_issues + [f"Missing Keywords: {missing_keywords}"]
        suggestions = " | ".join(all_suggestions)
        logger.info(f"Analysis complete. Final Score: {score}")

        download_link = None
        if generate_download.lower() == "yes":
            logger.info("📥 Download requested. Generating enhanced resume PDF.")
            
            # Get missing keywords for GPT enhancement
            resume_words = set(resume_text.lower().split())
            missing_hard_skills = hard_skills - resume_words
            missing_soft_skills = soft_skills - resume_words
            missing_keywords_set = missing_hard_skills.union(missing_soft_skills)
            
            logger.info(f"🔍 Analysis complete:")
            logger.info(f"   📊 Hard skills in JD: {len(hard_skills)}")
            logger.info(f"   📊 Soft skills in JD: {len(soft_skills)}")
            logger.info(f"   📊 Missing hard skills: {len(missing_hard_skills)} - {list(missing_hard_skills)[:10]}")
            logger.info(f"   📊 Missing soft skills: {len(missing_soft_skills)} - {list(missing_soft_skills)[:10]}")
            logger.info(f"   🎯 Total missing keywords for GPT: {len(missing_keywords_set)}")
            
            if missing_keywords_set:
                logger.info(f"🤖 Starting GPT enhancement for missing keywords...")
                enhanced_text = enhance_resume_text(resume_text, missing_keywords_set)
                
                # Verify enhancement worked
                original_length = len(resume_text)
                enhanced_length = len(enhanced_text)
                length_increase = enhanced_length - original_length
                
                logger.info(f"📈 Text length: {original_length} → {enhanced_length} (+{length_increase} chars)")
                
                if length_increase > 100:  # Should have significant increase if bullets were added
                    logger.info("✅ GPT enhancement appears successful - significant text increase detected")
                else:
                    logger.warning(f"⚠️ GPT enhancement may have failed - only {length_increase} character increase")
            else:
                logger.info("ℹ️ No missing keywords found - no GPT enhancement needed")
                enhanced_text = resume_text
            
            # Generate PDF with enhanced text
            logger.info("📄 Generating formatted PDF...")
            output_path = generate_formatted_resume_pdf(filename, enhanced_text, user_info)
            download_link = f"/download/{os.path.basename(output_path)}"
            logger.info(f"✅ Download link created: {download_link}")

        return templates.TemplateResponse("index.html", {
            "request": request,
            "score": score,
            "suggestions": suggestions,
            "download_link": download_link,
            "hard_score": hard_score,
            "soft_score": soft_score,
            "search_score": search_score,
            "generate_download": generate_download
        })
    except Exception as e:
        logger.error(f"An error occurred during /upload processing: {e}", exc_info=True)
        return templates.TemplateResponse("index.html", {
            "request": request, "score": 0, "suggestions": f"An unexpected error occurred: {str(e)}. Please check the logs."
        })

@router.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(GENERATED_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=filename, media_type='application/pdf')
    raise HTTPException(status_code=404, detail="File not found")
