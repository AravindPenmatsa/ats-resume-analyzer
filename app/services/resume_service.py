import re
import shutil
import os
import fitz
import docx2txt
import logging
from pathlib import Path
from fastapi import UploadFile
from app.core.config import UPLOAD_DIR, ACTION_VERBS, HARD_KEYWORDS, SOFT_KEYWORDS, CACHE_PATH, GENERATED_DIR
from app.services.openai_service import openai_service
import json
import datetime
import sys


logger = logging.getLogger("app")

def split_into_logical_bullets(text_block: list) -> list:
    """
    Splits a list of lines (potentially flattened) into logical bullet points.
    Preserves complete sentences/paragraphs that are already properly formatted.
    """
    # If text_block is already a list of complete lines, check if they're already proper bullets
    if isinstance(text_block, list) and len(text_block) > 0:
        # Special handling for KEY STRENGTHS or similar sections with symbols
        has_symbols = any(re.search(r'^[✅✓•\-\*]\s*', line.strip()) for line in text_block if line.strip())
        if has_symbols:
            # Each line with a symbol should be a separate bullet
            bullets = []
            for line in text_block:
                stripped = line.strip()
                if stripped and (re.search(r'^[✅✓•\-\*]\s*', stripped) or len(stripped.split()) > 3):
                    # Clean up extra tabs/spaces but preserve the symbol
                    cleaned = re.sub(r'\s+', ' ', stripped).strip()
                    bullets.append(cleaned)
            return bullets
        
        # Special handling for sections like certifications where each meaningful line should be a bullet
        non_empty_lines = [line.strip() for line in text_block if line.strip() and not re.match(r'^\s*\t*\s*$', line)]
        if len(non_empty_lines) >= 2:
            # Check if lines look like separate items (certifications, achievements, etc.)
            # Each line should be reasonably substantial and look like a separate item
            separate_items = []
            for line in non_empty_lines:
                # Skip very short lines that are probably not complete items
                if len(line.split()) >= 2:  # At least 2 words
                    separate_items.append(line)
            
            # If we have multiple substantial lines, treat each as a separate bullet
            if len(separate_items) >= 2:
                return separate_items
        
        # Check if the lines appear to be already properly formatted (each line is a complete thought)
        properly_formatted = True
        for line in text_block:
            stripped = line.strip()
            # If lines are very short (except empty lines), they're probably not complete bullets
            if stripped and len(stripped.split()) < 5:
                properly_formatted = False
                break
        
        if properly_formatted:
            # Just clean and return the lines as they are
            return [line.strip() for line in text_block if line.strip() and len(line.strip().split()) > 2]
    
    # Otherwise, proceed with the original logic for poorly formatted text
    full_text = " ".join(text_block).strip()
    if not full_text:
        return []

    # If the text contains explicit bullet characters, split by those
    if re.search(r'[•\-\*]\s+', full_text):
        # Split by bullet characters
        bullets = re.split(r'[•\-\*]\s+', full_text)
        return [b.strip() for b in bullets if b.strip() and len(b.strip().split()) > 2]
    
    # If text appears to be paragraphs separated by periods, preserve complete sentences
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z])', full_text)
    if len(sentences) > 1:
        # Check if these look like complete thoughts
        valid_sentences = []
        for sent in sentences:
            sent = sent.strip()
            if sent and len(sent.split()) > 5:  # Reasonable sentence length
                valid_sentences.append(sent)
        if len(valid_sentences) > 0:
            return valid_sentences
    
    # Fallback: return the full text as a single bullet if it's substantial enough
    if len(full_text.split()) > 5:
        return [full_text]
    
    return []

def is_likely_company_location_date(line: str) -> bool:
    line = line.strip()
    
    # Pattern to find a date range (e.g., "May 2023 - Present" or "September2023")
    date_range_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4})\b'
    if not re.search(date_range_pattern, line, re.IGNORECASE):
        return False # A date range is a strong indicator, must be present

    # Pattern to find a company name (starts with Capital, might have Inc/LLC etc or "Client: ")
    # This is simplified. It assumes if a date is found, and it's not a bullet, it's likely a company line.
    # More specifically, look for "Client: Company" or just "Company" at the start.
    company_lead_pattern = r'^(?:Client:\s*)?[A-Z][a-zA-Z0-9\s,&.-]+(?:Inc|LLC|Corp|Ltd|Group|Solutions|Technologies|Network|Holdings)?\b'
    if not re.search(company_lead_pattern, line):
        return False # Must start with something that looks like a company name

    # Avoid lines that are clearly roles or responsibilities
    if re.search(r'^\s*(?:Role|Responsibilities):\s*', line, re.IGNORECASE):
        return False
    
    # Avoid lines that start with explicit bullets
    if re.match(r'^[•\-\*]', line):
        return False

    return True

def is_experience_company_line(line: str) -> bool:
    """Detects lines containing a Company, Location, and Date range."""
    # Pattern: Optional (Company Name, ) Optional (City, State) Month Year to Month Year/Present
    # This is a complex pattern to capture common variations.
    # It prioritizes matching a company-like name, then a location, then a date range.
    
    # 1. Look for a date range (most reliable indicator of job entry)
    date_range_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4})\b'
    if not re.search(date_range_pattern, line, re.IGNORECASE):
        return False
        
    # 2. Look for a company name (starts with Capital, might have Inc/LLC etc) followed by comma or space
    company_name_pattern = r'^[A-Z][a-zA-Z0-9\s,&.-]+(?:Inc|LLC|Corp|Ltd|Group|Solutions|Technologies)?\b'
    if not re.search(company_name_pattern, line):
        return False # Must start with something that looks like a company name

    # 3. Look for a location (City, State/Country) or just a city/state
    # This is less strict to allow for variations
    location_pattern = r'\b(?:Dallas|Austin|Houston|TX|New York|NY|CA|San Francisco|Chicago|Hyderabad|India|Remote|USA)\b'
    if not re.search(location_pattern, line, re.IGNORECASE) and ',' not in line:
        # If no specific city/state is found, at least a comma is expected for "Company, Location"
        # However, for "Company Dates" directly, we might not have a comma.
        # This part requires a balance. Let's rely more on the date pattern and the general structure.
        pass # Not a strict requirement for every line, as location can be implied or missing.

    # Exclude lines that are clearly just role or responsibilities lines despite having dates
    if re.search(r'^\s*Role:\s*|^\s*Responsibilities:\s*', line, re.IGNORECASE):
        return False
    
    return True

def is_experience_role_line(line: str) -> bool:
    # Enhanced to handle "Role:" prefix and various role formats
    return bool(re.search(r'^\s*Role:\s*([A-Za-z0-9\s,/\-]+(?:engineer|analyst|developer|manager|contractor|fulltime|sdet|qa|specialist)\b)?', line, re.IGNORECASE))

def is_responsibility_heading(line: str) -> bool:
    return bool(re.search(r'^\s*(?:Responsibilities|Environment):\s*', line, re.IGNORECASE))

def split_experience_line(line: str) -> list:
    """
    Splits a single long line into multiple bullets based on Action Verbs and sentence boundaries.
    """
    bullets = []
    
    # First, split by likely sentence boundaries (period followed by space and Capital letter)
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z])', line)
    
    current_bullet = []
    for sent in sentences:
        # Check if this sentence starts with an action verb
        first_word = sent.split()[0].lower() if sent.split() else ""
        # Remove non-alphabetic chars from first word for better matching
        first_word = re.sub(r'[^a-z]', '', first_word)
        
        if first_word in ACTION_VERBS or (current_bullet and len(current_bullet) > 0 and first_word in ACTION_VERBS):
            # It's a new bullet
            if current_bullet:
                bullets.append(" ".join(current_bullet).strip())
            current_bullet = [sent]
        else:
            # Continuation or non-action sentence
            if not current_bullet:
                 current_bullet = [sent]
            else:
                 current_bullet.append(sent)
                 
    if current_bullet:
        bullets.append(" ".join(current_bullet).strip())
        
    return bullets

def process_general_bullets(content_lines: list) -> list:
    """
    Parses content for sections like Professional Summary and Achievements,
    returning a simple list of bullet strings.
    """
    processed_bullets = []
    current_bullet_segment = []
    
    # Regex for new bullet point detection for general text.
    # It's similar to the one in parse_experience_section but might be less strict
    # if summaries can start with non-verbs.
    summary_bullet_start_regex = re.compile(
        r'^(?:[•\-\*]\s*|' + # Explicit bullet characters
        r'(?:\b(?:' + '|'.join(re.escape(s) for s in ACTION_VERBS) + r')\b(?!\s*:)|' + # Action verb not followed by colon
        r'\b(?:Around|Experience|Proficient|Extensive|Good|Strong|Well|Accustomed|Involved|Knowledge|Hands)\b))', # Common summary starters
        re.IGNORECASE
    )

    for line in content_lines:
        stripped_line = line.strip()
        if not stripped_line:
            if current_bullet_segment: # If an empty line, finalize current bullet
                processed_bullets.append(" ".join(current_bullet_segment).strip())
                current_bullet_segment = []
            continue

        # Check if this line should start a *new* bullet point
        is_new_bullet = False
        if summary_bullet_start_regex.search(stripped_line):
            is_new_bullet = True
        elif len(current_bullet_segment) > 0 and current_bullet_segment[-1].strip().endswith(('.', '!', '?')):
            # If the previous line ended a sentence, assume this is a new bullet
            is_new_bullet = True

        if is_new_bullet and current_bullet_segment:
            processed_bullets.append(" ".join(current_bullet_segment).strip())
            current_bullet_segment = [stripped_line]
        else:
            current_bullet_segment.append(stripped_line)
            
    if current_bullet_segment: # Add the last collected bullet
        processed_bullets.append(" ".join(current_bullet_segment).strip())
            
    # Final filtering: remove any empty or very short items that don't make sense as bullets
    return [bullet for bullet in processed_bullets if bullet and len(bullet.split()) > 3]


def add_header_section(resume_text):
    # Limit the search to the first N characters (e.g., 500) to ensure it only captures header info.
    header_search_area = resume_text[:500]  # Increased to handle more varied formats

    # Try multiple patterns for name extraction
    name = None
    
    # First try to get the first line as name (most reliable)
    first_line = header_search_area.split('\n')[0].strip()
    if first_line and len(first_line.split()) <= 4 and re.match(r'^[A-Z][a-zA-Z\s]+$', first_line):
        name = first_line
    else:
        # Fallback to pattern matching
        # Pattern 0: All caps name on its own line (e.g., "NARENDRA REDDY DEVARAPALLI")
        name_pattern0 = re.search(r'^([A-Z][A-Z\s]+[A-Z])$', header_search_area.strip(), re.MULTILINE)
        
        # Pattern 1: Name with middle initial (e.g., "Akhil V Sakhineti")
        name_pattern1 = re.search(r'^([A-Z][a-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][a-z]+)+)', header_search_area.strip(), re.MULTILINE)
        
        # Pattern 2: Single name on its own line (e.g., "Aravind" or "John")
        name_pattern2 = re.search(r'^([A-Z][a-z]+)\s*$', header_search_area.strip(), re.MULTILINE)
        
        # Pattern 3: Full name with pipe separator (e.g., "John Doe | email")
        name_pattern3 = re.search(r'^([A-Z][a-z]+\s+[A-Z][a-z]+)\s*\|', header_search_area.strip(), re.MULTILINE)
        
        if name_pattern0:
            # Convert all caps to proper case
            name = name_pattern0.group(1).title()
        elif name_pattern1:
            name = name_pattern1.group(1)
        elif name_pattern3:
            name = name_pattern3.group(1)
        elif name_pattern2:
            # If we only found a first name, try to find a last name nearby
            first_name = name_pattern2.group(1)
            # Look for another capitalized word that might be a last name
            lines = header_search_area.split('\n')
            for i, line in enumerate(lines):
                if first_name in line:
                    # Check next few lines for potential last name
                    for j in range(i+1, min(i+5, len(lines))):
                        if lines[j].strip() and not lines[j].strip().isupper():
                            # Check if it's a name-like word
                            potential_last = re.search(r'^([A-Z][a-z]+)', lines[j].strip())
                            if potential_last:
                                name = f"{first_name} {potential_last.group(1)}"
                                break
                    break
            if not name:
                name = first_name  # Use just first name if no last name found
    
    # Extract skills in parentheses - find the longest one with commas (most likely to be skills)
    paren_matches = re.findall(r'\(\s*([^)]+)\s*\)', header_search_area)
    skills = ""
    if paren_matches:
        # Filter out phone number patterns (like "513", area codes, etc.)
        filtered_matches = []
        for match in paren_matches:
            # Skip if it looks like a phone number (only digits, or short digit sequences)
            if not (match.strip().isdigit() and len(match.strip()) <= 4):
                filtered_matches.append(match)
        
        if filtered_matches:
            # Find the longest match that contains commas (likely to be skills list)
            skills_candidates = [match for match in filtered_matches if ',' in match]
            if skills_candidates:
                skills = max(skills_candidates, key=len).strip()
            else:
                # Fallback to longest non-phone-number match
                skills = max(filtered_matches, key=len).strip()
    
    email_pattern = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', header_search_area)
    
    # Enhanced phone number extraction to get the full number
    phone_pattern = re.search(r'(?:Phone\s*:?\s*)?(\+?1?\s*\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', header_search_area, re.IGNORECASE)
    
    linkedin_pattern = re.search(r'https?://(www\.)?linkedin\.com/in/[^\s]+', header_search_area)
    
    # Apply location and title patterns
    location_pattern = re.search(r'(Dallas|Austin|Houston|TX|New York|NY|CA|San Francisco|Chicago|Hyderabad|India|LittleElm)', header_search_area)
    
    # Enhanced location extraction to handle "City,State" format
    location = ""
    if location_pattern:
        location = location_pattern.group(0)
        # Check if we have a city,state pattern nearby
        city_state_pattern = re.search(r'(Dallas|Austin|Houston|New York|San Francisco|Chicago|Hyderabad|LittleElm)\s*,?\s*(TX|NY|CA|India)?', header_search_area, re.IGNORECASE)
        if city_state_pattern:
            city = city_state_pattern.group(1)
            state = city_state_pattern.group(2) if city_state_pattern.group(2) else ""
            if city and state:
                location = f"{city}, {state}"
            else:
                location = city
    
    title_pattern = re.search(r'(SDET|QA Engineer|Software Engineer in Test|Senior Software Engineer in Test|Automation Engineer|Full Stack Developer|Sr\.Full Stack Developer|Software Development Engineer in Test|Senior Application Developer|Senior Software Developer)', header_search_area, re.IGNORECASE)

    return {
        "name": name if name else "Candidate Name",
        "title": title_pattern.group(1) if title_pattern else "",
        "subtitle": skills,  # Add skills as subtitle
        "email": email_pattern.group(0) if email_pattern else "",
        "phone": phone_pattern.group(1).strip() if phone_pattern else "",
        "linkedin": linkedin_pattern.group(0) if linkedin_pattern else "",
        "location": location
    }

def extract_text_from_file(upload_file: UploadFile) -> str:
    if not upload_file.filename:
        logger.error("Upload file has no filename.")
        raise ValueError("Upload file must have a filename")

    file_path = UPLOAD_DIR / upload_file.filename
    logger.info(f"Saving uploaded file to: {file_path}")
    with open(file_path, "wb") as f:
        shutil.copyfileobj(upload_file.file, f)

    file_ext = os.path.splitext(file_path)[1].lower()
    logger.info(f"Extracting text from '{file_path}' with extension '{file_ext}'")

    if file_ext == ".txt":
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    elif file_ext == ".pdf":
        doc = fitz.open(file_path)
        text = "".join(page.get_text() for page in doc)  # type: ignore
        logger.info(f"Extracted {len(text)} characters from PDF.")
        return text
    elif file_ext == ".docx":
        text = docx2txt.process(file_path)
        logger.info(f"Extracted {len(text)} characters from DOCX.")
        return text
    
    logger.warning(f"Unsupported file type: {file_ext}. Returning empty string.")
    return ""

# --- Helper Functions ---

def is_duration_line(line: str) -> bool:
    """Identifies a line containing a typical project date range."""
    date_pattern = re.search(
        r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4})',
        line,
        re.I
    )
    return bool(date_pattern)

def is_project_header_line(line: str) -> bool:
    """Robustly identifies a line containing a project header date."""
    date_pattern = re.search(
        r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4})',
        line,
        re.I
    )
    return bool(date_pattern)

def is_title_line(line: str) -> bool:
    """Dynamically identifies a job title line using keyword and length heuristics."""
    stripped_line = line.strip()

    # Heuristic 1: The line must be short to be a title.
    if len(stripped_line.split()) > 7:
        return False

    # Heuristic 2: The line contains a common job title keyword.
    title_keywords = [
        'SDET', 'QA', 'Engineer', 'Developer', 'Analyst', 'Architect',
        'Consultant', 'Manager', 'Lead', 'Tester', 'Specialist'
    ]
    if any(re.search(f"\\b{keyword}\\b", stripped_line, re.I) for keyword in title_keywords):
        # Heuristic 3: It shouldn't contain action verbs, which indicates a responsibility.
        if not any(verb in stripped_line.lower() for verb in ACTION_VERBS):
            return True

    return False

def is_date_line(line: str) -> bool:
    """Finds a line containing a date range like 'Mon YYYY - Mon YYYY' or 'Mon YYYY - Present'."""
    # Enhanced regex to handle various date formats including "September2023" without space
    date_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4})'
    return bool(re.search(date_pattern, line, re.I))

def has_start_date(line: str) -> bool:
    """Detects a line containing a start date (e.g., 'March 2022' or 'Company March 2022')."""
    # Look for month + year at the end of line or before separator
    start_date_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}(?:\s*$|(?=\s*[–-])|(?=\s*to))'
    return bool(re.search(start_date_pattern, line, re.I))

def has_end_date(line: str) -> bool:
    """Detects a line containing an end date (e.g., '– May 2023' or 'to Present')."""
    # Look for end date patterns starting with separator
    end_date_pattern = r'^\s*[–-]\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4})|^\s*to\s+(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4})'
    return bool(re.search(end_date_pattern, line, re.I))

def combine_split_date_range(line1: str, line2: str) -> str:
    """Combines two lines that together form a complete date range."""
    # Extract start date from line1
    start_match = re.search(r'\b((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4})', line1, re.I)
    start_date = start_match.group(1) if start_match else ""
    
    # Extract end date from line2  
    end_match = re.search(r'[–-]\s*((?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}))', line2, re.I)
    if not end_match:
        end_match = re.search(r'to\s+((?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}))', line2, re.I)
    
    end_date = end_match.group(1) if end_match else ""
    
    if start_date and end_date:
        return f"{start_date} – {end_date}"
    
    return ""

def is_company_line(line: str) -> bool:
    """
    Final corrected function to dynamically identify a company line.
    This version correctly handles dates appearing on the same line as the location.
    """
    stripped_line = line.strip()

    # Heuristic 1 (Primary): The line contains a location pattern.
    # The end-of-line anchor '$' has been REMOVED to make the pattern more flexible.
    location_pattern = r',\s*[A-Za-z\s]+,?\s*[A-Z]{2}'
    if re.search(location_pattern, stripped_line):
        return True

    # Heuristic 2 (Secondary): Contains a common company suffix and is not a sentence.
    # This remains a good fallback for company names without a location.
    company_suffixes = [
        'LLC', 'Inc', 'Corp', 'Ltd', 'Services', 'Solutions',
        'Technologies', 'Group', 'Labs', 'Network', 'Association'
    ]
    if len(stripped_line.split()) < 8 and any(re.search(f"\\b{suffix}\\b", stripped_line, re.I) for suffix in company_suffixes):
        # Ensure it's not a responsibility sentence by checking for action verbs.
        if not stripped_line.lower().split()[0] in ACTION_VERBS:
            return True

    return False


def parse_resume_to_structure(enhanced_text: str, user_info: dict) -> dict:
    resume_data = {
        "name": user_info.get("name", "Candidate Name"),
        "title": user_info.get("title", ""),
        "subtitle": user_info.get("subtitle", ""),  # Add subtitle (skills) from user_info
        "email": user_info.get("email", ""),
        "phone": user_info.get("phone", ""),
        "linkedin": user_info.get("linkedin", ""),
        "location": user_info.get("location", ""),
        "sections": []
    }

    VALID_SECTIONS = [
        "PROFILE SUMMARY", "PROFESSIONAL SUMMARY", "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE", "PROJECTS",
        "EDUCATION", "TECHNICAL SKILLS", "CERTIFICATIONS", "ACHIEVEMENTS", "EXPERTISE IN", "KEY STRENGTHS", "SUMMARY", "EDUCATION & CERTIFICATIONS"
    ]
    
    sections = {}
    current_section = None
    current_content = []
    
    lines = enhanced_text.split('\n')
    for idx, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line:
            if current_section and current_content:
                sections[current_section] = current_content.copy()
            continue
        
        # Remove decorative characters and punctuation for section detection
        # This handles cases like "---PROFESSIONAL SUMMARY---" or "PROFESSIONAL SUMMARY:"
        cleaned_line = re.sub(r'^[-=\s]*', '', stripped_line)  # Remove leading dashes/equals
        cleaned_line = re.sub(r'[-=\s]*$', '', cleaned_line)   # Remove trailing dashes/equals
        normalized_line = re.sub(r'[.:,;!?]+$', '', cleaned_line).strip()
        
        if normalized_line.upper() in VALID_SECTIONS:
            if current_section and current_content:
                sections[current_section] = current_content.copy()
            
            current_section = normalized_line.upper()
            current_content = []
            logger.debug(f"Detected section: {current_section}")
        else:
            if current_section:
                current_content.append(line)
    
    if current_section and current_content:
        sections[current_section] = current_content.copy()
    
    logger.debug(f"Parsed sections: {sections.keys()}")

    desired_sections = [
        "PROFESSIONAL SUMMARY", 
        "PROFILE SUMMARY",      
        "EDUCATION", 
        "TECHNICAL SKILLS",
        "CERTIFICATIONS",
        "ACHIEVEMENTS",         
        "KEY STRENGTHS",         
        "PROFESSIONAL EXPERIENCE",
        "WORK EXPERIENCE", 
        "PROJECTS",
        "SUMMARY",
        "EDUCATION & CERTIFICATIONS"
    ]
    
    for section_name in desired_sections:
        section_content = sections.get(section_name.upper())
        if section_content:
            if section_name.upper() in ['PROFESSIONAL EXPERIENCE', 'WORK EXPERIENCE']:
                processed_content = parse_experience_section(section_content)
                
                # Debug logging
                logger.info(f"Parsed {len(processed_content)} experience entries")
                for idx, entry in enumerate(processed_content):
                    logger.info(f"Entry {idx}: header='{entry.get('header', '')}', role='{entry.get('role', '')}'")
                
                # Transform to template format (company, location, duration, role, responsibilities)
                transformed_content = []
                for entry in processed_content:
                    # Parse header to extract company, location, and duration
                    header = entry.get('header', '')
                    role = entry.get('role', '')
                    
                    # Try to split header by | to get company/location and duration
                    if '|' in header:
                        parts = header.split('|')
                        company_location = parts[0].strip() if len(parts) > 0 else ''
                        duration = parts[1].strip() if len(parts) > 1 else ''
                        
                        # Further split company_location by comma to separate company and location
                        if ',' in company_location:
                            company_parts = company_location.split(',')
                            company = company_parts[0].strip()
                            location = ','.join(company_parts[1:]).strip()
                        else:
                            company = company_location
                            location = ''
                    else:
                        # No pipe separator, try to extract duration from end
                        date_match = re.search(r'((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))$', header, re.I)
                        if date_match:
                            duration = date_match.group(1)
                            company_location = header[:date_match.start()].strip()
                            
                            if ',' in company_location:
                                company_parts = company_location.split(',')
                                company = company_parts[0].strip()
                                location = ','.join(company_parts[1:]).strip()
                            else:
                                company = company_location
                                location = ''
                        else:
                            company = header
                            location = ''
                            duration = ''
                    
                    transformed_entry = {
                        'company': company,
                        'location': location,
                        'duration': duration,
                        'role': role,
                        'responsibilities': entry.get('responsibilities', []),
                        'environment': entry.get('environment', '')
                    }
                    transformed_content.append(transformed_entry)
                
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "professional_experience",
                    "content": transformed_content
                })
                # For Jinja2 template compatibility
                if "experience" not in resume_data:
                    resume_data["experience"] = transformed_content
                else:
                    resume_data["experience"].extend(transformed_content)
                    
            elif section_name.upper() == 'PROJECTS':
                # Check if PROJECTS section looks like professional experience format
                # (has "Responsibilities:" or multiple date lines indicating job entries)
                has_responsibilities = any('responsibilities:' in line.lower() for line in section_content if line.strip())
                date_count = sum(1 for line in section_content if line.strip() and is_date_line(line.strip()))
                
                # If it has "Responsibilities:" or multiple date entries, treat as professional experience
                if has_responsibilities or date_count > 1:
                    logger.info("PROJECTS section detected as professional experience format")
                    processed_content = parse_experience_section(section_content)
                    
                    # Debug logging
                    logger.info(f"Parsed {len(processed_content)} experience entries from PROJECTS")
                    for idx, entry in enumerate(processed_content):
                        logger.info(f"Entry {idx}: header='{entry.get('header', '')}', role='{entry.get('role', '')}'")
                    
                    # Transform to template format (company, location, duration, role, responsibilities)
                    transformed_content = []
                    for entry in processed_content:
                        # Parse header to extract company, location, and duration
                        header = entry.get('header', '')
                        role = entry.get('role', '')
                        
                        # Try to split header by | to get company/location and duration
                        if '|' in header:
                            parts = header.split('|')
                            company_location = parts[0].strip() if len(parts) > 0 else ''
                            duration = parts[1].strip() if len(parts) > 1 else ''
                            
                            # Further split company_location by comma to separate company and location
                            if ',' in company_location:
                                company_parts = company_location.split(',')
                                company = company_parts[0].strip()
                                location = ','.join(company_parts[1:]).strip()
                            else:
                                company = company_location
                                location = ''
                        else:
                            # No pipe separator, try to extract duration from end
                            date_match = re.search(r'((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))$', header, re.I)
                            if date_match:
                                duration = date_match.group(1)
                                company_location = header[:date_match.start()].strip()
                                
                                if ',' in company_location:
                                    company_parts = company_location.split(',')
                                    company = company_parts[0].strip()
                                    location = ','.join(company_parts[1:]).strip()
                                else:
                                    company = company_location
                                    location = ''
                            else:
                                company = header
                                location = ''
                                duration = ''
                        
                        transformed_entry = {
                            'company': company,
                            'location': location,
                            'duration': duration,
                            'role': role,
                            'responsibilities': entry.get('responsibilities', []),
                            'environment': entry.get('environment', '')
                        }
                        transformed_content.append(transformed_entry)
                        logger.info(f"Transformed entry {idx}: company='{company}', location='{location}', duration='{duration}'")
                    
                    resume_data["sections"].append({
                        "name": "Professional Experience",
                        "type": "professional_experience",
                        "content": transformed_content
                    })
                    # For Jinja2 template compatibility - add to experience, not projects
                    if "experience" not in resume_data:
                        resume_data["experience"] = transformed_content
                    else:
                        resume_data["experience"].extend(transformed_content)
                else:
                    # Traditional project format
                    processed_content = parse_projects_content(section_content)
                    resume_data["sections"].append({
                        "name": section_name.title(),
                        "type": "projects",
                        "content": processed_content
                    })
                    # For Jinja2 template compatibility
                    resume_data["projects"] = processed_content
                    
            elif section_name.upper() in ['PROFESSIONAL SUMMARY', 'PROFILE SUMMARY', 'SUMMARY']:
                # For Professional Summary, convert paragraphs to bullet points
                processed_content = parse_professional_summary(section_content)
                # Always keep as bullets for these sections - don't convert to plain text
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "bullets",
                    "content": processed_content
                })
                # For Jinja2 template compatibility
                resume_data["summary"] = processed_content
                
            elif section_name.upper() == 'TECHNICAL SKILLS':
                # For Technical Skills, handle key-value format
                processed_content = parse_technical_skills(section_content)
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "plain_text_block",
                    "content": processed_content
                })
                # For Jinja2 template compatibility (needs parsing to dict)
                skills_dict = {}
                if isinstance(processed_content, str):
                    for line in processed_content.split('\n'):
                        if ':' in line:
                            key, val = line.split(':', 1)
                            skills_dict[key.strip()] = val.strip()
                        elif line.strip():
                            skills_dict["Other"] = line.strip()
                resume_data["skills"] = skills_dict
                
            elif section_name.upper() in ['ACHIEVEMENTS', 'CERTIFICATIONS', 'KEY STRENGTHS']:
                # These sections use the general bullet parser
                processed_content = split_into_logical_bullets(section_content) 
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "bullets",
                    "content": processed_content
                })
                if section_name.upper() == 'CERTIFICATIONS':
                    resume_data["certifications"] = processed_content
                    
            elif section_name.upper() == 'EDUCATION':
                # Simple education parser for now
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "plain_text_block",
                    "content": "\n".join([l.strip() for l in section_content if l.strip()])
                })
                # For Jinja2 template compatibility (basic list of dicts)
                edu_list = []
                for line in section_content:
                    if line.strip():
                        edu_list.append({"university": line.strip(), "degree": "", "year": ""})
                resume_data["education"] = edu_list

            else:
                # Default case for other sections
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "plain_text_block",
                    "content": "\n".join([l.strip() for l in section_content if l.strip()])
                })

    return resume_data

def parse_professional_summary(content_lines: list) -> list:
    """
    Parses Professional Summary content, handling both bulleted and paragraph formats.
    Returns a list of bullet points or paragraphs.
    """
    bullets = []
    current_paragraph = []
    
    # Check if the content appears to be continuous prose without explicit bullets
    has_bullets = any(re.match(r'^[•\-\*]\s*', line.strip()) for line in content_lines if line.strip())
    
    # For continuous prose (like Aravind's resume), split into logical bullets
    if not has_bullets:
        for line in content_lines:
            stripped_line = line.strip()
            if stripped_line:
                # Check if this line starts with a capital letter and is a complete thought
                # This helps split continuous text into meaningful bullets
                if (current_paragraph and 
                    stripped_line[0].isupper() and 
                    len(current_paragraph[-1].split()) > 5):
                    # Previous paragraph seems complete, start a new bullet
                    bullets.append(" ".join(current_paragraph).strip())
                    current_paragraph = [stripped_line]
                else:
                    current_paragraph.append(stripped_line)
        
        # Add the last paragraph
        if current_paragraph:
            bullets.append(" ".join(current_paragraph).strip())
            
        # If we ended up with just one long bullet, try to split by sentences
        if len(bullets) == 1:
            text = bullets[0]
            # Split by periods followed by capital letters, but preserve the periods
            sentences = re.split(r'(?<=\.)\s+(?=[A-Z])', text)
            if len(sentences) > 1:
                bullets = sentences
        
        return bullets
    
    # Original bullet-handling logic for resumes with explicit bullets
    for line in content_lines:
        stripped_line = line.strip()
        if not stripped_line:
            if current_paragraph:
                bullets.append(" ".join(current_paragraph).strip())
                current_paragraph = []
            continue
        
        # Check if this line already has bullet markers
        if re.match(r'^[•\-\*]\s*', stripped_line):
            if current_paragraph:
                bullets.append(" ".join(current_paragraph).strip())
                current_paragraph = []
            # Remove bullet marker and add as new bullet
            cleaned_line = re.sub(r'^[•\-\*]\s*', '', stripped_line).strip()
            bullets.append(cleaned_line)
        else:
            # For paragraph-style text (like Akhil's resume)
            # Each non-empty line that's sufficiently long is treated as a separate bullet
            if len(stripped_line.split()) > 10:  # Paragraph threshold
                if current_paragraph:
                    bullets.append(" ".join(current_paragraph).strip())
                    current_paragraph = []
                bullets.append(stripped_line)
            else:
                current_paragraph.append(stripped_line)
    
    if current_paragraph:
        bullets.append(" ".join(current_paragraph).strip())
    
    return [b for b in bullets if b and len(b.split()) > 3]

def parse_technical_skills(content_lines: list) -> str:
    """
    Parses Technical Skills section, handling various formats robustly.
    Returns formatted text block.
    """
    formatted_lines = []
    current_category = None
    
    # Check if it's a single-line format like "Skills: ..."
    if len(content_lines) == 1 or (len(content_lines) > 0 and any(line.strip().startswith(('Skills:', 'Technical Skills:', 'Core Skills:')) for line in content_lines)):
        # Handle single-line or paragraph format
        for line in content_lines:
            stripped_line = line.strip()
            if stripped_line:
                # If it starts with "Skills:" or similar, keep it as is
                if re.match(r'^(Skills|Technical Skills|Core Skills)\s*:', stripped_line, re.I):
                    formatted_lines.append(stripped_line)
                else:
                    # Otherwise, add it as a continuation
                    formatted_lines.append(stripped_line)
        return "\n".join(formatted_lines)
    
    # Preprocess: Combine "Category:" lines with their values on the next line
    # This handles table format where docx2txt extracts category and value as separate lines
    preprocessed_lines = []
    i = 0
    while i < len(content_lines):
        line = content_lines[i].strip()
        if not line:
            i += 1
            continue
        
        # Check if this line ends with a colon (table category format)
        if line.endswith(':') and i + 1 < len(content_lines):
            next_line = content_lines[i + 1].strip()
            if next_line and not next_line.endswith(':'):
                # Combine category with its value
                preprocessed_lines.append(f"{line} {next_line}")
                i += 2  # Skip the next line since we've already processed it
                continue
        
        preprocessed_lines.append(line)
        i += 1
    
    # Use preprocessed lines for the rest of the parsing
    content_lines = preprocessed_lines
    
    # List of likely categories to help disambiguate
    LIKELY_CATEGORIES = [
        'Languages', 'Databases', 'Tools', 'Technologies', 'Frameworks', 'Web', 'Cloud',
        'Operating Systems', 'Methodologies', 'IDE', 'Servers', 'Version Control', 'Testing',
        'Application Servers', 'Database editors', 'Bug reporting tools', 'Reporting Tools',
        'Testing Tools', 'Programming Languages', 'Web Technologies'
    ]
    
    # Handle multi-line key-value format (like Akhil's resume)
    for i, line in enumerate(content_lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue
        
        # Check if this looks like a category (short line, no commas, capitalized)
        # Allow for longer category names containing & or / characters
        is_likely_category = False
        
        # First, check basic category characteristics
        looks_like_category = False
        if any(cat.lower() == stripped_line.lower() for cat in LIKELY_CATEGORIES):
            looks_like_category = True
        elif (len(stripped_line.split()) <= 5 and 
            ',' not in stripped_line and 
            ':' not in stripped_line and  # Categories don't have colons
            not stripped_line.startswith('•') and
            stripped_line[0].isupper() and
            not any(tech in stripped_line for tech in ['Java', 'Python', 'React', 'Angular', 'AWS', 'Docker', 'SQL', 'Linux', 'Windows', 'Selenium', 'JIRA', 'Maven', 'Jenkins'])):  # Avoid common tech names
            looks_like_category = True
        
        # If it looks like a category, apply additional checks
        if looks_like_category:
            # If we have a current category and this line is very short (1-2 words),
            # check the next line to disambiguate
            if current_category and len(stripped_line.split()) <= 2:
                # Check if the next line exists and looks like a value (has commas or is longer)
                if i + 1 < len(content_lines):
                    next_line = content_lines[i + 1].strip()
                    if next_line and (',' in next_line or '.' in next_line or len(next_line.split()) > 5):
                        # Next line looks like a value, so current line is probably also a value
                        is_likely_category = False
                    else:
                        # Next line looks like a category, so current line might be a category too
                        # But only if it's in our known list
                        is_likely_category = any(cat.lower() == stripped_line.lower() for cat in LIKELY_CATEGORIES)
                else:
                    # No next line, treat as value if we have a current category
                    is_likely_category = False
            else:
                # If we are already in a category, be stricter about switching to a new one
                # unless it's clearly a known category
                if current_category and any(cat.lower() == current_category.lower() for cat in LIKELY_CATEGORIES):
                    if not any(cat.lower() == stripped_line.lower() for cat in LIKELY_CATEGORIES):
                        # Current is known, New is unknown -> Treat New as VALUE
                        is_likely_category = False
                    else:
                        # Current is known, New is known -> Switch
                        is_likely_category = True
                else:
                    is_likely_category = True
            
        if is_likely_category:
            # Don't save the previous category if it has no value - just discard it
            # This handles cases where categories appear without values
            
            # This is a new category - save it for the next skills line
            current_category = stripped_line
        else:
            # This is likely the skills list for the category
            if current_category:
                # Check if we already have a value for this category
                # If so, append to it; otherwise create new entry
                if formatted_lines and formatted_lines[-1].startswith(f"{current_category}:"):
                    # Append to existing category line
                    formatted_lines[-1] += f", {stripped_line}"
                else:
                    # Format as "Category: skills"
                    formatted_lines.append(f"{current_category}: {stripped_line}")
                # Don't reset current_category - allow multiple values for same category
            else:
                # If no category was found, check if the line already has a colon (is already formatted)
                if ':' in stripped_line:
                    formatted_lines.append(stripped_line)
                else:
                    # Just add the line as is
                    formatted_lines.append(stripped_line)
    
    # Don't append the last category if it has no value - just discard it
    # (This handles cases where the last line is a category without skills)
    
    return "\n".join(formatted_lines)




def process_experience_bullets(content_lines: list) -> list:
    """
    Parses professional experience content, separating company/role headers from bullet points.
    Returns a list of dictionaries, each representing an experience entry.
    """
    experience_entries = []
    current_entry = {
        'header_lines': [],
        'bullets': [],
        'environment': '' # Assuming environment might appear here too
    }
    
    # Combined list of common action verbs for bullet splitting
    bullet_starters = ACTION_VERBS.union([
        "participated", "coordinated", "working", "developed", "designed", "implemented",
        "used", "wrote", "extensively", "prepared", "expertise", "role", "responsibilities"
    ])

    # Regex to detect lines that start with an action verb/starter keyword (case-insensitive, at word boundary)
    # or common bullet characters.
    bullet_start_regex = re.compile(
        r'^(?:' + '|'.join(re.escape(s) for s in ['•', '-', '*']) + r'\s*|' +
        r'\b(?:' + '|'.join(re.escape(s) for s in bullet_starters) + r')\b)',
        re.IGNORECASE
    )

    for line_num, line in enumerate(content_lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue # Skip empty lines

        # Check for start of a new experience entry
        # This is typically signaled by a line with Company, Location, Dates
        if is_experience_company_line(stripped_line):
            if current_entry['header_lines'] or current_entry['bullets']: # If building a previous entry, finalize it
                experience_entries.append(current_entry)
                current_entry = {'header_lines': [], 'bullets': [], 'environment': ''} # Start new entry
            current_entry['header_lines'].append(stripped_line) # Add the company/date line to header
            continue # Move to next line

        # Check for role line
        if is_experience_role_line(stripped_line):
            current_entry['header_lines'].append(stripped_line) # Add role line to header
            continue # Move to next line

        # Check for environment line (optional, if you want to pull it out separately)
        if stripped_line.lower().startswith('environment:'):
            current_entry['environment'] = re.sub(r'^environment\s*:\s*', '', stripped_line, flags=re.I).strip()
            continue # Move to next line

        # Check for the start of a new bullet point within the current responsibilities
        # This is where the core logic for splitting paragraphs into bullets goes.
        # We need to consider that docx2txt often flattens bullets onto fewer lines.
        is_new_bullet_item = False
        if bullet_start_regex.search(stripped_line):
            is_new_bullet_item = True
        
        # Heuristic for multi-line bullets flattened by docx2txt:
        # If the current_entry['bullets'] is empty, and this is the first content line, it's a bullet.
        # If the line starts with a capitalized word, and the previous line didn't end with punctuation
        # (might indicate a sentence split by docx2txt)
        if not is_new_bullet_item and current_entry['bullets']:
            # Check if this line is clearly a continuation of the previous bullet, not a new one
            last_bullet_line = current_entry['bullets'][-1] if current_entry['bullets'] else ""
            if last_bullet_line and not last_bullet_line.strip().endswith(('.', '!', '?')):
                 # If the last line didn't end a sentence, and this line is not a new bullet start,
                 # it's likely a continuation. Concatenate.
                 current_entry['bullets'][-1] += " " + stripped_line
                 continue
            # Else, if it didn't end a sentence, but this line IS a new bullet starter, then it's a new bullet.
        
        # If we reach here, it's either a confirmed new bullet, or the first line of content.
        # Clean potential leading bullet characters from docx2txt before adding.
        clean_bullet = re.sub(r'^[•\-\*]\s*', '', stripped_line).strip()
        if clean_bullet:
            current_entry['bullets'].append(clean_bullet)

    if current_entry['header_lines'] or current_entry['bullets']: # Add the last collected entry
        experience_entries.append(current_entry)
            
    # Final filtering/cleaning for robustness (optional, based on desired output)
    clean_entries = []
    for entry in experience_entries:
        if entry['header_lines'] or [b for b in entry['bullets'] if b]: # Ensure entry has content
            # Clean up the bullets one more time: join multi-line bullets if they were split mid-sentence
            final_bullets_for_entry = []
            temp_bullet = []
            for bullet_line in entry['bullets']:
                if bullet_start_regex.search(bullet_line) and temp_bullet:
                    final_bullets_for_entry.append(" ".join(temp_bullet).strip())
                    temp_bullet = [bullet_line]
                else:
                    temp_bullet.append(bullet_line)
            if temp_bullet:
                final_bullets_for_entry.append(" ".join(temp_bullet).strip())

            entry['bullets'] = [b for b in final_bullets_for_entry if b] # Filter empty final bullets
            clean_entries.append(entry)

    return clean_entries

def process_professional_experience_content(content_lines: list) -> list:
    """
    Attempts to re-introduce bullet points for Professional Experience content.
    It looks for lines starting with an action verb or appearing to be a new logical point.
    """
    processed_bullets = []
    current_bullet = []
    
    # Define keywords that typically start a new bullet point within professional experience
    # This includes common action verbs and phrases.
    bullet_start_keywords = set(ACTION_VERBS) # Re-use your existing ACTION_VERBS set
    bullet_start_keywords.update([
        "participated", "coordinated", "working", "developed", "designed", "implemented",
        "used", "wrote", "extensively", "prepared", "expertise", "role", "responsibilities"
    ])

    for line in content_lines:
        stripped_line = line.strip()
        if not stripped_line:
            if current_bullet: # If there was a bullet being built, close it
                processed_bullets.append(" ".join(current_bullet).strip())
                current_bullet = []
            continue

        # Check if this line looks like the start of a new bullet point
        # Heuristics: Starts with a capitalized word (likely verb), or very short lines
        # or contains explicit bullet characters if they somehow survived extraction
        is_new_bullet_start = False
        first_word = stripped_line.split(' ')[0].strip('.,').lower()
        if first_word in bullet_start_keywords or stripped_line.startswith(('•', '-', '*')):
            is_new_bullet_start = True
        elif stripped_line and stripped_line[0].isupper() and len(stripped_line.split()) < 5: # Short capitalized phrase
            is_new_bullet_start = True
        elif re.search(r'^\s*\w+\s*:', stripped_line): # "Role:", "Responsibilities:"
            is_new_bullet_start = True

        if is_new_bullet_start and current_bullet:
            processed_bullets.append(" ".join(current_bullet).strip())
            current_bullet = [stripped_line]
        else:
            current_bullet.append(stripped_line)
            
    if current_bullet: # Add the last collected bullet
        processed_bullets.append(" ".join(current_bullet).strip())
            
    # Filter out empty or extremely short/irrelevant lines that might have been processed
    return [bullet for bullet in processed_bullets if bullet and len(bullet.split()) > 2]


def parse_experience_section(content_lines: list) -> list:
    """
    Enhanced parsing for professional experience content into structured entries,
    handling multiple formats: company-first, role-first, and multi-line structures.
    """
    experience_entries = []
    i = 0
    
    while i < len(content_lines):
        line = content_lines[i].strip()
        if not line:
            i += 1
            continue
            
        entry_data = {
            'header': '',
                'role': '',
                'responsibilities': [],
                'environment': ''
            }
        
        extra_responsibility = None
        
        # Check if current line has a date (most reliable experience indicator)
        current_has_date = is_date_line(line)
        
        # Check for split date ranges (start date on current line, end date on next line)
        split_date_range = ""
        if not current_has_date and has_start_date(line):
            # Current line has start date, check if next line has end date
            if i + 1 < len(content_lines) and has_end_date(content_lines[i + 1]):
                split_date_range = combine_split_date_range(line, content_lines[i + 1])
                if split_date_range:
                    current_has_date = True  # Treat as if we found a complete date
        
        if current_has_date:
            # Current line has date - check structure
            if split_date_range:
                # Handle split date range case
                dates_part = split_date_range
                # Remove the start date from the current line to get company/role part
                start_date_match = re.search(r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}', line, re.I)
                if start_date_match:
                    before_dates = line[:start_date_match.start()].strip()
                else:
                    before_dates = line.strip()
                before_dates = re.sub(r'\s+', ' ', before_dates).strip()
                date_match = True  # Indicate we have a valid date
                line_has_content_and_dates = len(before_dates.split()) > 3  # Adjusted threshold for split dates
            else:
                # Handle normal single-line date range case
                # Try to match date at end of line first, then anywhere in line
                date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)+\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}))$', line, re.I)
                if not date_match:
                    # Date might be in the middle (merged line case)
                    date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)+\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}))', line, re.I)
                if date_match:
                    dates_part = date_match.group(1)
                    before_dates = line[:date_match.start()].strip()
                    before_dates = re.sub(r'\s+', ' ', before_dates).strip()
                    
                    # Check for merged content (Description | Company)
                    if '|' in before_dates and len(before_dates) > 60:
                        parts = before_dates.rsplit('|', 1)
                        if len(parts) == 2:
                            extra_responsibility = parts[0].strip()
                            before_dates = parts[1].strip()

                    line_has_content_and_dates = len(line.split()) > 6
                else:
                    dates_part = ""
                    before_dates = ""
                    line_has_content_and_dates = False
            
            if date_match and line_has_content_and_dates:
                # Line contains both content and dates
                
                # Check if there's a role line above this company+date line
                role_above = None
                if i > 0:
                    prev_line = content_lines[i-1].strip()
                    if prev_line and not is_date_line(prev_line) and not any(month in prev_line for month in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']):
                        # Previous line might be a role
                        if (len(prev_line.split()) <= 10 and 
                            any(title_word in prev_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test'])):
                            role_above = prev_line
                
                if role_above:
                    # Structure: Role (above) + Company+Date (current)
                    if '|' not in before_dates:
                        entry_data['header'] = f"{before_dates} | {dates_part}"
                    else:
                        entry_data['header'] = f"{before_dates} {dates_part}"
                    entry_data['role'] = role_above
                else:
                    # Structure: Company+Date on same line, look for role below
                    if '|' not in before_dates:
                        entry_data['header'] = f"{before_dates} | {dates_part}"
                    else:
                        entry_data['header'] = f"{before_dates} {dates_part}"
            else:
                # Date is on its own line or short line
                # Look for company and role in surrounding lines
                company_parts = []
                role_line = None
                
                # Look backwards for company and role (skip empty lines but go further back)
                for j in range(i-1, max(i-5, -1), -1):  # Increased range from 3 to 5
                    if j >= 0 and content_lines[j].strip():
                        potential_line = content_lines[j].strip()
                        # Enhanced company detection including common company patterns
                        if (',' in potential_line or 
                            any(suffix in potential_line.lower() for suffix in ['inc', 'corp', 'llc', 'ltd', 'group', 'solutions', 'services', 'client:', 'bank', 'financial', 'partners', 'permanente', 'america']) or
                            any(location in potential_line.lower() for location in ['dallas', 'tx', 'new york', 'charlotte', 'usa', 'india', 'hyderabad', 'california', 'ca']) or
                            # Additional pattern: lowercase company names with locations
                            re.search(r'^[a-z\s]+,\s*[A-Za-z\s,]+$', potential_line)):
                            company_parts.insert(0, potential_line)
                        else:
                            # Could be a role line
                            if (not role_line and 
                                len(potential_line.split()) <= 10 and
                                any(title_word in potential_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test'])):
                                role_line = potential_line
                
                # Assemble header
                if company_parts:
                    company_text = ' '.join(company_parts)
                    if '|' not in company_text and '|' not in line:
                        entry_data['header'] = f"{company_text} | {line}"
                    else:
                        entry_data['header'] = f"{company_text} {line}"
                else:
                    entry_data['header'] = line
                    
                if role_line:
                    entry_data['role'] = role_line
                
                # Add the extracted responsibility if any
                if extra_responsibility:
                    split_bullets = split_experience_line(extra_responsibility)
                    entry_data['responsibilities'].extend(split_bullets)
        
        else:
            # Current line doesn't have date - check if it's a role line followed by company+date
            next_date_line_idx = None
            for j in range(i+1, min(i+4, len(content_lines))):
                if j < len(content_lines) and is_date_line(content_lines[j]):
                    next_date_line_idx = j
                    break
            
            if next_date_line_idx:
                # Found a date line ahead - check if current line is a role
                next_line = content_lines[next_date_line_idx].strip()
                
                # Check if current line looks like a role
                if (len(line.split()) <= 10 and 
                    any(title_word in line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test'])):
                    
                    # This is likely: Role (current) + Company+Date (next)
                    # Try to match date at end of line first, then anywhere in line
                    date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)+\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}))$', next_line, re.I)
                    if not date_match:
                        # Date might be in the middle (merged line case)
                        date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}\s*(?:to|–|-)+\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{2,4}))', next_line, re.I)
                    
                    if date_match:
                        dates_part = date_match.group(1)
                        company_part = next_line[:date_match.start()].strip()
                        company_part = re.sub(r'\s+', ' ', company_part).strip()
                        
                        # Check for merged content (Description | Company)
                        if '|' in company_part and len(company_part) > 60:
                            parts = company_part.rsplit('|', 1)
                            if len(parts) == 2:
                                extra_responsibility = parts[0].strip()
                                company_part = parts[1].strip()
                        
                        if '|' not in company_part:
                            entry_data['header'] = f"{company_part} | {dates_part}"
                        else:
                            entry_data['header'] = f"{company_part} {dates_part}"
                        entry_data['role'] = line
                        
                        # Add the extracted responsibility if any
                        if extra_responsibility:
                            split_bullets = split_experience_line(extra_responsibility)
                            entry_data['responsibilities'].extend(split_bullets)
                        
                        i = next_date_line_idx  # Skip to the date line
                    else:
                        # Skip this line, not a clear experience start
                        i += 1
                        continue
                else:
                    # Skip this line, not a clear experience start
                    i += 1
                    continue
            else:
                # No date found ahead, skip this line
                i += 1
                continue

        # Skip the end date line if we processed a split date range
        if split_date_range:
            i += 1  # Skip the end date line
        
        # Now collect responsibilities until next experience entry
        i += 1
        while i < len(content_lines):
            resp_line = content_lines[i].strip()
            
            # Stop if we hit another experience entry (date line or role line followed by date)
            if resp_line and is_date_line(resp_line):
                break
            
            # Check if this looks like start of next experience (role followed by date)
            if (resp_line and 
                len(resp_line.split()) <= 10 and 
                any(title_word in resp_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'test']) and
                i+1 < len(content_lines) and is_date_line(content_lines[i+1])):
                break
            
            # Collect responsibilities and environment
            if resp_line:
                if resp_line.lower().startswith('environment:'):
                    entry_data['environment'] = re.sub(r'^environment\s*:\s*', '', resp_line, flags=re.I).strip()
                elif resp_line.startswith('Role:'):
                    # Handle "Role:" prefix if not already captured
                    if not entry_data['role']:
                        entry_data['role'] = re.sub(r'^\s*Role:\s*', '', resp_line, flags=re.I).strip()
                elif not resp_line.lower().startswith('responsibilities:'):
                    # Check if this is the first line after header and looks like a role
                    # (and we don't already have a role)
                    if (not entry_data['role'] and 
                        len(entry_data['responsibilities']) == 0 and
                        len(resp_line.split()) <= 10 and
                        any(title_word in resp_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'architect', 'designer', 'tester'])):
                        # This looks like a role line - extract it as the role
                        entry_data['role'] = resp_line.lstrip('•- ').strip()
                    else:
                        # Split the line into multiple bullets if needed
                        split_bullets = split_experience_line(resp_line.lstrip('•- ').strip())
                        entry_data['responsibilities'].extend(split_bullets)
            
            i += 1
        
        # Clean up the data
        entry_data['header'] = entry_data['header'].strip()
        entry_data['role'] = entry_data['role'].strip()
        
        if entry_data['header'] or entry_data['role']:
            experience_entries.append(entry_data)

    return experience_entries

def parse_projects_content(content_lines: list) -> list:
    logger.info("Starting to parse project content with enhanced parser for multiple formats...")
    projects = []
    
    # Enhanced approach: Handle both role-first and company-first structures
    i = 0
    while i < len(content_lines):
        line = content_lines[i].strip()
        if not line:
            i += 1
            continue
            
        project_data = {
            'header': [],
            'responsibilities': [],
            'environment': ''
        }
        
        # Check if current line has a date (most reliable project indicator)
        current_has_date = is_date_line(line)
        
        # Check for split date ranges (start date on current line, end date on next line)
        split_date_range = ""
        if not current_has_date and has_start_date(line):
            # Current line has start date, check if next line has end date
            if i + 1 < len(content_lines) and has_end_date(content_lines[i + 1]):
                split_date_range = combine_split_date_range(line, content_lines[i + 1])
                if split_date_range:
                    current_has_date = True  # Treat as if we found a complete date
        
        if current_has_date:
            # Current line has date - check structure
            if split_date_range:
                # Handle split date range case
                dates_part = split_date_range
                # Remove the start date from the current line to get company/role part
                start_date_match = re.search(r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}', line, re.I)
                if start_date_match:
                    before_dates = line[:start_date_match.start()].strip()
                else:
                    before_dates = line.strip()
                before_dates = re.sub(r'\s+', ' ', before_dates).strip()
                date_match = True  # Indicate we have a valid date
                line_has_content_and_dates = len(before_dates.split()) > 3  # Adjusted threshold for split dates
            else:
                # Handle normal single-line date range case
                date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))$', line, re.I)
                if date_match:
                    dates_part = date_match.group(1)
                    before_dates = line[:date_match.start()].strip()
                    before_dates = re.sub(r'\s+', ' ', before_dates).strip()
                    line_has_content_and_dates = len(line.split()) > 6
                else:
                    dates_part = ""
                    before_dates = ""
                    line_has_content_and_dates = False
            
            if date_match and line_has_content_and_dates:
                # Line contains both content and dates (e.g., "Role ... Date Range" or "Company, Location Date Range")
                # dates_part and before_dates already set above based on split_date_range or normal case
                
                # Determine if before_dates is a role or company by checking patterns
                is_role_line = (len(before_dates.split()) <= 10 and 
                              any(title_word in before_dates.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test']))
                
                is_company_line = (',' in before_dates or 
                                 any(suffix in before_dates.lower() for suffix in ['inc', 'corp', 'llc', 'ltd', 'group', 'solutions', 'services', 'client:', 'bank', 'financial', 'partners', 'permanente', 'america']) or
                                 any(location in before_dates.lower() for location in ['dallas', 'tx', 'new york', 'charlotte', 'usa', 'india', 'hyderabad', 'california', 'ca']))
                
                if is_role_line and not is_company_line:
                    # before_dates is a role, look for company above
                    company_line = None
                    for j in range(i-1, max(i-5, -1), -1):
                        if j >= 0 and content_lines[j].strip():
                            potential_line = content_lines[j].strip()
                            if (',' in potential_line or 
                                any(suffix in potential_line.lower() for suffix in ['inc', 'corp', 'llc', 'ltd', 'group', 'solutions', 'services', 'client:', 'bank', 'financial', 'partners', 'permanente', 'america']) or
                                any(location in potential_line.lower() for location in ['dallas', 'tx', 'new york', 'charlotte', 'usa', 'india', 'hyderabad', 'california', 'ca']) or
                                re.search(r'^[a-z\s]+,\s*[A-Za-z\s,]+$', potential_line)):
                                company_line = potential_line
                                break
                    
                    if company_line:
                        # Structure: Company (above) + Role+Date (current)
                        project_data['header'] = [company_line, before_dates, dates_part]
                    else:
                        # Structure: Role+Date only
                        project_data['header'] = [before_dates, dates_part]
                        
                elif is_company_line:
                    # before_dates is a company, check for role above or below
                    role_above = None
                    if i > 0:
                        prev_line = content_lines[i-1].strip()
                        if (prev_line and not is_date_line(prev_line) and not any(month in prev_line for month in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']) and
                            len(prev_line.split()) <= 10 and
                            any(title_word in prev_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test'])):
                            role_above = prev_line
                    
                    role_below = None
                    if i + 1 < len(content_lines):
                        next_line = content_lines[i + 1].strip()
                        if (next_line and not next_line.lower().startswith('responsibilities:') and 
                            not next_line.startswith('•') and not next_line.startswith('-') and
                            len(next_line.split()) <= 10 and
                            any(title_word in next_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test'])):
                            role_below = next_line
                    
                    if role_above:
                        # Structure: Role (above) + Company+Date (current)
                        project_data['header'] = [before_dates, role_above, dates_part]
                    elif role_below:
                        # Structure: Company+Date (current) + Role (below)
                        project_data['header'] = [before_dates, role_below, dates_part]
                        i += 1  # Skip the role line since we've consumed it
                    else:
                        # Structure: Company+Date on same line (no explicit role found)
                        project_data['header'] = [before_dates, dates_part]
                else:
                    # Ambiguous case - could be either, default to treating as company
                    project_data['header'] = [before_dates, dates_part]
            else:
                # Date is on its own line or short line
                # Look for company and role in previous lines
                company_line = None
                role_line = None
                
                # Look backwards for company (skip empty lines but go further back)
                for j in range(i-1, max(i-5, -1), -1):  # Increased range from 3 to 5
                    if j >= 0 and content_lines[j].strip():
                        potential_line = content_lines[j].strip()
                        # Enhanced company detection including common company patterns
                        if (',' in potential_line or 
                            any(suffix in potential_line.lower() for suffix in ['inc', 'corp', 'llc', 'ltd', 'group', 'solutions', 'services', 'client:', 'bank', 'financial', 'partners', 'permanente', 'america']) or
                            any(location in potential_line.lower() for location in ['dallas', 'tx', 'new york', 'charlotte', 'usa', 'india', 'hyderabad', 'california', 'ca']) or
                            # Additional pattern: lowercase company names with locations
                            re.search(r'^[a-z\s]+,\s*[A-Za-z\s,]+$', potential_line)):
                            if not company_line:
                                company_line = potential_line
                        else:
                            # Could be a role line
                            if (not role_line and 
                                len(potential_line.split()) <= 10 and  # Increased from 8 to 10
                                any(title_word in potential_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'test'])):
                                role_line = potential_line
                
                # Also check for role line after the date
                if not role_line and i + 1 < len(content_lines):
                    next_line = content_lines[i + 1].strip()
                    if (next_line and not next_line.lower().startswith('responsibilities:') and 
                        not next_line.startswith('•') and not next_line.startswith('-') and
                        len(next_line.split()) <= 10 and
                        any(title_word in next_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead', 'junior', 'jr', 'senior', 'sr'])):
                        role_line = next_line
                        i += 1  # Skip the role line since we've consumed it
                
                # Assemble header based on what we found
                if company_line and role_line:
                    project_data['header'] = [company_line, role_line, line]
                elif company_line:
                    project_data['header'] = [company_line, line]
                else:
                    project_data['header'] = [line]
        
        else:
            # Current line doesn't have date - check if next few lines do
            next_date_line_idx = None
            for j in range(i+1, min(i+4, len(content_lines))):
                if j < len(content_lines) and is_date_line(content_lines[j]):
                    next_date_line_idx = j
                    break

            if next_date_line_idx:
                # Found a date line ahead - this could be role-first structure
                next_line = content_lines[next_date_line_idx].strip()
                
                # Check if current line looks like a role
                if (len(line.split()) <= 8 and 
                    any(title_word in line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist', 'consultant', 'manager', 'lead'])):
                    
                    # This is likely: Role (current) + Company+Date (next)
                    date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))$', next_line, re.I)
                    
                    if date_match:
                        dates_part = date_match.group(1)
                        company_part = next_line[:date_match.start()].strip()
                        company_part = re.sub(r'\s+', ' ', company_part).strip()
                        
                        project_data['header'] = [company_part, line, dates_part]
                        i = next_date_line_idx  # Skip to the date line
                    else:
                        # Skip this line, not a clear project start
                        i += 1
                        continue
                else:
                    # Skip this line, not a clear project start
                    i += 1
                    continue
            else:
                # No date found ahead, skip this line
                i += 1
                continue
        
        # Skip the end date line if we processed a split date range
        if split_date_range:
            i += 1  # Skip the end date line
        
        # Now collect responsibilities until next project
        i += 1
        while i < len(content_lines):
            resp_line = content_lines[i].strip()
            
            # Stop if we hit another project (date line or role line followed by date)
            if resp_line and is_date_line(resp_line):
                break
            
            # Check if this looks like start of next project (role followed by date)
            if (resp_line and 
                len(resp_line.split()) <= 8 and 
                any(title_word in resp_line.lower() for title_word in ['engineer', 'developer', 'analyst', 'sdet', 'qa', 'intern', 'specialist']) and
                i+1 < len(content_lines) and is_date_line(content_lines[i+1])):
                break
            
            # Collect responsibilities and environment
            if resp_line:
                if resp_line.lower().startswith('environment:'):
                    project_data['environment'] = re.sub(r'^environment\s*:\s*', '', resp_line, flags=re.I).strip()
                elif not resp_line.lower().startswith('responsibilities:'):
                    project_data['responsibilities'].append(resp_line.lstrip('•- ').strip())
            
            i += 1
        
        # Clean up header
        project_data['header'] = [h.lstrip('•- ').strip() for h in project_data['header'] if h]
        
        if project_data['header']:
            projects.append(project_data)
    
    logger.info(f"Enhanced parsing complete. Found {len(projects)} projects.")
    return projects


def generate_profile_summary(resume_data: dict) -> list:
    """
    Generates a profile summary based on the resume data.
    """
    summary_points = []
    
    # Extract key info
    title = resume_data.get('title', 'Software Engineer')
    years_exp = "8+" # Default or calculate from dates
    
    # Calculate years of experience if possible
    # This is a placeholder - robust calculation would need date parsing from all experience entries
    
    summary_points.append(f"Experienced {title} with over {years_exp} years of expertise in software development lifecycle (SDLC).")
    
    # Add skills-based points
    tech_skills = ""
    for section in resume_data.get('sections', []):
        if section['name'] == 'TECHNICAL SKILLS':
            tech_skills = section['content']
            break
            
    if tech_skills:
        # Extract some top skills
        skills_list = []
        for line in tech_skills.split('\n'):
            if ':' in line:
                skills = line.split(':', 1)[1].strip()
                skills_list.extend([s.strip() for s in skills.split(',')[:3]])
        
        if skills_list:
            top_skills = ", ".join(skills_list[:5])
            summary_points.append(f"Proficient in {top_skills}, with a strong background in building scalable applications.")

    return summary_points

def find_current_project_by_date(resume_text: str) -> tuple:
    """
    Identifies the most recent/current project based on date in the text.
    Returns (start_line, end_line, info_dict)
    """
    lines = resume_text.split('\n')
    current_year = datetime.datetime.now().year
    
    # Regex for dates
    date_pattern = re.compile(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*\d{4}|present|current|now', re.IGNORECASE)
    
    best_start = None
    best_end = None
    best_info = {}
    
    # Scan for the most recent date
    for i, line in enumerate(lines):
        # Check for date patterns
        if date_pattern.search(line):
            # Check if it looks like a header (short line, maybe has company/role)
            if len(line.strip()) < 100 and len(line.strip()) > 5:
                # If it contains "Present" or current year, it's a strong candidate
                if 'present' in line.lower() or 'current' in line.lower() or str(current_year) in line:
                    best_start = i
                    best_info = {'date_match': line.strip()}
                    
                    # Find end (next header or empty lines followed by header)
                    # Simple heuristic: look for next line that looks like a date header
                    for j in range(i + 1, len(lines)):
                        if date_pattern.search(lines[j]) and len(lines[j].strip()) < 100:
                             # Found next date header
                             best_end = j - 1
                             break
                    else:
                        best_end = len(lines)
                    
                    # We found the current one, stop searching
                    return best_start, best_end, best_info

    return None, None, None

def enhance_resume_text(resume_text: str, missing_keywords: set) -> str:
    """
    Enhance resume by adding bullet points for missing keywords in the current job/project.
    """
    
    # Note: Cleanup functions will be applied later in generate_formatted_resume_pdf after all processing
    
    if not missing_keywords:
        logging.info("No missing keywords to enhance")
        return resume_text
    
    # Limit the number of keywords to enhance (prevent too many additions)
    max_keywords = 5
    limited_keywords = list(missing_keywords)[:max_keywords]
    
    logging.info(f"🎯 Starting resume enhancement for {len(limited_keywords)} missing keywords: {limited_keywords}")
    
    # Find the current project by analyzing dates
    current_project_start, current_project_end, current_project_info = find_current_project_by_date(resume_text)
    
    # Initialize target_section for all code paths
    target_section = None
    
    if current_project_start is None:
        logging.warning("❌ Could not identify current project by date - falling back to section-based detection")
        
        # Fallback to original section-based logic
        lines = resume_text.split('\n')
        
        # Check for both sections and prioritize based on content
        has_projects = any('PROJECT' in line.upper() for line in lines)
        has_professional_experience = any('PROFESSIONAL EXPERIENCE' in line.upper() or 'WORK EXPERIENCE' in line.upper() for line in lines)
        
        # Determine target section based on what's available
        if has_projects:
            target_section = "PROJECTS"
            logging.info("🎯 Target: Adding keywords to PROJECTS section (current project)")
        elif has_professional_experience:
            target_section = "PROFESSIONAL EXPERIENCE" 
            logging.info("🎯 Target: Adding keywords to PROFESSIONAL EXPERIENCE/WORK EXPERIENCE section (current position)")
        else:
            logging.warning("❌ No suitable section found for keyword enhancement")
            return resume_text
    else:
        # For date-based detection, we'll use "CURRENT PROJECT" as the target description
        target_section = "CURRENT PROJECT"
        logging.info(f"🎯 Target: Adding keywords to CURRENT PROJECT identified by date analysis")
        logging.info(f"📍 Current project range: lines {current_project_start+1}-{current_project_end+1}")
        logging.info(f"📋 Project info: {current_project_info['date_match']}")
    
    # Generate targeted bullets for current project with missing keywords
    keywords_list = list(missing_keywords)
    
    # Generate focused bullet points for current project
    current_project_bullets = []
    processed_keywords = set()  # Track which keywords we've already processed
    
    if keywords_list:
        # Limit to maximum 5 bullets total to avoid overwhelming the resume  
        max_keywords = min(len(keywords_list), 5)
        limited_keywords = keywords_list[:max_keywords]
        
        logging.info(f"🎯 Generating targeted bullet points for {target_section} with keywords: {limited_keywords}")
        
        for i, keyword in enumerate(limited_keywords, 1):
            # Skip if we've already successfully processed this keyword
            if keyword in processed_keywords:
                logging.info(f"⏭️ Skipping already processed keyword: '{keyword}'")
                continue
            
            # Hard limit: only allow one bullet per keyword
            if len(current_project_bullets) >= len(limited_keywords):
                logging.info(f"🛑 Reached maximum bullets ({len(limited_keywords)}), stopping")
                break
                
            logging.info(f"🔄 Processing keyword {i}/{len(limited_keywords)}: '{keyword}'")
            
            success = False
            try:
                # Generate bullet point for this specific keyword
                project_bullet = openai_service.generate_project_bullet_point(keyword, 1).lstrip("•- ").strip()
                logging.info(f"🤖 Generated raw bullet for '{keyword}': {project_bullet[:100]}...")
                
                # Validate bullet content before adding
                if project_bullet and len(project_bullet) > 15:  # Ensure meaningful content
                    # Ensure the keyword is actually mentioned in the bullet point
                    if keyword.lower() in project_bullet.lower():
                        formatted_bullet = f"• {project_bullet}"
                        current_project_bullets.append(formatted_bullet)
                        processed_keywords.add(keyword)  # Mark as processed
                        success = True
                        logging.info(f"✅ Added bullet for '{keyword}' to {target_section}: {project_bullet[:60]}...")
                    else:
                        logging.warning(f"⚠️ Generated bullet for '{keyword}' doesn't contain the keyword - regenerating...")
                        # Try once more with explicit keyword inclusion
                        project_bullet = openai_service.generate_project_bullet_point(f"{keyword} technology", 1).lstrip("•- ").strip()
                        if keyword.lower() in project_bullet.lower() and len(project_bullet) > 15:
                            formatted_bullet = f"• {project_bullet}"
                            current_project_bullets.append(formatted_bullet)
                            processed_keywords.add(keyword)  # Mark as processed
                            success = True
                            logging.info(f"✅ Added regenerated bullet for '{keyword}': {project_bullet[:60]}...")
                        else:
                            logging.warning(f"❌ Failed to generate valid bullet for '{keyword}' after retry")
                else:
                    logging.warning(f"⚠️ Generated bullet for '{keyword}' was too short ({len(project_bullet)} chars) - skipping")
                    
            except Exception as e:
                logging.error(f"❌ Exception generating bullet for '{keyword}': {e}")
                import traceback
                logging.error(f"Full traceback: {traceback.format_exc()}")
            
            # Ensure we only process each keyword once
            if success:
                logging.info(f"🎯 Successfully processed '{keyword}' - moving to next keyword")
            else:
                logging.warning(f"⚠️ Failed to process '{keyword}' - moving to next keyword")
        
        # Summary of what was generated
        target_info = f"CURRENT PROJECT" if current_project_start is not None else target_section
        logging.info(f"📊 SUMMARY: Generated {len(current_project_bullets)} bullets for {target_info}")
        for i, bullet in enumerate(current_project_bullets, 1):
            # Extract which keyword this bullet is for
            bullet_text = bullet.lower()
            matched_keywords = [kw for kw in limited_keywords if kw.lower() in bullet_text]
            logging.info(f"   {i}. Keywords {matched_keywords}: {bullet[:80]}...")

    if not current_project_bullets:
        target_info = f"CURRENT PROJECT" if current_project_start is not None else target_section
        logging.warning(f"❌ No GPT bullets were generated for {target_info}. Returning original text.")
        logging.warning(f"❌ Original missing keywords were: {list(missing_keywords)}")
        return resume_text

    target_info = f"CURRENT PROJECT" if current_project_start is not None else target_section
    logging.info(f"🎯 Total bullets to add to {target_info}: {len(current_project_bullets)}")
    updated_text = resume_text
    
    # Enhance ONLY the current project/position with missing keyword bullets
    if current_project_bullets:
        if current_project_start is not None:
            # Date-based current project insertion
            logging.info(f"📝 Adding {len(current_project_bullets)} targeted bullets to CURRENT PROJECT...")
            
            lines = updated_text.split('\n')
            
            # Find the best insertion point within the current project
            insert_line = current_project_start
            bullets_found = 0
            role_line_found = False
            
            # Strategy: Find role line by looking for job titles AFTER date information
            date_line_found = False
            
            # First pass: Find where the date information ends
            for i in range(current_project_start, min(current_project_end, len(lines))):
                line = lines[i].strip()
                
                # Check if this line contains date information
                if (any(date_word in line.lower() for date_word in ['present', 'current', 'till date']) or 
                    re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}', line.lower()) or
                    re.search(r'\d{4}\s+to\s+', line.lower())):
                    date_line_found = True
                    logging.info(f"📅 Found date line at {i+1}: {line[:50]}...")
                    # Start looking for role line from the next line
                    role_search_start = i + 1
                    break
            
            # If no clear date line found, start from the beginning of the project
            if not date_line_found:
                role_search_start = current_project_start
            
            # Second pass: Find the role/title line after date information
            for i in range(role_search_start, min(current_project_end, len(lines))):
                line = lines[i].strip()
                
                # Skip empty lines
                if not line:
                    continue
                
                # Skip company/location lines that still have commas
                if (',' in line and any(state in line.upper() for state in ['CA', 'NY', 'TX', 'FL', 'WA', 'IL', 'GA', 'NC', 'VA', 'OH'])):
                    continue
                
                # Skip additional date lines
                if (any(date_word in line.lower() for date_word in ['present', 'current', 'till date']) or 
                    re.search(r'\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}', line.lower())):
                    continue
                
                # Look for role/title line (job titles, but not bullet points)
                if (any(word in line.lower() for word in ['engineer', 'developer', 'analyst', 'manager', 'lead', 'senior', 'sdet', 'qa', 'tester', 'architect', 'consultant', 'specialist', 'coordinator']) and 
                    not line.startswith('•') and not line.startswith('-') and not line.startswith('*') and
                    not (',' in line and len(line.split(',')) >= 2)):  # Avoid company/location lines
                    role_line_found = True
                    insert_line = i + 1  # Insert after the role line
                    logging.info(f"📌 Found role line at {i+1}: {line[:50]}...")
                    break
            
            # Look for existing bullets after the role line
            if role_line_found:
                for i in range(insert_line, min(current_project_end, len(lines))):
                    line = lines[i].strip()
                    if line.startswith('•') or line.startswith('-') or line.startswith('*'):
                        bullets_found += 1
                        insert_line = i + 1  # Insert after the last bullet found
                        logging.info(f"📌 Found existing bullet at line {i+1}: {line[:40]}...")
            
            # If no role line found, fallback to original logic but be more careful
            if not role_line_found:
                # Look for the first non-empty line that could be after header info
                for i in range(current_project_start + 2, min(current_project_end, len(lines))):
                    line = lines[i].strip()
                    if line and not (',' in line and len(line.split(',')) >= 2):  # Skip company/location lines
                        insert_line = i
                        break
            
            logging.info(f"🎯 Inserting bullets at line {insert_line + 1} within current project")
            
            # Insert the bullets
            bullets_text = '\n' + '\n'.join(current_project_bullets) + '\n'
            
            if insert_line < len(lines):
                updated_lines = lines[:insert_line] + [''] + current_project_bullets + [''] + lines[insert_line:]
            else:
                updated_lines = lines + [''] + current_project_bullets + ['']
            
            updated_text = '\n'.join(updated_lines)
            
            # Verify the enhancement for date-based insertion
            original_bullet_count = resume_text.count('•')
            enhanced_bullet_count = updated_text.count('•')
            added_bullets = enhanced_bullet_count - original_bullet_count
            
            logging.info(f"✅ Current project enhancement complete! Original bullets: {original_bullet_count}, Enhanced bullets: {enhanced_bullet_count}, Added: {added_bullets}")
            
            # Verify missing keywords are now included
            enhanced_lower = updated_text.lower()
            found_keywords = []
            still_missing = []
            
            for keyword in missing_keywords:
                if keyword.lower() in enhanced_lower:
                    found_keywords.append(keyword)
                else:
                    still_missing.append(keyword)
            
            logging.info(f"🔍 Keyword verification:")
            logging.info(f"   ✅ Found in enhanced resume: {found_keywords}")
            if still_missing:
                logging.warning(f"   ❌ Still missing: {still_missing}")
            else:
                logging.info(f"   🎉 All missing keywords successfully added!")
            
            # Return enhanced text for date-based insertion
            logging.info(f"📈 Text length: {len(resume_text)} → {len(updated_text)} (+{len(updated_text) - len(resume_text)} chars)")
            if len(updated_text) <= len(resume_text):
                logging.warning("⚠️ GPT enhancement may have failed - no significant increase in text length")
            else:
                logging.info("✅ GPT enhancement appears successful - significant text increase detected")
            
            return updated_text
            
        else:
            # Fallback to section-based insertion (original logic)
            logging.info(f"📝 Adding {len(current_project_bullets)} targeted bullets to {target_section}...")
            
            # Find the target section with more flexible pattern that handles trailing spaces and newlines
            # Handle both exact matches and variations with trailing spaces/punctuation
            if target_section == "PROFESSIONAL EXPERIENCE":
                section_patterns = [
                    r'PROFESSIONAL\s+EXPERIENCE\s*[:\s]*',
                    r'WORK\s+EXPERIENCE\s*[:\s]*',
                    r'EMPLOYMENT\s+HISTORY\s*[:\s]*'
                ]
            elif target_section == "PROJECTS":
                section_patterns = [
                    r'PROJECTS?\s*[:\s]*',
                    r'PROJECT\s+EXPERIENCE\s*[:\s]*',
                    r'KEY\s+PROJECTS?\s*[:\s]*'
                ]
            else:
                section_patterns = [rf'{re.escape(target_section)}\s*[:\s]*']
        
        section_start = None
        for pattern in section_patterns:
            section_start = re.search(pattern, updated_text, re.IGNORECASE)
            if section_start:
                logging.info(f"✅ Found {target_section} section with pattern '{pattern}' at position {section_start.start()}")
                break
        
        if section_start:
            logging.info(f"✅ Found {target_section} section at position {section_start.start()}")
            try:
                # Find the end of the line containing the section heading
                section_line_start = section_start.start()
                section_line_end = updated_text.find('\n', section_line_start)
                if section_line_end == -1:
                    section_line_end = len(updated_text)
                section_start_pos = section_line_end + 1  # Position after the newline
                
                # Find the text after the target section
                after_section = updated_text[section_start_pos:]
                lines = after_section.split('\n')
                
                logging.info(f"📄 Found {len(lines)} lines in {target_section} section")
                
                # Smart insertion: Find the end of the CURRENT project/position bullets
                insert_position = 0
                current_entry_bullets_ended = False
                in_current_entry = False
                header_lines_seen = 0
                
                for i, line in enumerate(lines):
                    stripped_line = line.strip()
                    
                    # Skip empty lines
                    if not stripped_line:
                        continue
                    
                    # Count header lines (company, role, dates) to identify current entry
                    if not in_current_entry and header_lines_seen < 3:
                        header_lines_seen += 1
                        if header_lines_seen >= 2:  # After company and role
                            in_current_entry = True
                            logging.info(f"📋 Identified {target_section} around line {i}")
                        continue
                    
                    # We're in the current project/position area
                    if in_current_entry:
                        # If this is a bullet point, we're still in the current entry's bullets
                        if stripped_line.startswith('•') or stripped_line.startswith('-') or stripped_line.startswith('*'):
                            insert_position = i + 1  # Keep updating to insert after the last bullet
                            logging.info(f"📌 Found existing bullet in {target_section} at line {i}: {stripped_line[:40]}...")
                        # If this looks like a new job/section (company name, dates, etc), stop
                        elif (any(indicator in stripped_line.lower() for indicator in ['limited', 'ltd', 'inc', 'corp', 'pvt', 'technologies', 'systems']) or
                              re.search(r'\d{4}.*\d{4}|\w{3}\s+\d{4}', stripped_line) or  # Date patterns
                              stripped_line.isupper() and len(stripped_line.split()) <= 4):  # Section headers
                            logging.info(f"🛑 Detected end of {target_section} at line {i}: {stripped_line[:40]}...")
                            current_entry_bullets_ended = True
                            break
                        # If we found bullets before and now see non-bullet content, it might be the end
                        elif insert_position > 0:
                            # Allow a few non-bullet lines (like "Environment:" or spacing)
                            if not any(word in stripped_line.lower() for word in ['environment', 'tools', 'technologies']):
                                logging.info(f"🛑 End of {target_section} bullets detected at line {i}")
                                current_entry_bullets_ended = True
                                break
                
                # Validate and adjust insertion position
                if insert_position == 0:
                    # Find a safe position after header information
                    safe_position = 0
                    for i, line in enumerate(lines[:10]):  # Look at first 10 lines only
                        if line.strip() and not line.strip().startswith('•'):
                            safe_position = i + 1
                        if i >= 2:  # After at least 3 lines
                            break
                    insert_position = safe_position
                    logging.info(f"⚠️ No existing bullets found in {target_section}, inserting after header at position {insert_position}")
                else:
                    # Ensure we're not inserting in the middle of existing content
                    while (insert_position < len(lines) and 
                           lines[insert_position].strip() and 
                           not lines[insert_position].strip().startswith('•') and
                           len(lines[insert_position].strip()) < 50):  # Skip short lines that might be headers
                        insert_position += 1
                    logging.info(f"✅ Will insert keyword bullets at position {insert_position} (after {target_section} existing bullets)")
                
                # Validate that we have meaningful bullets to insert
                if not current_project_bullets:
                    logging.warning(f"⚠️ No valid bullets to insert for {target_section}")
                    return resume_text
                
                logging.info(f"📍 Inserting {len(current_project_bullets)} validated bullets at position {insert_position}")
                logging.info(f"📍 Context around insertion point:")
                context_start = max(0, insert_position - 2)
                context_end = min(len(lines), insert_position + 3)
                for i in range(context_start, context_end):
                    marker = " >>> INSERT HERE <<<" if i == insert_position else ""
                    line_text = lines[i] if i < len(lines) else "[END]"
                    logging.info(f"   Line {i}: {line_text[:60]}...{marker}")
                
                # Insert the current project bullets cleanly with proper spacing
                enhanced_lines = lines[:insert_position]
                
                # Add the new bullets with clear marking
                logging.info(f"📝 Adding {len(current_project_bullets)} bullets for missing keywords:")
                for i, bullet in enumerate(current_project_bullets):
                    enhanced_lines.append(bullet)
                    logging.info(f"   Added: {bullet[:70]}...")
                
                # Add remaining lines
                enhanced_lines.extend(lines[insert_position:])
                
                # Reconstruct the text
                enhanced_after_section = '\n'.join(enhanced_lines)
                updated_text = updated_text[:section_start_pos] + enhanced_after_section
                
                # Clean up any empty bullet points that might have been created
                updated_text = re.sub(r'\n\s*•\s*\n', '\n', updated_text)  # Remove empty bullets
                updated_text = re.sub(r'\n\s*•\s*$', '\n', updated_text, flags=re.MULTILINE)  # Remove trailing empty bullets
                updated_text = re.sub(r'^\s*•\s*\n', '', updated_text, flags=re.MULTILINE)  # Remove leading empty bullets
                
                logging.info("🧹 Cleaned up any empty bullet points")
                
                # Verify the enhancement
                original_bullet_count = resume_text.count('•')
                enhanced_bullet_count = updated_text.count('•')
                added_bullets = enhanced_bullet_count - original_bullet_count
                
                logging.info(f"✅ {target_section.title()} enhancement complete! Original bullets: {original_bullet_count}, Enhanced bullets: {enhanced_bullet_count}, Added: {added_bullets}")
                
                if added_bullets != len(current_project_bullets):
                    logging.warning(f"⚠️ Mismatch: Expected to add {len(current_project_bullets)} bullets to {target_section}, but only added {added_bullets}")
                
                # Verify missing keywords are now included
                enhanced_lower = updated_text.lower()
                found_keywords = []
                still_missing = []
                
                for keyword in missing_keywords:
                    if keyword.lower() in enhanced_lower:
                        found_keywords.append(keyword)
                    else:
                        still_missing.append(keyword)
                
                logging.info(f"🔍 Keyword verification:")
                logging.info(f"   ✅ Found in enhanced resume: {found_keywords}")
                if still_missing:
                    logging.warning(f"   ❌ Still missing: {still_missing}")
                else:
                    logging.info(f"   🎉 All missing keywords successfully added!")
                
                return updated_text
                
            except Exception as e:
                logging.error(f"❌ Error inserting bullets into {target_section}: {e}")
                import traceback
                logging.error(traceback.format_exc())
                return resume_text
        else:
            logging.warning(f"⚠️ Could not find {target_section} section to add bullets")
            return resume_text


def clean_resume_formatting_issues(resume_text: str) -> str:
    """
    Cleans up common formatting issues in the generated resume text.
    """
    logging.info("🧹 Starting final resume formatting cleanup...")
    
    # Fix 1: Remove duplicate "Environment:" lines
    # Pattern: Environment: ... \n Environment: ... -> Environment: ...
    lines = resume_text.split('\n')
    cleaned_lines = []
    skip_next = False
    
    for i, line in enumerate(lines):
        if skip_next:
            skip_next = False
            continue
            
        if i < len(lines) - 1:
            current_line = line.strip()
            next_line = lines[i+1].strip()
            
            if current_line.startswith("Environment:") and next_line.startswith("Environment:"):
                # If duplicate environment lines, keep the longer one
                if len(next_line) > len(current_line):
                    cleaned_lines.append(next_line)
                else:
                    cleaned_lines.append(current_line)
                skip_next = True
                logging.info("🔧 Removed duplicate Environment line")
                continue
        
        cleaned_lines.append(line)
    
    resume_text = '\n'.join(cleaned_lines)
    
    # Fix 2: Ensure proper spacing around bullet points
    # Pattern: Text•Text -> Text\n• Text
    lines = resume_text.split('\n')
    cleaned_lines = []
    
    for i, line in enumerate(lines):
        # Fix bullet points stuck to previous text
        if '•' in line and not line.strip().startswith('•'):
            parts = line.split('•')
            if len(parts) == 2 and len(parts[0].strip()) > 0:
                cleaned_lines.append(parts[0].strip())
                cleaned_lines.append(f"• {parts[1].strip()}")
                logging.info("🔧 Fixed stuck bullet point")
                continue
        
        cleaned_lines.append(line)
        i += 1
    
    resume_text = '\n'.join(cleaned_lines)
    
    # Fix 3: Clean up Technical Skills section with improper colons
    tech_skills_match = re.search(
        r'(TECHNICAL SKILLS.*?)(?=\n[A-Z][A-Z\s]+\n|\Z)', 
        resume_text, 
        re.DOTALL | re.IGNORECASE
    )
    
    if tech_skills_match:
        tech_section = tech_skills_match.group(1)
        original_tech_section = tech_section
        
        # Fix patterns like "Redis: Package Manager:" -> "Redis\nPackage Manager:"
        tech_section = re.sub(
            r'([A-Za-z\s/]+):\s*([A-Z][A-Za-z\s]+):\s*\n',
            r'\1\n\2:\n',
            tech_section
        )
        
        # Fix patterns like "ACS: Methodologies:" -> "ACS\nMethodologies:"
        tech_section = re.sub(
            r'([A-Z]{2,}[A-Za-z]*)\s*:\s*([A-Z][A-Za-z\s]+):\s*\n',
            r'\1\n\2:\n',
            tech_section
        )
        
        # Fix patterns like "(TDD): Testing tools:" -> "(TDD)\nTesting tools:"
        tech_section = re.sub(
            r'\(TDD\)\s*:\s*([A-Z][A-Za-z\s]+):\s*\n',
            r'(TDD)\n\1:\n',
            tech_section
        )
        
        if tech_section != original_tech_section:
            resume_text = resume_text.replace(original_tech_section, tech_section)
            logging.info("🔧 Fixed Technical Skills section formatting")
    
    # Fix 4: Remove more empty bullet patterns
    lines = resume_text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        stripped = line.strip()
        
        # Remove completely empty lines with just bullets
        if stripped in ['•', '• ', ' •', '•  ', '  •']:
            logging.info("🔧 Removed additional empty bullet point")
            continue
        
        # Fix role formatting - remove bullet from role lines
        if re.match(r'^\s*•\s*(Senior|Junior|Lead|Principal|API)?\s*(Software|Web|Full Stack|Backend|Frontend)?\s*(Developer|Engineer|Analyst|SDET|Tester|Manager)', line.strip(), re.I):
            line = re.sub(r'^\s*•\s*', '', line)
            logging.info(f"🔧 Removed bullet from role line: {line.strip()}")
        
        cleaned_lines.append(line)
    
    resume_text = '\n'.join(cleaned_lines)
    
    # Fix 5: Clean up project/company header formatting
    resume_text = re.sub(
        r'(\w+)\.\s*([A-Z][\w\s]*-?\s*[A-Z][a-z]+,\s*[A-Z][a-z]+,\s*[A-Z]{2})\s*\n',
        r'\1.\n\2\n',
        resume_text
    )
    
    # Fix 6: Remove excessive empty lines (but preserve some structure)
    resume_text = re.sub(r'\n\s*\n\s*\n\s*\n+', '\n\n', resume_text)
    
    logging.info("✅ Resume formatting cleanup completed")
    return resume_text


def clean_professional_experience_bullets(resume_text: str) -> str:
    """
    Remove inappropriate bullets from company/role lines in Professional Experience section.
    Works with both text and HTML content.
    """
    logging.info("🧹 Cleaning up Professional Experience bullet formatting...")
    
    # Detect if this is HTML content
    is_html = '<html>' in resume_text or '<div>' in resume_text or '<p>' in resume_text
    
    if is_html:
        # For HTML content, remove bullets from specific elements
        import re
        
        # Remove bullets from role/company lines in HTML
        # Pattern: <h3>• Senior Software Developer</h3> -> <h3>Senior Software Developer</h3>
        resume_text = re.sub(r'(<h[1-6][^>]*>)\s*•\s*([^<]+</h[1-6]>)', r'\1\2', resume_text)
        
        # Remove bullets from paragraph elements that look like roles/companies
        resume_text = re.sub(r'(<p[^>]*>)\s*•\s*((?:Senior|Junior|Lead|Principal|API)?\s*(?:Software|Web|Full Stack|Backend|Frontend)?\s*(?:Developer|Engineer|Analyst|SDET|Tester|Manager)[^<]*</p>)', r'\1\2', resume_text, flags=re.IGNORECASE)
        
        logging.info("✅ HTML Professional Experience bullet cleanup completed")
        return resume_text
    
    # Original text-based cleanup for non-HTML content
    lines = resume_text.split('\n')
    in_prof_exp = False
    cleaned_lines = []
    
    for line in lines:
        if 'PROFESSIONAL EXPERIENCE' in line.upper() or 'WORK EXPERIENCE' in line.upper():
            in_prof_exp = True
            cleaned_lines.append(line)
            continue
        elif in_prof_exp and line.strip() and line[0].isupper() and len(line.split()) <= 4:
            # Likely a new section header
            in_prof_exp = False
        
        if in_prof_exp:
            # Check if this line looks like a company or role line with inappropriate bullets
            stripped = line.strip()
            if (stripped.startswith('•') and 
                (re.search(r'(Senior|Junior|Lead|Principal|API)?\s*(Software|Web|Full Stack|Backend|Frontend)?\s*(Developer|Engineer|Analyst|SDET|Tester|Manager)', stripped, re.I) or
                 re.search(r'[A-Za-z\s]+,\s*[A-Za-z\s]+,\s*[A-Z]{2}', stripped))):  # Company location pattern
                line = re.sub(r'^\s*•\s*', '', line)
                logging.info(f"🔧 Removed bullet from role line: {line.strip()}")
        
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)



