# Import required modules from FastAPI and other libraries
from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docx import Document
from docx.shared import Pt, Inches
import openai
import subprocess
import shutil, os, fitz, docx2txt, spacy, re, json
from docx.oxml import OxmlElement
from pathlib import Path
from dotenv import load_dotenv
from docx.oxml.ns import qn
from openai import OpenAI
from .utils import validate_resume_format, extract_keywords  # Reusable utility functions
from docx.enum.style import WD_STYLE_TYPE

try:
    spacy.load("en_core_web_sm")
except OSError:
    subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"], check=True)
    
# Initialize FastAPI app
app = FastAPI()

# Load .env from project root
load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env")

# Fetch API key from environment
api_key = os.getenv("OPENAI_API_KEY")

# ✅ Initialize OpenAI client with error handling
if api_key:
    client = OpenAI(api_key=api_key)
else:
    print("Warning: OPENAI_API_KEY not found in environment variables")
    client = None

if client:
    models = client.models.list()
    for model in models.data:
        print(model.id)
else:
    print("OpenAI client not initialized - API key required")

app.debug = True
load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env")
CACHE_PATH = Path("bullet_cache.json")
#print("Loaded OpenAI Key:", os.getenv("OPENAI_API_KEY"))
print("✅ OpenAI API key loaded.")

# Load cache on startup
if CACHE_PATH.exists():
    with open(CACHE_PATH, "r", encoding="utf-8") as f:
        bullet_cache = json.load(f)
else:
    bullet_cache = {}

# Serve static files and HTML templates
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")
templates.env.auto_reload = True  # ✅ force reload

# Create folders for storing uploaded and generated resumes
UPLOAD_DIR = "uploads"
GENERATED_DIR = "generated_resumes"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)

# Load spaCy model for NLP tasks
nlp = spacy.load("en_core_web_sm")

# Define sets of known hard and soft skills, and action verbs
HARD_KEYWORDS = {
    "python", "java", "javascript", "typescript", "c", "c++", "c#", "ruby", "php", "go", "rust", "kotlin", "swift", "r", "matlab", "scala", "perl", "sql", "bash", "powershell", "dart", "julia",
    "react", "angular", "vue", "node", "django", "flask", "spring", "ruby on rails", "laravel", "asp.net", "express",
    "flutter", "react native", "xamarin", "ionic", "android sdk", "ios sdk", "swiftui", "uikit",
    "tensorflow", "pytorch", "scikit-learn", "keras", "pandas", "numpy", "matplotlib", "seaborn", "apache spark",
    "unity", "unreal engine", "godot",
    "selenium", "junit", "testng", "cypress", "jest", "mocha",
    "mysql", "postgresql", "oracle", "microsoft sql server", "sqlite", "mongodb", "cassandra", "dynamodb", "redis", "couchdb", "snowflake", "bigquery", "redshift",
    "aws", "azure", "gcp", "ibm cloud", "oracle cloud", "ec2", "s3", "lambda", "azure virtual machines", "blob storage", "google compute engine", "cloud storage", "microservices", "serverless", "docker", "kubernetes", "helm", "terraform", "cloudformation", "iam", "key management", "network security groups",
    "jenkins", "gitlab ci/cd", "circleci", "travis ci", "bamboo", "git", "github", "gitlab", "bitbucket", "prometheus", "grafana", "elasticsearch", "logstash", "kibana", "splunk", "ansible", "puppet", "chef", "saltstack",
    "statistical analysis", "data visualization", "etl", "tableau", "power bi", "looker", "qlikview", "hadoop", "apache kafka", "hive", "pig",
    "linear regression", "logistic regression", "decision trees", "random forests", "neural networks", "gradient boosting", "xgboost", "lightgbm", "hugging face", "openai api", "google ai platform", "microsoft cognitive services", "mlflow", "kubeflow", "onnx", "tensorrt",
    "penetration testing", "metasploit", "burp suite", "nmap", "wireshark", "ssl/tls", "oauth", "openid connect", "saml", "kali linux", "nessus", "owasp zap", "ethical hacking", "vulnerability assessment", "incident response", "cryptography", "aes", "rsa",
    "tcp/ip", "dns", "dhcp", "vpn", "firewalls", "load balancing", "linux", "windows server", "active directory", "vmware", "hyper-v", "http", "https", "ftp", "sftp", "snmp", "ldap",
    "html5", "css3", "es6", "webassembly", "figma", "adobe xd", "sketch", "invision", "wcag", "aria",
    "api", "rest", "graphql", "grpc", "soap", "nginx", "apache", "iis", "rabbitmq", "memcached",
    "agile", "scrum", "kanban", "tdd", "bdd", "mvc", "singleton", "factory", "code review", "refactoring",
    "vhdl", "verilog", "assembly", "arduino", "raspberry pi", "fpga", "i2c", "spi", "uart", "rtos",
    "blockchain", "ethereum", "hyperledger", "smart contracts", "solidity", "qiskit", "cirq", "arkit", "arcore", "oculus sdk", "mqtt", "coap", "zigbee",
    "manual testing", "automated testing", "performance testing", "jmeter", "loadrunner", "security testing", "unit testing", "integration testing", "regression testing",
    "technical writing", "api documentation", "swagger", "postman", "regex", "seo", "web performance optimization", "data modeling", "business intelligence", "jira"
}
SOFT_KEYWORDS = {"communication", "leadership", "teamwork", "collaboration", "adaptability", "problem-solving", "critical thinking", "flexibility"}

# --- ADD THIS SET DEFINITION ---
ACTION_VERBS = {
        "designed", "developed", "solved", "continuously", "conducted", "executed", "engaged", 
        "documented", "actively", "managed", "validated", "performed", "applied", "explored", 
        "gained", "contributed", "worked", "utilized", "configured", "implemented", "automated", 
        "created", "monitored", "used", "collaborated", "logged", "integrated", "tested", 
        "prepared", "good", "strong", "use", "demonstrated", "simulated", "analyzed", "tuned",
        "built", "performed", "participated", "verified"
}

# Function to extract plain text from uploaded resume file
def extract_text_from_file(upload_file: UploadFile) -> str:
    if not upload_file.filename:
        raise ValueError("Upload file must have a filename")
    file_path = os.path.join(UPLOAD_DIR, upload_file.filename)
    with open(file_path, "wb") as f:
        shutil.copyfileobj(upload_file.file, f)

    if file_path.endswith(".txt"):
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    elif file_path.endswith(".pdf"):
        doc = fitz.open(file_path)
        return "".join(page.get_text() for page in doc)  # type: ignore
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
        suggestions.append("⚠️ Use consistent date format like 'MM/YYYY' or 'MonthYYYY'.")
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

def generate_bullet_point_from_gpt(keyword: str) -> str:
    global client
    
    if client is None:
        return f"• {keyword.title()} experience (OpenAI API key not configured)"

    keyword = keyword.strip().lower()

    # ✅ Return cached result if available
    if keyword in bullet_cache:
        return bullet_cache[keyword]

    prompt = (
        f"You are a resume expert for QA automation engineers. Write 1 concise, powerful bullet point "
        f"demonstrating real project experience using '{keyword}' (e.g., scripting, test execution, integration). "
        f"Use strong verbs. Avoid general phrases. Limit to 30 words. Output ONLY the bullet."
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=80,
            temperature=0.7,
        )

        bullet = response.choices[0].message.content
        if bullet is None:
            return f"• {keyword.title()} experience (no content received)"
        bullet = bullet.strip()

        # ✅ Ensure bullet formatting
        if not bullet.startswith("•"):
            bullet = "• " + bullet

        # ✅ Cache and return
        bullet_cache[keyword] = bullet
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(bullet_cache, f, indent=2)

        return bullet

    except Exception as e:
        print(f"❌ GPT Exception: {e}")
        return f"• {keyword.title()} experience (could not fetch GPT response)"
        
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

def add_bottom_border(paragraph, color="87CEEB", size="18"):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    p = paragraph._p
    p_borders = OxmlElement('w:pBdr')
    bottom_border = OxmlElement('w:bottom')
    bottom_border.set(qn('w:val'), 'single')
    bottom_border.set(qn('w:sz'), size)
    bottom_border.set(qn('w:space'), '1')
    bottom_border.set(qn('w:color'), color)
    p_borders.append(bottom_border)
    p.get_or_add_pPr().append(p_borders)

def generate_formatted_resume_docx(filename: str, enhanced_text: str) -> str:
    doc = Document()
    output_path = os.path.join("generated_resumes", f"{os.path.splitext(filename)[0]}_formatted.docx")

    # ✅ Add horizontal line below header (No change)
    def add_horizontal_line(doc):
        p = doc.add_paragraph()
        p.alignment = 1  # Centered
        p_paragraph = p._p
        p_borders = OxmlElement('w:pBdr')
        bottom_border = OxmlElement('w:bottom')
        bottom_border.set(qn('w:val'), 'single')
        bottom_border.set(qn('w:sz'), '18')
        bottom_border.set(qn('w:space'), '1')
        bottom_border.set(qn('w:color'), '87CEEB')
        p_borders.append(bottom_border)
        p_props = p_paragraph.get_or_add_pPr()
        p_props.append(p_borders)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)

    # ✅ Add contact header (No change)
    def add_header_section(doc):
        name = "Aravind Penmatsa"
        title = "SDET"
        subtitle = "(JAVA, Selenium, Protractor, Web Services, )"
        email = "aravind.raju541@gmail.com"
        phone = "614-940-9680"
        linkedin = "https://www.linkedin.com/in/aravind-penmatsa/"
        location = "Dallas,TX"

        def add_centered_text(text, bold=False, size=11):
            p = doc.add_paragraph()
            p.alignment = 1  # Centered
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(size)
            p.paragraph_format.space_after = Pt(0)  # Tight spacing
            return p

        add_centered_text(name, bold=True, size=16)
        add_centered_text(title, bold=True, size=12)
        add_centered_text(subtitle)
        add_centered_text(email)
        add_centered_text(phone)
        add_centered_text(linkedin)
        add_centered_text(location)
        doc.add_paragraph()      # spacer

    # ✅ Insert the header section
    add_header_section(doc)

    # ✅ Apply base styles
    style = doc.styles['Normal']
    style.font.name = 'Calibri'  # type: ignore
    style.font.size = Pt(11)  # type: ignore

    # Remove repeated header content from enhanced_text
    header_keywords = {
        "aravind penmatsa", "sdet", "protractor", "614-940-9680",
        "gmail.com", "linkedin.com", "dallas"
    }
    table_keywords = {
        "expertise in", "java", "selenium", "web service/api", "rest assured", "cucumber", "apache poi", "mobile",
        "appium", "jasmine", "karma", "android", "ios", "bdd/tdd", "hybrid", "devops", "aws", "maven", "cicd",
        "git/github", "jenkins", "sql", "sap abap"
    }
    lines = enhanced_text.strip().splitlines()
    # Skip the first line if it matches the name exactly (to avoid stray name above Profile Summary)
    if lines and lines[0].strip().lower() in ["aravind", "aravind penmatsa"]:
        lines = lines[1:]
    filtered_lines = [
        line for i, line in enumerate(lines)
        if not (i < 15 and any(k in line.lower() for k in header_keywords))
        and not any(line.strip().lower().startswith(k) for k in table_keywords)
        and line.strip()
    ]
    # ✅ Build resume body
    current_section = ""
    # Removed "Responsibilities" from skip_keywords as we want to process it
    skip_keywords = {"Environment"} 
    seen_sections = set()

    # Track paragraphs for each section to apply keep_with_next/keep_together
    section_paragraphs = []
    # Track lines for plain-text sections
    plain_section_lines = []
    plain_sections = ["EDUCATION", "TECHNICAL SKILLS", "CERTIFICATIONS"]

    if 'SingleSpace' not in [s.name for s in doc.styles]:
        single_space_style = doc.styles.add_style('SingleSpace', WD_STYLE_TYPE.PARAGRAPH)
        single_space_style.font.name = 'Calibri'  # type: ignore
        single_space_style.font.size = Pt(11)  # type: ignore
        single_space_style.paragraph_format.line_spacing = 1.0  # type: ignore
        single_space_style.paragraph_format.space_after = Pt(0)  # type: ignore
        single_space_style.paragraph_format.space_before = Pt(0)  # type: ignore
    
    project_header_buffer = [] 
    
    for idx, line in enumerate(filtered_lines):
        # Important: Temporarily strip potential leading bullets for header detection
        temp_stripped = line.strip().lstrip('•').lstrip('-').strip() 
        stripped = line.strip()

        # Section headings
        if stripped.isupper() and len(stripped.split()) < 6:
            # Before switching section, flush any collected plain section lines
            if current_section.upper() in plain_sections and plain_section_lines:
                p_section = doc.add_paragraph('\n'.join(plain_section_lines), style='SingleSpace')
                p_section.paragraph_format.space_after = Pt(0)
                p_section.paragraph_format.space_before = Pt(0)
                p_section.paragraph_format.line_spacing = 1.0
                section_paragraphs.append(p_section)
                # Add a blank line for spacing after the section
                p_blank = doc.add_paragraph()
                p_blank.paragraph_format.space_after = Pt(0)
                p_blank.paragraph_format.space_before = Pt(0)
                section_paragraphs.append(p_blank)
                section_paragraphs = []
                plain_section_lines = []
            
            # If we were in PROJECTS and had buffered lines, flush them as bold now
            if current_section.upper() == "PROJECTS" and project_header_buffer:
                for header_line in project_header_buffer:
                    p = doc.add_paragraph()
                    run = p.add_run(header_line)
                    run.bold = True
                    p.paragraph_format.space_after = Pt(0)
                project_header_buffer = [] # Clear buffer after flushing

            current_section = stripped
            # Remove any name or extra paragraph before Profile Summary
            if stripped.upper() == "PROFILE SUMMARY" and len(doc.paragraphs) > 0:
                last_para = doc.paragraphs[-1]
                if last_para.text.strip().lower() in ["aravind", "aravind penmatsa"]:
                    doc.paragraphs.pop()
            p_heading = doc.add_paragraph()
            run_heading = p_heading.add_run(stripped.title())
            run_heading.bold = True
            run_heading.font.size = Pt(14) # Section header size
            run_heading.font.name = "Calibri"
            p_heading.paragraph_format.space_after = Pt(0)
            if stripped.title() != "Projects": 
                add_bottom_border(p_heading)
            
            p_heading.paragraph_format.keep_with_next = True
            p_heading.paragraph_format.keep_together = True
            
            section_paragraphs = [p_heading] # Start new section_paragraphs for this heading
            continue

        # For plain-text sections, collect lines
        if current_section.upper() in plain_sections and stripped:
            plain_section_lines.append(stripped)
            # Only flush at the end of the section or before a new section
            is_last_line = idx + 1 == len(filtered_lines) or \
                             (idx + 1 < len(filtered_lines) and \
                              (filtered_lines[idx + 1].strip().isupper() and len(filtered_lines[idx + 1].strip().split()) < 6)) # Next is new section
            if is_last_line:
                p_section = doc.add_paragraph('\n'.join(plain_section_lines), style='SingleSpace')
                p_section.paragraph_format.space_after = Pt(0)
                p_section.paragraph_format.space_before = Pt(0)
                p_section.paragraph_format.line_spacing = 1.0
                section_paragraphs.append(p_section)
                # Add a blank line for spacing after the section
                p_blank = doc.add_paragraph()
                p_blank.paragraph_format.space_after = Pt(0)
                p_blank.paragraph_format.space_before = Pt(0)
                section_paragraphs.append(p_blank)
                section_paragraphs = []
                plain_section_lines = []
            continue

        # Logic for PROJECTS section
        if current_section.upper() == "PROJECTS":
            # Determine if the current line is a project header component using temp_stripped
            is_project_header_component = is_company_line(temp_stripped) or \
                                          is_title_line(temp_stripped) or \
                                          is_duration_line(temp_stripped) or \
                                          ("Remote" in temp_stripped and len(temp_stripped.split()) < 3) # Specific for the "Remote" line

            # Determine if the current line starts project responsibilities/details
            # Check for "Responsibilities:", explicit bullets, or action verbs on original stripped line
            is_responsibilities_heading = stripped.lower().startswith('responsibilities:')
            is_explicit_bullet = stripped.startswith(('•', '-'))
            is_action_verb_start = stripped and stripped.split()[0].lower().strip(',.:') in ACTION_VERBS

            is_project_content_start = is_responsibilities_heading or is_explicit_bullet or is_action_verb_start

            if is_project_header_component and not is_project_content_start:
                # If it's a header component and not yet content, add to buffer
                # Store the original stripped line in the buffer
                project_header_buffer.append(stripped) 
            elif is_project_content_start or (stripped and not project_header_buffer):
                # If it's the start of content, or if we are processing content lines and buffer is empty
                
                # First, flush any buffered headers as bold
                if project_header_buffer:
                    for header_line in project_header_buffer:
                        p_header = doc.add_paragraph()
                        run_header = p_header.add_run(header_line)
                        run_header.bold = True
                        p_header.paragraph_format.space_after = Pt(0)
                    project_header_buffer = [] # Clear buffer after flushing
                
                # --- START: MODIFIED CODE BLOCK ---
                if is_responsibilities_heading:
                    # The "Responsibilities:" heading itself.
                    p_content = doc.add_paragraph(stripped)
                elif not stripped.lower().startswith('responsibilities:') and (is_explicit_bullet or is_action_verb_start):
                    # This handles lines that are identified as bullet points.
                    # It strips any existing bullet character but does NOT add a new one.
                    line_content = stripped.lstrip('•- ').strip()
                    p_content = doc.add_paragraph(line_content)
                else: 
                    # Fallback for any other lines of content.
                    p_content = doc.add_paragraph(stripped)
                # --- END: MODIFIED CODE BLOCK ---

                p_content.paragraph_format.space_after = Pt(0)
                p_content.paragraph_format.left_indent = Inches(0.25) 
            elif not stripped: # Handle empty lines (for spacing)
                # Only add a blank paragraph if it's not part of an ongoing header block or immediately after a header block
                if not project_header_buffer: 
                    p_blank = doc.add_paragraph()
                    p_blank.paragraph_format.space_after = Pt(0)
                    p_blank.paragraph_format.space_before = Pt(0)
            
            # Any line processed within the PROJECTS section means we continue
            continue 

        # Skip duplicate/irrelevant sections (This part remains largely the same)
        if any(keyword.lower() in stripped.lower() for keyword in skip_keywords):
            if stripped in seen_sections:
                continue
            seen_sections.add(stripped)
            continue  

        # Bullet points (for sections outside PROJECTS) - these lines are not processed if in PROJECTS section
        if stripped.startswith("•") or stripped.startswith("-"):
            p = doc.add_paragraph(stripped, style='List Bullet')
            p.paragraph_format.space_after = Pt(0)  
        elif stripped:
            p = doc.add_paragraph(stripped)
            p.paragraph_format.space_after = Pt(0)  

    # This ensures the header of the VERY LAST project is written correctly if it wasn't flushed before
    if project_header_buffer:
        for header_line in project_header_buffer:
            p = doc.add_paragraph()
            run = p.add_run(header_line)
            run.bold = True
            p.paragraph_format.space_after = Pt(0)
        project_header_buffer = []

    doc.save(output_path)
    return output_path

# Enhances the resume text by appending keyword content to the Profile Summary and CMS section
def generate_profile_summary(job_title: str, skills: set) -> str:
    """
    Generates a professional profile summary with bullet points highlighting key skills and qualifications.
    
    Args:
        job_title: The target job title/role
        skills: Set of relevant skills to highlight
        
    Returns:
        Formatted profile summary text with bullet points
    """
    # Group skills into categories
    technical_skills = [s for s in skills if s.lower() in HARD_KEYWORDS]
    soft_skills = [s for s in skills if s.lower() in SOFT_KEYWORDS]
    
    # Generate base summary
    summary = f"PROFILE SUMMARY\n\n"
    summary += f"Results-driven {job_title} with expertise in "
    
    # Add technical skills
    if technical_skills:
        summary += ", ".join(technical_skills[:3]) + ". "
    
    # Add soft skills
    if soft_skills:
        summary += f"Demonstrated strengths in {', '.join(soft_skills[:2])}. "
    
    # Add bullet points
    summary += "\n\nKey Qualifications:\n"
    
    # Technical bullet points
    for skill in technical_skills[:3]:
        summary += f"• Proficient in {skill} with hands-on experience in development and implementation\n"
    
    # Soft skill bullet points
    for skill in soft_skills[:2]:
        summary += f"• Strong {skill} abilities enabling effective collaboration and project delivery\n"
    
    # Add achievement bullet point
    summary += "• Track record of delivering high-quality solutions that meet business objectives\n"
    
    return summary

def enhance_resume_text(resume_text: str, missing_keywords: set) -> str:
    # 1. Generate bullet-style achievements using GPT and format them with bullet points
    real_world_bullets = [
        generate_bullet_point_from_gpt(k).lstrip("•- ").strip().capitalize() for k in missing_keywords
    ]

    # 2. Insert bullet points at the END of the PROFILE SUMMARY section with normalized spacing
    updated_text = re.sub(
        r"(PROFILE SUMMARY\s*\n)(.*?)(?=\n[A-Z][A-Za-z ]+\n|\Z)",
        lambda m: (
            f"{m.group(1)}"
            + "\n".join(
                [line.strip() for line in m.group(2).strip().splitlines() if line.strip()] +
                ([f"• {bullet}" for bullet in real_world_bullets if bullet.strip()] if real_world_bullets else [])
            )
        ),
        resume_text,
        flags=re.DOTALL | re.IGNORECASE
    )

    # 3. Modify or remove enhancement note to CMS section (Removed as per request)
    # The user specifically asked to "not want to write this point in responsibilities under project 'integrated tools like c'"
    # So, we will remove this block entirely.
    # if missing_keywords:
    #     updated_text = re.sub(
    #         r"(Centers for Medicare & Medicaid Services,.*?Responsibilities:\s+)",
    #         lambda m: m.group(1) + "• Integrated tools like: " + ", ".join(sorted(missing_keywords)) + ".\n",
    #         updated_text,
    #         flags=re.DOTALL | re.IGNORECASE
    #     )

    return updated_text


# Endpoint to serve upload form
@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Endpoint to handle resume and job description upload and analysis
@app.post("/upload", response_class=HTMLResponse)
async def upload_resume(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc_text: str = Form(...),
    generate_download: str = Form("no")
):
    resume_text = extract_text_from_file(resume)
    jd_text = jobdesc_text

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

    # Extract skills
    hard_skills, soft_skills = categorize_keywords(jd_text)
    score, hard_score, soft_score, search_score, missing_keywords, matched_hard, matched_soft = score_resume(resume_text, hard_skills, soft_skills)

    filename = resume.filename or "unknown_file"
    formatting_issues = ats_formatting_warnings(resume_text, filename)
    content_issues = ats_content_warnings(resume_text, matched_hard, matched_soft)
    all_suggestions = formatting_issues + content_issues + [f"Missing Keywords: {missing_keywords}"]
    suggestions = " | ".join(all_suggestions)

    download_link = None
    if generate_download.lower() == "yes":
        # Pass an empty set if you don't want any keywords to be added by enhance_resume_text in the CMS section
        enhanced_text = enhance_resume_text(resume_text, set()) # Pass empty set to prevent "integrated tools like c"
        # This will now attempt PDF generation and fall back to DOCX if needed
        output_path = generate_formatted_resume_pdf(filename, enhanced_text)
        download_link = f"/download/{os.path.basename(output_path)}"

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

def parse_resume_to_structure(enhanced_text: str) -> dict:
    """Parse resume text into structured data for HTML template"""
    # Default header data
    resume_data = {
        "name": "Aravind Penmatsa",
        "title": "SDET",
        "email": "aravind.raju541@gmail.com",
        "phone": "614-940-9680",
        "linkedin": "https://www.linkedin.com/in/aravind-penmatsa/",
        "location": "Dallas,TX",
        "sections": []
    }
    
    # ✅ Define the desired section order
    desired_sections = [
        "PROFILE SUMMARY",
        "EDUCATION", 
        "TECHNICAL SKILLS",
        "CERTIFICATIONS",
        "PROFESSIONAL EXPERIENCE",
        "PROJECTS"
    ]
    
    # Parse all sections first
    sections = {}
    current_section = None
    current_content = []
    
    for line in enhanced_text.split('\n'):
        line = line.strip()
        if not line:
            continue
            
        # Detect section headers
        if line.isupper() and len(line.split()) < 6:
            if current_section and current_content:
                sections[current_section.upper()] = current_content.copy()
            current_section = line
            current_content = []
        else:
            if current_section:
                current_content.append(line)
    
    # Save the last section
    if current_section and current_content:
        sections[current_section.upper()] = current_content.copy()
    
    # ✅ Process sections in the desired order
    for section_name in desired_sections:
        if section_name in sections:
            if section_name == 'PROJECTS':
                # Parse projects into structured format
                projects = parse_projects_content(sections[section_name])
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "projects",
                    "content": projects
                })
            else:
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "plain",
                    "content": sections[section_name]
                })
    
    return resume_data

ACTION_VERBS = {
    "developed", "implemented", "designed", "created", "led", "automated", "executed", "optimized", 
    "tested", "analyzed", "built", "streamlined", "coordinated", "improved", "monitored", "reduced",
    "identified", "documented", "enhanced", "configured", "integrated", "validated"
}

def is_company_line(line: str) -> bool:
    return any(keyword in line.lower() for keyword in ["inc", "llc", "technologies", "solutions", "corp", "company", "services", "labs"])

def is_title_line(line: str) -> bool:
    return any(title in line.lower() for title in ["sdet", "qa", "test", "tester", "engineer", "analyst", "lead"])

def is_duration_line(line: str) -> bool:
    return bool(re.search(r"\b(20\d{2})\b", line)) and bool(re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)", line, re.IGNORECASE))


def parse_projects_content(lines):
    projects = []
    current_project = None
    capture_unstyled_bullets = False

    known_project_headers = [
        "Centers for Medicare & Medicaid Services, Dallas, TX",
        "Dish Network, Denver, CO",
        "United Services Automobile Association (USAA), San Antonio, TX",
        "Infolob Solutions , Irving , TX"
    ]

    def finish_project():
        nonlocal current_project, projects, capture_unstyled_bullets
        if current_project:
            if current_project.get("company") or current_project.get("responsibilities"):
                projects.append(current_project)
        current_project = None
        capture_unstyled_bullets = False

    for line in lines:
        line = line.strip().replace("\t", "")
        if not line:
            continue

        if any(line.startswith(h) for h in known_project_headers):
            finish_project()
            current_project = {
                "company": line,
                "title": "",
                "duration": "",
                "responsibilities": []
            }
        elif is_title_line(line):
            if current_project and not current_project.get("title"):
                current_project["title"] = line
        elif is_duration_line(line):
            if current_project and not current_project.get("duration"):
                current_project["duration"] = line
        elif line.lower().startswith("responsibilities"):
            capture_unstyled_bullets = True
        elif line.startswith("•") or line.startswith("-"):
            if current_project:
                clean_bullet = re.sub(r'^[•\-]+', '', line).strip()
                current_project["responsibilities"].append(clean_bullet)
        elif capture_unstyled_bullets:
            if current_project:
                current_project["responsibilities"].append(line.strip())
        else:
            if current_project and current_project.get("responsibilities"):
                current_project["responsibilities"][-1] += " " + line

    finish_project()
    return projects

def generate_resume_html(resume_data: dict) -> str:
    """Generate HTML resume that can be converted to PDF"""
    template = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { font-family: Arial, sans-serif; margin: 1in; }
            .header { text-align: center; border-bottom: 2px solid #87CEEB; }
            .section { margin-top: 20px; }
            .section-title { font-size: 16px; font-weight: bold; border-bottom: 1px solid #ccc; }
            .project-header { font-weight: normal; }
            .bullet { margin-left: 20px; }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>{{ name }}</h1>
            <h2>{{ title }}</h2>
            <p>{{ email }} | {{ phone }} | {{ linkedin }} | {{ location }}</p>
        </div>
        
        {% for section in sections %}
        <div class="section">
            <div class="section-title">{{ section.name }}</div>
            {% if section.type == 'projects' %}
                {% for project in section.content %}
                <div class="project-header">{{ project.company }}, {{ project.location }} | {{ project.title }} | {{ project.duration }}</div>
                {% for bullet in project.responsibilities %}
                <div class="bullet">• {{ bullet }}</div>
                {% endfor %}
                {% endfor %}
            {% else %}
                {% for item in section.content %}
                <div>{{ item }}</div>
                {% endfor %}
            {% endif %}
        </div>
        {% endfor %}
    </body>
    </html>
    """
    
    from jinja2 import Template
    return Template(template).render(**resume_data)

def generate_formatted_resume_pdf(filename: str, enhanced_text: str) -> str:
    """Generate PDF resume using HTML template, with a fallback to DOCX."""
    try:
        # First, attempt to generate a PDF
        resume_data = parse_resume_to_structure(enhanced_text)
        html_content = generate_resume_html(resume_data)
        
        # Import weasyprint inside try block to handle missing dependencies
        try:
            from weasyprint import HTML
            output_path = os.path.join("generated_resumes", f"{os.path.splitext(filename)[0]}_formatted.pdf")
            HTML(string=html_content).write_pdf(output_path)
            return output_path
        except (ImportError, OSError) as e:
            print(f"weasyprint not available ({e}), falling back to .docx generation.")
            return generate_formatted_resume_docx(filename, enhanced_text)
    except Exception as e:
        print(f"PDF generation failed ({e}), falling back to .docx generation.")
        return generate_formatted_resume_docx(filename, enhanced_text)