# Import required modules from FastAPI and other libraries
import logging
import sys
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

# 1. Configure the root logger to catch everything at DEBUG level.
#    This is important to capture logs from third-party libraries.
logging.basicConfig(
    level=logging.DEBUG,  # Set the lowest level to capture all messages
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler()
    ]
)

# 2. Get the specific loggers for your application and third-party libraries.
#    This allows you to control the verbosity of each part of the system.
app_logger = logging.getLogger(__name__)
app_logger.setLevel(logging.INFO)  # Keep your app's logs at INFO level for clarity

# 3. Set the logging level for noisy third-party libraries.
#    'weasyprint' and 'fontTools' are the libraries generating the detailed font logs.
#    Setting them to DEBUG will produce the output you want.
logging.getLogger("weasyprint").setLevel(logging.DEBUG)
logging.getLogger("fontTools").setLevel(logging.DEBUG)

# 4. Silence overly verbose libraries if needed. (Optional)
#    For example, uvicorn's access logs are useful but can be noisy.
#    We can keep them at INFO level.
logging.getLogger("uvicorn.access").setLevel(logging.INFO)
# --- END: Advanced Logging Configuration ---

try:
    spacy.load("en_core_web_sm")
except OSError:
    logging.info("spaCy model 'en_core_web_sm' not found. Downloading...")
    subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"], check=True)
    logging.info("spaCy model downloaded successfully.")

    
# Initialize FastAPI app
app = FastAPI()

# Load .env from project root
load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env")

# Fetch API key from environment
api_key = os.getenv("OPENAI_API_KEY")

# ✅ Initialize OpenAI client with error handling
if api_key:
    client = OpenAI(api_key=api_key)
    logging.info("✅ OpenAI client initialized successfully with API key.")
else:
    logging.warning("⚠️ Warning: OPENAI_API_KEY not found in environment variables. GPT features will be disabled.")
    client = None

if client:
    try:
        models = client.models.list()
        # Log the first few model IDs as a confirmation
        model_ids = [model.id for model in models.data[:5]]
        logging.info(f"Successfully connected to OpenAI. Available models include: {model_ids}")
    except Exception as e:
        logging.error(f"❌ Could not connect to OpenAI API: {e}")
        client = None
else:
    logging.warning("OpenAI client not initialized - API key required.")

app.debug = True
CACHE_PATH = Path("bullet_cache.json")

# Load cache on startup
if CACHE_PATH.exists():
    with open(CACHE_PATH, "r", encoding="utf-8") as f:
        bullet_cache = json.load(f)
    logging.info(f"✅ Bullet point cache loaded from {CACHE_PATH} with {len(bullet_cache)} items.")
else:
    bullet_cache = {}
    logging.info("No existing cache found. A new cache will be created.")

# Serve static files and HTML templates
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")
templates.env.auto_reload = True  # ✅ force reload

# Create folders for storing uploaded and generated resumes
UPLOAD_DIR = "uploads"
GENERATED_DIR = "generated_resumes"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)
logging.info(f"Required directories '{UPLOAD_DIR}' and '{GENERATED_DIR}' are ready.")

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
    "built", "participated", "coordinated", "involved", "modernized", "dealt", "responsible", 
    "reporting", "wrote", "extensively", "prepared", "expertise",
    "proficient", "experienced", "knowledge", "accustomed", "hands", "strong", "well",
    "education" # Added for education parsing if it uses bullet-like structure
}

# --- NEW/REFINED HELPERS FOR EXPERIENCE PARSING ---
def is_likely_company_location_date(line: str) -> bool:
    line = line.strip()
    
    date_range_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})\b'
    if not re.search(date_range_pattern, line, re.IGNORECASE):
        return False
        
    company_name_pattern = r'^[A-Z][a-zA-Z0-9\s,&.-]+(?:Inc|LLC|Corp|Ltd|Group|Solutions|Technologies)?\b'
    if not re.search(company_name_pattern, line):
        return False

    if re.search(r'^\s*Role:\s*|^\s*Responsibilities:\s*', line, re.IGNORECASE):
        return False
    
    return True

def is_experience_company_line(line: str) -> bool:
    """Detects lines containing a Company, Location, and Date range."""
    # Pattern: Optional (Company Name, ) Optional (City, State) Month Year to Month Year/Present
    # This is a complex pattern to capture common variations.
    # It prioritizes matching a company-like name, then a location, then a date range.
    
    # 1. Look for a date range (most reliable indicator of job entry)
    date_range_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\s*(?:to|–|-)\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})\b'
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
    return bool(re.search(r'^\s*Role:\s*([A-Za-z0-9\s,/\-]+(?:engineer|analyst|developer|manager|contractor|fulltime|sdet|qa|specialist)\b)?', line, re.IGNORECASE))

def is_responsibility_heading(line: str) -> bool:
    return bool(re.search(r'^\s*(?:Responsibilities|Environment):\s*', line, re.IGNORECASE))

# --- NEW FUNCTION: process_general_bullets ---
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
    # This prevents picking up titles/locations from professional summaries or experience sections.
    header_search_area = resume_text[:70] 

    name_pattern = re.search(r'(?i)(?:Name[:\-]?)?\s*([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)', header_search_area)
    email_pattern = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', header_search_area)
    phone_pattern = re.search(r'(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', header_search_area)
    linkedin_pattern = re.search(r'https?://(www\.)?linkedin\.com/in/[^\s]+', header_search_area)
    
    # Apply location and title patterns only within the restricted search area
    location_pattern = re.search(r'(Dallas|Austin|Houston|TX|New York|NY|CA|San Francisco|Chicago)', header_search_area)
    title_pattern = re.search(r'(SDET|QA Engineer|Software Engineer in Test|Automation Engineer)', header_search_area, re.IGNORECASE)

    return {
        "name": name_pattern.group(1) if name_pattern else "Candidate Name",
        "title": title_pattern.group(1).upper() if title_pattern else "",
        "subtitle": "",
        "email": email_pattern.group(0) if email_pattern else "",
        "phone": phone_pattern.group(0) if phone_pattern else "",
        "linkedin": linkedin_pattern.group(0) if linkedin_pattern else "",
        "location": location_pattern.group(0) if location_pattern else ""
    }
# Function to extract plain text from uploaded resume file
def extract_text_from_file(upload_file: UploadFile) -> str:
    if not upload_file.filename:
        logging.error("Upload file has no filename.")
        raise ValueError("Upload file must have a filename")

    file_path = os.path.join(UPLOAD_DIR, upload_file.filename)
    logging.info(f"Saving uploaded file to: {file_path}")
    with open(file_path, "wb") as f:
        shutil.copyfileobj(upload_file.file, f)

    file_ext = os.path.splitext(file_path)[1].lower()
    logging.info(f"Extracting text from '{file_path}' with extension '{file_ext}'")

    if file_ext == ".txt":
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    elif file_ext == ".pdf":
        doc = fitz.open(file_path)
        text = "".join(page.get_text() for page in doc)  # type: ignore
        logging.info(f"Extracted {len(text)} characters from PDF.")
        return text
    elif file_ext == ".docx":
        text = docx2txt.process(file_path)
        logging.info(f"Extracted {len(text)} characters from DOCX.")
        return text
    
    logging.warning(f"Unsupported file type: {file_ext}. Returning empty string.")
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
    logging.info(f"Categorized keywords: {len(hard_skills)} hard skills, {len(soft_skills)} soft skills.")
    return hard_skills, soft_skills

# Perform basic formatting checks on the resume text
def ats_formatting_warnings(resume_text: str, resume_filename: str):
    suggestions = []
    logging.info("Running ATS formatting checks...")

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

    logging.info(f"Found {len(suggestions)} formatting suggestions.")
    return suggestions

# Content-based ATS validations (keywords, action verbs, length, etc.)
def ats_content_warnings(resume_text: str, matched_hard: set, matched_soft: set):
    suggestions = []
    logging.info("Running ATS content checks...")

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

    logging.info(f"Found {len(suggestions)} content suggestions.")
    return suggestions

# Scoring logic based on keyword match and readability
def score_resume(resume_text: str, hard_skills, soft_skills):
    logging.info("Scoring resume against job description keywords...")
    resume_words = set(resume_text.lower().split())
    matched_hard = hard_skills.intersection(resume_words)
    matched_soft = soft_skills.intersection(resume_words)

    hard_score = round(len(matched_hard) / len(hard_skills) * 100, 2) if hard_skills else 0
    soft_score = round(len(matched_soft) / len(soft_skills) * 100, 2) if soft_skills else 0
    search_score = 90 if len(resume_text) > 300 and "-" in resume_text else 50

    final_score = round((0.5 * hard_score) + (0.3 * soft_score) + (0.2 * search_score), 2)
    missing_keywords = (hard_skills | soft_skills) - resume_words

    logging.info(f"Scoring complete. Final Score: {final_score}, Hard: {hard_score}, Soft: {soft_score}, Missing: {len(missing_keywords)} keywords.")
    return final_score, hard_score, soft_score, search_score, ", ".join(sorted(missing_keywords)), matched_hard, matched_soft

def generate_bullet_point_from_gpt(keyword: str) -> str:
    global client
    
    if client is None:
        logging.warning("OpenAI client not configured. Returning placeholder bullet point.")
        return f"• {keyword.title()} experience (OpenAI API key not configured)"

    keyword = keyword.strip().lower()

    # ✅ Return cached result if available
    if keyword in bullet_cache:
        logging.info(f"Cache hit for keyword: '{keyword}'. Returning cached bullet point.")
        return bullet_cache[keyword]

    logging.info(f"Cache miss for '{keyword}'. Calling GPT-4o API.")
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
            logging.warning("GPT response content was null.")
            return f"• {keyword.title()} experience (no content received)"
        bullet = bullet.strip()

        # ✅ Ensure bullet formatting
        if not bullet.startswith("•"):
            bullet = "• " + bullet

        # ✅ Cache and return
        logging.info(f"Successfully generated bullet for '{keyword}'. Caching result.")
        bullet_cache[keyword] = bullet
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(bullet_cache, f, indent=2)

        return bullet

    except Exception as e:
        logging.error(f"❌ GPT API Exception for keyword '{keyword}': {e}", exc_info=True)
        return f"• {keyword.title()} experience (could not fetch GPT response)"
        
# Save optimized resume with suggestions into a downloadable file
def save_optimized_resume(filename: str, resume_text: str, suggestions: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    base_name = os.path.splitext(filename)[0]
    output_path = os.path.join(GENERATED_DIR, f"{base_name}{ext if ext == '.docx' else '.txt'}")
    logging.info(f"Saving optimized resume to {output_path}")

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

    logging.info("Optimized resume saved successfully.")
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
    output_path = os.path.join("generated_resumes", f"{os.path.splitext(filename)[0]}_formatted.docx")
    logging.info(f"Generating formatted DOCX resume: {output_path}")
    doc = Document()

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

    # ✅ Insert the header section
    add_header_section(enhanced_text)

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
    plain_sections = ["EDUCATION","EDUCATION QUALIFICATIONS", "TECHNICAL SKILLS", "CERTIFICATIONS"]

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

    try:
        doc.save(output_path)
        logging.info("Formatted DOCX resume saved successfully.")
    except Exception as e:
        logging.error(f"Failed to save formatted DOCX: {e}", exc_info=True)

    return output_path

# Enhances the resume text by appending keyword content to the Profile Summary and CMS section
def generate_profile_summary(job_title: str, skills: set) -> str:
    logging.info(f"Generating profile summary for job title: {job_title}")
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
    
    logging.info("Profile summary generated.")
    return summary

def enhance_resume_text(resume_text: str, missing_keywords: set) -> str:
    logging.info(f"Enhancing resume text with {len(missing_keywords)} missing keywords.")    # 1. Generate bullet-style achievements using GPT and format them with bullet points
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
    logging.info("Resume text enhanced successfully.")
    return updated_text


# Endpoint to serve upload form
@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    logging.info("GET /upload - Serving upload form.")
    return templates.TemplateResponse("index.html", {"request": request})

# Endpoint to handle resume and job description upload and analysis
@app.post("/upload", response_class=HTMLResponse)
async def upload_resume(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc_text: str = Form(...),
    generate_download: str = Form("no")
):
    logging.info(f"POST /upload - Received new request. Resume: '{resume.filename}', Generate Download: '{generate_download}'")
    try:
        resume_text = extract_text_from_file(resume)
        user_info = add_header_section(resume_text)

        resume_data =   {
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
            logging.warning("Insufficient text from Resume or JD. Aborting analysis.")
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
        logging.info(f"Analysis complete. Final Score: {score}")

        download_link = None
        if generate_download.lower() == "yes":
            logging.info("Download requested. Generating enhanced resume file.")
            enhanced_text = enhance_resume_text(resume_text, set()) 
            output_path = generate_formatted_resume_pdf(filename, enhanced_text, user_info)
            download_link = f"/download/{os.path.basename(output_path)}"
            logging.info(f"Download link created: {download_link}")

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
        logging.error(f"An error occurred during /upload processing: {e}", exc_info=True)
        return templates.TemplateResponse("index.html", {
            "request": request, "score": 0, "suggestions": "An unexpected error occurred. Please check the logs."
        })

# Endpoint to serve the optimized resume file for download
@app.get("/download/{filename}")
async def download_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    logging.info(f"GET /download/{filename} - Attempting to serve file from {path}")
    if os.path.exists(path):
        logging.info("File found. Sending response.")
        return FileResponse(path, media_type="application/octet-stream", filename=filename)
    
    logging.error(f"File not found at path: {path}")
    return {"detail": "File not found"}

# Root endpoint that redirects to the upload form
@app.get("/", include_in_schema=False)
async def root():
    logging.info("GET / - Redirecting to /upload")
    return HTMLResponse('<script>window.location.replace("/upload")</script>')

# --- MODIFIED parse_resume_to_structure function ---
def parse_resume_to_structure(enhanced_text: str, user_info: dict) -> dict:
    resume_data = {
        "name": user_info.get("name", "Candidate Name"),
        "title": user_info.get("title", ""),
        "email": user_info.get("email", ""),
        "phone": user_info.get("phone", ""),
        "linkedin": user_info.get("linkedin", ""),
        "location": user_info.get("location", ""),
        "sections": []
    }

    VALID_SECTIONS = [
        "PROFILE SUMMARY", "PROFESSIONAL SUMMARY", "PROFESSIONAL EXPERIENCE", "PROJECTS",
        "EDUCATION", "TECHNICAL SKILLS", "CERTIFICATIONS", "ACHIEVEMENTS", "EXPERTISE IN"
    ]
    
    sections = {}
    current_section = None
    current_content = []
    
    lines = enhanced_text.split('\n')
    for idx, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line:
            if current_section and current_section not in ["PROJECTS", "PROFESSIONAL EXPERIENCE"]:
                if current_content and current_content[-1] != '\n':
                    current_content.append('\n')
            continue
        
        normalized_stripped_line = re.sub(r'[.:,;!?-]+$', '', stripped_line).strip()
        
        if normalized_stripped_line.upper() in VALID_SECTIONS:
            if current_section and current_content:
                sections[current_section] = current_content.copy()
            
            current_section = normalized_stripped_line.upper()
            current_content = []
            logging.debug(f"Detected section: {current_section}")
        else:
            if current_section:
                current_content.append(line)
    
    if current_section and current_content:
        sections[current_section] = current_content.copy()
    
    logging.debug(f"Parsed sections: {sections.keys()}")

    desired_sections = [
        "PROFESSIONAL SUMMARY", 
        "PROFILE SUMMARY",      
        "EDUCATION", 
        "TECHNICAL SKILLS",
        "CERTIFICATIONS",
        "ACHIEVEMENTS",         
        "PROFESSIONAL EXPERIENCE", 
        "PROJECTS"
    ]
    
    for section_name in desired_sections:
        section_content = sections.get(section_name.upper())
        if section_content:
            if section_name.upper() == 'PROFESSIONAL EXPERIENCE':
                processed_content = parse_experience_section(section_content) # Use specialized parser
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "professional_experience", # Specific type for rendering
                    "content": processed_content # This will be list of dicts
                })
            elif section_name.upper() == 'PROFESSIONAL SUMMARY' or \
                 section_name.upper() == 'ACHIEVEMENTS':
                # Use the simpler general bullet parser for these
                processed_content = process_general_bullets(section_content) # <--- NEW FUNCTION CALL
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "bullets", # Render as regular bullet list
                    "content": processed_content
                })
            elif section_name == 'PROJECTS':
                projects = parse_projects_content(section_content)
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "projects",
                    "content": projects
                })
            else: # For Education, Technical Skills, Certifications
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "plain_text_block",
                    "content": "\n".join([l.strip() for l in section_content if l.strip()])
                })
    
    return resume_data
# --- END MODIFIED parse_resume_to_structure function ---

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
        "used", "wrote", "extensively", "prepared", "expertise", "proficient", "experienced",
        "knowledge", "accustomed", "involved", "hands on", "strong sense", "responsible", "reporting",
        "updated", "automated", "created"
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
# --- NEW FUNCTION: process_professional_experience_content ---
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
# --- END NEW FUNCTION ---
ACTION_VERBS = {
    "developed", "implemented", "designed", "created", "led", "automated", "executed", "optimized", 
    "tested", "analyzed", "built", "streamlined", "coordinated", "improved", "monitored", "reduced",
    "identified", "documented", "enhanced", "configured", "integrated", "validated"
}

# Helper: Detects a line containing a date range.
def is_duration_line(line: str) -> bool:
    """Identifies a line containing a typical project date range."""
    date_pattern = re.search(
        r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})',
        line,
        re.I
    )
    return bool(date_pattern)

# --- PASTE THIS ENTIRE BLOCK INTO YOUR CODE ---

# --- PASTE THIS ENTIRE BLOCK INTO YOUR CODE ---
# This replaces the old parse_projects_content and its helpers.

# Helper 1: Detects a line containing a date range, which signals a project header.
def is_project_header_line(line: str) -> bool:
    """Robustly identifies a line containing a project header date."""
    date_pattern = re.search(
        r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\s*(?:to|–|-)\s*(?:Present|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})',
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

# 1. Add this new, reliable helper function to your code.
def is_date_line(line: str) -> bool:
    """Finds a line containing a date range like 'Mon YYYY - Mon YYYY' or 'Mon YYYY - Present'."""
    # This regex is specific and robust for identifying project date ranges.
    date_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}\s*(?:to|–|-)\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})'
    return bool(re.search(date_pattern, line, re.I))

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

# --- RESTRUCTURED parse_experience_section FUNCTION (more robust state machine) ---
def parse_experience_section(content_lines: list) -> list:
    """
    Parses professional experience content into structured entries,
    distinguishing boldable headers (Company, Dates, Role) from bullet points.
    Strictly expects company/role lines.
    """
    experience_entries = []
    current_entry = None # Will be a dict for the current job experience
    
    bullet_start_detection_regex = re.compile(
        r'^(?:[•\-\*]\s*|' + # Explicit bullet characters
        r'\b(?:' + '|'.join(re.escape(s) for s in ACTION_VERBS) + r')\b(?!\s*:)|' + # Action verb not followed by colon
        r'\b(?:Role|Responsibilities|Environment):\s*)', # Specific headers like "Role:" (should be caught and processed)
        re.IGNORECASE
    )

    for line_num, line in enumerate(content_lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue

        # Step 1: Detect start of a NEW Experience Entry (Company, Location, Date line)
        if is_likely_company_location_date(stripped_line):
            if current_entry: # Save previous entry if exists
                experience_entries.append(current_entry)
            current_entry = {
                'company_info': stripped_line,
                'role_info': '',
                'bullets': [],
                'environment': ''
            }
            continue

        # If we are expecting a new entry but haven't found a company line yet
        if current_entry is None:
            # Skip lines until a company line is found, or if it's a role line right at the start
            if is_experience_role_line(stripped_line):
                # This could be a role line appearing before a company line in a weird format.
                # Treat it as part of the header if it's the very first line of the section,
                # or if a company line wasn't found before.
                current_entry = {
                    'company_info': '', # Company info is missing or embedded in role line
                    'role_info': stripped_line,
                    'bullets': [],
                    'environment': ''
                }
                continue
            else:
                continue # Discard lines until a company or role line is found for a new entry

        # Step 2: Detect Role Line (if not already captured in company line)
        if is_experience_role_line(stripped_line):
            current_entry['role_info'] = stripped_line
            continue

        # Step 3: Detect Environment Line
        if stripped_line.lower().startswith('environment:'):
            current_entry['environment'] = re.sub(r'^environment\s*:\s*', '', stripped_line, flags=re.I).strip()
            continue
        
        # Step 4: Process Bullet Points / Responsibilities
        # Any remaining line should be treated as a bullet point.
        clean_bullet = re.sub(r'^[•\-\*]\s*', '', stripped_line).strip() # Remove explicit bullets
        
        is_new_bullet_start = False
        if bullet_start_detection_regex.search(stripped_line): # Check if it matches a known bullet start pattern
            is_new_bullet_start = True
        elif current_entry['bullets'] and current_entry['bullets'][-1].strip().endswith(('.', '!', '?')):
            # If the previous bullet ended a sentence, the current line (if not just a continuation) is a new bullet
            is_new_bullet_start = True
        
        if is_new_bullet_start and clean_bullet:
            current_entry['bullets'].append(clean_bullet)
        elif current_entry['bullets']: # Concatenate if not new bullet start AND not ended previous sentence
            if not current_entry['bullets'][-1].strip().endswith(('.', '!', '?')):
                current_entry['bullets'][-1] += " " + clean_bullet
            else: # previous bullet ended a sentence, so this must be a new bullet
                current_entry['bullets'].append(clean_bullet)
        else: # First bullet for this entry
            if clean_bullet:
                current_entry['bullets'].append(clean_bullet)

    if current_entry: # Add the last collected entry after the loop finishes
        experience_entries.append(current_entry)

    # Final cleanup and filtering for robustness
    final_parsed_entries = []
    for entry in experience_entries:
        # Filter empty bullets and ensure job entries have at least some info
        filtered_bullets = [b for b in entry['bullets'] if b.strip() and len(b.split()) > 3] # Min 4 words for bullet

        if entry['company_info'] or entry['role_info'] or filtered_bullets:
            entry['bullets'] = filtered_bullets
            final_parsed_entries.append(entry)

    return final_parsed_entries

    
# You can keep your existing `is_date_line` function.

def parse_projects_content(content_lines: list) -> list:
    app_logger.info("Starting to parse project content with new universal parser...")
    projects = []
    project_indices = []

    for i, line in enumerate(content_lines):
        if is_date_line(line):
            project_indices.append(i)

    if not project_indices:
        return []

    for i, date_line_index in enumerate(project_indices):
        current_date_line = content_lines[date_line_index].strip()
        header_lines = []
        start_of_resp_index = date_line_index + 1

        if len(current_date_line.split()) < 6:
            start_of_header_index = max(0, date_line_index - 2)
            header_lines = [content_lines[j].strip() for j in range(start_of_header_index, date_line_index + 1)]
        else:
            header_lines.append(current_date_line)
            if date_line_index + 1 < len(content_lines):
                if not is_date_line(content_lines[date_line_index + 1]):
                    header_lines.append(content_lines[date_line_index + 1].strip())
                    start_of_resp_index += 1

        cleaned_header = [h.lstrip('•- ').strip() for h in header_lines if h]

        if i + 1 < len(project_indices):
            end_of_resp_index = project_indices[i+1]
            next_date_line = content_lines[end_of_resp_index].strip()
            if len(next_date_line.split()) < 6:
                end_of_resp_index = max(0, end_of_resp_index - 2)
        else:
            end_of_resp_index = len(content_lines)

        responsibility_lines = [
            line.strip().lstrip('•- ').strip()
            for line in content_lines[start_of_resp_index:end_of_resp_index]
            if line.strip() and not line.strip().lower().startswith('environment:') and not line.strip().lower().startswith('responsibilities:')
        ]
        
        environment_line = ""
        for line in content_lines[start_of_resp_index:end_of_resp_index]:
             if line.strip().lower().startswith('environment:'):
                 environment_line = re.sub(r'^environment\s*:\s*', '', line.strip(), flags=re.I).strip()
                 break

        projects.append({
            'header': cleaned_header,
            'responsibilities': [line for line in responsibility_lines if line],
            'environment': environment_line
        })

    app_logger.info(f"Finished parsing. Found {len(projects)} projects.")
    return projects

# --- PASTE THIS FUNCTION INTO YOUR CODE ---

# Assuming the rest of your code is unchanged, only the generate_resume_html is modified.

def generate_resume_html(resume_data: dict) -> str:
    template = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { font-family: Calibri, Arial, sans-serif; margin: 0.7in; line-height: 1.2; font-size: 11pt; }
            .header { }
            .header h1 { margin: 0; font-size: 20pt; }
            .header h2 { margin: 0; font-size: 14pt; font-weight: normal; }
            .header p { margin: 2px 0; font-size: 10pt; }
            .section { margin-top: 15px; }
            .section-title { font-size: 14pt; font-weight: bold; color: #4F81BD; border-bottom: 1px solid #B0C4DE; padding-bottom: 3px; margin-bottom: 8px;}
            
            .project { margin-bottom: 18px; }
            
            .project-header { margin-top: 0; padding-top: 0; margin-bottom: 8px; }
            .project-header-line { margin: 0; padding: 0; font-weight: normal; line-height: 1.3; }
            .project-header-line:first-child { margin-top: 0; padding-top: 0; font-weight: bold; font-size: 12pt; }
            .project-title-and-date { font-style: italic; }

            .bullet-list { list-style-position: outside; padding-left: 22px; margin: 0; }
            .bullet-list li { margin-bottom: 6px; }
            .project-environment { margin-top: 8px; padding: 4px; background-color: #F2F2F2; font-size: 9pt; font-style: italic; }
            .project-environment b { font-style: normal; }

            /* Style for plain text content that forms a single block */
            .plain-text-block-content {
                margin-top: 0;
                margin-bottom: 10px;
                white-space: pre-wrap; /* Preserves explicit newlines */
                line-height: 1.3;
            }

            /* New style for Professional Experience headers */
            .experience-main-header {
                font-weight: bold;
                margin-bottom: 5px; /* Space between header and first bullet */
                line-height: 1.3;
                font-size: 11.5pt; /* Slightly larger than normal body text */
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>{{ name }}</h1>
            {% if title %}<h2>{{ title }}</h2>{% endif %}
            <p>
                {% set contact_parts = [] %}
                {% if email %}{% set _ = contact_parts.append(email) %}{% endif %}
                {% if phone %}{% set _ = contact_parts.append(phone) %}{% endif %}
                {% if linkedin %}{% set _ = contact_parts.append('<a href="' ~ linkedin ~ '">' ~ linkedin ~ '</a>') %}{% endif %}
                {% if location %}{% set _ = contact_parts.append(location) %}{% endif %}
                
                {{ contact_parts | join(' | ') }}
            </p>
        </div>
        {% for section in sections %}
        <div class="section">
            <div class="section-title">{{ section.name }}</div>
            
            {% if section.type == 'projects' %}
                {% for project in section.content %}
                <div class="project">
                    <div class="project-header">
                        {%- for line in project.header %}
                            {%- if loop.index == 1 %}
                                <div class="project-header-line">{{ line }}</div>
                            {%- else %}
                                <div class="project-header-line project-title-and-date">{{ line }}</div>
                            {%- endif %}
                        {%- endfor %}
                    </div>

                    {% if project.responsibilities %}
                    <ul class="bullet-list">
                        {% for bullet in project.responsibilities %}
                        <li>{{ bullet }}</li>
                        {% endfor %}
                    </ul>
                    {% endif %}
                    
                    {% if project.environment %}
                    <div class="project-environment"><b>Environment:</b> {{ entry.environment }}</div>
                    {% endif %}
                </div>
                {% endfor %}
            {% elif section.type == 'bullets' %} {# Applies to Professional Summary and Achievements now #}
                <ul class="bullet-list">
                    {% for bullet in section.content %}
                    <li>{{ bullet }}</li>
                    {% endfor %}
                </ul>
            {% elif section.type == 'professional_experience' %} {# NEW RENDERING FOR PROFESSIONAL EXPERIENCE #}
                {% for entry in section.content %}
                    <div class="experience-entry" style="margin-bottom: 15px;">
                        {% if entry.company_info or entry.role_info %}
                            <p class="experience-main-header">
                                {# Render company info and role info, separated by a pipe if both exist #}
                                {% if entry.company_info %}{{ entry.company_info }}{% endif %}
                                {% if entry.company_info and entry.role_info %} | {% endif %}
                                {% if entry.role_info %}{{ entry.role_info }}{% endif %}
                            </p>
                        {% endif %}
                        {% if entry.bullets %}
                            <ul class="bullet-list">
                                {% for bullet in entry.bullets %}
                                <li>{{ bullet }}</li>
                                {% endfor %}
                            </ul>
                        {% endif %}
                        {% if entry.environment %}
                            <div class="project-environment"><b>Environment:</b> {{ entry.environment }}</div>
                        {% endif %}
                    </div>
                {% endfor %}
            {% elif section.type == 'plain_paragraph' or section.type == 'plain_text_block' %}
                <p class="plain-text-block-content">
                    {{ section.content }}
                </p>
            {% endif %}
        </div>
        {% endfor %}
    </body>
    </html>
    """
    from jinja2 import Template
    return Template(template).render(**resume_data)

# Modified generate_formatted_resume_pdf to accept user_info
def generate_formatted_resume_pdf(filename: str, enhanced_text: str, user_info: dict) -> str:
    """Generate PDF resume using HTML template, with a fallback to DOCX."""
    logging.info("Attempting to generate PDF resume.")
    try:
        # Pass user_info to parse_resume_to_structure
        resume_data = parse_resume_to_structure(enhanced_text, user_info) # Modified call
        html_content = generate_resume_html(resume_data)
        
        try:
            # Set library path for macOS if necessary
            if 'darwin' in str(sys.platform):
                os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ.get('DYLD_LIBRARY_PATH', '')
            
            from weasyprint import HTML
            output_path = os.path.join("generated_resumes", f"{os.path.splitext(filename)[0]}_formatted.pdf")
            HTML(string=html_content).write_pdf(output_path)
            logging.info(f"PDF generated successfully: {output_path}")
            return output_path
        except (ImportError, OSError) as e:
            logging.warning(f"weasyprint not available ({e}), falling back to .docx generation.")
            return generate_formatted_resume_docx(filename, enhanced_text)
            
    except Exception as e:
        logging.error(f"PDF generation failed ({e}), falling back to .docx generation.", exc_info=True)
        return generate_formatted_resume_docx(filename, enhanced_text)