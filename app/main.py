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
from pathlib import Path
from dotenv import load_dotenv
from openai import OpenAI
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

# Test WeasyPrint availability at startup
try:
    import weasyprint
    logging.info("✅ WeasyPrint is available for PDF generation")
    
    # Test basic functionality
    try:
        from weasyprint import HTML
        test_html = "<html><body><p>WeasyPrint test</p></body></html>"
        HTML(string=test_html).write_pdf("/tmp/weasyprint_startup_test.pdf")
        logging.info("✅ WeasyPrint startup test successful")
    except Exception as test_error:
        logging.warning(f"⚠️ WeasyPrint available but test failed: {test_error}")
        
except ImportError as e:
    logging.warning(f"⚠️ WeasyPrint not available: {e}. Will use ReportLab fallback.")
except Exception as e:
    logging.error(f"❌ WeasyPrint error during startup: {e}")
    
# Test ReportLab availability
try:
    import reportlab
    logging.info("✅ ReportLab is available for fallback PDF generation")
except ImportError:
    logging.error("❌ ReportLab not available - PDF generation may fail!")

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

# Create a temporary directory for storing uploaded and generated resumes
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
GENERATED_DIR = BASE_DIR / "generated_resumes"

# Ensure directories exist
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
    "education", "engineered", "migrated", "created", "integrated", "automated", "designed", 
    "architected", "facilitated", "strengthened", "implemented", "conducted", "built", 
    "authored", "piloted", "refactored", "managed", "deployed", "delivered", "converted", 
    "maintained", "assisted", "contributed" # Added more from Akhil's resume
}

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

# --- NEW/REFINED HELPERS FOR EXPERIENCE PARSING ---
def is_likely_company_location_date(line: str) -> bool:
    line = line.strip()
    
    # Pattern to find a date range (e.g., "May 2023 - Present" or "September2023")
    date_range_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4})\b'
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
    date_range_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4})\b'
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
        f"Use strong verbs. Avoid general phrases. Limit to 30 words. Output ONLY the bullet"
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

def generate_summary_bullet_from_gpt(keyword: str, variation: int = 1) -> str:
    """Generate varied bullet points for Profile/Professional Summary section."""
    global client
    
    if client is None:
        logging.warning("OpenAI client not configured. Returning placeholder bullet point.")
        return f"• Experienced with {keyword.title()} technologies and best practices (OpenAI API key not configured)"

    keyword = keyword.strip().lower()
    
    # Use variation-specific cache key
    summary_cache_key = f"summary_{keyword}_v{variation}"
    
    # ✅ Return cached result if available
    if summary_cache_key in bullet_cache:
        logging.info(f"Cache hit for summary keyword: '{keyword}' variation {variation}. Returning cached bullet point.")
        return bullet_cache[summary_cache_key]

    logging.info(f"Cache miss for summary keyword '{keyword}' variation {variation}. Calling GPT-4o API.")
    
    # Different prompts for variations to ensure diversity
    prompts = [
        f"Write a professional summary bullet point highlighting expertise and experience with '{keyword}'. Focus on skills and capabilities. Use confident language. Limit to 30 words.",
        f"Write a professional summary bullet point showing proven track record and proficiency in '{keyword}'. Emphasize achievements and competence. Limit to 30 words.", 
        f"Write a professional summary bullet point demonstrating advanced knowledge and hands-on experience with '{keyword}'. Focus on practical application. Limit to 30 words."
    ]
    
    prompt = prompts[(variation - 1) % len(prompts)]

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=80,
            temperature=0.8,  # Higher temperature for more variation
        )

        bullet = response.choices[0].message.content
        if bullet is None:
            logging.warning("GPT response content was null.")
            return f"• Experienced with {keyword.title()} technologies and methodologies (no content received)"
        bullet = bullet.strip()

        # ✅ Ensure bullet formatting
        if not bullet.startswith("•"):
            bullet = "• " + bullet

        # ✅ Cache and return
        logging.info(f"Successfully generated summary bullet for '{keyword}' variation {variation}. Caching result.")
        bullet_cache[summary_cache_key] = bullet
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(bullet_cache, f, indent=2)

        return bullet

    except Exception as e:
        logging.error(f"❌ GPT API Exception for summary keyword '{keyword}' variation {variation}: {e}", exc_info=True)
        return f"• Experienced with {keyword.title()} technologies and methodologies (could not fetch GPT response)"

def generate_project_bullet_point_from_gpt(keyword: str, variation: int = 1) -> str:
    """Generate varied project-specific bullet points for missing keywords."""
    global client
    
    if client is None:
        logging.warning("OpenAI client not configured. Returning placeholder bullet point.")
        return f"• Implemented {keyword.title()} solutions in project development (OpenAI API key not configured)"

    keyword = keyword.strip().lower()
    
    # Use variation-specific cache key for project bullets
    project_cache_key = f"project_{keyword}_v{variation}"
    
    # ✅ Return cached result if available
    if project_cache_key in bullet_cache:
        logging.info(f"Cache hit for project keyword: '{keyword}' variation {variation}. Returning cached bullet point.")
        return bullet_cache[project_cache_key]

    logging.info(f"Cache miss for project keyword '{keyword}' variation {variation}. Calling GPT-4o API.")
    
    # Different prompts for variations to ensure diversity
    prompts = [
        f"Write a project bullet point showing specific implementation and development work with '{keyword}'. Focus on technical execution and results. Use action verbs. Limit to 35 words.",
        f"Write a project bullet point highlighting optimization, integration, or enhancement achieved using '{keyword}'. Emphasize improvements and problem-solving. Limit to 35 words.",
        f"Write a project bullet point demonstrating testing, deployment, or automation involving '{keyword}'. Focus on technical processes and quality assurance. Limit to 35 words."
    ]
    
    prompt = prompts[(variation - 1) % len(prompts)]

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=90,
            temperature=0.8,  # Higher temperature for more variation
        )

        bullet = response.choices[0].message.content
        if bullet is None:
            logging.warning("GPT response content was null.")
            return f"• Implemented {keyword.title()} solutions to enhance project functionality (no content received)"
        bullet = bullet.strip()

        # ✅ Ensure bullet formatting
        if not bullet.startswith("•"):
            bullet = "• " + bullet

        # ✅ Cache and return
        logging.info(f"Successfully generated project bullet for '{keyword}' variation {variation}. Caching result.")
        bullet_cache[project_cache_key] = bullet
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(bullet_cache, f, indent=2)

        return bullet

    except Exception as e:
        logging.error(f"❌ GPT API Exception for project keyword '{keyword}' variation {variation}: {e}", exc_info=True)
        return f"• Implemented {keyword.title()} solutions to enhance project functionality (could not fetch GPT response)"
        
# Save optimized resume with suggestions into a downloadable file
def save_optimized_resume(filename: str, resume_text: str, suggestions: str) -> str:
    base_name = os.path.splitext(filename)[0]
    output_path = os.path.join(GENERATED_DIR, f"{base_name}_optimized.pdf")
    logging.info(f"Saving optimized resume to {output_path}")

    # Use WeasyPrint to generate a PDF with the resume text and suggestions
    try:
        from weasyprint import HTML
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{ font-family: Calibri, Arial, sans-serif; margin: 0.7in; line-height: 1.2; font-size: 11pt; }}
                h1 {{ font-size: 14pt; }}
                h2 {{ font-size: 12pt; }}
            </style>
        </head>
        <body>
            <h1>Optimized Resume Content</h1>
            <pre>{resume_text}</pre>
            <h2>Suggestions</h2>
            <p>{suggestions}</p>
        </body>
        </html>
        """
        if 'darwin' in str(sys.platform):
            os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ.get('DYLD_LIBRARY_PATH', '')
        HTML(string=html_content).write_pdf(output_path)
        logging.info("Optimized resume PDF saved successfully.")
        return output_path
    except Exception as e:
        logging.error(f"Failed to save optimized resume PDF: {e}", exc_info=True)
        raise Exception("PDF generation failed. Please ensure WeasyPrint is properly installed and configured.")

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
            .section-title { font-size: 14pt; font-weight: bold; color: #4F81BD; border-bottom: 1px solid #B0C4DE; padding-bottom: 3px; margin-bottom: 8px; }
            
            .experience-entry { margin-bottom: 18px; }
            .experience-header { 
                margin: 0; 
                padding: 0; 
                display: flex; 
                justify-content: space-between; 
                width: 100%; 
            }
            .experience-header .company-location { 
                margin-right: auto; 
                padding-right: 20px; 
                font-weight: bold; 
            }
            .experience-header .duration { 
                text-align: right; 
                flex-shrink: 0; 
                white-space: nowrap; 
                font-weight: bold;
                letter-spacing: normal;
                word-spacing: normal;
                min-width: 150px;
            }
            .experience-role { 
                margin: 5px 0 5px 0; 
                font-style: italic; 
                font-weight: bold; 
            }
            .project-company {
                font-weight: bold;
                margin: 0;
                padding: 0;
            }
            .project-duration {
                margin: 2px 0;
                padding: 0;
                font-size: 10pt;
                font-weight: bold;
            }
            .project-role {
                font-weight: bold;
                font-style: italic;
                margin: 2px 0 8px 0;
                padding: 0;
            }
            .bullet-list { list-style-position: outside; padding-left: 22px; margin: 0; }
            .bullet-list li { margin-bottom: 6px; }
            .project-environment { margin-top: 8px; padding: 4px; background-color: #F2F2F2; font-size: 9pt; font-style: italic; }
            .project-environment b { font-style: normal; }
            .plain-text-block-content { 
                margin-top: 0; 
                margin-bottom: 10px; 
                white-space: pre-wrap; 
                line-height: 1.3; 
                padding-left: 0;
                text-align: left;
            }
            .skills-table { 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 5px;
            }
            .skills-table td { 
                padding: 3px 0; 
                vertical-align: top; 
                line-height: 1.3;
            }
            .skills-table td:first-child { 
                font-weight: bold; 
                width: 200px;
                padding-right: 15px;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>{{ name }}</h1>
            {% if title %}<h2>{{ title }}</h2>{% endif %}
            {% if subtitle %}<h3 style="font-size: 12pt; font-weight: normal; margin: 2px 0;">{{ subtitle }}</h3>{% endif %}
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
            {# Rename Projects to Professional Experience #}
            {% if section.name == 'Projects' %}
                <div class="section-title">Professional Experience</div>
            {% else %}
            <div class="section-title">{{ section.name }}</div>
            {% endif %}
            {% if section.type == 'professional_experience' %}
                {% for entry in section.content %}
                <div class="experience-entry">
                    {% if entry.header %}
                        {# Split header to separate company from duration #}
                        {% set header_parts = entry.header.split('|') %}
                        {% if header_parts|length >= 2 %}
                            {% set company_location = header_parts[0].strip() %}
                            {% set duration_raw = header_parts[1].strip() %}
                            {% set duration = duration_raw | replace('  ', ' ') | replace('  ', ' ') | trim %}
                        {% else %}
                            {# Fallback if no | separator #}
                            {% set header_words = entry.header.split() %}
                            {% set duration_raw = header_words[-3:] | join(' ') %}
                            {% set duration = duration_raw | replace('  ', ' ') | replace('  ', ' ') | trim %}
                            {% set company_location = ' '.join(header_words[:-3]) %}
                    {% endif %}
                        
                        {# 3-line format: Company/Location, Duration, Role #}
                        <div class="project-company"><b>{{ company_location }}</b></div>
                        <div class="project-duration">{{ duration }}</div>
                    {% if entry.role %}
                            <div class="project-role"><b>{{ entry.role }}</b></div>
                        {% endif %}
                    {% endif %}
                    {% if entry.responsibilities %}
                        <ul class="bullet-list">
                            {% for bullet in entry.responsibilities %}
                            <li>{{ bullet }}</li>
                            {% endfor %}
                        </ul>
                    {% endif %}
                    {% if entry.environment %}
                        <div class="project-environment"><b>Environment:</b> {{ entry.environment }}</div>
                    {% endif %}
                </div>
                {% endfor %}
            {% elif section.type == 'projects' %}
                {# Render projects with 3-line format: Company/Location, Duration, Role #}
                {% for entry in section.content %}
                <div class="experience-entry">
                    {% if entry.header %}
                        {# Handle project header format #}
                        {% if entry.header|length > 0 %}
                            {% set company_line = entry.header[0] if entry.header|length > 0 else '' %}
                            {% set role_line = entry.header[1] if entry.header|length > 1 else '' %}
                            {% set date_line_raw = entry.header[2] if entry.header|length > 2 else entry.header[-1] %}
                            {% set date_line = date_line_raw | replace('  ', ' ') | replace('  ', ' ') | trim %}
                            
                            {# 3-line format: Company/Location, Duration, Role #}
                            <div class="project-company"><b>{{ company_line }}</b></div>
                            <div class="project-duration">{{ date_line }}</div>
                            {% if role_line and role_line != date_line %}
                                <div class="project-role"><b>{{ role_line }}</b></div>
                            {% endif %}
                        {% endif %}
                    {% endif %}
                    {% if entry.responsibilities %}
                        <ul class="bullet-list">
                            {% for bullet in entry.responsibilities %}
                            <li>{{ bullet }}</li>
                            {% endfor %}
                        </ul>
                    {% endif %}
                    {% if entry.environment %}
                        <div class="project-environment"><b>Environment:</b> {{ entry.environment }}</div>
                    {% endif %}
                </div>
                {% endfor %}
            {% elif section.type == 'bullets' %}
                <ul class="bullet-list">
                    {% for bullet in section.content %}
                    <li>{{ bullet }}</li>
                    {% endfor %}
                </ul>
            {% elif section.type == 'plain_paragraph' or section.type == 'plain_text_block' %}
                {% if section.name == 'Technical Skills' and ':' in section.content %}
                    {# Render as table if it contains key-value pairs #}
                    <table class="skills-table">
                    {% for line in section.content.split('\n') %}
                        {% if line.strip() and ':' in line %}
                            {% set parts = line.split(':', 1) %}
                            <tr>
                                <td>{{ parts[0].strip() }}:</td>
                                <td>{{ parts[1].strip() }}</td>
                            </tr>
                        {% elif line.strip() %}
                            <tr>
                                <td colspan="2">{{ line.strip() }}</td>
                            </tr>
                        {% endif %}
                    {% endfor %}
                    </table>
                {% else %}
                    {# For other plain text, check if it should be rendered as continuous text #}
                    {% if section.name == 'Profile Summary' %}
                        <p style="margin: 10px 0; line-height: 1.5;">{{ section.content }}</p>
                    {% else %}
                        <p class="plain-text-block-content">{{ section.content }}</p>
                    {% endif %}
                {% endif %}
            {% endif %}
        </div>
        {% endfor %}
    </body>
    </html>
    """
    from jinja2 import Template
    return Template(template).render(**resume_data)

# Modified to ensure PDF-only output
def generate_formatted_resume_pdf(filename: str, enhanced_text: str, user_info: dict) -> str:
    """Generate PDF resume using HTML template, enforcing PDF output."""
    logging.info("Attempting to generate PDF resume.")
    try:
        # Parse resume data using user_info
        resume_data = parse_resume_to_structure(enhanced_text, user_info)
        html_content = generate_resume_html(resume_data)
        
        # Set library path for macOS if necessary
        if 'darwin' in str(sys.platform):
            os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ.get('DYLD_LIBRARY_PATH', '')
        
        output_path = os.path.join(GENERATED_DIR, f"{os.path.splitext(filename)[0]}_formatted.pdf")
        
        # Try WeasyPrint first with enhanced error handling
        try:
            # Set additional environment variables for WeasyPrint if not set
            if not os.getenv('FONTCONFIG_PATH'):
                os.environ['FONTCONFIG_PATH'] = '/nix/store/*/etc/fonts:/usr/share/fonts'
            
            from weasyprint import HTML
            logging.info("WeasyPrint import successful, attempting PDF generation...")
            
            # Test WeasyPrint with a simple document first
            try:
                HTML(string="<html><body><h1>Test</h1></body></html>").write_pdf("/tmp/test.pdf")
                logging.info("WeasyPrint test document generation successful")
            except Exception as test_error:
                logging.warning(f"WeasyPrint test failed: {test_error}")
            
            # Generate actual PDF
            HTML(string=html_content).write_pdf(output_path)
            logging.info(f"✅ PDF generated successfully using WeasyPrint: {output_path}")
            return output_path
        except ImportError as import_error:
            logging.warning(f"WeasyPrint import failed: {import_error}. Trying ReportLab as fallback...")
        except Exception as weasyprint_error:
            logging.warning(f"WeasyPrint failed: {weasyprint_error}. Trying ReportLab as fallback...")
            
            # Fallback to ReportLab (pure Python, no system dependencies)
            try:
                from reportlab.lib.pagesizes import letter
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.units import inch
                from reportlab.lib import colors
                from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
                
                # Create PDF using ReportLab with enhanced formatting
                doc = SimpleDocTemplate(output_path, pagesize=letter, 
                                      rightMargin=0.7*inch, leftMargin=0.7*inch,
                                      topMargin=0.7*inch, bottomMargin=0.7*inch)
                
                styles = getSampleStyleSheet()
                
                # Enhanced custom styles to match WeasyPrint
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Title'],
                    fontSize=20,
                    fontName='Helvetica-Bold',
                    spaceAfter=2,
                    alignment=TA_LEFT,
                    textColor=colors.black
                )
                
                subtitle_style = ParagraphStyle(
                    'CustomSubtitle',
                    parent=styles['Normal'],
                    fontSize=14,
                    fontName='Helvetica',
                    spaceAfter=2,
                    alignment=TA_LEFT,
                    textColor=colors.black
                )
                
                contact_style = ParagraphStyle(
                    'ContactStyle',
                    parent=styles['Normal'],
                    fontSize=10,
                    fontName='Helvetica',
                    spaceAfter=15,
                    alignment=TA_LEFT,
                    textColor=colors.black
                )
                
                section_title_style = ParagraphStyle(
                    'SectionTitle',
                    parent=styles['Heading2'],
                    fontSize=14,
                    fontName='Helvetica-Bold',
                    spaceBefore=15,
                    spaceAfter=8,
                    textColor=colors.HexColor('#4F81BD'),
                    borderWidth=1,
                    borderColor=colors.HexColor('#B0C4DE'),
                    borderPadding=3
                )
                
                company_style = ParagraphStyle(
                    'CompanyStyle',
                    parent=styles['Normal'],
                    fontSize=11,
                    fontName='Helvetica-Bold',
                    spaceBefore=8,
                    spaceAfter=2,
                    alignment=TA_LEFT
                )
                
                duration_style = ParagraphStyle(
                    'DurationStyle',
                    parent=styles['Normal'],
                    fontSize=10,
                    fontName='Helvetica-Bold',
                    spaceAfter=2,
                    alignment=TA_LEFT
                )
                
                role_style = ParagraphStyle(
                    'RoleStyle',
                    parent=styles['Normal'],
                    fontSize=11,
                    fontName='Helvetica-BoldOblique',
                    spaceAfter=8,
                    alignment=TA_LEFT
                )
                
                bullet_style = ParagraphStyle(
                    'BulletStyle',
                    parent=styles['Normal'],
                    fontSize=11,
                    fontName='Helvetica',
                    spaceBefore=3,
                    spaceAfter=3,
                    leftIndent=22,
                    bulletIndent=10,
                    alignment=TA_JUSTIFY
                )
                
                environment_style = ParagraphStyle(
                    'EnvironmentStyle',
                    parent=styles['Normal'],
                    fontSize=9,
                    fontName='Helvetica-Oblique',
                    spaceBefore=8,
                    spaceAfter=5,
                    leftIndent=10,
                    rightIndent=10,
                    backColor=colors.HexColor('#F2F2F2'),
                    borderWidth=0.5,
                    borderColor=colors.HexColor('#E0E0E0'),
                    borderPadding=4
                )
                
                content = []
                
                # Enhanced header section
                content.append(Paragraph(resume_data['name'], title_style))
                if resume_data.get('title'):
                    content.append(Paragraph(resume_data['title'], subtitle_style))
                if resume_data.get('subtitle'):
                    content.append(Paragraph(resume_data['subtitle'], contact_style))
                
                # Enhanced contact info
                contact_parts = []
                if resume_data.get('email'):
                    contact_parts.append(resume_data['email'])
                if resume_data.get('phone'):
                    contact_parts.append(resume_data['phone'])
                if resume_data.get('location'):
                    contact_parts.append(resume_data['location'])
                if resume_data.get('linkedin'):
                    contact_parts.append(resume_data['linkedin'])
                
                if contact_parts:
                    content.append(Paragraph(' | '.join(contact_parts), contact_style))
                
                content.append(Spacer(1, 0.1*inch))
                
                # Enhanced sections with professional styling
                for section in resume_data['sections']:
                    # Section title with underline
                    section_title = section['name']
                    if section_title == 'Projects':
                        section_title = 'Professional Experience'
                    
                    content.append(Paragraph(section_title, section_title_style))
                    content.append(HRFlowable(width="100%", thickness=1, 
                                            color=colors.HexColor('#B0C4DE'), 
                                            spaceBefore=3, spaceAfter=8))
                    
                    if section['type'] == 'professional_experience':
                        for entry in section['content']:
                            # Enhanced company and duration formatting
                            header = entry.get('header', '')
                            if '|' in header:
                                parts = header.split('|')
                                company = parts[0].strip()
                                duration = parts[1].strip() if len(parts) > 1 else ''
                            else:
                                company = header
                                duration = ''
                            
                            content.append(Paragraph(company, company_style))
                            if duration:
                                content.append(Paragraph(duration, duration_style))
                            if entry.get('role'):
                                content.append(Paragraph(entry['role'], role_style))
                            
                            # Enhanced responsibilities with proper bullet formatting
                            for resp in entry.get('responsibilities', []):
                                content.append(Paragraph(f"• {resp}", bullet_style))
                            
                            # Environment section if present
                            if entry.get('environment'):
                                env_text = f"<b>Environment:</b> {entry['environment']}"
                                content.append(Paragraph(env_text, environment_style))
                            
                            content.append(Spacer(1, 0.15*inch))
                    
                    elif section['type'] == 'bullets':
                        for bullet in section['content']:
                            content.append(Paragraph(f"• {bullet}", bullet_style))
                        content.append(Spacer(1, 0.1*inch))
                    
                    elif section['type'] == 'plain_text_block':
                        # Handle Technical Skills as table if it contains colons
                        if section['name'] == 'Technical Skills' and ':' in section['content']:
                            lines = section['content'].split('\n')
                            skills_data = []
                            for line in lines:
                                if line.strip() and ':' in line:
                                    parts = line.split(':', 1)
                                    skills_data.append([f"{parts[0].strip()}:", parts[1].strip()])
                                elif line.strip():
                                    skills_data.append([line.strip(), ''])
                            
                            if skills_data:
                                skills_table = Table(skills_data, colWidths=[2*inch, 4*inch])
                                skills_table.setStyle(TableStyle([
                                    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                                    ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                                    ('FONTSIZE', (0, 0), (-1, -1), 11),
                                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                    ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                    ('RIGHTPADDING', (0, 0), (0, -1), 15),
                                    ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                                ]))
                                content.append(skills_table)
                        else:
                            # Regular plain text
                            lines = section['content'].split('\n')
                            for line in lines:
                                if line.strip():
                                    content.append(Paragraph(line.strip(), bullet_style))
                        content.append(Spacer(1, 0.1*inch))
                    
                    elif section['type'] == 'projects':
                        for project in section['content']:
                            header = project.get('header', [])
                            if isinstance(header, list) and header:
                                # Company
                                content.append(Paragraph(header[0], company_style))
                                # Duration (usually the last element)
                                if len(header) > 1:
                                    content.append(Paragraph(header[-1], duration_style))
                                # Role (middle element if exists)
                                if len(header) > 2:
                                    content.append(Paragraph(header[1], role_style))
                            
                            # Enhanced responsibilities
                            for resp in project.get('responsibilities', []):
                                content.append(Paragraph(f"• {resp}", bullet_style))
                            
                            # Environment section
                            if project.get('environment'):
                                env_text = f"<b>Environment:</b> {project['environment']}"
                                content.append(Paragraph(env_text, environment_style))
                            
                            content.append(Spacer(1, 0.15*inch))
                
                # Build PDF
                doc.build(content)
                
                logging.info(f"PDF generated successfully using ReportLab: {output_path}")
                return output_path
                
            except Exception as reportlab_error:
                logging.error(f"Both WeasyPrint and ReportLab failed. WeasyPrint: {weasyprint_error}, ReportLab: {reportlab_error}")
                raise Exception("Failed to generate PDF resume. Both WeasyPrint and ReportLab failed.")
        
    except Exception as e:
        logging.error(f"PDF generation failed: {e}", exc_info=True)
        raise Exception("Failed to generate PDF resume. Please ensure PDF generation libraries are properly installed and configured.")

# Modified upload_resume endpoint to ensure PDF output
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
            logging.info("Download requested. Generating enhanced resume PDF.")
            # Get missing keywords for GPT enhancement
            resume_words = set(resume_text.lower().split())
            missing_hard_skills = hard_skills - resume_words
            missing_soft_skills = soft_skills - resume_words
            missing_keywords = missing_hard_skills.union(missing_soft_skills)
            
            logging.info(f"Found {len(missing_keywords)} missing keywords for enhancement: {list(missing_keywords)[:5]}...")
            enhanced_text = enhance_resume_text(resume_text, missing_keywords)
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
            "request": request, "score": 0, "suggestions": f"An unexpected error occurred: {str(e)}. Please check the logs."
        })

# Endpoint to serve the optimized resume file for download
@app.get("/download/{filename}")
async def download_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    logging.info(f"GET /download/{filename} - Attempting to serve file from {path}")
    if os.path.exists(path):
        logging.info("File found. Sending response.")
        return FileResponse(path, media_type="application/pdf", filename=filename)
    
    logging.error(f"File not found at path: {path}")
    return {"detail": "File not found"}

# Endpoint to serve the upload form
@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    logging.info("GET /upload - Serving upload form")
    return templates.TemplateResponse("index.html", {"request": request})

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
        "subtitle": user_info.get("subtitle", ""),  # Add subtitle (skills) from user_info
        "email": user_info.get("email", ""),
        "phone": user_info.get("phone", ""),
        "linkedin": user_info.get("linkedin", ""),
        "location": user_info.get("location", ""),
        "sections": []
    }

    VALID_SECTIONS = [
        "PROFILE SUMMARY", "PROFESSIONAL SUMMARY", "PROFESSIONAL EXPERIENCE", "PROJECTS",
        "EDUCATION", "TECHNICAL SKILLS", "CERTIFICATIONS", "ACHIEVEMENTS", "EXPERTISE IN", "KEY STRENGTHS"
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
        "KEY STRENGTHS",         
        "PROFESSIONAL EXPERIENCE", 
        "PROJECTS"
    ]
    
    for section_name in desired_sections:
        section_content = sections.get(section_name.upper())
        if section_content:
            if section_name.upper() == 'PROFESSIONAL EXPERIENCE':
                processed_content = parse_experience_section(section_content)
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "professional_experience",
                    "content": processed_content
                })
            elif section_name.upper() in ['PROFESSIONAL SUMMARY', 'PROFILE SUMMARY']:
                # For Professional Summary, convert paragraphs to bullet points
                processed_content = parse_professional_summary(section_content)
                # Always keep as bullets for these sections - don't convert to plain text
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "bullets",
                    "content": processed_content
                })
            elif section_name.upper() == 'TECHNICAL SKILLS':
                # For Technical Skills, handle key-value format
                processed_content = parse_technical_skills(section_content)
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "plain_text_block",
                    "content": processed_content
                })
            elif section_name.upper() in ['ACHIEVEMENTS', 'CERTIFICATIONS', 'KEY STRENGTHS']:
                # These sections use the general bullet parser
                processed_content = split_into_logical_bullets(section_content) 
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "bullets",
                    "content": processed_content
                })
            elif section_name.upper() == 'PROJECTS':
                processed_content = parse_projects_content(section_content)
                resume_data["sections"].append({
                    "name": section_name.title(),
                    "type": "projects",
                    "content": processed_content
                })
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
    
    # Handle multi-line key-value format (like Akhil's resume)
    for i, line in enumerate(content_lines):
        stripped_line = line.strip()
        if not stripped_line:
            continue
        
        # Check if this looks like a category (short line, no commas, capitalized)
        # Allow for longer category names containing & or / characters
        if (len(stripped_line.split()) <= 5 and 
            ',' not in stripped_line and 
            ':' not in stripped_line and  # Categories don't have colons
            not stripped_line.startswith('•') and
            stripped_line[0].isupper() and
            not any(tech in stripped_line for tech in ['Java', 'Python', 'React', 'Angular', 'AWS', 'Docker'])):  # Avoid common tech names
            current_category = stripped_line
        else:
            # This is likely the skills list for the category
            if current_category:
                formatted_lines.append(f"{current_category}: {stripped_line}")
                current_category = None
            else:
                # If no category was found, just add the line as is
                formatted_lines.append(stripped_line)
    
    # Handle case where last line was a category with no value
    if current_category:
        formatted_lines.append(current_category)
    
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

# Helper 1: Detects a line containing a project header date.
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
    # Enhanced regex to handle various date formats including "September2023" without space
    date_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4})'
    return bool(re.search(date_pattern, line, re.I))

def has_start_date(line: str) -> bool:
    """Detects a line containing a start date (e.g., 'March 2022' or 'Company March 2022')."""
    # Look for month + year at the end of line or before separator
    start_date_pattern = r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}(?:\s*$|(?=\s*[–-])|(?=\s*to))'
    return bool(re.search(start_date_pattern, line, re.I))

def has_end_date(line: str) -> bool:
    """Detects a line containing an end date (e.g., '– May 2023' or 'to Present')."""
    # Look for end date patterns starting with separator
    end_date_pattern = r'^\s*[–-]\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4})|^\s*to\s+(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4})'
    return bool(re.search(end_date_pattern, line, re.I))

def combine_split_date_range(line1: str, line2: str) -> str:
    """Combines two lines that together form a complete date range."""
    # Extract start date from line1
    start_match = re.search(r'\b((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4})', line1, re.I)
    start_date = start_match.group(1) if start_match else ""
    
    # Extract end date from line2  
    end_match = re.search(r'[–-]\s*((?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))', line2, re.I)
    if not end_match:
        end_match = re.search(r'to\s+((?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))', line2, re.I)
    
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

# --- RESTRUCTURED parse_experience_section FUNCTION (more robust state machine) ---
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
                    date_match = re.search(r'(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}\s*(?:to|–|-)*\s*(?:Present|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*\d{4}))$', next_line, re.I)
                    
                    if date_match:
                        dates_part = date_match.group(1)
                        company_part = next_line[:date_match.start()].strip()
                        company_part = re.sub(r'\s+', ' ', company_part).strip()
                        
                        if '|' not in company_part:
                            entry_data['header'] = f"{company_part} | {dates_part}"
                        else:
                            entry_data['header'] = f"{company_part} {dates_part}"
                        entry_data['role'] = line
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
                    entry_data['responsibilities'].append(resp_line.lstrip('•- ').strip())
            
            i += 1
        
        # Clean up the data
        entry_data['header'] = entry_data['header'].strip()
        entry_data['role'] = entry_data['role'].strip()
        
        if entry_data['header'] or entry_data['role']:
            experience_entries.append(entry_data)

    return experience_entries

def parse_projects_content(content_lines: list) -> list:
    app_logger.info("Starting to parse project content with enhanced parser for multiple formats...")
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
    
    app_logger.info(f"Enhanced parsing complete. Found {len(projects)} projects.")
    return projects

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
    logging.info(f"Enhancing resume text with {len(missing_keywords)} missing keywords.")
    
    if not missing_keywords:
        logging.info("No missing keywords to enhance.")
        return resume_text
    
    # Validate that we have the required sections
    has_projects = bool(re.search(r'PROJECTS\s*\n', resume_text, re.IGNORECASE))
    
    if not has_projects:
        logging.warning("No Projects section found. Cannot enhance first project.")
        return resume_text
        
    # Give ALL keywords to the first project (current job)
    keywords_list = list(missing_keywords)
    
    # Generate 3 bullet points for each missing skill in first project (current job)
    first_project_bullets = []
    if keywords_list:
        logging.info(f"Generating 3 bullet points each for {len(keywords_list)} skills in first project (current job)...")
        for keyword in keywords_list:
            # Generate 3 different project-specific bullet points for each skill
            for i in range(3):
                project_bullet = generate_project_bullet_point_from_gpt(keyword, i+1).lstrip("•- ").strip()
                if project_bullet:
                    first_project_bullets.append(f"• {project_bullet}")

    updated_text = resume_text
    
    # Enhance ONLY the first project in PROJECTS section (current job)
    if first_project_bullets:
        logging.info(f"Adding {len(first_project_bullets)} GPT bullets to current project...")
        
        # Find the PROJECTS section
        projects_start = re.search(r'PROJECTS\s*\n', updated_text, re.IGNORECASE)
        
        if projects_start:
            try:
                # Get the position after "PROJECTS" heading
                projects_start_pos = projects_start.end()
                
                # Find the text after PROJECTS section
                after_projects = updated_text[projects_start_pos:]
                lines = after_projects.split('\n')
                
                # Find where to insert bullets - after existing bullet points in first project
                insert_position = 0
                found_bullets = False
                
                # Look for existing bullet points (• or - at start of line)
                for i, line in enumerate(lines):
                    stripped_line = line.strip()
                    if stripped_line.startswith('•') or stripped_line.startswith('-') or stripped_line.startswith('*'):
                        found_bullets = True
                        insert_position = i + 1  # Insert after this bullet
                    elif found_bullets and stripped_line and not stripped_line.startswith('•') and not stripped_line.startswith('-') and not stripped_line.startswith('*'):
                        # We've found the end of bullet points
                        break
                    elif not stripped_line:
                        # Empty line - might be end of bullets or spacing
                        if found_bullets:
                            insert_position = i
                
                # If no bullets found, insert after first few non-empty lines (header info)
                if not found_bullets:
                    header_lines = 0
                    for i, line in enumerate(lines):
                        stripped_line = line.strip()
                        if stripped_line:
                            header_lines += 1
                            if header_lines >= 3:  # After company, role, date info
                                insert_position = i + 1
                                break
                
                # Insert the GPT bullets
                enhanced_lines = (
                    lines[:insert_position] +
                    [''] +  # Add empty line before GPT bullets
                    first_project_bullets +
                    lines[insert_position:]
                )
                
                # Reconstruct the text
                enhanced_after_projects = '\n'.join(enhanced_lines)
                updated_text = updated_text[:projects_start_pos] + enhanced_after_projects
                
                logging.info(f"Successfully enhanced first project with {len(first_project_bullets)} bullet points.")
                
            except Exception as e:
                logging.error(f"Failed to enhance first project: {e}")
        else:
            logging.warning("PROJECTS section not found. No project enhancement performed.")

    logging.info("Resume text enhanced successfully.")
    return updated_text