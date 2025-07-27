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
app.mount("/static", StaticFiles(directory=str(Path(__file__).resolve().parent.parent / "static")), name="static")
templates = Jinja2Templates(directory=str(Path(__file__).resolve().parent.parent / "templates"))
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

def verify_openai_connection():
    """Verify that OpenAI client is properly configured and working."""
    global client
    
    if client is None:
        logging.error("❌ OpenAI client not initialized - API key missing or invalid")
        return False
    
    try:
        # Test with a simple API call
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": "Say 'test' in one word."}],
            max_tokens=5,
            temperature=0.1
        )
        
        if response and response.choices and response.choices[0].message.content:
            logging.info("✅ OpenAI client verified - API connection working")
            return True
        else:
            logging.error("❌ OpenAI API returned empty response")
            return False
            
    except Exception as e:
        logging.error(f"❌ OpenAI API test failed: {e}")
        return False

def generate_project_bullet_point_from_gpt(keyword: str, variation: int = 1) -> str:
    """Generate varied project-specific bullet points for missing keywords."""
    global client
    
    if client is None:
        logging.warning("⚠️ OpenAI client not configured. Returning placeholder bullet point.")
        return f"• Implemented {keyword.title()} solutions in project development (OpenAI API key not configured)"

    keyword = keyword.strip().lower()
    
    # Use variation-specific cache key for project bullets
    project_cache_key = f"project_{keyword}_v{variation}"
    
    # ✅ Return cached result if available
    if project_cache_key in bullet_cache:
        logging.info(f"💾 Cache hit for project keyword: '{keyword}' variation {variation}. Returning cached bullet point.")
        return bullet_cache[project_cache_key]

    logging.info(f"🔄 Cache miss for project keyword '{keyword}' variation {variation}. Calling GPT-4o API...")
    
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

        if not response or not response.choices or not response.choices[0].message.content:
            logging.warning(f"⚠️ GPT returned empty response for keyword '{keyword}' variation {variation}")
            return f"• Implemented {keyword.title()} solutions to enhance project functionality (empty response)"

        bullet = response.choices[0].message.content.strip()

        # ✅ Ensure bullet formatting
        if not bullet.startswith("•"):
            bullet = "• " + bullet

        # ✅ Cache and return
        logging.info(f"✅ Successfully generated project bullet for '{keyword}' variation {variation}. Caching result.")
        bullet_cache[project_cache_key] = bullet
        
        try:
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                json.dump(bullet_cache, f, indent=2)
        except Exception as cache_error:
            logging.warning(f"⚠️ Failed to save cache: {cache_error}")

        return bullet

    except Exception as e:
        logging.error(f"❌ GPT API Exception for keyword '{keyword}' variation {variation}: {e}")
        
        # Verify if it's an API key issue
        if "api" in str(e).lower() and ("key" in str(e).lower() or "auth" in str(e).lower()):
            logging.error("❌ This appears to be an API key authentication issue!")
            
        return f"• Implemented {keyword.title()} solutions to enhance project functionality (API error: {str(e)[:50]})"

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
    """
    Generate a professionally formatted resume PDF with enhanced content.
    """
    
    # Parse the enhanced text into structured data
    resume_data = parse_resume_to_structure(enhanced_text, user_info)
    
    # Generate HTML content first
    html_content = generate_resume_html(resume_data)
    
    # CRITICAL: Apply comprehensive cleanup to the HTML content to fix all formatting issues
    logging.info("🧹 Applying comprehensive formatting fixes to HTML content...")
    html_content = clean_resume_formatting_issues(html_content)
    html_content = clean_professional_experience_bullets(html_content)
    
    logging.info("Attempting to generate PDF resume.")
    try:
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
            
            # Generate actual PDF using cleaned HTML content
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
            logging.info("📥 Download requested. Generating enhanced resume PDF.")
            
            # Get missing keywords for GPT enhancement
            resume_words = set(resume_text.lower().split())
            missing_hard_skills = hard_skills - resume_words
            missing_soft_skills = soft_skills - resume_words
            missing_keywords = missing_hard_skills.union(missing_soft_skills)
            
            logging.info(f"🔍 Analysis complete:")
            logging.info(f"   📊 Hard skills in JD: {len(hard_skills)}")
            logging.info(f"   📊 Soft skills in JD: {len(soft_skills)}")
            logging.info(f"   📊 Missing hard skills: {len(missing_hard_skills)} - {list(missing_hard_skills)[:10]}")
            logging.info(f"   📊 Missing soft skills: {len(missing_soft_skills)} - {list(missing_soft_skills)[:10]}")
            logging.info(f"   🎯 Total missing keywords for GPT: {len(missing_keywords)}")
            
            if missing_keywords:
                logging.info(f"🤖 Starting GPT enhancement for missing keywords...")
                enhanced_text = enhance_resume_text(resume_text, missing_keywords)
                
                # Verify enhancement worked
                original_length = len(resume_text)
                enhanced_length = len(enhanced_text)
                length_increase = enhanced_length - original_length
                
                logging.info(f"📈 Text length: {original_length} → {enhanced_length} (+{length_increase} chars)")
                
                if length_increase > 100:  # Should have significant increase if bullets were added
                    logging.info("✅ GPT enhancement appears successful - significant text increase detected")
                else:
                    logging.warning(f"⚠️ GPT enhancement may have failed - only {length_increase} character increase")
            else:
                logging.info("ℹ️ No missing keywords found - no GPT enhancement needed")
                enhanced_text = resume_text
            
            # Generate PDF with enhanced text
            logging.info("📄 Generating formatted PDF...")
            output_path = generate_formatted_resume_pdf(filename, enhanced_text, user_info)
            download_link = f"/download/{os.path.basename(output_path)}"
            logging.info(f"✅ Download link created: {download_link}")

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
            
            # This is a category - save it for the next skills line
            current_category = stripped_line
        else:
            # This is likely the skills list for the category
            if current_category:
                # Format as "Category: skills"
                formatted_lines.append(f"{current_category}: {stripped_line}")
                current_category = None
            else:
                # If no category was found, check if the line already has a colon (is already formatted)
                if ':' in stripped_line:
                    formatted_lines.append(stripped_line)
                else:
                    # Just add the line as is
                    formatted_lines.append(stripped_line)
    
    # Handle case where last line was a category with no value - add it as a standalone line
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

def find_current_project_by_date(resume_text: str) -> tuple:
    """
    Find the current/most recent project by analyzing dates in the resume.
    Returns (project_start_line, project_end_line, project_info) or (None, None, None) if not found.
    """
    import datetime
    import re
    
    lines = resume_text.split('\n')
    current_year = datetime.datetime.now().year
    
    # Enhanced date patterns that capture "to" separators
    date_patterns = [
        r'\b\w+\s+\d{4}\s+to\s+Present\b',                    # Sep 2022 to Present
        r'\b\w+\s+\d{4}\s+to\s+Current\b',                    # Sep 2022 to Current
        r'\b\w+\s+\d{4}\s+to\s+Till\s+Date\b',               # Sep 2022 to Till Date
        r'\b\w+\s+\d{4}\s+to\s+Ongoing\b',                    # Sep 2022 to Ongoing
        r'\b\w{3}\s+\d{4}\s+to\s+Present\b',                 # Sep 2022 to Present (short month)
        r'\b\d{4}\s+to\s+Present\b',                          # 2022 to Present
        r'\b\w+\s+\d{4}\s*[-–]\s*Present\b',                 # Sep 2022 - Present
        r'\b\w+\s+\d{4}\s*[-–]\s*Current\b',                 # Sep 2022 - Current
        r'\b\w{3}\s+\d{4}\s*[-–]\s*\w{3}\s+\d{4}\b',        # Sep 2022 - Dec 2024
        r'\b\w+\s+\d{4}\s+to\s+\w+\s+\d{4}\b',              # Sep 2022 to Dec 2024
    ]
    
    found_projects = []
    
    for i, line in enumerate(lines):
        for pattern in date_patterns:
            matches = re.findall(pattern, line, re.IGNORECASE)
            if matches:
                for match in matches:
                    # Determine if this is current/recent
                    is_current = False
                    latest_year = 0
                    
                    if any(word in match.lower() for word in ['present', 'current', 'ongoing', 'till date']):
                        is_current = True
                        latest_year = current_year
                    else:
                        # Extract years from the date match
                        years = re.findall(r'\d{4}', match)
                        if years:
                            latest_year = max(int(year) for year in years)
                            if latest_year >= current_year - 1:  # Current or previous year
                                is_current = True
                    
                    # Find the project context (company/project name)
                    project_context = []
                    # Look at lines before the date for company/project info
                    context_start = max(0, i - 5)
                    for j in range(context_start, i + 2):
                        if j < len(lines) and lines[j].strip():
                            project_context.append(lines[j].strip())
                    
                    found_projects.append({
                        'line_num': i + 1,
                        'date_match': match,
                        'is_current': is_current,
                        'latest_year': latest_year,
                        'context': project_context
                    })
                    
                    logging.info(f"📅 Found project date at line {i+1}: {match} (Current: {'YES' if is_current else 'NO'})")
    
    # Find the most current project
    if found_projects:
        # Sort by latest year and whether it's current
        current_projects = [p for p in found_projects if p['is_current']]
        if current_projects:
            # Use the first current project found
            current_project = current_projects[0]
            logging.info(f"🎯 IDENTIFIED CURRENT PROJECT: Line {current_project['line_num']} - {current_project['date_match']}")
            
            # Find the start and end of this project section
            project_line = current_project['line_num'] - 1  # Convert to 0-based
            
            # Find project start (look backwards for company/project header)
            project_start = project_line
            for j in range(project_line, max(0, project_line - 10), -1):
                line = lines[j].strip()
                if line and (any(word in line for word in ['Company', 'Corp', 'Inc', 'Ltd', 'Solutions', 'Services', 'Technologies']) or 
                           line.isupper() and len(line.split()) <= 4):
                    project_start = j
                    break
            
            # Find project end (look forwards for next project/section)
            project_end = min(len(lines), project_line + 50)  # Default end
            for j in range(project_line + 1, len(lines)):
                line = lines[j].strip()
                # Stop at next date pattern or obvious section header
                if any(re.search(pattern, line, re.IGNORECASE) for pattern in date_patterns):
                    project_end = j
                    break
                elif line.isupper() and len(line) > 5 and not any(char in line for char in ['•', '-', ':']):
                    project_end = j
                    break
            
            return project_start, project_end, current_project
        else:
            # If no current projects, use the most recent one
            most_recent = max(found_projects, key=lambda x: x['latest_year'])
            logging.info(f"🎯 USING MOST RECENT PROJECT: Line {most_recent['line_num']} - {most_recent['date_match']}")
            return most_recent['line_num'] - 5, most_recent['line_num'] + 20, most_recent
    
    logging.warning("❌ No current/recent projects found with date patterns")
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
        has_professional_experience = any('PROFESSIONAL EXPERIENCE' in line.upper() for line in lines)
        
        # Determine target section based on what's available
        if has_projects:
            target_section = "PROJECTS"
            logging.info("🎯 Target: Adding keywords to PROJECTS section (current project)")
        elif has_professional_experience:
            target_section = "PROFESSIONAL EXPERIENCE" 
            logging.info("🎯 Target: Adding keywords to PROFESSIONAL EXPERIENCE section (current position)")
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
                project_bullet = generate_project_bullet_point_from_gpt(keyword, 1).lstrip("•- ").strip()
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
                        project_bullet = generate_project_bullet_point_from_gpt(f"{keyword} technology", 1).lstrip("•- ").strip()
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
                
            except Exception as e:
                logging.error(f"❌ Failed to enhance {target_section} in {target_section}: {e}")
                import traceback
                logging.error(f"Full error: {traceback.format_exc()}")
                return resume_text
        else:
            logging.warning(f"❌ {target_section} section not found during enhancement.")
            # Debug: Show what sections we can find
            all_lines = updated_text.split('\n')
            section_lines = [line.strip() for line in all_lines if line.strip() and line.strip().isupper()]
            logging.warning(f"📋 Available sections detected: {section_lines[:10]}")
            # Try alternative patterns
            for pattern in ["PROFESSIONAL EXPERIENCE", "EXPERIENCE", "PROJECTS", "PROJECT"]:
                if pattern.upper() in updated_text.upper():
                    logging.warning(f"🔍 Found alternative section '{pattern}' in text")
                    break
            logging.warning("❌ No enhancement performed.")
            return resume_text

    # Count final bullets for verification
    enhanced_bullet_count = len(re.findall(r'^\s*•', updated_text, re.MULTILINE))
    logging.info(f"✅ Enhancement complete. Total bullets in enhanced resume: {enhanced_bullet_count}")
    
    target_info = f"CURRENT PROJECT" if current_project_start is not None else target_section
    logging.info(f"🎉 {target_info} enhancement completed successfully with missing keyword bullets!")
    
    # Final text length verification
    logging.info(f"📈 Text length: {len(resume_text)} → {len(updated_text)} (+{len(updated_text) - len(resume_text)} chars)")
    if len(updated_text) <= len(resume_text):
        logging.warning("⚠️ GPT enhancement may have failed - no significant increase in text length")
    else:
        logging.info("✅ GPT enhancement appears successful - significant text increase detected")
    
    return updated_text

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
            <pre>{{ resume_text }}</pre>
            <h2>Suggestions</h2>
            <p>{{ suggestions }}</p>
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

# Verify OpenAI connection at startup
if verify_openai_connection():
    logging.info("🚀 All systems ready - OpenAI GPT enhancement fully operational!")
else:
    logging.warning("⚠️ OpenAI verification failed - GPT enhancement may not work properly")

# For local development
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="127.0.0.1", port=8000, reload=False)

def clean_resume_formatting_issues(resume_text: str) -> str:
    """
    Clean up specific formatting issues in the resume text or HTML content.
    """
    logging.info("🧹 Cleaning up resume formatting issues...")
    
    # Detect if this is HTML content
    is_html = '<html>' in resume_text or '<div>' in resume_text or '<p>' in resume_text
    
    if is_html:
        # For HTML content, we need to clean up the text within HTML tags
        
        # Fix 0: CRITICAL - Remove empty bullet points from HTML
        # Pattern: <li>•</li> or <li> • </li> or similar
        resume_text = re.sub(r'<li[^>]*>\s*•\s*</li>', '', resume_text)
        resume_text = re.sub(r'<li[^>]*>\s*</li>', '', resume_text)  # Completely empty list items
        
        # Fix 1: Remove empty paragraphs with just bullets
        resume_text = re.sub(r'<p[^>]*>\s*•\s*</p>', '', resume_text)
        
        # Fix 2: Clean up Redis: Package Manager: pattern in HTML
        resume_text = re.sub(r'Redis:\s*</[^>]+>\s*<[^>]+>\s*Package Manager:', 'Redis Package Manager:', resume_text)
        
        # Fix 3: Remove excessive empty elements
        resume_text = re.sub(r'<p[^>]*>\s*</p>', '', resume_text)  # Empty paragraphs
        resume_text = re.sub(r'<div[^>]*>\s*</div>', '', resume_text)  # Empty divs
        
        logging.info("✅ HTML formatting cleanup completed")
        return resume_text
    
    # Original text-based cleanup for non-HTML content
    # Fix 0: CRITICAL - Remove massive empty bullet point spam (this is the major issue!)
    # Remove lines that are ONLY bullets with optional whitespace
    lines = resume_text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        stripped = line.strip()
        # Skip lines that are just bullets with nothing else
        if stripped == '•' or stripped == '• ' or stripped == ' •' or not stripped:
            logging.info("🔧 Removed empty bullet point")
            continue
        cleaned_lines.append(line)
    
    resume_text = '\n'.join(cleaned_lines)
    
    # Fix 1: Clean up duplicate company names in Professional Experience
    lines = resume_text.split('\n')
    cleaned_lines = []
    i = 0
    
    while i < len(lines):
        current_line = lines[i]
        
        # Special handling for Professional Experience company duplicates
        if i + 2 < len(lines):
            line1 = lines[i].strip()
            line2 = lines[i + 1].strip() if i + 1 < len(lines) else ""
            line3 = lines[i + 2].strip() if i + 2 < len(lines) else ""
            line4 = lines[i + 3].strip() if i + 3 < len(lines) else ""
            
            # Check if line1 has company info without complete date
            if (re.search(r'[A-Za-z\s]+,\s*[A-Za-z\s]+,\s*[A-Z]{2}\s+\w+\s+\d{4}$', line1) and
                re.match(r'^[–-]\s*\w+\s+\d{4},?$', line2)):
                
                # Check if line3 is a duplicate of the combined pattern
                if line3.startswith(line1) and re.search(r'\d{4}\s*[–-]\s*\w+\s+\d{4},?$', line3):
                    # Skip the duplicate lines
                    cleaned_lines.append(lines[i])  # Keep original line1
                    cleaned_lines.append(lines[i + 1])  # Keep original line2
                    # Skip line3 (duplicate)
                    if line4:  # Add line4 if it exists (likely the additional date range)
                        cleaned_lines.append(line4)
                    logging.info(f"🔧 Removed duplicate company entry: {line3[:50]}...")
                    i += 4  # Skip past all processed lines
                    continue
        
        # Also check for same-line duplicates
        if i + 1 < len(lines):
            line = lines[i].strip()
            next_line = lines[i + 1].strip()
            
            # Check for duplicate company pattern on consecutive lines
            if (line and next_line and 
                re.search(r'\w+\s+\d{4}\s*[–-]\s*\w+\s+\d{4},?\s*$', line) and
                next_line.startswith(line.rstrip(',')) and
                line != next_line):
                # Keep only the first line, skip the duplicate
                cleaned_lines.append(lines[i])
                logging.info(f"🔧 Removed duplicate company line: {next_line[:50]}...")
                i += 2  # Skip the duplicate line
                continue
        
        cleaned_lines.append(lines[i])
        i += 1
    
    resume_text = '\n'.join(cleaned_lines)
    
    # Fix 2: Clean up ALL duplicate Environment sections (there are 7 of them!)
    lines = resume_text.split('\n')
    cleaned_lines = []
    environment_content = []
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        if line.startswith('Environment:'):
            # Collect all Environment lines
            env_content = line[len('Environment:'):].strip()
            if env_content and env_content not in environment_content:
                environment_content.append(env_content)
            
            # Skip subsequent Environment lines within a reasonable distance
            j = i + 1
            while j < len(lines) and j < i + 10:  # Look ahead max 10 lines
                if lines[j].startswith('Environment:'):
                    next_env = lines[j][len('Environment:'):].strip()
                    if next_env and next_env not in environment_content:
                        environment_content.append(next_env)
                    j += 1
                elif lines[j].strip():  # Non-empty non-Environment line
                    break
                else:
                    j += 1
            
            # Create merged Environment line
            if environment_content:
                merged_env = "Environment: " + ", ".join(environment_content)
                cleaned_lines.append(merged_env)
                logging.info(f"🔧 Merged {len(environment_content)} Environment sections")
                environment_content = []  # Reset for next batch
            
            i = j  # Jump past all processed Environment lines
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
        if 'PROFESSIONAL EXPERIENCE' in line.upper():
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