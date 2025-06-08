# Import required modules from FastAPI and other libraries
from fastapi import FastAPI, File, UploadFile, Form, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docx import Document
from docx.shared import Pt
import openai
import shutil, os, fitz, docx2txt, spacy, re, json
from docx.oxml import OxmlElement
from pathlib import Path
from dotenv import load_dotenv
from docx.oxml.ns import qn
from openai import OpenAI
from .utils import validate_resume_format, extract_keywords  # Reusable utility functions

# Initialize FastAPI app
app = FastAPI()
load_dotenv()
client = OpenAI()

models = client.models.list()
for model in models.data:
    print(model.id)

app.debug = True
load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env")
CACHE_PATH = Path("bullet_cache.json")
print("Loaded OpenAI Key:", os.getenv("OPENAI_API_KEY"))

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
app.template_folder = "templates"

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
    from openai import OpenAI
    client = OpenAI(api_key=openai.api_key)

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

        bullet = response.choices[0].message.content.strip()

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

def generate_formatted_resume_docx(filename: str, enhanced_text: str) -> str:
    doc = Document()
    output_path = os.path.join("generated_resumes", f"{os.path.splitext(filename)[0]}_formatted.docx")

    # ✅ Add horizontal line below header
    def add_horizontal_line(doc):
        p = doc.add_paragraph()
        p.alignment = 1
        p_paragraph = p._p
        p_borders = OxmlElement('w:pBdr')
        bottom_border = OxmlElement('w:bottom')
        bottom_border.set(qn('w:val'), 'single')
        bottom_border.set(qn('w:sz'), '6')
        bottom_border.set(qn('w:space'), '1')
        bottom_border.set(qn('w:color'), 'auto')
        p_borders.append(bottom_border)
        p_props = p_paragraph.get_or_add_pPr()
        p_props.append(p_borders)

    # ✅ Add contact header
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
        doc.add_paragraph()       # spacer

    # ✅ Insert the header section
    add_header_section(doc)

    # ✅ Apply base styles
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

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
    filtered_lines = [
        line for line in lines
        if not any(k in line.lower() for k in header_keywords)
        and not any(line.strip().lower().startswith(k) for k in table_keywords)
    ]
    # ✅ Build resume body
    current_section = ""
    skip_keywords = {"Environment", "Responsibilities"}
    seen_sections = set()

    for idx, line in enumerate(filtered_lines):
        stripped = line.strip()

        # Modified regex to match Company, City, State, and optional Duration on the same line
        # This will correctly identify lines like "Dish Network, Denver, CO Nov 2021 to Sep 2022"
        company_line_match = re.match(
            r"^([A-Za-z0-9 &().-]+)\s*,\s*([A-Za-z .-]+)\s*,\s*([A-Z]{2})\s*([A-Za-z]{3,9}\s+\d{4}(?:\s*(?:to|–|-)\s*[A-Za-z]{3,9}\s+\d{4})?)?$",
            stripped
        )
        
        if company_line_match:
            company = company_line_match.group(1).strip()
            city = company_line_match.group(2).strip()
            state = company_line_match.group(3).strip()
            duration_on_line = company_line_match.group(4) # This will be None if duration not present on this line

            # Try to get job title and duration from next lines if not found on current line
            job_title = ""
            duration = duration_on_line if duration_on_line else "" # Use duration from line if present

            # Check next lines for job title and duration if not on current line
            if not duration_on_line: # Only look for duration on next lines if it wasn't on current line
                if idx + 1 < len(filtered_lines):
                    next_line = filtered_lines[idx + 1].strip()
                    if next_line and not next_line.isupper() and not next_line.startswith(("•", "-")) and not re.match(r"^[A-Za-z]{3,9}\s+\d{4}", next_line):
                        job_title = next_line
                        idx_offset = 2
                    else:
                        idx_offset = 1
                    if idx + idx_offset < len(filtered_lines):
                        duration_candidate = filtered_lines[idx + idx_offset].strip()
                        if re.match(r"^[A-Za-z]{3,9}\s+\d{4}\s*(to|–|-)\s*[A-Za-z]{3,9}\s+\d{4}$", duration_candidate) or re.match(r"^[A-Za-z]{3,9}\s+\d{4}$", duration_candidate):
                            duration = duration_candidate

            # Add two blank lines before company block
            doc.add_paragraph()
            doc.add_paragraph()
            table = doc.add_table(rows=1, cols=2)
            table.autofit = False
            table.columns[0].width = Pt(400)
            table.columns[1].width = Pt(200)

            # Left cell: company, job title
            cell_left = table.cell(0, 0)
            p_left = cell_left.paragraphs[0]
            
            # Apply bolding only to the company name
            run_company_name = p_left.add_run(company)
            run_company_name.bold = True
            run_company_name.font.size = Pt(16)
            run_company_name.font.name = "Calibri"
            r_name = run_company_name._element
            r_name.rPr.rFonts.set(qn('w:eastAsia'), "Calibri")

            # Add city and state without bolding, maintaining same font size and name
            run_location = p_left.add_run(f", {city}, {state}")
            run_location.font.size = Pt(16)
            run_location.font.name = "Calibri"
            r_location = run_location._element
            r_location.rPr.rFonts.set(qn('w:eastAsia'), "Calibri")

            if job_title:
               p_left.add_run("\n")
               run_title = p_left.add_run(job_title)
               run_title.italic = True
               run_title.font.size = Pt(12)
               run_title.font.name = "Calibri"
               r2 = run_title._element
               r2.rPr.rFonts.set(qn('w:eastAsia'), "Calibri")

            # Right cell: duration
            cell_right = table.cell(0, 1)
            p_right = cell_right.paragraphs[0]
            p_right.alignment = 2
            if duration:
                run_right = p_right.add_run(duration)
                run_right.bold = True
                run_right.font.size = Pt(12)
                run_right.font.name = "Calibri"
                r3 = run_right._element
                r3.rPr.rFonts.set(qn('w:eastAsia'), "Calibri")

            continue  # Important to skip rest of loop for company lines

        # Section headings
        if stripped.isupper() and len(stripped.split()) < 6:
            current_section = stripped
            doc.add_heading(stripped.title(), level=1)
            add_horizontal_line(doc)
            continue

        # Skip duplicate/irrelevant sections
        if any(keyword.lower() in stripped.lower() for keyword in skip_keywords):
            if stripped in seen_sections:
                continue
            seen_sections.add(stripped)
            continue  # Skip processing this line further

        # Bullet points
        if stripped.startswith("•") or stripped.startswith("-"):
            p = doc.add_paragraph(stripped, style='List Bullet')
            p.paragraph_format.space_after = Pt(0)  # Remove space after bullet
        elif stripped:
            p = doc.add_paragraph(stripped)
            p.paragraph_format.space_after = Pt(0)  # Remove space after normal paragraph

    doc.save(output_path)
    return output_path

# Enhances the resume text by appending keyword content to the Profile Summary and CMS section

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
                  ([bullet for bullet in real_world_bullets if bullet.strip()] if real_world_bullets else [])
             )
    ),
    resume_text,
    flags=re.DOTALL | re.IGNORECASE
)

    # 3. Add enhancement note to CMS section
    if missing_keywords:
        updated_text = re.sub(
            r"(Centers for Medicare & Medicaid Services,.*?Responsibilities:\s+)",
            lambda m: m.group(1) + "• Integrated tools like: " + ", ".join(sorted(missing_keywords)) + ".\n",
            updated_text,
            flags=re.DOTALL | re.IGNORECASE
        )

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

    formatting_issues = ats_formatting_warnings(resume_text, resume.filename)
    content_issues = ats_content_warnings(resume_text, matched_hard, matched_soft)
    all_suggestions = formatting_issues + content_issues + [f"Missing Keywords: {missing_keywords}"]
    suggestions = " | ".join(all_suggestions)

    download_link = None
    if generate_download.lower() == "yes":
        enhanced_text = enhance_resume_text(resume_text, set(missing_keywords.split(", ")))
        output_path = generate_formatted_resume_docx(resume.filename, enhanced_text)
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
