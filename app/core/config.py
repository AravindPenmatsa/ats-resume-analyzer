import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env from project root
# Assuming this file is in app/core/config.py, so project root is ../../
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
load_dotenv(dotenv_path=PROJECT_ROOT / ".env")

API_KEY = os.getenv("OPENAI_API_KEY")

# Directories
BASE_DIR = PROJECT_ROOT / "app"
UPLOAD_DIR = BASE_DIR / "uploads"
GENERATED_DIR = BASE_DIR / "generated_resumes"
STATIC_DIR = PROJECT_ROOT / "static"
TEMPLATES_DIR = PROJECT_ROOT / "templates"
CACHE_PATH = PROJECT_ROOT / "bullet_cache.json"
JD_SKILLS_CACHE_PATH = PROJECT_ROOT / "jd_skills_cache.json"  # Cache for extracted JD skills

# Ensure directories exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)

# Constants
# DEPRECATED: HARD_KEYWORDS and SOFT_KEYWORDS are no longer used for scoring.
# Skills are now extracted dynamically from job descriptions using AI.
# These are kept for backward compatibility only.
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
    "maintained", "assisted", "contributed"
}

from fastapi.templating import Jinja2Templates
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
templates.env.auto_reload = True

