import re

def validate_resume_format(text: str) -> str:
    errors = []

    # Check for tables or non-ATS-friendly formatting (basic heuristic)
    if re.search(r'\+|\||={2,}', text):
        errors.append("⚠️ Avoid using tables or complex formatting.")

    # Check for non-standard fonts or symbols (basic heuristic)
    if re.search(r'[★✓■●♦▪]', text):
        errors.append("⚠️ Use standard fonts and avoid special symbols like ★ or ✓.")

    # Check minimum length (very short resumes or parsing failures)
    if len(text.strip().split()) < 100:
        errors.append("⚠️ Resume text appears too short. Ensure full content is extracted.")

    return " ".join(errors)


def extract_keywords(text: str) -> set:
    return set(word.lower().strip(".,()") for word in text.split() if len(word) > 2)


def validate_resume_formatting(text: str) -> list:
    warnings = []

    if not re.search(r"(Work Experience|Professional Experience)", text, re.IGNORECASE):
        warnings.append("⚠️ Missing standard section 'Work Experience'.")

    if re.search(r"[•▪▶❖➤]", text):
        warnings.append("❌ Use standard bullet symbols like '-' instead of decorative icons.")

    if not re.search(r"(Summary|Objective)", text, re.IGNORECASE):
        warnings.append("⚠️ Consider adding a professional Summary or Objective section.")

    if re.search(r"\d{1,2}/\d{2,4}", text):
        warnings.append("⚠️ Use consistent date format like MM/YYYY or Month YYYY.")

    if len(re.findall(r"([A-Z][A-Za-z\s]+):", text)) < 3:
        warnings.append("⚠️ Use proper section headings like 'Skills', 'Education', etc.")

    return warnings


def validate_readability(text: str) -> list:
    issues = []

    if len(text.split()) < 150:
        issues.append("✏️ Resume seems short. Consider expanding with accomplishments.")

    if "lorem ipsum" in text.lower() or "dummy text" in text.lower():
        issues.append("❌ Remove placeholder or dummy text.")

    if not any(verb in text.lower() for verb in ["developed", "implemented", "managed", "led"]):
        issues.append("🔧 Consider adding action verbs to bullet points.")

    return issues
