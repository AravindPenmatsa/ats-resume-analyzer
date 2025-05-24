# Logic to compare resume and JD
from collections import Counter
import re

def extract_keywords(text):
    # Basic keyword extraction: lowercased words, removing common stopwords
    stopwords = {"the", "and", "to", "of", "in", "for", "on", "with", "at", "by", "an", "be", "is", "are", "was", "as"}
    words = re.findall(r"\b\w+\b", text.lower())
    return [word for word in words if word not in stopwords]

def calculate_ats_score(resume_text, jd_text):
    resume_keywords = extract_keywords(resume_text)
    jd_keywords = extract_keywords(jd_text)

    resume_counts = Counter(resume_keywords)
    jd_counts = Counter(jd_keywords)

    matched_keywords = set(resume_keywords) & set(jd_keywords)
    total_keywords = len(set(jd_keywords))

    if total_keywords == 0:
        return 0

    match_score = len(matched_keywords) / total_keywords * 100

    suggestions = list(set(jd_keywords) - set(resume_keywords))

    return {
        "score": round(match_score, 2),
        "matched_keywords": list(matched_keywords),
        "missing_keywords": suggestions
    }
