import logging
import json
import hashlib
from pathlib import Path
from typing import Tuple, Set
from app.services.openai_service import openai_service
from app.core.config import PROJECT_ROOT

logger = logging.getLogger("app")

# Cache file for JD skill extraction
JD_SKILLS_CACHE_PATH = PROJECT_ROOT / "jd_skills_cache.json"


class JDSkillExtractor:
    """Service for extracting technical and soft skills from job descriptions using AI."""
    
    def __init__(self):
        self.cache = {}
        self._load_cache()
    
    def _load_cache(self):
        """Load cached skill extractions from disk."""
        if JD_SKILLS_CACHE_PATH.exists():
            try:
                with open(JD_SKILLS_CACHE_PATH, "r", encoding="utf-8") as f:
                    self.cache = json.load(f)
                logger.info(f"✅ JD skills cache loaded with {len(self.cache)} entries")
            except Exception as e:
                logger.warning(f"⚠️ Failed to load JD skills cache: {e}")
                self.cache = {}
        else:
            logger.info("No existing JD skills cache found. Starting fresh.")
            self.cache = {}
    
    def _save_cache(self):
        """Save skill extraction cache to disk."""
        try:
            with open(JD_SKILLS_CACHE_PATH, "w", encoding="utf-8") as f:
                json.dump(self.cache, f, indent=2)
            logger.info(f"💾 JD skills cache saved with {len(self.cache)} entries")
        except Exception as e:
            logger.warning(f"⚠️ Failed to save JD skills cache: {e}")
    
    def _generate_cache_key(self, jd_text: str) -> str:
        """Generate a hash-based cache key for the job description."""
        # Normalize text: lowercase, strip whitespace
        normalized = jd_text.lower().strip()
        # Create SHA256 hash
        return hashlib.sha256(normalized.encode('utf-8')).hexdigest()[:16]
    
    def extract_skills_from_jd(self, jd_text: str) -> Tuple[Set[str], Set[str]]:
        """
        Extract technical (hard) and soft skills from a job description.
        
        Args:
            jd_text: The job description text
            
        Returns:
            Tuple of (hard_skills_set, soft_skills_set)
        """
        if not jd_text or len(jd_text.strip()) < 20:
            logger.warning("Job description too short for skill extraction")
            return set(), set()
        
        # Check cache first
        cache_key = self._generate_cache_key(jd_text)
        if cache_key in self.cache:
            logger.info(f"💾 Cache hit for JD skills (key: {cache_key})")
            cached_data = self.cache[cache_key]
            return set(cached_data["hard_skills"]), set(cached_data["soft_skills"])
        
        logger.info(f"🔄 Cache miss for JD skills (key: {cache_key}). Extracting with AI...")
        
        # Extract skills using OpenAI
        try:
            result = openai_service.extract_skills_from_job_description(jd_text)
            
            if result and "hard_skills" in result and "soft_skills" in result:
                # Normalize skills (lowercase, strip whitespace)
                hard_skills = {skill.lower().strip() for skill in result["hard_skills"]}
                soft_skills = {skill.lower().strip() for skill in result["soft_skills"]}
                
                # Remove empty strings
                hard_skills = {s for s in hard_skills if s}
                soft_skills = {s for s in soft_skills if s}
                
                logger.info(f"✅ Extracted {len(hard_skills)} hard skills and {len(soft_skills)} soft skills")
                logger.info(f"   Hard skills: {list(hard_skills)[:10]}{'...' if len(hard_skills) > 10 else ''}")
                logger.info(f"   Soft skills: {list(soft_skills)[:10]}{'...' if len(soft_skills) > 10 else ''}")
                
                # Cache the results
                self.cache[cache_key] = {
                    "hard_skills": list(hard_skills),
                    "soft_skills": list(soft_skills)
                }
                self._save_cache()
                
                return hard_skills, soft_skills
            else:
                logger.error("❌ Invalid response format from OpenAI skill extraction")
                return set(), set()
                
        except Exception as e:
            logger.error(f"❌ Failed to extract skills from JD: {e}", exc_info=True)
            return set(), set()


# Singleton instance
jd_skill_extractor = JDSkillExtractor()


def extract_skills_from_jd(jd_text: str) -> Tuple[Set[str], Set[str]]:
    """
    Convenience function to extract skills from job description.
    
    Args:
        jd_text: The job description text
        
    Returns:
        Tuple of (hard_skills_set, soft_skills_set)
    """
    return jd_skill_extractor.extract_skills_from_jd(jd_text)
