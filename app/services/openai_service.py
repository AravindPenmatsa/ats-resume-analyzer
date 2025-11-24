import logging
import json
from openai import OpenAI
from app.core.config import API_KEY, CACHE_PATH

logger = logging.getLogger("app")

class OpenAIService:
    def __init__(self):
        self.client = None
        self.bullet_cache = {}
        self._initialize_client()
        self._load_cache()

    def _initialize_client(self):
        if API_KEY:
            self.client = OpenAI(api_key=API_KEY)
            logger.info("✅ OpenAI client initialized successfully with API key.")
            try:
                models = self.client.models.list()
                # Log the first few model IDs as a confirmation
                model_ids = [model.id for model in models.data[:5]]
                logger.info(f"Successfully connected to OpenAI. Available models include: {model_ids}")
            except Exception as e:
                logger.error(f"❌ Could not connect to OpenAI API: {e}")
                self.client = None
        else:
            logger.warning("⚠️ Warning: OPENAI_API_KEY not found in environment variables. GPT features will be disabled.")

    def _load_cache(self):
        if CACHE_PATH.exists():
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                self.bullet_cache = json.load(f)
            logger.info(f"✅ Bullet point cache loaded from {CACHE_PATH} with {len(self.bullet_cache)} items.")
        else:
            self.bullet_cache = {}
            logger.info("No existing cache found. A new cache will be created.")

    def _save_cache(self):
        try:
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                json.dump(self.bullet_cache, f, indent=2)
        except Exception as cache_error:
            logger.warning(f"⚠️ Failed to save cache: {cache_error}")

    def verify_connection(self):
        """Verify that OpenAI client is properly configured and working."""
        if self.client is None:
            logger.error("❌ OpenAI client not initialized - API key missing or invalid")
            return False
        
        try:
            # Test with a simple API call
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": "Say 'test' in one word."}],
                max_tokens=5,
                temperature=0.1
            )
            
            if response and response.choices and response.choices[0].message.content:
                logger.info("✅ OpenAI client verified - API connection working")
                return True
            else:
                logger.error("❌ OpenAI API returned empty response")
                return False
                
        except Exception as e:
            logger.error(f"❌ OpenAI API test failed: {e}")
            return False

    def generate_project_bullet_point(self, keyword: str, variation: int = 1, job_role: str = None) -> str:
        """Generate varied project-specific bullet points for missing keywords with role context."""
        if self.client is None:
            logger.warning("⚠️ OpenAI client not configured. Returning placeholder bullet point.")
            return f"• Implemented {keyword.title()} solutions in project development (OpenAI API key not configured)"

        keyword = keyword.strip().lower()
        
        # Include job_role in cache key for role-specific caching
        if job_role:
            project_cache_key = f"project_{keyword}_v{variation}_role_{job_role.lower().replace(' ', '_')}"
        else:
            project_cache_key = f"project_{keyword}_v{variation}"
        
        # ✅ Return cached result if available
        if project_cache_key in self.bullet_cache:
            logger.info(f"💾 Cache hit for project keyword: '{keyword}' variation {variation}. Returning cached bullet point.")
            return self.bullet_cache[project_cache_key]

        logger.info(f"🔄 Cache miss for project keyword '{keyword}' variation {variation}. Calling GPT-4o API...")
        
        # Role-specific prompts that sound more authentic
        if job_role:
            # Extract key role type for better prompts
            role_lower = job_role.lower()
            
            if any(term in role_lower for term in ['sdet', 'qa', 'test', 'quality']):
                prompts = [
                    f"As an {job_role}, write a realistic bullet point about testing or automation work you did with '{keyword}'. Include specific testing activities like writing test cases, automation, or quality assurance. Sound like you actually did this work. Limit to 30 words.",
                    f"As an {job_role}, describe how you used '{keyword}' for test automation, framework development, or quality improvements. Be specific and authentic. Limit to 30 words.",
                    f"As an {job_role}, explain your hands-on experience with '{keyword}' in testing, CI/CD, or test infrastructure. Write as if describing your actual work. Limit to 30 words."
                ]
            elif any(term in role_lower for term in ['developer', 'engineer', 'programmer']):
                prompts = [
                    f"As a {job_role}, write a realistic bullet point about development work you did with '{keyword}'. Include specific coding, implementation, or architecture work. Sound authentic. Limit to 30 words.",
                    f"As a {job_role}, describe how you used '{keyword}' to build features, optimize performance, or solve technical problems. Be specific. Limit to 30 words.",
                    f"As a {job_role}, explain your hands-on experience with '{keyword}' in development, deployment, or system design. Write naturally. Limit to 30 words."
                ]
            else:
                # Generic professional role
                prompts = [
                    f"As a {job_role}, write a realistic bullet point about work you did with '{keyword}'. Be specific about your contributions and impact. Sound authentic. Limit to 30 words.",
                    f"As a {job_role}, describe how you used '{keyword}' to improve processes, deliver results, or solve problems. Be specific. Limit to 30 words.",
                    f"As a {job_role}, explain your practical experience with '{keyword}'. Focus on real work activities and outcomes. Limit to 30 words."
                ]
        else:
            # Fallback prompts without role context
            prompts = [
                f"Write a professional bullet point showing specific implementation and development work with '{keyword}'. Focus on technical execution and results. Use action verbs. Limit to 30 words.",
                f"Write a professional bullet point highlighting optimization, integration, or enhancement achieved using '{keyword}'. Emphasize improvements and problem-solving. Limit to 30 words.",
                f"Write a professional bullet point demonstrating testing, deployment, or automation involving '{keyword}'. Focus on technical processes and quality assurance. Limit to 30 words."
            ]
        
        prompt = prompts[(variation - 1) % len(prompts)]

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=90,
                temperature=0.7,  # Slightly lower for more focused, authentic responses
            )

            if not response or not response.choices or not response.choices[0].message.content:
                logger.warning(f"⚠️ GPT returned empty response for keyword '{keyword}' variation {variation}")
                return f"• Implemented {keyword.title()} solutions to enhance project functionality (empty response)"

            bullet = response.choices[0].message.content.strip()

            # ✅ Ensure bullet formatting
            if not bullet.startswith("•"):
                bullet = "• " + bullet

            # ✅ Cache and return
            logger.info(f"✅ Successfully generated project bullet for '{keyword}' variation {variation}. Caching result.")
            self.bullet_cache[project_cache_key] = bullet
            self._save_cache()

            return bullet

        except Exception as e:
            logger.error(f"❌ GPT API Exception for keyword '{keyword}' variation {variation}: {e}")
            
            # Verify if it's an API key issue
            if "api" in str(e).lower() and ("key" in str(e).lower() or "auth" in str(e).lower()):
                logger.error("❌ This appears to be an API key authentication issue!")
                
            return f"• Implemented {keyword.title()} solutions to enhance project functionality (API error: {str(e)[:50]})"

    def generate_summary_bullet(self, keyword: str, variation: int = 1) -> str:
        """Generate varied bullet points for Profile/Professional Summary section."""
        if self.client is None:
            logger.warning("OpenAI client not configured. Returning placeholder bullet point.")
            return f"• Experienced with {keyword.title()} technologies and best practices (OpenAI API key not configured)"

        keyword = keyword.strip().lower()
        
        # Use variation-specific cache key
        summary_cache_key = f"summary_{keyword}_v{variation}"
        
        # ✅ Return cached result if available
        if summary_cache_key in self.bullet_cache:
            logger.info(f"Cache hit for summary keyword: '{keyword}' variation {variation}. Returning cached bullet point.")
            return self.bullet_cache[summary_cache_key]

        logger.info(f"Cache miss for summary keyword '{keyword}' variation {variation}. Calling GPT-4o API.")
        
        # Different prompts for variations to ensure diversity
        prompts = [
            f"Write a professional summary bullet point highlighting expertise and experience with '{keyword}'. Focus on skills and capabilities. Use confident language. Limit to 30 words.",
            f"Write a professional summary bullet point showing proven track record and proficiency in '{keyword}'. Emphasize achievements and competence. Limit to 30 words.", 
            f"Write a professional summary bullet point demonstrating advanced knowledge and hands-on experience with '{keyword}'. Focus on practical application. Limit to 30 words."
        ]
        
        prompt = prompts[(variation - 1) % len(prompts)]

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=80,
                temperature=0.8,  # Higher temperature for more variation
            )

            bullet = response.choices[0].message.content
            if bullet is None:
                logger.warning("GPT response content was null.")
                return f"• Experienced with {keyword.title()} technologies and methodologies (no content received)"
            bullet = bullet.strip()

            # ✅ Ensure bullet formatting
            if not bullet.startswith("•"):
                bullet = "• " + bullet

            # ✅ Cache and return
            logger.info(f"Successfully generated summary bullet for '{keyword}' variation {variation}. Caching result.")
            self.bullet_cache[summary_cache_key] = bullet
            self._save_cache()

            return bullet

        except Exception as e:
            logger.error(f"❌ GPT API Exception for summary keyword '{keyword}' variation {variation}: {e}", exc_info=True)
            return f"• Experienced with {keyword.title()} technologies and methodologies (could not fetch GPT response)"

    def extract_skills_from_job_description(self, jd_text: str) -> dict:
        """
        Extract technical (hard) and soft skills from a job description using GPT-4o.
        
        Args:
            jd_text: The job description text
            
        Returns:
            Dictionary with 'hard_skills' and 'soft_skills' lists
            Example: {"hard_skills": ["python", "react", "aws"], "soft_skills": ["communication", "teamwork"]}
        """
        if self.client is None:
            logger.warning("⚠️ OpenAI client not configured. Cannot extract skills from JD.")
            return {"hard_skills": [], "soft_skills": []}
        
        # Truncate very long job descriptions to avoid token limits
        max_chars = 4000
        if len(jd_text) > max_chars:
            logger.info(f"Truncating JD from {len(jd_text)} to {max_chars} characters")
            jd_text = jd_text[:max_chars]
        
        prompt = f"""Analyze the following job description and extract all technical skills (hard skills) and soft skills mentioned.

Job Description:
{jd_text}

Instructions:
1. Extract ALL technical skills including: programming languages, frameworks, tools, platforms, databases, cloud services, methodologies, certifications, etc.
2. Extract ALL soft skills including: communication, leadership, teamwork, problem-solving, etc.
3. Normalize skill names to lowercase
4. Use common abbreviations where appropriate (e.g., "aws" instead of "amazon web services")
5. Return as JSON with two arrays: "hard_skills" and "soft_skills"

Return ONLY valid JSON in this exact format:
{{
  "hard_skills": ["skill1", "skill2", ...],
  "soft_skills": ["skill1", "skill2", ...]
}}"""

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=1500,
                temperature=0.3,  # Lower temperature for more consistent extraction
            )
            
            if not response or not response.choices or not response.choices[0].message.content:
                logger.error("❌ GPT returned empty response for JD skill extraction")
                return {"hard_skills": [], "soft_skills": []}
            
            result_text = response.choices[0].message.content.strip()
            
            # Parse JSON response
            try:
                result = json.loads(result_text)
                
                # Validate structure
                if "hard_skills" not in result or "soft_skills" not in result:
                    logger.error(f"❌ Invalid JSON structure from GPT: {result_text[:200]}")
                    return {"hard_skills": [], "soft_skills": []}
                
                # Ensure lists
                if not isinstance(result["hard_skills"], list):
                    result["hard_skills"] = []
                if not isinstance(result["soft_skills"], list):
                    result["soft_skills"] = []
                
                logger.info(f"✅ Successfully extracted {len(result['hard_skills'])} hard skills and {len(result['soft_skills'])} soft skills from JD")
                
                return result
                
            except json.JSONDecodeError as je:
                logger.error(f"❌ Failed to parse JSON from GPT response: {je}")
                logger.error(f"Response was: {result_text[:500]}")
                return {"hard_skills": [], "soft_skills": []}
                
        except Exception as e:
            logger.error(f"❌ GPT API Exception during JD skill extraction: {e}", exc_info=True)
            return {"hard_skills": [], "soft_skills": []}

# Singleton instance
openai_service = OpenAIService()
