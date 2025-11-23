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

    def generate_project_bullet_point(self, keyword: str, variation: int = 1) -> str:
        """Generate varied project-specific bullet points for missing keywords."""
        if self.client is None:
            logger.warning("⚠️ OpenAI client not configured. Returning placeholder bullet point.")
            return f"• Implemented {keyword.title()} solutions in project development (OpenAI API key not configured)"

        keyword = keyword.strip().lower()
        
        # Use variation-specific cache key for project bullets
        project_cache_key = f"project_{keyword}_v{variation}"
        
        # ✅ Return cached result if available
        if project_cache_key in self.bullet_cache:
            logger.info(f"💾 Cache hit for project keyword: '{keyword}' variation {variation}. Returning cached bullet point.")
            return self.bullet_cache[project_cache_key]

        logger.info(f"🔄 Cache miss for project keyword '{keyword}' variation {variation}. Calling GPT-4o API...")
        
        # Different prompts for variations to ensure diversity
        prompts = [
            f"Write a project bullet point showing specific implementation and development work with '{keyword}'. Focus on technical execution and results. Use action verbs. Limit to 35 words.",
            f"Write a project bullet point highlighting optimization, integration, or enhancement achieved using '{keyword}'. Emphasize improvements and problem-solving. Limit to 35 words.",
            f"Write a project bullet point demonstrating testing, deployment, or automation involving '{keyword}'. Focus on technical processes and quality assurance. Limit to 35 words."
        ]
        
        prompt = prompts[(variation - 1) % len(prompts)]

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=90,
                temperature=0.8,  # Higher temperature for more variation
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

# Singleton instance
openai_service = OpenAIService()
