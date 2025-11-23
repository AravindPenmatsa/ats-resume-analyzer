import logging
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from app.core.logging import setup_logging
from app.core.config import STATIC_DIR
from app.routers.resume import router as resume_router
from app.services.openai_service import openai_service

# Setup logging
setup_logging()
logger = logging.getLogger("app")

# Initialize FastAPI app
app = FastAPI(title="ATS Resume Analyzer")

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Include routers
app.include_router(resume_router)

# Startup event to verify OpenAI connection
@app.on_event("startup")
async def startup_event():
    if openai_service.verify_connection():
        logger.info("🚀 All systems ready - OpenAI GPT enhancement fully operational!")
    else:
        logger.warning("⚠️ OpenAI verification failed - GPT enhancement may not work properly")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)