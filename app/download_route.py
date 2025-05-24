
from fastapi import APIRouter
from fastapi.responses import FileResponse
import os

router = APIRouter()

@router.get("/download/{filename}")
async def download_resume(filename: str):
    filepath = os.path.join("generated_resumes", filename)
    return FileResponse(
        filepath,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=filename
    )
