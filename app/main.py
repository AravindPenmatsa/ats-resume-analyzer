from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import os

app = FastAPI()
templates = Jinja2Templates(directory="templates")

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/upload", response_class=HTMLResponse)
async def show_upload_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def handle_upload(
    request: Request,
    resume: UploadFile = File(...),
    jobdesc: UploadFile = File(...)
):
    resume_contents = await resume.read()
    jobdesc_contents = await jobdesc.read()

    return templates.TemplateResponse("result.html", {
        "request": request,
        "resume_name": resume.filename,
        "jobdesc_name": jobdesc.filename,
        "message": "Files uploaded successfully!"
    })
