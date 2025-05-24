
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse
from app.routes import router

app = FastAPI()
app.include_router(router)

@app.get("/", response_class=HTMLResponse)
def homepage():
    return """
    <h2>✅ Resume Analyzer App is Live</h2>
    <a href='/upload'>Click here to upload a resume</a>
    """

@app.get("/upload", response_class=HTMLResponse)
def upload_page():
    return """
    <h3>Upload Your Resume (TXT or PDF):</h3>
    <form action="/evaluate" method="post" enctype="multipart/form-data">
        Resume: <input type="file" name="resume"><br><br>
        Job Description: <input type="file" name="jobdesc"><br><br>
        <input type="submit" value="Evaluate">
    </form>
    """
