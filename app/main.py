
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from app.routes import router as resume_router
from app.download_route import router as download_router

app = FastAPI()

app.mount("/static", StaticFiles(directory="templates"), name="static")
app.include_router(resume_router)
app.include_router(download_router)

@app.get("/", response_class=HTMLResponse)
async def root():
    return '''
    <html>
        <head><title>Resume Analyzer</title></head>
        <body style="font-family:sans-serif; text-align:center; padding-top:50px;">
            <h2>✅ Resume Analyzer App is Live</h2>
            <p><a href="/upload">Click here to upload a resume</a></p>
        </body>
    </html>
    '''
