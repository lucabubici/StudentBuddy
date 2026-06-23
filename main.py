from fastapi import FastAPI, UploadFile, File
from fastapi.responses import PlainTextResponse
from fastapi.middleware.cors import CORSMiddleware
from converter import extract_pptx_to_md

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/convert", response_class=PlainTextResponse)
async def convert(file: UploadFile = File(...)):
    contents = await file.read()
    return extract_pptx_to_md(contents, file.filename)

@app.get("/")
def root():
    return {"message": "StudentBuddy is running"}