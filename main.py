from fastapi import FastAPI, UploadFile, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from parser import parse_excel_to_csv
import os
import smtplib
from email.message import EmailMessage

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app = FastAPI()
templates = Jinja2Templates(directory=TEMPLATE_DIR)

@app.get("/")
def home():
    return {"message": "Dienstplan App çalışıyor!"}

@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})

@app.post("/upload")
async def upload_file(file: UploadFile, email: str = Form(...)):
    content = await file.read()
    csv_data = parse_excel_to_csv(content)

    msg = EmailMessage()
    msg["Subject"] = "Muhammet Sariaslan – Dienstplan"
    msg["From"] = os.getenv("EMAIL_ADDRESS")
    msg["To"] = email
    msg.set_content("Dienstplan CSV eklendi.")
    msg.add_attachment(csv_data, filename="Dienstplan.csv", maintype="text", subtype="csv")

    with smtplib.SMTP_SSL(os.getenv("SMTP_SERVER"), int(os.getenv("SMTP_PORT"))) as smtp:
        smtp.login(os.getenv("EMAIL_ADDRESS"), os.getenv("EMAIL_PASSWORD"))
        smtp.send_message(msg)

    return {"status": "ok", "detail": "E‑mail gönderildi"}
