from fastapi import FastAPI, UploadFile, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from parser import parse_excel_to_csv
import smtplib
from email.message import EmailMessage
import os

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/")
def home():
    return {"message": "Dienstplan App çalışıyor!"}

# ✅ HTML form sayfası
@app.get("/upload", response_class=HTMLResponse)
async def upload_form(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})

# ✅ Excel dosyasını işleme kısmı
@app.post("/upload")
async def upload_file(file: UploadFile, email: str = Form(...)):
    content = await file.read()
    csv_data = parse_excel_to_csv(content)

    msg = EmailMessage()
    msg["Subject"] = "Muhammet Sariaslan – Dienstplan"
    msg["From"] = os.getenv("EMAIL_ADDRESS")
    msg["To"] = email
    msg.set_content("Dienstplan CSV ektedir.")
    msg.add_attachment(csv_data, filename="Dienstplan.csv",
                       maintype="text", subtype="csv")

    with smtplib.SMTP_SSL(os.getenv("SMTP_SERVER"), int(os.getenv("SMTP_PORT"))) as smtp:
        smtp.login(os.getenv("EMAIL_ADDRESS"), os.getenv("EMAIL_PASSWORD"))
        smtp.send_message(msg)

    return {"status": "ok", "detail": "E-mail gönderildi"}
