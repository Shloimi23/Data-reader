import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

RECIPIENT_EMAIL = "sh5684770@gmail.com"


def reading_from_the_link(url):
    response = requests.get(url)

    soup = BeautifulSoup(response.text, "html.parser")
    rows = soup.find_all("tr", {"class": "data"})

    data = []

    for row in rows:
        name = row.find("td", {"class": "qName"}).find("a").text.strip()
        price = row.find("td", {"class": "last"}).text.strip()
        time = row.find("td", {"class": "lrtime"}).find("small").text.strip()
        change_element = row.find("span", {"class": "red bgr"})  # או לפי class אחר
        if change_element:
            change_value = change_element.text.strip()
        print(f"שם מניה: {name}, מחיר: {price}, סוג מסחר: {time}, שינוי יומי: {change_value}")
        data.append({"שם מניה": name, "מחיר": price, "שעה": time, "שינוי יומי": change_value })
    create_excel_file(data)


def create_excel_file(data):
    df = pd.DataFrame(data)

    current_time = datetime.now().strftime("%H_%M")
    file_name = f"snp_{current_time}.xlsx"

    df.to_excel(file_name, index=False, engine="openpyxl")

    wb = load_workbook(file_name)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        try:
            change_cell = row[3]
            if change_cell.value is not None:
                change = float(change_cell.value.replace("%", "").strip())
                if change > 1:
                    for cell in row:
                        cell.fill = fill
        except ValueError:
            continue

    wb.save(file_name)

    print(f"הנתונים נשמרו בהצלחה בקובץ {file_name}")
    print(f"קובץ האקסל נשמר בתיקייה: {os.getcwd()}")
    sort_excel_by_change(file_name)



def sort_excel_by_change(file_name):
    # קריאת קובץ ה-Excel כ-DataFrame
    df = pd.read_excel(file_name)

    # ניקוי עמודת השינוי והמרה ל-float
    df["שינוי יומי"] = df["שינוי יומי"].str.replace("%", "").astype(float)

    # מיון לפי עמודת "שינוי" בסדר עולה
    df = df.sort_values(by="שינוי יומי", ascending=False)

    # שמירת הטבלה הממוינת לקובץ חדש או דריסת הקיים
    sorted_file_name = file_name.replace(".xlsx", "_sorted.xlsx")
    df.to_excel(sorted_file_name, index=False, engine="openpyxl")

    print(f"הקובץ מוין לפי עמודת 'שינוי יומי' ונשמר בשם {sorted_file_name}")
    send_email_with_attachment(file_name)


def send_email_with_attachment(file_path):
    # הגדרות שליחת המייל
    sender_email = "tzviprus@gmail.com"  # כתובת השולח
    sender_password = "uzil nvms yxmk ltoi"  # סיסמת האפליקציה של השולח
    subject = "קובץ מניות חדש נוצר"
    body = "שלום,\n\nמצורף קובץ המניות שנוצר כעת.\n\nבברכה,\nהמערכת"

    # יצירת הודעת מייל
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = RECIPIENT_EMAIL
    msg["Subject"] = subject

    # הוספת גוף ההודעה
    msg.attach(MIMEText(body, "plain"))

    # הוספת הקובץ המצורף
    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(file_path)}",  # שם הקובץ המצורף
        )
        msg.attach(part)

    # שליחת המייל
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()  # חיבור מאובטח
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print(f"המייל נשלח בהצלחה לכתובת {RECIPIENT_EMAIL}")
    except Exception as e:
        print(f"שגיאה בשליחת המייל: {e}")

def waiting_for_a_full_hour():
    now = datetime.now()
    next_hour = (now + timedelta(hours=1)).replace(minute=0,second=0,microsecond=0)
    wait_time = (next_hour - now).total_seconds()
    print(f"ממתין {wait_time / 60:.2f} דקות עד לשעה {next_hour.strftime('%H:%M')}")
    time.sleep(wait_time)


def main(url):
    while True:
        print(f"מתחיל את התהליך בשעה: {datetime.now().strftime('%H:%M:%S')}")
        waiting_for_a_full_hour()
        reading_from_the_link(url)

# URL של העמוד
url = "https://www.globes.co.il/portal/instrument.aspx?instrumentid=373853&feeder=1&mode=composition&showAll=true#jt40991"
main(url)

