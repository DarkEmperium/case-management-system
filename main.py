import webview
import sqlite3
import smtplib
import os
import sys
import random
import string
import webbrowser
import urllib.parse
from email.message import EmailMessage
from datetime import datetime

# --- CONFIGURATION ---
SENDER_EMAIL = "chuamicrotech@gmail.com"
SENDER_PASSWORD = "gvap zipq qack wbaj"


def get_path(relative_path, internal=True):
  
    if internal:
        try:
            # PyInstaller temp folder
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
    else:
        # Actual folder where the EXE is sitting
        if getattr(sys, "frozen", False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


class Api:
    def __init__(self):
        # Database sits OUTSIDE the exe so it's not read-only
        self.db_path = get_path("database.db", internal=False)
        self.init_db()

    def init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """CREATE TABLE IF NOT EXISTS tickets 
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                    case_id TEXT UNIQUE, phone TEXT, email TEXT NULL, model TEXT, 
                    status TEXT, remarks TEXT, date TEXT)"""
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_search ON tickets(phone, email, case_id)"
            )

    def send_professional_email(self, recipient_email, model, case_id, status, remarks):
        # 1. Check for empty email explicitly
        if not recipient_email or str(recipient_email).strip().lower() in ["none", ""]:
            print(f"DEBUG >>> Skipping email for {case_id}: No valid address.")
            return

        template_path = get_path("email_template.html")
        try:
            if os.path.exists(template_path):
                with open(template_path, "r", encoding="utf-8") as f:
                    html_template = f.read()
            else:
                html_template = "<html><body><h1>Service Update</h1><p>{status}: {model} [{case_id}]</p></body></html>"

            final_html = html_template.format(
                status=status.upper(),
                model=model,
                case_id=case_id,
                remarks=remarks if remarks else "Laboratory analysis in progress.",
            )

            msg = EmailMessage()
            msg.set_content(f"Technical Status Update: {status} for your {model}")
            msg.add_alternative(final_html, subtype="html")
            msg["Subject"] = f"Service Update: {model} [{case_id}]"
            msg["From"] = f"Chua Micro Tech <{SENDER_EMAIL}>"
            msg["To"] = recipient_email

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
                smtp.send_message(msg)
            
            # This will now trigger only on actual success
            print(f"SUCCESS >>> Email sent to {recipient_email} for Case {case_id}")

        except Exception as e:
            print(f"TRANSMISSION ERROR >>> Case {case_id}: {e}")

    def open_whatsapp(self, phone, case_id, status, model):
        clean_phone = "".join(filter(str.isdigit, phone))
        if clean_phone.startswith("01"):
            clean_phone = "6" + clean_phone

        text = (
            f"*CHUA MICRO TECH | TECHNICAL STATUS*\n\n"
            f"*Device Model:* {model}\n"
            f"*Case Number:* {case_id}\n"
            f"*Current Status:* {status.upper()}\n\n"
            f"Your unit is currently being processed by our technical team. Kindly check your email for the service report and latest updates."
        )

        encoded_text = urllib.parse.quote(text)
        protocol_url = f"whatsapp://send?phone={clean_phone}&text={encoded_text}"

        try:
            os.startfile(protocol_url)
        except Exception:
            web_url = f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_text}"
            webbrowser.open(web_url)

    def add_ticket(self, phone, email, model, remarks):
        random_id = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
        case_id = f"CMT-{random_id}"
        date = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Check if email is just whitespace or empty
        email_val = email.strip() if email and email.strip() else None

        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute(
                    "INSERT INTO tickets (case_id, phone, email, model, status, remarks, date) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (case_id, phone, email_val, model, "Case Logged", remarks, date),
                )

            # Only attempt to send email if email_val is not None
            if email_val:
                self.send_professional_email(
                    email_val, model, case_id, "Case Logged", remarks
                )
            else:
                print(f"Case {case_id} logged without email. Skipping notification.")

            return {"status": "success", "case_id": case_id}
        except Exception as e:
            print(f"Database Error: {e}")
            return {"status": "error"}

    def get_tickets(self, search="", view_type="active"):
        with sqlite3.connect(self.db_path) as conn:
            operator = "=" if view_type == "completed" else "!="
            query = f"SELECT * FROM tickets WHERE status {operator} 'Completed' "
            params = []
            if search:
                query += "AND (phone LIKE ? OR email LIKE ? OR case_id LIKE ? OR model LIKE ?) "
                search_param = f"%{search}%"
                params.extend([search_param] * 4)
            query += "ORDER BY id DESC"
            return conn.execute(query, params).fetchall()

    def update_status(self, db_id, new_status):
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                # Fetch current info to see if we have an email
                cursor.execute(
                    "SELECT email, model, case_id, remarks FROM tickets WHERE id=?", (db_id,)
                )
                t = cursor.fetchone()
                
                if t:
                    conn.execute("UPDATE tickets SET status=? WHERE id=?", (new_status, db_id))
                    
                    # Log the database change
                    print(f"DATABASE >>> Case {t[2]} status updated to {new_status}")
                    
                    # Pass the email through our checker
                    self.send_professional_email(t[0], t[1], t[2], new_status, t[3])
                    return True
            return False
        except Exception as e:
            print(f"UPDATE FAILURE >>> {e}")
            return False

    def delete_ticket(self, db_id):
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.execute("DELETE FROM tickets WHERE id=?", (db_id,))
            return True
        except Exception as e:
            return False
        

if __name__ == "__main__":
    # Essential for some Windows environments
    os.environ["WEBKIT_DISABLE_COMPOSITING_MODE"] = "1"

    api = Api()
    # GUI is internal (bundled)
    gui_path = get_path("gui.html")

    window = webview.create_window(
        "CHUA MICRO TECH SYSTEM",
        gui_path,
        js_api=api,
        width=1350,
        height=850,
        background_color="#0b0e14",  # Matches your CSS --bg
    )
    webview.start()
