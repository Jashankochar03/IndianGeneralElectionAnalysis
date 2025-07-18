import time
import yagmail
import os
import shutil
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import tkinter as tk
from tkinter import messagebox


# === CONFIG ===
TEMP_FOLDER = r"C:\Users\jashk\AppData\Local\Temp\Power BI Desktop"
PDF_NAME = "GeneralElection2024.pdf"
SAFE_COPY = r"C:\ElectionProject\report.pdf"
EMAIL_USER = "jaskochar2003@gmail.com"
EMAIL_PASS = "ahnd mhwq sivg mwzt"  # Gmail App Password
EMAIL_TO = ["jashkochar@gmail.com", "jaskochar2003@gmail.com"]

LOG_FILE = r"C:\ElectionProject\automation_log.txt"

# === EMAIL SENDER ===
def send_pdf(copied_path):
    try:
        # Confirmation popup
        root = tk.Tk()
        root.withdraw()
        answer = messagebox.askyesno(
            "Send Power BI Report?",
            "Do you want to send the exported Power BI report via email?"
        )
        root.destroy()

        if not answer:
            print("🚫 User canceled email.")
            return

        # Send email
        print(f"📤 Sending: {copied_path}")
        yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASS)
        yag.send(
            to=EMAIL_TO,
            subject="📊 Power BI Report - Auto Export",
            contents="Hi,\n\nPlease find the attached Power BI report.\n\nRegards,\nAutomation",
            attachments=copied_path
        )
        os.remove(copied_path)
        print("✅ Email sent and file deleted.")

        with open(LOG_FILE, "a") as f:
            f.write(f"{datetime.now()} - Sent: {copied_path}\n")

    except Exception as e:
        print(f"❌ Email failed: {e}")


# === FILE WATCHER ===
class PowerBIPDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith(PDF_NAME):
            print(f"📄 Detected export: {event.src_path}")
            time.sleep(5)  # Wait for file write to finish
            try:
                shutil.copy(event.src_path, SAFE_COPY)
                print(f"📥 Copied to: {SAFE_COPY}")
                send_pdf(SAFE_COPY)
            except Exception as e:
                print(f"❌ Copy/Send error: {e}")

# === MAIN RUN ===
if __name__ == "__main__":
    print(f"👁️ Watching: {TEMP_FOLDER}")
    os.makedirs(os.path.dirname(SAFE_COPY), exist_ok=True)

    observer = Observer()
    observer.schedule(PowerBIPDFHandler(), TEMP_FOLDER, recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("🛑 Stopped.")
    observer.join()
