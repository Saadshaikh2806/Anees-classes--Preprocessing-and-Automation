import pandas as pd
import webbrowser
import urllib.parse
import time
import pyautogui
import logging
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import os
import threading

# Speed optimization: Set pyautogui to be faster
pyautogui.PAUSE = 0.1  # Reduce default pause between actions
pyautogui.FAILSAFE = True

def setup_logging():
    logging.basicConfig(
        filename='whatsapp_sender.log',
        level=logging.INFO,
        format='%(asctime)s - %(message)s'
    )

def remove_trailing_zeros(number):
    if isinstance(number, float) and number.is_integer():
        return int(number)
    return number

def send_whatsapp_message_via_url(phone_number, message, name, recipient):
    """Optimized message sending function with minimal delays."""
    try:
        logging.info(f"Sending to {name}'s {recipient}")
        
        # Speed optimization: Pre-encode message
        encoded_message = urllib.parse.quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
        
        # Speed optimization: Use new=2 for faster tab opening
        webbrowser.open(url, new=2, autoraise=False)
        
        # Speed optimization: Dynamic wait based on system performance
        if os.name == 'nt':  # Windows
            time.sleep(6)  # Windows typically needs less time

        
        # Speed optimization: Quick send and close
        pyautogui.press("enter")
        time.sleep(2)  # Minimum wait for send
        pyautogui.hotkey("ctrl", "w")
        
        return True
            
    except Exception as e:
        logging.error(f"Failed for {name}'s {recipient}: {str(e)}")
        try:
            pyautogui.hotkey("ctrl", "w")
        except:
            pass
        return False

def create_message(name, tests, recipient_type):
    """Create message with pre-formatted strings for speed."""
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    if recipient_type == "student":
        message = f"{greeting}Dear {name},\n\n🧾 Your Academic progress for the following tests is as below: 🧾\n\n"
    else:
        message = f"{greeting}Dear Parent,\n\n🧾 The Academic progress detail of your ward {name} for the following tests is as below: 🧾\n\n"
    
    message += "📊 Subjective test details -\n"

    for test_date, subjects_marks in tests.items():
        if test_date != "N/A":
            message += f"📅 Date: {test_date}\n"
            for subject, marks in subjects_marks.items():
                if pd.notna(subject) and pd.notna(marks):
                    message += f"Subject: {subject}, Marks: {marks}\n"
            message += "\n"

    message += ("Regards,\nTeam ADCI\n\n"
                "📌 Note- \n"
                "✏Do visit the Academy on a regular basis for your ward's progress.\n"
                "✏Check the Official ADCI Parents-Students WhatsApp group daily for new informative updates from ADCI.")
    
    return message

def format_date(date_value):
    try:
        if pd.isna(date_value):
            return "N/A"
        if isinstance(date_value, pd.Timestamp):
            return date_value.strftime('%d-%m-%Y')
        return str(date_value)
    except:
        return "N/A"

def process_data(df):
    """Pre-process data for faster message sending."""
    processed_data = []
    
    for _, row in df.iterrows():
        name = row['NAME']
        
        # Speed optimization: Process phone numbers once
        phone_numbers = {
            "student": f"+91{remove_trailing_zeros(row['SELF NO'])}" if pd.notna(row['SELF NO']) else None,
            "father": f"+91{remove_trailing_zeros(row['FATHER NO'])}" if pd.notna(row['FATHER NO']) else None,
            "mother": f"+91{remove_trailing_zeros(row['MOTHER NO'])}" if pd.notna(row['MOTHER NO']) else None
        }
        
        # Speed optimization: Process tests once
        tests = {}
        for i in range(1, 11):
            date_col = f'Subjective Date{i}'
            subj_col = f'Subject{i}'
            marks_col = f'Subjective Marks{i}'
            
            if all(col in df.columns for col in [date_col, subj_col, marks_col]):
                test_date = format_date(row[date_col])
                if test_date != "N/A":
                    subject = row[subj_col]
                    marks = remove_trailing_zeros(row[marks_col])
                    if pd.isna(marks):
                        marks = "Absent"
                    if test_date not in tests:
                        tests[test_date] = {}
                    tests[test_date][subject] = marks
        
        # Speed optimization: Pre-generate messages
        messages = {
            "student": create_message(name, tests, "student"),
            "parent": create_message(name, tests, "parent")
        }
        
        processed_data.append((name, phone_numbers, messages))
    
    return processed_data

def send_messages(file_path, status_label):
    try:
        # Load and pre-process data
        df = pd.read_csv(file_path) if file_path.endswith('.csv') else pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        
        processed_data = process_data(df)
        messages_sent = 0
        total = len(processed_data)
        
        for idx, (name, phone_numbers, messages) in enumerate(processed_data, 1):
            status_label.config(text=f"Processing {name} ({idx}/{total})")
            status_label.update()
            
            for recipient, phone in phone_numbers.items():
                if phone:
                    message = messages["student"] if recipient == "student" else messages["parent"]
                    if send_whatsapp_message_via_url(phone, message, name, recipient):
                        messages_sent += 1
        
        status_label.config(text=f"Complete! Messages sent: {messages_sent}")
        
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")
        logging.error(str(e))

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fast WhatsApp Sender")
        self.geometry("600x400")
        self.configure(bg="#2c3e50")

        # Speed optimization: Create widgets once
        self.create_widgets()
        self.file_path = None

    def create_widgets(self):
        self.label = tk.Label(
            self, 
            text="Drop CSV/XLSX file here",
            width=60, height=10,
            bg="#ecf0f1", fg="#2c3e50",
            relief="groove"
        )
        self.label.pack(padx=10, pady=20)
        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.on_drop)

        self.status_label = tk.Label(
            self,
            text="Ready...",
            bg="#2c3e50", fg="white",
            font=("Helvetica", 10)
        )
        self.status_label.pack(pady=5)

        self.send_button = tk.Button(
            self,
            text="Send Messages",
            command=self.on_send,
            bg="#3498db", fg="white",
            font=("Helvetica", 14)
        )
        self.send_button.pack(padx=10, pady=10)

    def on_drop(self, event):
        self.file_path = event.data.strip("{}")
        self.label.config(text=f"File: {os.path.basename(self.file_path)}")

    def on_send(self):
        if not self.file_path:
            messagebox.showwarning("No file", "Please drop a file first")
            return
        
        self.send_button.config(state="disabled")
        threading.Thread(target=self.process_sending, daemon=True).start()

    def process_sending(self):
        try:
            send_messages(self.file_path, self.status_label)
        finally:
            self.send_button.config(state="normal")

if __name__ == '__main__':
    setup_logging()
    app = App()
    app.mainloop()