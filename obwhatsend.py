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
            time.sleep(8)  # Windows typically needs less time
        else:
            time.sleep(10)  # Other OS might need slightly more time
        
        # Speed optimization: Quick send and close
        pyautogui.press("enter")
        time.sleep(1)  # Minimum wait for send
        pyautogui.hotkey("ctrl", "w")
        
        return True
            
    except Exception as e:
        logging.error(f"Failed for {name}'s {recipient}: {str(e)}")
        try:
            pyautogui.hotkey("ctrl", "w")
        except:
            pass
        return False

def create_nda_message(name, roll_no, exams, recipient_type):
    """Create NDA message with pre-formatted strings for speed."""
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    if recipient_type == "student":
        message = f"{greeting}Dear {name},\n\n🧾 Your Academic progress for the following tests is as below: 🧾\n\n"
    else:
        message = f"{greeting}Dear Parent,\n\n🧾 The Academic progress detail of your ward {name} for the following tests is as below: 🧾\n\n"
    
    message += f"📝 Roll No: {roll_no}\n\n"

    for exam_name, marks in exams.items():
        total = "300" if "MATHS" in exam_name.upper() else "600"
        message += f"📊 {exam_name} Exam details -\nTotal Marks - {marks}/{total}\n\n"

    message += ("Regards,\nTeam ADCI\n\n"
                "📌 Note- NDA WEEKLY MATHS/GAT- OBJECTIVE TESTS\n"
                "✏PARENTS Do visit the Academy on a regular basis for your ward's progress.\n"
                "✏Check the Official ADCI Parents-Students WhatsApp group daily for new informative updates from ADCI.")
    
    return message

def create_jee_neet_message(name, roll_no, exams, recipient_type):
    """Create JEE/NEET message with pre-formatted strings for speed."""
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    if recipient_type == "student":
        message = f"{greeting}Dear {name},\n\n🧾 Your Academic progress for the following weekly JEE/NEET tests is as below: 🧾\n\n"
    else:
        message = f"{greeting}Dear Parent,\n\n🧾 The Academic progress detail of your ward {name} for the following weekly JEE/NEET tests is as below: 🧾\n\n"
    
    message += f"📝 Roll No: {roll_no}\n\n"

    for exam_name, marks in exams.items():
        total = "360" if "JEE" in exam_name.upper() else "720"
        message += f"📊 {exam_name} Test details -\nTotal Marks - {marks}/{total}\n\n"

    message += ("Regards,\nTeam ADCI\n\n"
                "📌 Note- Weekly JEE (360 marks) and NEET (720 marks) tests are conducted to track progress\n"
                "✏PARENTS Do visit the Academy on a regular basis for your ward's progress.\n"
                "✏Check the Official ADCI Parents-Students WhatsApp group daily for new informative updates from ADCI.")
    
    return message

def process_data(df):
    """Pre-process data for faster message sending."""
    processed_data = []
    
    for _, row in df.iterrows():
        name = row['Name']
        roll_no = row['Roll No']
        
        # Speed optimization: Process phone numbers once
        phone_numbers = {
            "student": f"+91{remove_trailing_zeros(row['Student Contact No.'])}" if pd.notna(row['Student Contact No.']) else None,
            "father": f"+91{remove_trailing_zeros(row['Father/Guardian Contact No.'])}" if pd.notna(row['Father/Guardian Contact No.']) else None,
            "mother": f"+91{remove_trailing_zeros(row['Mother/Guardian Contact No.'])}" if pd.notna(row['Mother/Guardian Contact No.']) else None
        }
        
        # Speed optimization: Process exams once
        nda_exams = {}
        jee_neet_exams = {}
        
        for i in range(1, (len(df.columns) - 4) // 2 + 1):
            if f'Exam{i}' in df.columns and f'Total Marks{i}' in df.columns:
                exam_name = row[f'Exam{i}']
                marks = row[f'Total Marks{i}']
                
                if pd.notna(exam_name) and str(exam_name).strip():
                    marks = remove_trailing_zeros(marks) if pd.notna(marks) else "Absent"
                    
                    if isinstance(exam_name, str) and ("JEE" in exam_name.upper() or "NEET" in exam_name.upper()):
                        jee_neet_exams[exam_name] = marks
                    else:
                        nda_exams[exam_name] = marks
        
        # Speed optimization: Pre-generate messages
        messages = {
            "nda": {
                "student": create_nda_message(name, roll_no, nda_exams, "student") if nda_exams else None,
                "parent": create_nda_message(name, roll_no, nda_exams, "parent") if nda_exams else None
            },
            "jee_neet": {
                "student": create_jee_neet_message(name, roll_no, jee_neet_exams, "student") if jee_neet_exams else None,
                "parent": create_jee_neet_message(name, roll_no, jee_neet_exams, "parent") if jee_neet_exams else None
            }
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
                    for exam_type in ["nda", "jee_neet"]:
                        message = messages[exam_type]["student" if recipient == "student" else "parent"]
                        if message:
                            if send_whatsapp_message_via_url(phone, message, name, f"{recipient} ({exam_type.upper()})"):
                                messages_sent += 1
        
        status_label.config(text=f"Complete! Messages sent: {messages_sent}")
        
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")
        logging.error(str(e))

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Integrated Exam Message Sender (NDA, JEE, NEET)")
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
