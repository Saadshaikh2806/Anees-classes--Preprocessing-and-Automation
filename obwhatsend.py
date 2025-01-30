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
        
        encoded_message = urllib.parse.quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
        
        webbrowser.open(url, new=2, autoraise=False)
        
        if os.name == 'nt':  # Windows
            time.sleep(8)
        
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "w")
        
        return True
            
    except Exception as e:
        logging.error(f"Failed for {name}'s {recipient}: {str(e)}")
        try:
            pyautogui.hotkey("ctrl", "w")
        except:
            pass
        return False

def get_exam_total_marks(exam_name):
    """Determine total marks based on exam type."""
    exam_upper = exam_name.upper()
    if "MATHS" in exam_upper:
        return "300"
    elif "GAT" in exam_upper:
        return {"ENGLISH": "200", "GAT": "400"}  # Modified for GAT exam
    elif "JEE" in exam_upper:
        return "360"
    elif "NEET" in exam_upper:
        return "720"
    elif "CLAT" in exam_upper:
        return "120"
    elif "MHTCET" in exam_upper:
        return "150"
    return "50"  # Default case

def create_exam_message(name, exams, recipient_type, exam_category):
    """Generic message creator for all exam types."""
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    
    category_texts = {
        "nda": "NDA weekly tests",
        "jee_neet": "weekly JEE/NEET tests",
        "clat": "CLAT weekly tests",
        "mhtcet": "MHTCET weekly tests"
    }
    
    if recipient_type == "student":
        message = f"{greeting}Dear {name},\n\n🧾 Your Academic progress for the following {category_texts[exam_category]} is as below: 🧾\n\n"
    else:
        message = f"{greeting}Dear Parent,\n\n🧾 The Academic progress detail of your ward {name} for the following {category_texts[exam_category]} is as below: 🧾\n\n"

    for exam_name, marks_data in exams.items():
        if isinstance(marks_data, dict) and 'ENGLISH' in marks_data:  # For GAT exam with ENGLISH and GAT components
            message += f"📊 {exam_name} Test details -\n"
            if marks_data.get('ENGLISH') == "Absent" and marks_data.get('GAT') == "Absent":
                message += "Total Marks - Absent\n\n"
            else:
                message += f"ENGLISH Marks - {marks_data['ENGLISH']}/200\n"
                message += f"GAT Marks - {marks_data['GAT']}/400\n\n"
        else:  # For other exams including MATHS
            total_marks = get_exam_total_marks(exam_name)
            if marks_data == "Absent":
                message += f"📊 {exam_name} Test details -\nTotal Marks - Absent\n\n"
            else:
                message += f"📊 {exam_name} Test details -\nTotal Marks - {marks_data}/{total_marks}\n\n"

    notes = {
        "nda": "📌 Note- NDA WEEKLY MATHS/GAT- OBJECTIVE TESTS",
        "jee_neet": "📌 Note- Weekly JEE (360 marks) and NEET (720 marks) tests are conducted to track progress",
        "clat": "📌 Note- Weekly CLAT (120 marks) tests are conducted to track progress",
        "mhtcet": "📌 Note- Weekly MHTCET (150 marks) tests are conducted to track progress"
    }

    message += (f"Regards,\nTeam ADCI\n\n"
                f"{notes[exam_category]}\n"
                "✏PARENTS Do visit the Academy on a regular basis for your ward's progress.\n"
                "✏Check the Official ADCI Parents-Students WhatsApp group daily for new informative updates from ADCI.")
    
    return message

def process_data(df):
    """Pre-process data for faster message sending."""
    processed_data = []
    
    for _, row in df.iterrows():
        name = row['Name']
        
        phone_numbers = {
            "student": f"+91{remove_trailing_zeros(row['Student Contact No.'])}" if pd.notna(row['Student Contact No.']) else None,
            "father": f"+91{remove_trailing_zeros(row['Father/Guardian Contact No.'])}" if pd.notna(row['Father/Guardian Contact No.']) else None,
            "mother": f"+91{remove_trailing_zeros(row['Mother/Guardian Contact No.'])}" if pd.notna(row['Mother/Guardian Contact No.']) else None
        }
        
        exam_categories = {
            "nda": {},
            "jee_neet": {},
            "clat": {},
            "mhtcet": {}
        }
        
        num_exams = sum(1 for col in df.columns if col.startswith('Exam'))
        
        for i in range(1, num_exams + 1):
            
            exam_name = row.get(f'Exam{i}')
            
            if pd.notna(exam_name) and str(exam_name).strip():
                exam_upper = str(exam_name).upper()
                
                if "GAT" in exam_upper:
                    # Handle GAT exam with ENGLISH and GAT components
                    english_marks = row.get(f'ENGLISH{i}', None)
                    gat_marks = row.get(f'GAT{i}', None)
                    
                    # Debug logging
                    logging.info(f"Processing GAT exam for {name}:")
                    logging.info(f"English marks: {english_marks}, GAT marks: {gat_marks}")
                    
                    # Only mark as Absent if the marks are explicitly NaN or None
                    english_val = "Absent" if pd.isna(english_marks) else remove_trailing_zeros(english_marks)
                    gat_val = "Absent" if pd.isna(gat_marks) else remove_trailing_zeros(gat_marks)
                    
                    exam_categories["nda"][exam_name] = {
                        "ENGLISH": english_val,
                        "GAT": gat_val
                    }
                elif "MATHS" in exam_upper:
                    marks = row.get(f'Total Marks{i}', None)
                    marks_val = "Absent" if pd.isna(marks) else remove_trailing_zeros(marks)
                    exam_categories["nda"][exam_name] = marks_val
                elif "JEE" in exam_upper:
                    marks = row.get(f'Total Marks{i}', None)
                    marks_val = "Absent" if pd.isna(marks) else remove_trailing_zeros(marks)
                    exam_categories["jee_neet"][exam_name] = marks_val
                elif "NEET" in exam_upper:
                    marks = row.get(f'Total Marks{i}', None)
                    marks_val = "Absent" if pd.isna(marks) else remove_trailing_zeros(marks)
                    exam_categories["jee_neet"][exam_name] = marks_val
                elif "CLAT" in exam_upper:
                    marks = row.get(f'Total Marks{i}', None)
                    marks_val = "Absent" if pd.isna(marks) else remove_trailing_zeros(marks)
                    exam_categories["clat"][exam_name] = marks_val
                elif "MHTCET" in exam_upper:
                    marks = row.get(f'Total Marks{i}', None)
                    marks_val = "Absent" if pd.isna(marks) else remove_trailing_zeros(marks)
                    exam_categories["mhtcet"][exam_name] = marks_val
                else:
                    marks = row.get(f'Total Marks{i}', None)
                    marks_val = "Absent" if pd.isna(marks) else remove_trailing_zeros(marks)
                    exam_categories["nda"][exam_name] = marks_val
        
        # Generate messages for each category
        messages = {}
        for category, exams in exam_categories.items():
            if exams:  # Only create messages if there are exams in the category
                messages[category] = {
                    "student": create_exam_message(name, exams, "student", category),
                    "parent": create_exam_message(name, exams, "parent", category)
                }
            else:
                messages[category] = {"student": None, "parent": None}
        
        processed_data.append((name, phone_numbers, messages))
    
    return processed_data

def send_messages(file_path, status_label):
    try:
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
                    for exam_type in ["nda", "jee_neet", "clat", "mhtcet"]:
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
        self.title("Integrated Exam Message Sender (NDA, JEE, NEET, CLAT, MHTCET)")
        self.geometry("600x400")
        self.configure(bg="#2c3e50")
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