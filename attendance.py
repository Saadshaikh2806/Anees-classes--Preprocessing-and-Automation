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
from datetime import datetime

# Speed optimization: Set pyautogui to be faster
pyautogui.PAUSE = 0.1  # Reduce default pause between actions
pyautogui.FAILSAFE = True

def setup_logging():
    """Initialize logging with proper format."""
    logging.basicConfig(
        filename='attendance_sender.log',
        level=logging.INFO,
        format='%(asctime)s - %(message)s'
    )

def remove_trailing_zeros(number):
    """Remove trailing zeros from numbers."""
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

def get_ordinal_suffix(day):
    """Get ordinal suffix for day."""
    if 11 <= day <= 13:
        return 'th'
    return {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')

def create_attendance_message(name, roll_no, attendance_records, recipient_type, date_range):
    """Create optimized attendance message with pre-formatted strings."""
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    
    # Speed optimization: Use string formatting instead of concatenation
    message = (
        f"{greeting}"
        f"Dear {('Parent of ' + name) if recipient_type == 'parent' else name},\n\n"
        f"🧾 The attendance details {'of your ward ' if recipient_type == 'parent' else ''}"
        f"for the period {date_range} are as below: 🧾\n\n"
        f"📝 Name: {name}\n"
        f"📝 Roll No: {roll_no}\n\n"
    )

    # Speed optimization: Pre-calculate attendance stats
    total_lectures = 0
    total_present = 0
    
    # Build attendance details
    for date, lectures in attendance_records.items():
        date_obj = datetime.strptime(date, "%d-%m")
        formatted_date = date_obj.strftime(f"%d{get_ordinal_suffix(date_obj.day)} %b")
        message += f"📅 {formatted_date}:\n"
        
        for time, status in lectures.items():
            message += f"    {time}: {status}\n"
            total_lectures += 1
            if status == 'P':
                total_present += 1

    # Calculate attendance percentage
    attendance_percentage = (total_present / total_lectures * 100) if total_lectures > 0 else 0
    
    # Speed optimization: Use single f-string for stats
    message += (
        f"\nTotal Lectures: {total_lectures}\n"
        f"Total Present: {total_present}\n"
        f"Attendance Percentage: {attendance_percentage:.2f}%\n\n"
        "Regards,\nTeam ADCI\n\n"
        "✏️ PARENTS, please ensure your ward attends all classes regularly and stays updated on their progress.\n"
        "📌 Note- Daily attendance for multiple lectures is essential for monitoring academic performance."
    )
    
    return message

def process_attendance_data(df):
    """Pre-process attendance data for faster message sending."""
    processed_data = []
    
    # Extract date columns
    date_columns = [col for col in df.columns if '-' in col and ':' in col]
    date_range = f"{date_columns[0].split()[0]} to {date_columns[-1].split()[0]}"
    
    for _, row in df.iterrows():
        name = row['Name']
        roll_no = row['Roll No.']
        
        # Speed optimization: Process phone numbers once
        phone_numbers = {
            "student": f"+91{remove_trailing_zeros(row['Student Contact No.'])}" if pd.notna(row['Student Contact No.']) else None,
            "father": f"+91{remove_trailing_zeros(row['Father/Guardian Contact No.'])}" if pd.notna(row['Father/Guardian Contact No.']) else None,
            "mother": f"+91{remove_trailing_zeros(row['Mother/Guardian Contact No.'])}" if pd.notna(row['Mother/Guardian Contact No.']) else None
        }
        
        # Speed optimization: Process attendance records once
        attendance_records = {}
        for col in date_columns:
            date_time = col.split()
            if len(date_time) >= 2:
                date = date_time[0]
                time = ' '.join(date_time[1:])
                
                try:
                    date_obj = datetime.strptime(date, "%d-%m")
                    formatted_date = date_obj.strftime("%d-%m")
                    
                    if formatted_date not in attendance_records:
                        attendance_records[formatted_date] = {}
                    
                    attendance_records[formatted_date][time] = 'P' if row[col] == 'P' else 'A'
                except ValueError:
                    logging.warning(f"Invalid date format in column: {col}")
                    continue
        
        # Speed optimization: Pre-generate messages
        messages = {
            "student": create_attendance_message(name, roll_no, attendance_records, "student", date_range),
            "parent": create_attendance_message(name, roll_no, attendance_records, "parent", date_range)
        }
        
        processed_data.append((name, phone_numbers, messages))
    
    return processed_data

def send_attendance_messages(file_path, status_label):
    """Main function to process and send attendance messages."""
    try:
        # Load and pre-process data
        df = pd.read_csv(file_path) if file_path.endswith('.csv') else pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        
        processed_data = process_attendance_data(df)
        messages_sent = 0
        total = len(processed_data)
        
        for idx, (name, phone_numbers, messages) in enumerate(processed_data, 1):
            status_label.config(text=f"Processing {name} ({idx}/{total})")
            status_label.update()
            
            # Send to student
            if phone_numbers["student"]:
                if send_whatsapp_message_via_url(phone_numbers["student"], messages["student"], name, "student"):
                    messages_sent += 1
            
            # Send to parents
            for recipient in ["father", "mother"]:
                if phone_numbers[recipient]:
                    if send_whatsapp_message_via_url(phone_numbers[recipient], messages["parent"], name, recipient):
                        messages_sent += 1
        
        status_label.config(text=f"Complete! Messages sent: {messages_sent}")
        
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")
        logging.error(str(e))

class AttendanceApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fast Attendance Message Sender")
        self.geometry("600x400")
        self.configure(bg="#2c3e50")
        
        # Speed optimization: Create widgets once
        self.create_widgets()
        self.file_path = None

    def create_widgets(self):
        """Create all UI widgets with optimized settings."""
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
            text="Send Attendance Messages",
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
            send_attendance_messages(self.file_path, self.status_label)
        finally:
            self.send_button.config(state="normal")

if __name__ == '__main__':
    setup_logging()
    app = AttendanceApp()
    app.mainloop()