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

def convert_to_12hour(time_str):
    """Convert 24-hour time format to 12-hour format with AM/PM."""
    try:
        # Parse the time string
        time_obj = datetime.strptime(time_str, "%H:%M")
        # Convert to 12-hour format
        return time_obj.strftime("%I:%M %p")
    except ValueError:
        return time_str  # Return original if conversion fails

def read_attendance_data(file_path):
    """Read and process attendance data from Excel file."""
    try:
        df = pd.read_excel(file_path)
        
        # Get the base columns and time columns
        base_columns = ['EMP CODE', 'NAME', 'MOTHER NO', 'FATHER NO', 'SELF NO']
        time_columns = [col for col in df.columns if col.startswith('Time')]
        
        # Create a mapping for our standardized column names
        column_mapping = {
            'EMP CODE': 'EMP_CODE',
            'NAME': 'NAME',
            'MOTHER NO': 'MOTHER_NO',
            'FATHER NO': 'FATHER_NO',
            'SELF NO': 'SELF_NO'
        }
        # Add time columns to mapping
        for col in time_columns:
            column_mapping[col] = col.replace(' ', '_')
            
        # Rename columns using the mapping
        df = df.rename(columns=column_mapping)
        
        # Process contact information
        df['CONTACTS'] = df.apply(lambda row: {
            'mother': str(row['MOTHER_NO']) if pd.notna(row['MOTHER_NO']) else None,
            'father': str(row['FATHER_NO']) if pd.notna(row['FATHER_NO']) else None,
            'self': str(row['SELF_NO']) if pd.notna(row['SELF_NO']) else None
        }, axis=1)
        
        return df
    except Exception as e:
        logging.error(f"Error reading attendance data: {str(e)}")
        raise

def process_attendance_data(df):
    """Process attendance data and prepare messages."""
    messages = []
    current_date = datetime.now().strftime("%d-%m-%Y")
    
    # Get all time columns dynamically
    time_columns = [col for col in df.columns if col.startswith('Time_')]
    
    for _, row in df.iterrows():
        contacts = row['CONTACTS']
        
        # Get attendance records
        attendance_records = {}
        for time_col in time_columns:
            if pd.notna(row[time_col]):
                attendance_records[row[time_col]] = 'P'
        
        if attendance_records:
            # Create messages for each available contact
            for contact_type, phone_number in contacts.items():
                if phone_number and phone_number != 'nan' and phone_number != 'NO PHONE':
                    message = create_attendance_message(
                        name=row['NAME'],
                        roll_no=row['EMP_CODE'],
                        phone_number=phone_number,
                        attendance_records=attendance_records,
                        recipient_type=contact_type,
                        date=current_date
                    )
                    messages.append({
                        'name': row['NAME'],
                        'phone_number': f"+91{phone_number.replace('+91', '')}",  # Ensure proper format
                        'message': message,
                        'recipient': contact_type
                    })
    
    return messages

def create_attendance_message(name, roll_no, phone_number, attendance_records, recipient_type, date):
    """Create attendance message showing IN/OUT times in 12-hour format."""
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    
    # Customize greeting based on recipient
    recipient_greeting = {
        'mother': f'Dear Mother of {name}',
        'father': f'Dear Father of {name}',
        'self': f'Dear {name}'
    }.get(recipient_type, f'Dear {name}')
    
    message = (
        f"{greeting}"
        f"{recipient_greeting},\n\n"
        f"🧾 Attendance details for {date}:\n\n"
        f"📝 Name: {name}\n"
        f"📝 EMP Code: {roll_no}\n\n"
        f"Today's IN/OUT Times:\n"
    )
    
    # Sort times chronologically and convert to 12-hour format
    times = sorted(attendance_records.keys())
    for i, time in enumerate(times, 1):
        status = "IN" if i % 2 != 0 else "OUT"
        time_12hr = convert_to_12hour(time)
        message += f"⏰ {status} Time: {time_12hr}\n"
    
    message += "\nThank you for your attention to this matter.\n"
    message += "Best regards,\nADCI Team"
    
    return message

def send_attendance_messages(file_path, status_label):
    """Main function to process and send attendance messages."""
    try:
        # Load and pre-process data
        df = read_attendance_data(file_path)
        messages = process_attendance_data(df)
        messages_sent = 0
        total = len(messages)
        
        for idx, message in enumerate(messages, 1):
            status_label.config(text=f"Processing {message['name']} ({idx}/{total})")
            status_label.update()
            
            # Send to student
            if send_whatsapp_message_via_url(message['phone_number'], message['message'], message['name'], message['recipient']):
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
            text="Drop Excel file here",
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