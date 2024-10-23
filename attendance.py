import pandas as pd
import webbrowser
import urllib.parse
import time
import pyautogui
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
import threading
import os
import csv
import sys
import re
from datetime import datetime
import calendar

def remove_trailing_zeros(number):
    if isinstance(number, float) and number.is_integer():
        return int(number)
    return number

def send_whatsapp_message_via_url(phone_number, message, name, recipient, retries=3):
    for attempt in range(retries):
        try:
            print(f"Drafting message for {name}'s {recipient} phone: {phone_number}")
            encoded_message = urllib.parse.quote(message)
            webbrowser.open(f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}")
            time.sleep(6)  # Give enough time for WhatsApp Web to load and draft the message
            pyautogui.press("enter")  # Automatically press "Enter" to send the message
            time.sleep(1)  # Wait for the message to be sent
            pyautogui.hotkey("ctrl", "w")  # Close the current browser tab
            print(f"Message sent and tab closed for {name}'s {recipient} phone successfully!")
            
            return True
        except Exception as e:
            print(f"Attempt {attempt+1} failed: {str(e)}")
    
    print(f"Failed to send message and close tab for {name}'s {recipient} phone after {retries} attempts.")
    return False

def create_attendance_message(name, roll_no, attendance_records, recipient_type, date_range):
    greeting = "*🗓 Greetings from Anees Defence Career Institute Pune (ADCI) 🗓*\n\n"
    if recipient_type == "student":
        message = f"{greeting}Dear {name},\n\n"
    else:
        message = f"{greeting}Dear Parent of {name},\n\n"
    
    message += f"🧾 The attendance details {'of your ward ' if recipient_type == 'parent' else ''}for the period {date_range} are as below: 🧾\n\n"
    message += f"📝 Name: {name}\n"
    message += f"📝 Roll No: {roll_no}\n\n"
    
    for date, lectures in attendance_records.items():
        date_obj = datetime.strptime(date, "%d-%m")
        formatted_date = date_obj.strftime("%d{} %b").format(get_ordinal_suffix(date_obj.day))
        message += f"📅 {formatted_date}:\n"
        for time, status in lectures.items():
            # Try to parse and format the time, if it fails, use the original string
            try:
                
                time_obj = datetime.strptime(time, "%I:%M %p")
                formatted_time = time_obj.strftime("%I:%M %p")
            except ValueError:
                formatted_time = time  # Use the original time string if parsing fails
            message += f"    {formatted_time}: {status}\n"
    
    total_lectures = sum(len(lectures) for lectures in attendance_records.values())
    total_present = sum(sum(status == 'P' for status in lectures.values()) for lectures in attendance_records.values())
    attendance_percentage = (total_present / total_lectures) * 100 if total_lectures > 0 else 0
    
    message += f"\nTotal Lectures: {total_lectures}\n"
    message += f"Total Present: {total_present}\n"
    message += f"Attendance Percentage: {attendance_percentage:.2f}%\n\n"
    
    message += ("Regards,\nTeam ADCI\n\n"
                "✏️ PARENTS, please ensure your ward attends all classes regularly and stays updated on their progress.\n"
                "📌 Note- Daily attendance for multiple lectures is essential for monitoring academic performance.")
    
    return message

def get_ordinal_suffix(day):
    if 11 <= day <= 13:
        return 'th'
    else:
        return {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')

def process_attendance_file(file_path, status_label):
    try:
        # Check file extension
        _, ext = os.path.splitext(file_path)
        if ext.lower() == '.csv':
            df = pd.read_csv(file_path)
        elif ext.lower() in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file format. Please use CSV or Excel file.")

        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Print column names for debugging
        print("Columns in the file:", df.columns.tolist())
        
        # Extract date columns (assuming they're in the format "DD-MM HH:MM am/pm")
        date_columns = [col for col in df.columns if '-' in col and ':' in col]
        
        if not date_columns:
            raise ValueError("No date columns found. Please check the file format.")
        
        print("Date columns found:", date_columns)
        
        # Get the date range for the message
        date_range = f"{date_columns[0].split()[0]} to {date_columns[-1].split()[0]}"
        
        messages_sent = 0
        
        for _, row in df.iterrows():
            name = row['Name']
            roll_no = row['Roll No.']
            student_phone_number = remove_trailing_zeros(row['Student Contact No.'])
            father_phone_number = remove_trailing_zeros(row['Father/Guardian Contact No.'])
            mother_phone_number = remove_trailing_zeros(row['Mother/Guardian Contact No.'])
            
            # Collect attendance data for this student
            attendance_records = {}
            for col in date_columns:
                # Split the column name into date and time
                date_time = col.split()
                if len(date_time) >= 2:
                    date = date_time[0]  # This should be in "DD-MM" format
                    time = ' '.join(date_time[1:])  # This will include the time and AM/PM if present
                else:
                    continue  # Skip this column if it doesn't have both date and time

                try:
                    date_obj = datetime.strptime(date, "%d-%m")  # Parse the date
                    formatted_date = date_obj.strftime("%d-%m")  # Keep it as "DD-MM" for consistency
                    
                    if formatted_date not in attendance_records:
                        attendance_records[formatted_date] = {}
                    
                    # Store the original time string, we'll format it later in the message creation
                    attendance_records[formatted_date][time] = 'P' if row[col] == 'P' else 'A'
                except ValueError:
                    print(f"Skipping invalid date format in column: {col}")
                    continue
            
            # Create the attendance messages
            student_message = create_attendance_message(name, roll_no, attendance_records, "student", date_range)
            parent_message = create_attendance_message(name, roll_no, attendance_records, "parent", date_range)
            
            # Send message to student
            if pd.notna(student_phone_number) and str(student_phone_number).strip():
                full_student_phone_number = f"+91{student_phone_number}"
                if send_whatsapp_message_via_url(full_student_phone_number, student_message, name, "student (Attendance)"):
                    messages_sent += 1
            
            # Send message to father
            if pd.notna(father_phone_number) and str(father_phone_number).strip():
                full_father_phone_number = f"+91{father_phone_number}"
                if send_whatsapp_message_via_url(full_father_phone_number, parent_message, name, "father (Attendance)"):
                    messages_sent += 1
            
            # Send message to mother
            if pd.notna(mother_phone_number) and str(mother_phone_number).strip():
                full_mother_phone_number = f"+91{mother_phone_number}"
                if send_whatsapp_message_via_url(full_mother_phone_number, parent_message, name, "mother (Attendance)"):
                    messages_sent += 1
        
        success_message = f"Attendance messages sent successfully! Total sent: {messages_sent}"
        print(success_message)
        status_label.config(text=success_message)
    
    except Exception as e:
        error_message = f"Error processing file or sending messages: {str(e)}"
        print(error_message)
        status_label.config(text=error_message)
        
        # Print more detailed error information
        import traceback
        traceback.print_exc()

class AttendanceApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Attendance Message Sender")
        self.geometry("600x400")
        self.configure(bg="#2c3e50")

        self.label = tk.Label(self, text="Drag and drop a CSV or XLSX file here", width=60, height=10, bg="#ecf0f1", fg="#2c3e50", relief="groove", bd=2)
        self.label.pack(padx=10, pady=20)
        self.label.bind("<Configure>", self.rounded_label)

        self.file_path = None
        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.on_drop)

        self.send_button = tk.Button(self, text="Send Attendance Messages", command=self.on_send, bg="#3498db", fg="white", font=("Helvetica", 14), relief="raised", bd=3)
        self.send_button.pack(padx=10, pady=10)
        self.send_button.bind("<Configure>", self.rounded_button)

        self.status_label = tk.Label(self, text="", bg="#2c3e50", fg="white", font=("Helvetica", 10))
        self.status_label.pack(pady=5)

    def rounded_label(self, event):
        self.label.config(highlightthickness=3, highlightbackground="#3498db")
        self.label.update_idletasks()

    def rounded_button(self, event):
        self.send_button.config(highlightthickness=3, highlightbackground="#3498db")
        self.send_button.update_idletasks()

    def on_drop(self, event):
        self.file_path = event.data.strip("{}")
        self.label.config(text=f"File selected: {self.file_path}")

    def on_send(self):
        if self.file_path:
            self.status_label.config(text="Processing...")
            self.update()
            threading.Thread(target=process_attendance_file, args=(self.file_path, self.status_label)).start()
        else:
            messagebox.showwarning("No file selected", "Please drag and drop a CSV or XLSX file first")

if __name__ == '__main__':
    app = AttendanceApp()
    app.mainloop()
