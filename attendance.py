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

def process_attendance_data(df, date):
    """Process attendance data and prepare messages."""
    messages = []
    
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
                        date=date
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

def send_attendance_messages(file_path, status_label, date):
    """Main function to process and send attendance messages."""
    try:
        # Load and pre-process data
        df = read_attendance_data(file_path)
        messages = process_attendance_data(df, date)
        messages_sent = 0
        total = len(messages)
        
        for idx, message in enumerate(messages, 1):
            status_label.config(text=f"File: {os.path.basename(file_path)} - Processing {message['name']} ({idx}/{total})")
            status_label.update()
            
            # Send to student
            if send_whatsapp_message_via_url(message['phone_number'], message['message'], message['name'], message['recipient']):
                messages_sent += 1
        
        status_label.config(text=f"Complete! Messages sent: {messages_sent} for {os.path.basename(file_path)}")
        
    except Exception as e:
        status_label.config(text=f"Error processing {os.path.basename(file_path)}: {str(e)}")

class DateInputDialog(tk.Toplevel):
    def __init__(self, parent, files):
        super().__init__(parent)
        self.title("Enter Dates")
        self.geometry("500x400")
        self.configure(bg="#2c3e50")
        
        self.dates = {}  # Store file:date pairs
        self.files = files
        self.result = None
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        
        # Center the dialog
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        # Title
        title = tk.Label(
            self,
            text="Enter Date for Each File",
            font=("Helvetica", 14, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title.pack(pady=10)

        # Create a frame for the canvas and scrollbar
        container = tk.Frame(self, bg="#2c3e50")
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create a canvas
        canvas = tk.Canvas(container, bg="#2c3e50", highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg="#2c3e50")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the widgets
        canvas.pack(side="left", fill=tk.BOTH, expand=True)
        scrollbar.pack(side="right", fill="y")

        # Add date entries for each file
        self.entries = {}
        for file_path in self.files:
            frame = tk.Frame(self.scrollable_frame, bg="#2c3e50")
            frame.pack(fill="x", padx=5, pady=5)

            filename = os.path.basename(file_path)
            tk.Label(
                frame,
                text=filename,
                bg="#2c3e50",
                fg="white",
                wraplength=300
            ).pack(side=tk.LEFT, padx=5)

            entry = tk.Entry(frame, width=12)
            entry.insert(0, datetime.now().strftime("%d-%m-%Y"))
            entry.pack(side=tk.RIGHT, padx=5)
            self.entries[file_path] = entry

        # Buttons frame
        btn_frame = tk.Frame(self, bg="#2c3e50")
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="OK",
            command=self.on_ok,
            bg="#27ae60",
            fg="white",
            width=10
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            btn_frame,
            text="Cancel",
            command=self.on_cancel,
            bg="#e74c3c",
            fg="white",
            width=10
        ).pack(side=tk.LEFT, padx=5)

    def validate_dates(self):
        for file_path, entry in self.entries.items():
            date = entry.get()
            try:
                datetime.strptime(date, "%d-%m-%Y")
                self.dates[file_path] = date
            except ValueError:
                messagebox.showerror("Invalid Date", f"Invalid date format for file {os.path.basename(file_path)}\nPlease use DD-MM-YYYY format")
                return False
        return True

    def on_ok(self):
        if self.validate_dates():
            self.result = self.dates
            self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

class AttendanceApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fast Attendance Message Sender")
        self.geometry("600x500")
        self.configure(bg="#2c3e50")
        
        # Speed optimization: Create widgets once
        self.file_paths = []  # Store multiple file paths
        self.create_widgets()

    def create_widgets(self):
        # Title Label
        title_label = tk.Label(
            self,
            text="Fast Attendance Message Sender",
            font=("Helvetica", 16, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=20)

        # Drop Zone
        self.drop_label = tk.Label(
            self,
            text="Drop Excel Files Here\n(You can drop multiple files)",
            font=("Helvetica", 12),
            bg="#34495e",
            fg="white",
            width=40,
            height=4
        )
        self.drop_label.pack(pady=20)
        
        # Make the label a drop target
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        # File List
        self.file_listbox = tk.Listbox(
            self,
            width=50,
            height=5,
            bg="#34495e",
            fg="white",
            selectmode=tk.MULTIPLE
        )
        self.file_listbox.pack(pady=10)

        # Status Label
        self.status_label = tk.Label(
            self,
            text="Ready to process files...",
            font=("Helvetica", 10),
            bg="#2c3e50",
            fg="white",
            wraplength=500
        )
        self.status_label.pack(pady=20)

        # Buttons Frame
        btn_frame = tk.Frame(self, bg="#2c3e50")
        btn_frame.pack(pady=10)

        # Send Button
        self.send_button = tk.Button(
            btn_frame,
            text="Send Messages",
            command=self.on_send,
            font=("Helvetica", 12),
            bg="#27ae60",
            fg="white",
            width=20
        )
        self.send_button.pack(side=tk.TOP, pady=5)

        # Clear Button
        self.clear_button = tk.Button(
            btn_frame,
            text="Clear Files",
            command=self.clear_files,
            font=("Helvetica", 10),
            bg="#e74c3c",
            fg="white",
            width=15
        )
        self.clear_button.pack(side=tk.TOP, pady=5)

    def clear_files(self):
        self.file_paths = []
        self.file_listbox.delete(0, tk.END)
        self.status_label.config(text="File list cleared")

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file_path in files:
            if file_path.lower().endswith('.xlsx') or file_path.lower().endswith('.xls'):
                # Convert to normalized path
                normalized_path = os.path.normpath(file_path)
                if normalized_path not in self.file_paths:
                    self.file_paths.append(normalized_path)
                    self.file_listbox.insert(tk.END, os.path.basename(normalized_path))
            else:
                messagebox.showwarning("Invalid File", f"File {os.path.basename(file_path)} is not an Excel file")
        
        self.status_label.config(text=f"{len(self.file_paths)} files ready to process")

    def on_send(self):
        if not self.file_paths:
            messagebox.showwarning("No Files", "Please drop Excel files first")
            return

        # Show date input dialog
        dialog = DateInputDialog(self, self.file_paths)
        self.wait_window(dialog)
        
        if dialog.result is None:
            return  # User cancelled

        self.send_button.config(state=tk.DISABLED)
        threading.Thread(target=self.process_sending, args=(dialog.result,), daemon=True).start()

    def process_sending(self, file_dates):
        try:
            total_files = len(self.file_paths)
            for file_idx, file_path in enumerate(self.file_paths, 1):
                self.status_label.config(text=f"Processing file {file_idx}/{total_files}: {os.path.basename(file_path)}")
                date = file_dates[file_path]
                send_attendance_messages(file_path, self.status_label, date)
                
            self.status_label.config(text="All files processed successfully!")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
        finally:
            self.send_button.config(state=tk.NORMAL)

if __name__ == '__main__':
    setup_logging()
    app = AttendanceApp()
    app.mainloop()