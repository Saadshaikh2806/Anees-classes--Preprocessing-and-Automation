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

def setup_logging():
    """Initialize logging with proper format."""
    logging.basicConfig(
        filename='message_sender.log',
        level=logging.INFO,
        format='%(asctime)s - %(message)s'
    )

def send_whatsapp_message(phone_number, message):
    """Send a WhatsApp message using web.whatsapp.com."""
    try:
        logging.info(f"Sending message to {phone_number}")
        encoded_message = urllib.parse.quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
        webbrowser.open(url, new=2, autoraise=False)
        time.sleep(10)  # Wait for the WhatsApp web page to load
        pyautogui.press("enter")
        time.sleep(2)  # Wait for the message to send
        pyautogui.hotkey("ctrl", "w")  # Close the browser tab
        return True
    except Exception as e:
        logging.error(f"Failed to send message to {phone_number}: {str(e)}")
        return False

def read_message_data(file_path):
    """Read message data from an Excel file."""
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        logging.error(f"Error reading file {file_path}: {str(e)}")
        raise

def send_messages(file_path, status_label):
    """Process the Excel file and send messages."""
    try:
        df = read_message_data(file_path)
        messages_sent = 0
        total = len(df)

        # Hardcoded message
        hardcoded_message = "Hello! This is a hardcoded message sent via the automated system."

        for idx, row in df.iterrows():
            phone_number = str(row['Phone']).strip()
            status_label.config(text=f"Sending message {idx + 1}/{total} to {phone_number}")
            status_label.update()

            if send_whatsapp_message(phone_number, hardcoded_message):
                messages_sent += 1

        status_label.config(text=f"Complete! Messages sent: {messages_sent} out of {total}")
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")

class MessageSenderApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Simple WhatsApp Message Sender")
        self.geometry("500x400")
        self.configure(bg="#2c3e50")

        self.file_path = None
        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(
            self,
            text="Simple WhatsApp Message Sender",
            font=("Helvetica", 16, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=20)

        self.drop_label = tk.Label(
            self,
            text="Drop Excel File Here",
            font=("Helvetica", 12),
            bg="#34495e",
            fg="white",
            width=30,
            height=4
        )
        self.drop_label.pack(pady=20)

        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)

        self.status_label = tk.Label(
            self,
            text="Ready to process file...",
            font=("Helvetica", 10),
            bg="#2c3e50",
            fg="white",
            wraplength=400
        )
        self.status_label.pack(pady=20)

        send_button = tk.Button(
            self,
            text="Send Messages",
            command=self.on_send,
            font=("Helvetica", 12),
            bg="#27ae60",
            fg="white",
            width=20
        )
        send_button.pack(pady=10)

    def on_drop(self, event):
        file_path = self.tk.splitlist(event.data)[0]
        if file_path.lower().endswith(('.xlsx', '.xls')):
            self.file_path = os.path.normpath(file_path)
            self.status_label.config(text=f"File ready: {os.path.basename(self.file_path)}")
        else:
            messagebox.showwarning("Invalid File", "Please drop a valid Excel file.")

    def on_send(self):
        if not self.file_path:
            messagebox.showwarning("No File", "Please drop an Excel file first.")
            return

        threading.Thread(target=self.process_sending, daemon=True).start()

    def process_sending(self):
        self.status_label.config(text="Processing...")
        send_messages(self.file_path, self.status_label)

if __name__ == '__main__':
    setup_logging()
    app = MessageSenderApp()
    app.mainloop()
