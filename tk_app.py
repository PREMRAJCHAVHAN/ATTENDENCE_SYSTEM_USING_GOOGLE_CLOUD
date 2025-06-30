import tkinter as tk
from tkinter import messagebox
from google_sheet_manager import GoogleSheetsManager
from datetime import datetime

# Define the headers
HEADERS = ["Timestamp", "Name", "Email", "Phone", "College", "Branch", "Year"]
# Fixed Google Sheet ID
SPREADSHEET_ID = "1sHn3WcgSOBNkQ2VPbzAfrSORk_6IK6pykyqibRnviEo"

class DataEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Google Sheets Data Entry")
        self.entries = {}

        # Create form labels and entries
        fields = HEADERS[1:]  # Skip Timestamp
        for idx, field in enumerate(fields):
            label = tk.Label(root, text=field + ":")
            label.grid(row=idx, column=0, padx=10, pady=5, sticky='e')
            entry = tk.Entry(root, width=30)
            entry.grid(row=idx, column=1, padx=10, pady=5)
            self.entries[field] = entry

        # Submit button
        submit_btn = tk.Button(root, text="Submit", command=self.submit)
        submit_btn.grid(row=len(fields), column=0, columnspan=2, pady=10)

        # Google Sheets setup (always use the fixed ID)
        self.sheets = GoogleSheetsManager(spreadsheet_id=SPREADSHEET_ID)
        # Add headers if sheet is empty
        data = self.sheets.read_data()
        if not data:
            self.sheets.add_data(HEADERS)

    def submit(self):
        # Collect data
        values = [entry.get().strip() for entry in self.entries.values()]
        if not all(values):
            messagebox.showerror("Error", "Please fill all fields.")
            return
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row = [timestamp] + values
        try:
            self.sheets.add_data(row)
            messagebox.showinfo("Success", "Data added to Google Sheet!")
            for entry in self.entries.values():
                entry.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add data: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataEntryApp(root)
    root.mainloop()