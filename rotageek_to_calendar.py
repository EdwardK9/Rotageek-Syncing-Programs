import requests
import sys
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# --- CONFIG ---
URL = "https://screwfix.rotageek.com/api/graphql-userschedules"


class ShiftSelector(tk.Toplevel):
    def __init__(self, parent, items):
        super().__init__(parent)
        self.title("Select Shifts for Google Calendar")
        self.geometry("450x650")
        self.result = []
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        header_frame = tk.Frame(self, pady=10)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="Select items to add to Calendar:", font=("Arial", 10, "bold")).pack()

        bulk_frame = tk.Frame(self)
        bulk_frame.pack(fill="x", pady=5)
        tk.Button(bulk_frame, text="Select All", command=self.select_all).pack(side="left", padx=20)
        tk.Button(bulk_frame, text="Deselect All", command=self.deselect_all).pack(side="left")

        container = ttk.Frame(self)
        container.pack(expand=True, fill="both", padx=5, pady=5)
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.vars = []
        for item in items:
            s_dt = datetime.fromisoformat(item['start'].replace('Z', ''))
            e_dt = datetime.fromisoformat(item['end'].replace('Z', ''))
            is_leave = item.get('isLeave') or 'leaveType' in item

            type_label = "🌸 LEAVE" if is_leave else f"🕒 {s_dt.strftime('%H:%M')}-{e_dt.strftime('%H:%M')}"
            display_text = f"{s_dt.strftime('%a %d %b')}: {type_label}"

            var = tk.BooleanVar(self, value=True)
            cb = tk.Checkbutton(self.scrollable_frame, text=display_text, variable=var, anchor="w",
                                font=("Consolas", 10))
            cb.pack(fill="x", padx=10, pady=2)
            self.vars.append((var, item))

        canvas.pack(side="left", expand=True, fill="both")
        scrollbar.pack(side="right", fill="y")

        btn_frame = tk.Frame(self, pady=15)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Create Calendar File", command=self.confirm, bg="#4285F4", fg="white", width=20,
                  height=2).pack(side="right", padx=20)
        tk.Button(btn_frame, text="Cancel", command=self.on_close).pack(side="right")

    def select_all(self):
        for var, _ in self.vars: var.set(True)

    def deselect_all(self):
        for var, _ in self.vars: var.set(False)

    def on_close(self):
        self.destroy()
        sys.exit()

    def confirm(self):
        self.result = [item for var, item in self.vars if var.get()]
        self.destroy()


def get_credentials():
    root = tk.Tk();
    root.withdraw()
    instructions = (
        "HOW TO GET CREDENTIALS:\n"
        "1. Login to Rotageek (Chrome/Edge).\n"
        "2. Press F12 -> Network tab.\n"
        "3. Refresh the page.\n"
        "4. Click on 'graphql-userschedules' in the list.\n"
        "5. In 'Request Headers', find 'cookie' and 'requestverificationtoken'."
    )
    cookie = simpledialog.askstring("Cookie", f"{instructions}\n\nPaste 'cookie':")
    if not cookie: sys.exit()
    token = simpledialog.askstring("Token", "Paste 'requestverificationtoken':")
    if not token: sys.exit()
    return cookie, token


def create_calendar():
    cookie, token = get_credentials()

    # Fetch Data
    now = datetime.now()
    start_q, end_q = (now - timedelta(days=7)).strftime('%Y-%m-%dT00:00:00'), (now + timedelta(days=28)).strftime(
        '%Y-%m-%dT23:59:59')
    headers = {"content-type": "application/json", "cookie": cookie, "requestverificationtoken": token}
    payload = {
        "query": "query { dateRangeSchedule(start: \"" + start_q + "\", end: \"" + end_q + "\") { journals { start end isLeave } absences { start end leaveType { name } } } }"}

    try:
        r = requests.post(URL, headers=headers, json=payload)
        data = r.json().get('data', {}).get('dateRangeSchedule', {})
        raw_items = data.get('journals', []) + data.get('absences', [])
        raw_items.sort(key=lambda x: x['start'])
    except:
        messagebox.showerror("Error", "Check Connection/Credentials")
        sys.exit()

    root = tk.Tk();
    root.withdraw()
    selector = ShiftSelector(root, raw_items)
    root.wait_window(selector)
    selected_items = selector.result
    if not selected_items: sys.exit()

    # Format for Google Calendar CSV
    calendar_list = []
    for item in selected_items:
        s_dt = datetime.fromisoformat(item['start'].replace('Z', ''))
        e_dt = datetime.fromisoformat(item['end'].replace('Z', ''))
        is_leave = item.get('isLeave') or 'leaveType' in item

        calendar_list.append({
            'Subject': 'Screwfix: Annual Leave' if is_leave else 'Work: Screwfix',
            'Start Date': s_dt.strftime('%m/%d/%Y'),
            'Start Time': s_dt.strftime('%I:%M %p'),
            'End Date': e_dt.strftime('%m/%d/%Y'),
            'End Time': e_dt.strftime('%I:%M %p'),
            'All Day Event': 'True' if is_leave else 'False',
            'Description': 'Imported from Rotageek',
            'Location': 'Screwfix Store'
        })

    # Save CSV
    df = pd.DataFrame(calendar_list)
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                             title="Save Calendar Import File")
    if save_path:
        df.to_csv(save_path, index=False)
        messagebox.showinfo("Success", f"File saved!\nNow follow the steps to import to Google.")


if __name__ == "__main__":
    create_calendar()