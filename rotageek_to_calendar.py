import requests
import sys
import os
import pandas as pd
from ics import Calendar, Event
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# --- CONFIG ---
URL = "https://screwfix.rotageek.com/api/graphql-userschedules"


class ShiftSelector(tk.Toplevel):
    def __init__(self, parent, items):

        super().__init__(parent)
        self.title("Select Shifts to Export")
        self.geometry("450x650")
        self.result = []
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # Header
        tk.Label(self, text="Select items to include in your file:", font=("Arial", 10, "bold"), pady=10).pack()

        # Bulk Action Buttons
        bulk_frame = tk.Frame(self)
        bulk_frame.pack(fill="x", pady=5)
        tk.Button(bulk_frame, text="Select All", command=self.select_all).pack(side="left", padx=20)
        tk.Button(bulk_frame, text="Deselect All", command=self.deselect_all).pack(side="left")

        # Scrollable Area
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

        # Export Buttons
        export_label = tk.Label(self, text="Save as:", pady=5)
        export_label.pack()

        btn_frame = tk.Frame(self, pady=10)
        btn_frame.pack()
        tk.Button(btn_frame, text="Universal (.ics)", command=lambda: self.confirm("ics"), bg="#007bff", fg="white",
                  width=15).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Google CSV (.csv)", command=lambda: self.confirm("csv"), bg="#34a853", fg="white",
                  width=15).pack(side="left", padx=5)
        help_label = tk.Label(
            self,
            text="💡 Pro-Tip: Uncheck any shifts you've already added manually.",
            fg="gray",
            font=("Arial", 8, "italic")
        )
        help_label.pack(pady=5)

    def select_all(self):
        for var, _ in self.vars: var.set(True)

    def deselect_all(self):
        for var, _ in self.vars: var.set(False)

    def on_close(self):
        self.destroy()
        sys.exit()

    def confirm(self, fmt):
        self.export_format = fmt
        self.result = [item for var, item in self.vars if var.get()]
        self.destroy()


def get_credentials():
    root = tk.Tk();
    root.withdraw()
    cookie_help = (
    "Where to find this:\n"
        "1. Login to Rotageek (Chrome/Edge).\n"
    "2. Press F12 -> Network tab -> Refresh\n"
    "3. Click 'graphql-userschedules'\n"
    "4. Copy the 'cookie' from Request Headers"
    )
    c = simpledialog.askstring("Step 1: Rotageek Cookie", cookie_help)
    if not c: sys.exit()
    token_help = (
        "Where to find this:\n"
        "If you haven't opened it already for the cookie then:\n"
        "1. Login to Rotageek (Chrome/Edge).\n"
        "2. Press F12 -> Network tab -> Refresh\n"
        "3. Click 'graphql-userschedules'\n"
        "4. Copy the 'requestverificationtoken' from Request Headers"
    )
    t = simpledialog.askstring("Step 2: Rotageek Token", token_help)
    if not t: sys.exit()
    return c, t


def run_export():
    cookie, token = get_credentials()

    # 1. Fetch 4 weeks of data
    now = datetime.now()
    start_q = (now - timedelta(days=7)).strftime('%Y-%m-%dT00:00:00')
    end_q = (now + timedelta(days=28)).strftime('%Y-%m-%dT23:59:59')

    headers = {"content-type": "application/json", "cookie": cookie, "requestverificationtoken": token}
    payload = {
        "query": "query { dateRangeSchedule(start: \"" + start_q + "\", end: \"" + end_q + "\") { journals { start end isLeave } absences { start end leaveType { name } } } }"}

    try:
        r = requests.post(URL, headers=headers, json=payload)
        data = r.json().get('data', {}).get('dateRangeSchedule', {})
        raw_items = sorted(data.get('journals', []) + data.get('absences', []), key=lambda x: x['start'])
    except:
        messagebox.showerror("Error", "Could not fetch data from Rotageek.")
        sys.exit()

    # 2. Show Selector
    root = tk.Tk();
    root.withdraw()
    selector = ShiftSelector(root, raw_items)
    root.wait_window(selector)

    if not selector.result:
        return

    # 3. Handle Export
    if selector.export_format == "ics":
        save_path = filedialog.asksaveasfilename(defaultextension=".ics", filetypes=[("iCalendar", "*.ics")])
        if save_path:
            cal = Calendar()
            for item in selector.result:
                s_dt = datetime.fromisoformat(item['start'].replace('Z', ''))
                e_dt = datetime.fromisoformat(item['end'].replace('Z', ''))
                is_leave = item.get('isLeave') or 'leaveType' in item

                ev = Event()
                ev.name = "Screwfix: Leave" if is_leave else "Work: Screwfix"
                ev.begin = s_dt
                ev.end = e_dt
                if is_leave: ev.make_all_day()
                cal.events.add(ev)

            with open(save_path, 'w') as f:
                f.writelines(cal.serialize_iter())
            messagebox.showinfo("Success", "ICS file saved!")

    else:  # CSV Format
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if save_path:
            csv_data = []
            for item in selector.result:
                s_dt = datetime.fromisoformat(item['start'].replace('Z', ''))
                e_dt = datetime.fromisoformat(item['end'].replace('Z', ''))
                is_leave = item.get('isLeave') or 'leaveType' in item

                csv_data.append({
                    'Subject': 'Screwfix: Leave' if is_leave else 'Work: Screwfix',
                    'Start Date': s_dt.strftime('%m/%d/%Y'),
                    'Start Time': s_dt.strftime('%I:%M %p'),
                    'End Date': e_dt.strftime('%m/%d/%Y'),
                    'End Time': e_dt.strftime('%I:%M %p'),
                    'All Day Event': 'True' if is_leave else 'False'
                })
            pd.DataFrame(csv_data).to_csv(save_path, index=False)
            messagebox.showinfo("Success", "Google CSV saved!")


if __name__ == "__main__":
    run_export()