import os
import json
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from ics import Calendar, Event
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# --- CONFIG ---
CALENDAR_ID = 'primary'
SCOPES = ['https://www.googleapis.com/auth/calendar']


class ExportSelector(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Choose Export Method")
        self.geometry("400x300")
        self.choice = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        tk.Label(self, text="How would you like to export your shifts?", font=("Arial", 11, "bold"), pady=20).pack()

        tk.Button(self, text="Direct to Google Calendar", command=lambda: self.set_choice("google"),
                  bg="#4285F4", fg="white", width=25, pady=10).pack(pady=5)

        tk.Button(self, text="Save as Universal File (.ics)", command=lambda: self.set_choice("ics"),
                  bg="#34A853", fg="white", width=25, pady=10).pack(pady=5)

        tk.Button(self, text="Save as Google CSV (.csv)", command=lambda: self.set_choice("csv"),
                  bg="#FBBC05", fg="black", width=25, pady=10).pack(pady=5)

    def set_choice(self, choice):
        self.choice = choice
        self.destroy()

    def on_close(self):
        self.choice = None
        self.destroy()


def get_google_config_from_user():
    if os.path.exists('token.json'): return None

    msg = (
        "GOOGLE CALENDAR SETUP:\n"
        "1. Go to Google Cloud Console > Credentials.\n"
        "2. Copy your 'Client ID' and 'Client Secret'.\n"
    )
    cid = simpledialog.askstring("Step 1/2", f"{msg}\nPaste Client ID:")
    if not cid: return None
    sec = simpledialog.askstring("Step 2/2", "Paste Client Secret:")
    if not sec: return None

    return {
        "installed": {
            "client_id": cid.strip(), "client_secret": sec.strip(),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": ["http://localhost"]
        }
    }


def get_calendar_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            config = get_google_config_from_user()
            if not config: return None
            flow = InstalledAppFlow.from_client_config(config, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)


def main():
    root = tk.Tk();
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select Screwfix Excel", filetypes=[("Excel", "*.xlsx")])
    if not file_path: return

    # --- 1. PARSE EXCEL DATA (Your Original Logic) ---
    try:
        xl = pd.ExcelFile(file_path)
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        sheets = [s for s in xl.sheet_names if any(m in s for m in months) and "Overall" not in s]

        all_shifts = []
        all_holidays = []

        for sheet in sheets:
            df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
            header_idx = next((i for i, r in df_raw.iterrows() if r.astype(str).str.contains('DATE', case=False).any()),
                              -1)
            if header_idx == -1: continue
            df = pd.read_excel(file_path, sheet_name=sheet, skiprows=header_idx)
            df.columns = [str(c).strip() for c in df.columns]
            date_col = next((c for c in df.columns if 'DATE' in c.upper()), None)
            if not date_col: continue

            for _, row in df.dropna(subset=[date_col]).iterrows():
                try:
                    dt = pd.to_datetime(row[date_col])
                    date_str = dt.strftime('%Y-%m-%d')
                except:
                    continue
                start_val = str(row.get('Start time', 'nan')).strip()
                finish_val = str(row.get('Finish time', 'nan')).strip()
                if ":" in start_val:
                    all_shifts.append({'date': date_str, 'dt': dt, 'start': start_val, 'end': finish_val})
                elif any(kw in start_val.upper() for kw in ['ANNUAL', 'HOLIDAY', 'CLOSED']):
                    all_holidays.append({'date': date_str, 'dt': dt})
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel: {e}")
        return

    # --- 2. CHOOSE EXPORT METHOD ---
    selector = ExportSelector(root)
    root.wait_window(selector)
    if not selector.choice: return

    # --- 3. EXECUTE EXPORT ---
    if selector.choice == "google":
        service = get_calendar_service()
        if not service: return

        existing_dates = set()
        events_result = service.events().list(calendarId=CALENDAR_ID, q="Edward Work", singleEvents=True).execute()
        for e in events_result.get('items', []):
            start = e['start'].get('dateTime', e['start'].get('date'))
            existing_dates.add(start.split('T')[0])

        count = 0
        for h in all_holidays:
            if h['date'] not in existing_dates:
                event = {'summary': 'Edward Work - Holiday', 'start': {'date': h['date']},
                         'end': {'date': (h['dt'] + timedelta(days=1)).strftime('%Y-%m-%d')}}
                service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
                count += 1
        for s in all_shifts:
            if s['date'] not in existing_dates:
                st = s['start'] if len(s['start']) > 5 else f"{s['start']}:00"
                en = s['end'] if len(s['end']) > 5 else f"{s['end']}:00"
                event = {'summary': 'Edward Work', 'location': 'Screwfix',
                         'start': {'dateTime': f"{s['date']}T{st}", 'timeZone': 'Europe/London'},
                         'end': {'dateTime': f"{s['date']}T{en}", 'timeZone': 'Europe/London'}}
                service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
                count += 1
        messagebox.showinfo("Success", f"Directly added {count} entries to Google Calendar.")

    elif selector.choice == "ics":
        save_path = filedialog.asksaveasfilename(defaultextension=".ics", filetypes=[("iCalendar", "*.ics")])
        if not save_path: return
        cal = Calendar()
        for h in all_holidays:
            e = Event(name="Edward Work - Holiday", begin=h['date'])
            e.make_all_day();
            cal.events.add(e)
        for s in all_shifts:
            st = s['start'] if len(s['start']) > 5 else f"{s['start']}:00"
            en = s['end'] if len(s['end']) > 5 else f"{s['end']}:00"
            e = Event(name="Edward Work", begin=f"{s['date']} {st}", end=f"{s['date']} {en}");
            cal.events.add(e)
        with open(save_path, 'w') as f:
            f.writelines(cal.serialize_iter())
        messagebox.showinfo("Success", "ICS file saved!")

    elif selector.choice == "csv":
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not save_path: return
        rows = []
        for h in all_holidays:
            rows.append({'Subject': 'Edward Work - Holiday', 'Start Date': h['dt'].strftime('%m/%d/%Y'),
                         'All Day Event': 'True'})
        for s in all_shifts:
            rows.append({'Subject': 'Edward Work', 'Start Date': s['dt'].strftime('%m/%d/%Y'), 'Start Time': s['start'],
                         'End Date': s['dt'].strftime('%m/%d/%Y'), 'End Time': s['end'], 'All Day Event': 'False'})
        pd.DataFrame(rows).to_csv(save_path, index=False)
        messagebox.showinfo("Success", "Google CSV saved!")


if __name__ == '__main__':
    main()