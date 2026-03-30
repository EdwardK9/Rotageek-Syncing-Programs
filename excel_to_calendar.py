import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import timedelta

# --- CONFIG ---
CALENDAR_ID = '***ENTER CALENDAR ID HERE'
SCOPES = ['https://www.googleapis.com/auth/calendar']


def get_calendar_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)


def sync_excel_timesheet():
    root = tk.Tk();
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select Screwfix XLSX", filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    if not file_path: return

    service = get_calendar_service()

    # 1. Fetch existing "Edward Work" events to avoid duplicates
    existing_dates = set()
    page_token = None
    while True:
        events_result = service.events().list(calendarId=CALENDAR_ID, pageToken=page_token, q="Edward Work",
                                              singleEvents=True).execute()
        for e in events_result.get('items', []):
            start = e['start'].get('dateTime', e['start'].get('date'))
            existing_dates.add(start.split('T')[0])
        page_token = events_result.get('nextPageToken')
        if not page_token: break

    # 2. Extract all data from Excel
    xl = pd.ExcelFile(file_path)
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    sheets = [s for s in xl.sheet_names if any(m in s for m in months) and "Overall" not in s]

    all_shifts = []
    all_holidays = []

    for sheet in sheets:
        df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
        header_idx = next((i for i, r in df_raw.iterrows() if r.astype(str).str.contains('DATE', case=False).any()), -1)
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
                all_shifts.append({'date': date_str, 'start': start_val, 'end': finish_val})
            elif any(kw in start_val.upper() for kw in ['ANNUAL', 'HOLIDAY', 'CLOSED']):
                all_holidays.append(dt)

    # 3. Process Holidays (Merging Logic)
    all_holidays.sort()
    if all_holidays:
        print("Merging consecutive holidays...")
        holiday_blocks = []
        if all_holidays:
            start_date = all_holidays[0]
            end_date = all_holidays[0]

            for i in range(1, len(all_holidays)):
                if all_holidays[i] == end_date + timedelta(days=1):
                    end_date = all_holidays[i]
                else:
                    holiday_blocks.append((start_date, end_date))
                    start_date = all_holidays[i]
                    end_date = all_holidays[i]
            holiday_blocks.append((start_date, end_date))

        for start, end in holiday_blocks:
            s_str = start.strftime('%Y-%m-%d')
            # Google "All Day" end dates are exclusive, so add 1 day to the end
            e_str = (end + timedelta(days=1)).strftime('%Y-%m-%d')

            if s_str not in existing_dates:
                event = {
                    'summary': 'Edward Work - Holiday',
                    'start': {'date': s_str},
                    'end': {'date': e_str},
                }
                service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
                print(f"  + Added Holiday Block: {s_str} to {end.strftime('%Y-%m-%d')}")

    # 4. Process Shifts
    for shift in all_shifts:
        if shift['date'] not in existing_dates:
            s_time = shift['start'] if len(shift['start']) > 5 else f"{shift['start']}:00"
            f_time = shift['end'] if len(shift['end']) > 5 else f"{shift['end']}:00"
            event = {
                'summary': 'Edward Work',
                'location': 'Screwfix',
                'start': {'dateTime': f"{shift['date']}T{s_time}", 'timeZone': 'Europe/London'},
                'end': {'dateTime': f"{shift['date']}T{f_time}", 'timeZone': 'Europe/London'},
            }
            service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
            print(f"  + Added Shift: {shift['date']}")


if __name__ == '__main__':
    sync_excel_timesheet()