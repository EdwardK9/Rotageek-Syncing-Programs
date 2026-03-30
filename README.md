# 🕒 Rotageek & Excel Sync Tools

A collection of Python utilities to extract work schedules from **Rotageek** or existing **Excel** timesheets and sync them to any **Calendar App** (Google, Apple, Outlook) or your official spreadsheet.

---

## 🚀 Features

* **Multi-Source Parsing:** Extract shifts directly from the Rotageek API or from your saved Screwfix Excel workbooks.
* **Smart Selection GUI:** Review and verify shifts in a popup window before they are saved.
* **Three Export Options:** * **Direct Google Sync:** Pushes shifts directly to your Google account.
    * **Universal (.ics):** Works on iPhone, Android, and PC (Double-click to add).
    * **Google CSV:** For manual bulk uploads via the Google Calendar website.
* **Automatic Excel Sync:** Updates monthly sheets, handles unmerging, and colors holidays **Pink**.
* **Safety Protocol:** Closing the selector window kills the process to prevent accidental overwrites.

---

## 🔑 Setup Credentials

### 1. Rotageek (For API access)
1.  Log in to [Screwfix Rotageek](https://screwfix.rotageek.com) (Chrome/Edge).
2.  Press **F12** -> **Network** tab and refresh the page.
3.  Find the request named `graphql-userschedules`.
4.  In **Request Headers**, copy the `cookie` and the `requestverificationtoken`.

### 2. Google Calendar (For Direct Sync)
1.  Go to [Google Cloud Console](https://console.cloud.google.com/).
2.  Enable the **Google Calendar API** and create **OAuth 2.0 Client IDs** (Desktop App).
3.  Copy the **Client ID** and **Client Secret**. The script will ask for these via popup on the first run.

---

## 📁 Included Tools

### 1. Rotageek to Calendar (`rotageek_to_calendar.py`)
Fetches live data from Rotageek.
* **Action:** Paste credentials -> Confirm shifts -> Choose **Google**, **ICS**, or **CSV**.

### 2. Excel to Calendar (`excel_to_calendar.py`)
Parses your existing monthly Screwfix Excel file.
* **Action:** Select Excel file -> Confirm shifts -> Choose **Google**, **ICS**, or **CSV**.

### 3. Sync to Excel (`rotageek_to_excel.py`)
Fetches Rotageek data and writes it into your official monthly Excel workbook.
* **Action:** Paste credentials -> Select Excel file -> Confirm shifts.

---

## 📅 Adding to your Calendar

### **Universal (.ics) Method:**
Simply open/double-click the `.ics` file. Your device (iPhone/Mac/PC) will ask if you want to "Add All" events to your calendar.

### **Google Calendar (Web) Method:**
1.  Open [Google Calendar](https://calendar.google.com).
2.  Click **Settings (⚙️)** > **Import & Export**.
3.  Upload your file (.ics or .csv) and click **Import**.

---

## 🛠️ Installation

```bash
# Clone the repository
git clone [https://github.com/EdwardK9/rotageek-sync.git](https://github.com/EdwardK9/rotageek-sync.git)

# Install required libraries
pip install requests pandas openpyxl ics

# 'ics' requires a specific version of Tatsu. Run these:
pip uninstall tatsu -y
pip install tatsu==5.7.3
```

## DISCLAIMER
This project is an independent tool I developed for fun and is not affiliated with Rotageek or Screwfix. Use responsibly and keep your API credentials private.
