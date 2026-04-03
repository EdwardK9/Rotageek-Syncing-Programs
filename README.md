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

## ⚡ Quick Start Guide

### Step 1: Get your Rotageek "Keys" (Required once per session)
The program needs to "pretend" to be your browser to see your shifts.
1. Log in to [Screwfix Rotageek](https://screwfix.rotageek.com).
2. Press **F12** on your keyboard (opens Developer Tools).
3. Click the **Network** tab at the top of that side-panel.
4. Refresh the page (F5).
5. In the list, find the row named `graphql-userschedules`. 
6. Scroll down the right side to find **Request Headers**:
   * Copy the text after `cookie:`
   * Copy the text after `requestverificationtoken:`

### Step 2: Run the Program
1. Download the `.exe` from the [Releases](https://github.com/YOUR_USERNAME/rotageek-sync/releases) page.
2. Run it and paste your keys when prompted.
3. Select which shifts you want to keep.

### Step 3: Choose your Export
* **Google Calendar:** Sends them straight to your account (Requires a one-time API setup).
* **Universal (.ics):** Recommended! Saves a file. Double-click it on your PC or Phone to "Add All" events.

---

## 🛠️ One-Time Setup for Google Direct Sync
If you want the "One-Click to Google" feature, you need to tell Google it's okay for this app to talk to your account:
1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project named "WorkSync".
3. Enable **Google Calendar API**.
4. Create **OAuth Client ID** (Type: Desktop App).
5. Copy the **Client ID** and **Secret** into the app popups.

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
* **Action:** Create Excel doc, template included here. Create sheets with month names (e.g. Mar 2026). Run Script -> Paste credentials -> Select Excel file -> Confirm shifts.

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
Or download and run the `.exe` from the [Releases](https://github.com/YOUR_USERNAME/rotageek-sync/releases) page.
## DISCLAIMER
This project is an independent tool I developed for fun and is not affiliated with Rotageek or Screwfix. Use responsibly and keep your API credentials private.
