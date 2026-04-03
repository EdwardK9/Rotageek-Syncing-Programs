# 🕒 Rotageek & Excel Sync Tools



A collection of Python utilities to extract work schedules from **Rotageek** or existing **Excel** timesheets and sync them to any **Calendar App** (Google, Apple, Outlook) or your official spreadsheet.



---



## ⭐ Features



* **Multi-Source Parsing:** Extract shifts directly from the Rotageek API or from your saved Screwfix Excel workbooks.

* **Smart Selection GUI:** Review and verify shifts in a popup window before they are saved.

* **Three Export Options:**

    * **Direct Google Sync:** Pushes shifts directly to your Google account.

    * **Universal (.ics):** Works on iPhone, Android, and PC (Double-click to add).

    * **Google CSV:** For manual bulk uploads via the Google Calendar website.

* **Automatic Excel Sync:** Updates monthly sheets, handles unmerging, and colors holidays **Pink**.

* **Safety Protocol:** Closing the selector window kills the process to prevent accidental overwrites.



---



## 🚀 Quick Start Guide



### Step 1: Getting your "Browser Keys"

Because Rotageek is secure, the program needs to "borrow" your login session to see your shifts. **You only need to do this once per session.**



1.  Log in to [Screwfix Rotageek](https://screwfix.rotageek.com) on your computer.

2.  Press **F12** on your keyboard (opens the side panel).

3.  Click the **Network** tab at the top of that panel.

4.  **Refresh the page (F5)**.

5.  In the list that appears, find and click the row named `graphql-userschedules`.

6.  On the right-side panel, scroll down to the **Request Headers** section:

    * **Cookie:** Copy the massive block of text next to `cookie:`.

    * **Token:** Copy the text next to `requestverificationtoken:`.



### Step 2: Run the Tool

1.  Download the latest `.exe` from the [Releases](https://github.com/YOUR_USERNAME/rotageek-sync/releases) page.

2.  Run the program and paste your **Cookie** and **Token** when prompted.

3.  A window will appear with all your shifts. **Uncheck** any you don't want to sync.

4.  Select your **Export Method**:

    * **Universal (.ics):** Recommended! Saves a file. Just double-click it on your PC or Phone to "Add All" events.

    * **Google Direct:** Sends shifts straight to your Google account (Requires a one-time API setup).



---

## 📅 Adding to your Calendar



### **Universal (.ics) Method:**

Simply open/double-click the `.ics` file. Your device (iPhone/Mac/PC) will ask if you want to "Add All" events to your calendar.



### **Google Calendar (Web) Method:**

1.  Open [Google Calendar](https://calendar.google.com).

2.  Click **Settings (⚙️)** > **Import & Export**.

3.  Upload your file (.ics or .csv) and click **Import**.

---

## 📁 Which Tool Should I Use?

| File | What it does |
| :--- | :--- |
| **Rotageek_to_Calendar.exe** | Takes your live Rotageek schedule and puts it in your phone/PC calendar. |
| **Excel_to_Calendar.exe** | Scans an existing Screwfix Excel file and moves those shifts to your calendar. |
| **Sync_to_Excel.exe** | Grabs your Rotageek schedule and writes it directly into your official Excel workbook. |

---



## 🛠️ Installation (For Developers)

If you want to run the `.py` files manually:

1. `pip install requests pandas openpyxl ics`

2. `pip install tatsu==5.7.3`



---



# ⚠️ DISCLAIMER

This is an independent tool developed for personal convenience and is not affiliated with Rotageek or Screwfix. Use responsibly and never share your cookie, token or API credentials.

## DISCLAIMER

This project is an independent tool I developed for fun and is not affiliated with Rotageek or Screwfix. Use responsibly and keep your API credentials private.





