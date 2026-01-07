# Google Sheets to Calendar Sync with Targeted Reminders

This Google Apps Script automates the creation of Google Calendar events based on data entered into a Google Sheet. It is designed for "Achin Bindlish Couture Label" to manage client delivery schedules.

## 🚀 Features

1.  **Force Time:** Automatically sets all events to **11:00 AM - 12:00 PM** on the Delivery Date (regardless of when the script is run).
2.  **Dual Reminders:** Sets two specific popup reminders for every event:
    * **1 Week before** (7 Days)
    * **3 Days before**
3.  **Safe Write Mode:** Writes the "Event ID" and "Status" back to the Google Sheet immediately after processing each row. This prevents data loss if the script times out or hits an API limit on a large list.
4.  **Duplicate Prevention:** Checks for an existing `Event ID`. If found, it updates the existing calendar event instead of creating a duplicate.

## 📋 Sheet Structure Requirements

For the script to work, your Google Sheet tab must be named **"Calendar Events Data"** and contain the following exact column headers (row 1):

* `Client Name`
* `Outfit Type`
* `Date of Assigning`
* `Date of Delivery`
* `Job-Card No.`
* `Status`
* `Event ID`
* `Table No.` (Optional)

## 🛠️ Installation & Usage

1.  Open your Google Sheet.
2.  Go to **Extensions > Apps Script**.
3.  Delete any code in the editor and paste the content of `Code.js` from this repository.
4.  Click **Save**.
5.  Refresh your Google Sheet.
6.  You will see a new menu item called **"Reminder Manager"**.
7.  Click **Reminder Manager > Set Events (1 Week & 3 Day Reminders)**.

## ⚠️ Troubleshooting

* **Error: "Column missing":** Check your headers. They must match the spelling in the list above exactly (case-sensitive).
* **Script Timeout:** The script is optimized with "Safe Write" logic. If it stops, simply run it again; it will pick up where it left off (updating existing ones and creating new ones).# sheets-calendar-sync-automation
A robust Google Apps Script for managing delivery schedules. Syncs Sheets to Calendar, locks event times to 11:00 AM, and ensures data integrity with immediate write-back.
