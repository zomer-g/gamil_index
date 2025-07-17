# Gmail Message and Attachment Exporter 📩

This Google Apps Script project allows you to export Gmail threads to Google Sheets and Google Drive.  
Each **email** and **attachment** is saved as a separate file and indexed in a structured table with full metadata.

---

## ✨ Features

- 🔍 Search Gmail threads using any Gmail query (e.g. labels, senders, dates)
- 📄 Export each email as a PDF document to your Google Drive
- 📎 Export each attachment as a separate Drive file
- 🗂️ Automatically generate a Google Sheet index with:
  - Thread ID
  - Message ID
  - File type (message / attachment)
  - Date, From, To, Subject
  - Snippet of content
  - Drive File Name + URL
  - Message ID reference for attachments
- 📤 Sidebar UI with input form and launch button from within the Sheet

---

## 📋 Output Format

Each row in the sheet represents a file (message or attachment):

| Thread ID | Message ID | File Type | Date | From | To | Subject | Snippet | File Name | Drive URL | Attachment Message ID |
|-----------|------------|-----------|------|------|----|---------|---------|------------|------------|------------------------|

---

## 🚀 Getting Started

1. Open a new or existing [Google Sheets](https://sheets.new)
2. Go to **Extensions > Apps Script**
3. Paste the code from `Code.gs`
4. Add a new HTML file named `Sidebar.html` and paste the sidebar content
5. Save all and run `onOpen()` manually once
6. Refresh the sheet – a new menu will appear:  
   **📩 Gmail Exporter > 📤 Open Export Sidebar**

---

## 🧠 Usage

1. Click **📤 Open Export Sidebar**
2. Enter your Gmail search query  
   (e.g. `label:customers after:2023/01/01`)
3. Paste the URL of a Google Drive folder you have edit access to
4. Click **Start Export**
5. A new sheet named `Gmail Files Index` will be created with a full export log

---

## 🛡️ Permissions Required

- Access to Gmail threads and messages
- Access to create/edit Google Drive files
- Access to write to the current spreadsheet

---

## 📦 Sample Gmail Query Examples

```text
label:inbox has:attachment
from:support@example.com
label:legal after:2024/01/01 before:2024/06/30
