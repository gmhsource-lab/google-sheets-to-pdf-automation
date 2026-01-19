# 🚀 One-Click Document Engine (Google Apps Script)

This system automates the generation, archiving, and emailing of professional PDF documents directly from Google Sheets.

## ✨ Key Features
- **Dynamic Header Mapping:** Zero-maintenance logic that automatically detects new spreadsheet columns.
- **PDF Archiving:** Automatically converts documents to PDF and saves them to a specific Drive folder with custom naming conventions.
- **Gmail Integration:** Sends personalized emails with the PDF attachment in one click.
- **User Interface:** Adds a custom menu to the Google Sheets toolbar.

## 🛠️ Setup
1. Copy the `Code.gs` into your Google Apps Script editor.
2. Update the `⚙️ Settings` tab with your Template ID and Folder ID.
3. Refresh the sheet and use the custom menu.

## 👨‍💻 About the Author
Legacy Microsoft Partner (2012) specializing in enterprise-grade automation and solutions architecture.

## 📖 User Instructions
1. **Prepare the Sheet:** Create two tabs: `📊 Dashboard` and `⚙️ Settings`.
2. **Configure Settings:**
   - Place your Google Doc Template ID in cell **B2**.
   - Place your Destination Folder ID in cell **B3**.
   - Enter your Business Name in cell **B4**.
3. **Template Setup:** In your Google Doc, use `{{Column Name}}` tags that match your Dashboard headers exactly.
4. **Execution:** Use the `🚀 DOCUMENT ENGINE` menu to generate and email your PDF.
