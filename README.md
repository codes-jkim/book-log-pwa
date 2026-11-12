# 📚 Book Log PWA

A zero-cost, lightweight Progressive Web App (PWA) built for iOS Safari to archive finished books, memorable quotes, and personal reflections directly into Google Sheets. Co-designed and developed with Gemini through iterative AI-assisted prototyping.

By pasting a product URL from a specific online bookstore (such as Aladin, which embeds clean Schema.org JSON-LD metadata), book details—title, author, and cover image—are automatically parsed and formatted in your spreadsheet via Google Apps Script without requiring external book APIs.

---

## ✨ Key Features

- **📱 iOS Standalone PWA:** Installable via Safari's **"Add to Home Screen"** for a native full-screen app experience—no Apple Developer account or build step required.
- **⚡ Zero-API Metadata Scraping:** Parses product links from a certain bookstore (Aladin) that exposes structured JSON-LD (`application/ld+json`) and Open Graph tags, eliminating the need for third-party book API keys or rate-limit management.
- **📅 Custom Finished Date:** Defaults to today's date with a native mobile date picker for logging previously read books.
- **📊 Auto-Formatted Google Sheets DB:** Automatically renders cover thumbnails via `=IMAGE()`, adjusts row height and column width for clean visuals, and auto-increments row numbers.
- **🔒 Zero-Secret Repository:** No API keys, sheet IDs, or endpoints are stored in the codebase. The GAS Web App URL and `SECRET_TOKEN` live exclusively in the device's `localStorage`.

---

## 🛠️ Tech Stack

- **Frontend:** Vanilla HTML5, JavaScript (ES6+), Tailwind CSS (CDN)
- **Hosting:** Cloudflare Pages
- **Backend:** Google Apps Script (`UrlFetchApp`, `ContentService`, `PropertiesService`)
- **Database:** Google Sheets
- **AI Pair Programmer:** Google Gemini (Architecture design, GAS scraping logic & iOS Safari PWA optimization)

---

## 🚀 Setup Guide

### 1. Google Sheets & Apps Script
1. Create a new Google Sheet and set up Row 1 headers (`No`, `Date`, `Cover`, `Title`, `Author`, `Quote`, `Review`).
2. In cell **A2**, add the auto-numbering formula:
   ```excel
   =ARRAYFORMULA(IF(LEN(B2:B), TEXT(ROW(B2:B)-1, "#"), ""))
   ```
3. Open **Extensions > Apps Script** and paste the backend script (`Code.gs`).
4. Go to **Project Settings (⚙️) > Script Properties** and add:
   - **Property:** `SECRET_TOKEN`
   - **Value:** Your custom secret passphrase
5. Click **Deploy > New deployment > Web app**:
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`
6. Copy the generated Web App URL (`https://script.google.com/macros/s/.../exec`).

### 2. Frontend Deployment (Cloudflare Pages)
1. Fork or push `index.html` to a GitHub repository.
2. In Cloudflare Dashboard, go to **Workers & Pages > Create > Pages > Connect to Git**.
3. Select your repository with **Framework preset** set to `None` (leave build commands empty) and deploy.

### 3. iOS Device Setup
1. Open your deployed `*.pages.dev` URL in **Safari** on iPhone.
2. Tap **⚙️** in the top-right corner and enter your **GAS Web App URL** and **Secret Token**, then tap **Save Locally**.
3. Tap the Safari **Share** icon and select **Add to Home Screen**.

