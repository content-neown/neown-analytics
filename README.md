# neOwn Analytics Dashboard
Pure HTML · Vercel Serverless Functions · Chart.js · Claude AI

---

## Files
```
index.html    — main dashboard
style.css     — all styles
script.js     — frontend logic
api/
  sheets.js   — Google Sheets proxy (serverless)
  insights.js — Claude AI insights (serverless)
  chat.js     — Claude chatbot (serverless)
vercel.json   — Vercel config
```

---

## Deploy Steps

### 1. Upload to GitHub
- Go to github.com → New repository → name it `neown-analytics`
- Upload all these files (drag & drop the whole folder)

### 2. Connect to Vercel
- Go to vercel.com → Add New Project → Import your GitHub repo

### 3. Add Environment Variables (BEFORE deploying)
In Vercel, before clicking Deploy, scroll to **Environment Variables** and add:

| Name | Value |
|---|---|
| `GOOGLE_SHEETS_API_KEY` | your Google key |
| `ANTHROPIC_API_KEY` | your Anthropic key |
| `SHEET_ID` | `13q9R1M2RAxSQ5rsujRLDOtO-ip-eH6eH8Or-_9ZjiTQ` |

### 4. Get a Google Sheets API Key (free, 2 min)
1. Go to console.cloud.google.com
2. New Project → any name → Create
3. Search "Google Sheets API" → Enable
4. Credentials → Create Credentials → API Key → copy it

### 5. Make your Sheet public
Google Sheet → Share → Anyone with the link → Viewer

---

## Local testing
You need Vercel CLI for the API functions to work locally:
```
npm i -g vercel
vercel dev
```
Then open http://localhost:3000
