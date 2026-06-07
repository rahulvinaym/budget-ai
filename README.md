# Budget.AI 💰

An AI-powered personal finance tracker built entirely on Google Sheets, 
Apps Script and Groq's free LLM API. No servers. No subscriptions. 
Your data stays in your own Google account.

## What it does

- Upload any HDFC/SBI/ICICI bank statement CSV or Excel
- AI categorizes every transaction automatically
- Rule engine handles ~70% instantly, batch AI handles the rest
- Detects duplicate transactions
- Flags spending spikes (50%+ above your monthly average)
- Populates Fund Flow month by month (April to March FY)
- Live web dashboard with dark theme, charts and 3-year horizon view
- Quality Control tab shows uncategorized transactions for review

## Tech Stack

- Google Sheets — single source of truth
- Google Apps Script — automation engine
- Groq API (free) — Llama 3.3 for AI categorization
- Vanilla HTML/CSS/JS — web dashboard

## Setup

### 1. Get a free Groq API key
Go to console.groq.com → Sign up → API Keys → Create Key

### 2. Create Google Sheet
Create a new Google Sheet with these tabs:
- `Raw Imports`
- `Categorized`
- `Quality Control`
- `Status`

### 3. Set up Apps Script
Extensions → Apps Script → paste `Code.gs` contents
Replace `paste-your-groq-key-here` with your actual key

### 4. Add Dashboard
In Apps Script, click + → HTML → name it `dashboard`
Paste `dashboard.html` contents

### 5. Deploy as Web App
Deploy → New Deployment → Web App
Execute as: Me | Who has access: Only myself
Copy the URL → bookmark it

### 6. Run
Upload bank statement to Raw Imports tab
Click 💰 Budget Tools → 🚀 Run Everything

## Cost

Groq API free tier: 14,400 requests/day
Typical monthly usage: ~5-10 API calls
Effective cost: ₹0/month

## Roadmap

- [ ] Gmail API integration for auto-fetching CC statements
- [ ] CC vs bank transaction cross-matching
- [ ] Monthly budget targets with red/green tracking
- [ ] Multi-bank support (SBI, ICICI, Axis formats)
- [ ] Year-on-year comparison in dashboard

## Built by

Rahul — Chief Manager, Credit & Product Strategy
Building in public while learning AI and prompt engineering

LinkedIn: [Your LinkedIn URL]
