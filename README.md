# Aaruush'26 Mail Agent

AI-powered sponsorship email automation for Aaruush'26, SRM IST.

## What it does
- Upload your sponsor Excel/CSV list
- AI (Claude) personalizes each email using company data
- Opens pre-filled Gmail draft windows — ready to review and send
- Auto-adds all Aaruush CCs

---

## Deploy to Vercel (5 minutes)

### Option A — GitHub + Vercel (recommended)

1. Push this folder to a new GitHub repo
2. Go to [vercel.com](https://vercel.com) → New Project
3. Import your GitHub repo
4. Vercel auto-detects Create React App — just click **Deploy**
5. Share the URL with your team ✅

### Option B — Vercel CLI

```bash
npm install -g vercel
cd aaruush-mail-agent
vercel
```

---

## Local development

```bash
npm install
npm start
```

Open [http://localhost:3000](http://localhost:3000)

---

## How to use

1. **Upload** your Excel/CSV file (needs headers in row 1)
2. **Map** which column is company name and which is email
3. **Edit** the template if needed (uses `{{column_name}}` placeholders)
4. **Preview** each personalized email
5. **Generate** — allow popups from mail.google.com, then all drafts open in Gmail

## Excel file format

Your sheet should have at minimum:
- A column with company/brand names (any header name)
- A column with email addresses

Extra columns (industry, city, contact person, etc.) will be used by AI for personalization.

---

Built for Aaruush'26 Sponsorship & Marketing Team, SRM IST
