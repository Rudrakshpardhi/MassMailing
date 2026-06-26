# Aaruush'26 Mail Agent

A client-side sponsorship outreach tool for the Aaruush'26 Sponsorship & Marketing team (SRM IST, Chennai). Upload a sponsor list, fill in a template, and push ready-to-send Gmail drafts — with matching Invite and Proposal `.docx` files generated per company.

No backend, no database. Everything (spreadsheet parsing, template merging, `.docx` generation) runs in the browser; nothing is uploaded anywhere except the Gmail draft itself.

## What it does

- Upload a sponsor Excel/CSV list, or add companies one by one
- Map which column is the company name and which is the email address
- Edit a single email template with `{{column_name}}` placeholders, merged per row
- Preview every email exactly as it will look in Gmail, signature included
- Generate Invite and/or Proposal `.docx` files per company, with the company name swapped into the official template
- Push pre-filled Gmail compose drafts (To, CC, subject, body) for every selected company — opened as real Gmail tabs, never sent automatically
- Auto-CCs the sponsorship team on every draft
- Download a CSV backup of the drafts for the currently selected companies

## Logging in & roles

There's no real backend auth — this is a shared internal tool, and "login" is just a name + password check that runs in the browser. Two roles:

- **Committee Head** — if the uploaded sheet has a `CH` (or "Committee Head") column, the list is automatically filtered down to only the companies assigned to that person (matched loosely by first name).
- **Organizer** — sees every company in the sheet, unfiltered, with the `CH` column shown so they can see who's assigned to what.

Passwords follow the pattern `Aaruush@` + the last two digits of the person's phone number. Since this is all client-side, anyone with browser devtools access can read the full committee list (names, phone numbers, passwords) from the page — treat this as a lightweight convenience gate, not real security, and don't put anything more sensitive than what's already here into it.

The person signed in as **Rudraksh Pardhi** also sees an **Admin panel** (collapsed by default, top of the page) for adding/removing committee heads. It generates an updated code block — copy it, paste it over the `COMMITTEE_HEADS` array in `src/App.js` on GitHub, and commit. The change goes live for everyone after Vercel redeploys (a minute or so).

## How to use

1. **Add Companies** — upload an `.xlsx` / `.xls` / `.csv` (headers in row 1), or switch to "Add Manually" to type companies in one at a time.
2. **Map Columns** (upload only) — pick which column is the company name and which is the email. This also drives the Invite/Proposal filenames, so don't leave it unset.
3. **Email Template** — edit the subject and body. Use `{{column_name}}` for any column from your sheet (e.g. `{{COMPANY}}`, `{{CITY}}`).
4. **Preview** — step through each company to see the merged email exactly as it'll appear in Gmail, and optionally download that company's Invite/Proposal doc on its own.
5. **Generate** — pick which companies to send to, tick Invite/Proposal if you want the docs downloaded alongside each draft, then hit Generate. Allow popups from `mail.google.com` when prompted. Each draft opens in its own Gmail tab, about a second apart; nothing is sent until you personally hit Send in Gmail.

## Excel/CSV file format

At minimum, your sheet needs:
- A column with company/brand names (any header name — auto-detected, but double-check it in Step 2)
- A column with email addresses

Optional:
- A `CH` column (or one with "Committee Head" in the header) to auto-assign rows to specific Committee Heads
- Any other columns (city, contact person, notes, etc.) — usable as `{{placeholders}}` in the template

## Invite & Proposal docs

`public/templates/official-invite.docx` and `official-proposal.docx` are the official Word templates. When docs are generated, the app fetches these as raw bytes, swaps the placeholder company name baked into the template XML for the real company name (via `src/docxTemplates.js`, using PizZip — no server round-trip), and downloads:

- `AARUUSH X {Company} Invite.docx`
- `AARUUSH X {Company} Proposal A'26.docx`

Since Gmail's compose URL can't attach files, these land in your Downloads folder — drag them into the Gmail draft manually before sending.

If a generated doc renders with washed-out/illegible body text in Microsoft Word, it's a known Word bug tied to the dark/black Office Theme, not the file itself — switch Word's Office Theme (File → Options → General → Personalize your copy of Office) away from Black, or open the file in Google Docs instead.

## Local development

```bash
npm install
npm start
```

Open [http://localhost:3000](http://localhost:3000)

```bash
npm run build   # production build, output in build/
```

## Deploy

Auto-deploys on every push to `main` via Vercel (`vercel.json` is already configured for Create React App). To set up your own:

1. Push this repo to GitHub
2. [vercel.com](https://vercel.com) → New Project → import the repo → Deploy (zero config needed)

## Tech stack

Create React App (`react-scripts` 5), `xlsx` for spreadsheet parsing, `pizzip` for `.docx` XML manipulation. No backend, no API keys, no environment variables required.

---

Built for Aaruush'26 Sponsorship & Marketing Team, SRM IST
