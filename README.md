# Old version---Not updated---Client Handover & Account Transition Dashboard

A premium, single-page B2B tool for capturing everything that lives in an account manager's head — and exporting it as a fully formatted, executive-grade Excel workbook.

Link: https://client-transition-pro.lovable.app

---

## The Problem

When a client account changes hands, the knowledge rarely does. Handovers are scattered across Slack threads, inboxes, half-finished SOP docs, and the outgoing manager's memory. The result:

- Recurring deliverables get missed in the first 30 days
- Nobody knows who owns which platform login
- Open issues resurface with no context or history
- The client feels the churn — and account risk spikes

## The Solution

One structured workspace where the outgoing owner records the full account picture, and one click turns it into a polished, boardroom-ready `.xlsx` handover pack the incoming owner can actually use.

- **Upload existing docs** — SOPs, notes, emails, CSVs are parsed in the browser to auto-fill the form
- **Structured capture** — client info, stakeholders, recurring tasks, platforms, logins, issues, preferences, knowledge transfer
- **Editable 30-day transition plan** — checkbox milestones with owner, status and elaboration
- **Styled Excel export** — branded title bands, dark headers, zebra rows, colour-coded priority/status pills, frozen panes, auto-filters

No backend, no accounts, no API tokens. Everything runs client-side in the browser.

---

## Features

### 1. Document Upload (top banner)
Drag-and-drop or browse. Text-based files (`.txt`, `.md`, `.csv`, `.json`, `.log`, `.yml`) under 2 MB are read in-browser and pattern-matched to pre-populate:

- Client name, industry, region, services
- Platforms mentioned (Marketo, Salesforce, 6sense, ON24, Salesloft, HubSpot)
- Stakeholders inferred from email addresses plus nearby role keywords
- Open issues from lines starting with `issue:`, `risk:`, `bug:`, `blocker:`, `problem:`

A toast confirms how many fields were filled. Files appear as removable chips and ride along into the export.

### 2. Left Panel — Structured Form
- **Client Info** — name, industry, region, services, prepared by
- **Stakeholders** — name, role, email, notes
- **Recurring Tasks** — task, frequency, current owner, new owner, instructions
- **Platforms Checklist** — Marketo, Salesforce, 6sense, ON24, Salesloft + custom entries
- **Login Compilation** — platform, link, owner, username, password, notes
- **Open Issues** — issue, priority, status, details
- **Client Preferences** — communication style, reporting expectations, escalation path
- **Knowledge Transfer Notes** — tribal knowledge, watch-outs, historical context

### 3. Right Panel — 30-Day Transition Plan
An editable milestone list: tick when done, add a title, elaborate in the details field, assign an owner and set a status. Add or remove items freely. Flows straight into the workbook.

### 4. Export
A single **Generate & Download Excel** action assembles the whole workbook from live form state.

---

## Generated Workbook

| # | Sheet | Contents |
|---|-------|----------|
| 1 | Executive Summary | Client, prepared by, prepared date, headline counts, plan overview |
| 2 | Client Overview | Name, industry, region, services delivered |
| 3 | Stakeholders | Name, role, email, notes |
| 4 | Recurring Tasks | Task, frequency, current owner, new owner, instructions |
| 5 | Platforms | Ticked platforms only |
| 6 | Login Compilation | Platform, link, owner, username, password, notes |
| 7 | Open Issues | Issue, colour-coded priority, colour-coded status, details |
| 8 | Client Preferences | Communication, reporting, escalation |
| 9 | Knowledge Transfer Checklist | Category, notes, done column |
| 10 | Source Documents | Uploaded file names, size, type, content snippet *(only when files are attached)* |
| 11 | 30-Day Transition Plan | Milestone, details, owner, status *(final tab)* |

Formatting applied to every sheet: merged branded title band, subtitle row, dark column headers, alternating zebra rows, cell borders, tuned column widths, hidden gridlines, frozen header row and auto-filter.

---

## Tech Stack

| Layer | Choice |
|-------|--------|
| Framework | TanStack Start v1 (React 19, SSR-capable) |
| Build tool | Vite 7 |
| Language | TypeScript |
| Styling | Tailwind CSS v4 (`src/styles.css`, OKLCH theme tokens) |
| Icons | lucide-react |
| Excel engine | `xlsx-js-style` (styled workbook generation) |
| File download | `file-saver` |

### Design System
A professional slate-white and deep navy palette (`#1E3A5F`) defined as semantic tokens in `src/styles.css`. No hardcoded colour utilities in components — accents, focus rings, buttons and Excel title bands all read from the same tokens.

---

## Architecture

```text
src/
├── routes/
│   ├── __root.tsx     Root layout, head metadata, global shell
│   └── index.tsx      The entire dashboard: state, form, plan, export engine
├── styles.css         Tailwind v4 theme tokens (slate / navy / ink)
└── router.tsx         TanStack Router bootstrap
```

Single-route by design. The whole tool is one continuous workflow — splitting it across pages would break the "capture then export" flow and add navigation state for no benefit. All form data lives in React state; nothing is persisted or transmitted.

### Principles
- **Client-side only** — no server, no database, no API keys, no rate limits, no cost per use
- **Blank by default** — every field starts empty with a plain-language placeholder so the output is never polluted with sample data
- **Polish over features** — the deliverable is a document an executive will open; formatting is the product
- **Privacy by default** — uploaded documents are parsed in the browser and never leave the machine

---

## Running Locally

```bash
npm install
npm run dev      # http://localhost:8080
npm run build    # production build
```

## Notes & Limitations

- Login credentials are stored in browser memory and written into the exported file in plain text — treat the workbook as confidential and distribute accordingly.
- Auto-fill parsing covers text-based formats only. PDF and DOCX files are listed in the Source Documents tab but not parsed.
- Nothing is saved between page reloads; export before closing the tab.
