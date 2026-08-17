# Sunset---Not updated---Client Handover & Account Transition Dashboard

A premium, single-page tool for capturing everything that lives in an account manager's head, and exporting it as a fully formatted, executive-grade Excel workbook.

Live: https://client-transition-pro.lovable.app

## The Problem
When a client account changes hands, the knowledge rarely does. Handovers end up scattered across chat threads, inboxes, half-finished SOP docs, and the outgoing manager's memory. The result:

Recurring deliverables get missed in the first 30 days
Nobody knows who owns which platform access
Open issues resurface with no context or history
The client feels the churn, and account risk spikes
The Solution
One structured workspace where the outgoing owner records the full account picture, and one click turns it into a polished, boardroom-ready .xlsx handover pack the incoming owner can actually use.

Structured capture: client info, stakeholders, recurring tasks, platforms, access owners, open issues, preferences, knowledge transfer
Editable 30-day transition plan: checkbox milestones with owner, status and detail notes
Styled Excel export: branded title bands, dark headers, zebra rows, colour-coded priority and status pills, frozen panes, auto-filters
Customisable theme colours for the exported workbook
No backend, no accounts, no API keys. Everything runs in the browser.

Generated Workbook (10 tabs)
Executive Summary, Client Overview, Stakeholders, Recurring Tasks, Platforms (ticked only), Access & Ownership, Open Issues, Client Preferences, Knowledge Transfer Checklist, 30-Day Transition Plan.

Every sheet gets a merged branded title band, subtitle row, dark column headers, zebra rows, borders, tuned column widths, hidden gridlines, frozen header row and auto-filter.

Tech Stack
TanStack Start v1 (React 19), Vite 7, TypeScript, Tailwind CSS v4 with OKLCH theme tokens, lucide-react icons, xlsx-js-style for styled workbook generation, file-saver for download.

Design system: professional slate-white and deep navy palette defined as semantic tokens, so UI accents and Excel branding read from the same source.

Principles
Client-side only: no server, no database, no keys, no per-use cost
Blank by default: every field starts empty with plain-language placeholders
Polish over features: the deliverable is a document an executive will open
Privacy by default: nothing is transmitted; data stays in the browser session
Notes
Anything typed into the form is written into the exported file as-is, so treat the workbook as confidential and share it through your normal secure channels. Best practice is to record access owners and request routes rather than secrets themselves.
Nothing is saved between page reloads; export before closing the tab.

