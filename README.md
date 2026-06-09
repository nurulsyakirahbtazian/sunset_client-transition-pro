# Handover OS — B2B Client Handover & Account Transition Dashboard

> A premium, single-page executive dashboard that turns chaotic client handovers into a structured, auditable, and instantly shareable transition package — exportable as a fully styled, multi-tab Excel workbook.

Live preview: https://id-preview--80faa4ac-1459-4200-b5a2-eb708a82acd5.lovable.app  
Production: https://client-transition-pro.lovable.app

---

## 1. The Problem

In B2B service organizations (agencies, consultancies, managed-service providers, customer success teams), **client handovers are the single biggest source of churn, dropped balls, and internal friction**.

When an account manager leaves, gets promoted, or rotates off an account, knowledge typically lives in:

- Scattered Slack DMs and email threads
- Personal Notion/Google Doc pages nobody else can find
- "Tribal knowledge" only the outgoing AM remembers
- Half-finished SOPs in five different folders
- Mental models of stakeholder politics that are never written down

The result:
- **Clients feel the drop** — repeated questions, missed context, broken trust
- **New AMs ramp slowly** — often 30–60 days of low productivity
- **Issues fall through the cracks** — open tickets, renewal dates, billing quirks
- **Leadership has zero visibility** — no standard artifact to review or audit
- **Compliance risk** — no documented trail of credentials, access, or commitments

Most teams "solve" this with a blank Google Doc template. It doesn't work.

## 2. The Solution

**Handover OS** is an opinionated, executive-grade dashboard that enforces a complete handover in a single sitting and produces a professional deliverable the incoming AM, the client, and leadership can all use.

Key ideas:

1. **One screen, one source of truth.** Stakeholders, recurring tasks, platforms & credentials, open issues, client preferences, and a 30-day KT checklist — all captured in a single structured form.
2. **Pre-filled with realistic mock data.** The app loads 100% complete so users see exactly what "good" looks like before they type a single character.
3. **AI-polished preview.** A live right-hand panel shows how raw notes transform into formal, client-ready language, plus a generated 30-day onboarding roadmap.
4. **ROI made visible.** Executive metric cards quantify time saved and risk mitigated, so the tool sells itself to leadership.
5. **One-click, beautifully styled Excel export.** Not a plain CSV — a branded, multi-tab `.xlsx` workbook with title bands, colored section headers, zebra rows, color-coded priority/status pills, frozen headers, and auto-filters.
6. **Upload existing artifacts.** Drag-and-drop zone for existing SOPs, notes, and emails so nothing is lost in translation.

The output is a single `.xlsx` file that can be emailed, attached to a CRM record, or stored in a shared drive as the canonical handover document.

## 3. What's Inside the Generated Workbook

The "Generate & Download Excel" button reads every current form value and assembles **8 fully styled tabs**:

| # | Sheet | Contents |
|---|-------|----------|
| 1 | Executive Summary | Client snapshot, ROI metrics, 30-day transition plan |
| 2 | Client Overview | Company profile, contract value, tier, renewal date |
| 3 | Stakeholders | Name, role, influence, communication preferences |
| 4 | Recurring Tasks | Cadence, owner, SLA, last-completed |
| 5 | Platforms | Tools, access level, credential location, MFA owner |
| 6 | Open Issues | Title, priority pill, status pill, owner, due date |
| 7 | Client Preferences | Tone, meeting cadence, escalation path, do's & don'ts |
| 8 | KT Checklist | 30-day onboarding tasks with check column |

Styling highlights: branded `#FF0F0F` title bands, dark ink headers, alternating zebra rows, color-coded priority (High/Med/Low) and status (Open/In Progress/Blocked/Done) pills, merged title rows, frozen header row, gridlines off, tuned column widths, and per-sheet auto-filters.

## 4. Tech Stack

| Layer | Choice | Why |
|-------|--------|-----|
| Framework | **TanStack Start v1** (React 19 + Vite 7) | Modern full-stack React with SSR-ready routing |
| Language | **TypeScript (strict)** | Type-safe forms and Excel cell specs |
| Styling | **Tailwind CSS v4** via `@import` in `src/styles.css` | Native CSS tokens, OKLCH color space |
| Icons | **lucide-react** | Consistent, lightweight icon set |
| State | **React `useState`** | Local form state — no backend needed for the artifact flow |
| Excel | **xlsx-js-style** + **file-saver** | Per-cell styling (fonts, fills, borders, merges, freeze, autofilter) |
| Hosting | **Lovable** (Cloudflare Workers runtime) | Edge-deployed, zero-config publish |
| Build | **Vite 7** with TanStack Start plugin | Fast HMR, automatic route tree generation |

Brand system: executive slate-white + black with a single `#FF0F0F` accent. Typography uses a slate-ink scale defined as OKLCH tokens in `src/styles.css`.

## 5. Architecture

```text
src/
├── routes/
│   ├── __root.tsx        # Root layout (html/head/body shell)
│   └── index.tsx         # The entire Handover OS dashboard
├── styles.css            # Tailwind v4 + design tokens (OKLCH)
└── router.tsx            # TanStack Start router bootstrap
```

The dashboard is intentionally a single route (`/`) because the entire value of the tool is the one-screen, one-sitting workflow.

Key in-file modules inside `src/routes/index.tsx`:

- **Mock data block** — realistic Northwind Industrial Group dataset for instant demo value
- **Form state hooks** — one `useState` per section (stakeholders, tasks, platforms, issues, preferences, checklist)
- **`buildSheet` helper** — turns an array-of-arrays + per-cell `CellSpec` objects into a styled XLSX worksheet (column widths, row heights, merges, freeze, autofilter)
- **`makeTableSheet` helper** — standard styled table with header row + zebra body rows
- **`generateExcel`** — orchestrates all 8 sheets and triggers `file-saver` download
- **AI Preview tabs** — show raw → polished transformations and a generated 30-day roadmap
- **ROI cards** — Time Saved (6.5 hrs) and Risk Mitigation (98%)

## 6. Getting Started Locally

Prerequisites: **Bun** (or Node 20+ with npm).

```bash
# install
bun install

# run dev server
bun run dev

# build for production
bun run build
```

The dev server runs on Vite's default port. Open the URL it prints.

## 7. How to Use

1. **Open the app** — it loads pre-filled with mock B2B data so you can see the finished shape immediately.
2. **(Optional) Drop existing artifacts** — use the top-banner drag-and-drop to attach existing SOPs, notes, or email exports.
3. **Edit the left panel** — overwrite the mock data with real client information across all sections.
4. **Watch the right panel** — the AI-polished preview, 30-day roadmap, and ROI metrics update live.
5. **Click "Generate & Download Excel"** — a fully styled multi-tab `.xlsx` is generated client-side and downloaded.
6. **Share the file** — attach it to the CRM, email it to the incoming AM, or store it in the client folder.

No backend, no account, no data leaves the browser.

## 8. Design Principles

- **Executive aesthetic** — slate-white, black, single red accent. No gradients-for-gradient's-sake.
- **Show, don't tell** — pre-filled mock data demonstrates the tool's value in the first 3 seconds.
- **One artifact, not ten** — the Excel file is the deliverable. Everything else is scaffolding.
- **Polish over features** — a beautiful, branded XLSX beats a plain one with twice the columns.

## 9. Roadmap Ideas

- Persist drafts to Lovable Cloud (auto-save)
- Real AI polish via Lovable AI Gateway (Gemini) instead of static transforms
- PDF export with the same branding
- Multi-client library with search and version history
- Role-based access (outgoing AM, incoming AM, manager)
- Slack / Email delivery of the generated workbook

## 10. Deploying

This project is built and hosted with **Lovable**. To publish your own changes:

- **Frontend changes** — click **Publish → Update** in the Lovable editor
- **Custom domain** — Project Settings → Domains, after first publish

To self-host, the codebase is standard TanStack Start and can be deployed to any platform that supports a Vite-built TanStack Start app (Cloudflare Workers, Netlify, Vercel, Node).

## 11. Connecting to GitHub

In the Lovable editor: **Plus (+) → GitHub → Connect project**, authorize the Lovable GitHub App, and create a repository. Once connected, changes in Lovable push to GitHub in real time, and commits pushed to GitHub sync back to Lovable automatically.

## 12. License

Proprietary — all rights reserved by the project owner. Contact for licensing terms.

---

Built with Lovable.
