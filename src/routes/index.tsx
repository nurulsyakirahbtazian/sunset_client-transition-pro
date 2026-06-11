import { createFileRoute } from "@tanstack/react-router";
import { useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import { saveAs } from "file-saver";
import {
  Upload, FileSpreadsheet, Users, Briefcase, ListChecks, AlertTriangle,
  MessageSquare, BookOpen, Layers, Sparkles, Clock, ShieldCheck, Download,
  Calendar, CheckCircle2, FileText, Plus, Trash2,
} from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Handover OS — B2B Client Transition Dashboard" },
      { name: "description", content: "Premium B2B client handover and account transition platform with AI-polished insights and one-click Excel export." },
    ],
  }),
  component: Dashboard,
});

// ---------- Mock data ----------
type Stakeholder = { name: string; role: string; email: string; notes: string };
type RecurringTask = { task: string; frequency: string; instructions: string };
type Issue = { issue: string; priority: string; status: string };

const initialClient = {
  name: "",
  industry: "",
  region: "",
  services: "",
};

const initialStakeholders: Stakeholder[] = [];

const initialTasks: RecurringTask[] = [];

const initialPlatforms: Record<string, boolean> = {
  Marketo: false,
  Salesforce: false,
  "6sense": false,
  ON24: false,
  "HubSpot": false,
  "LinkedIn Campaign Manager": false,
};

const initialIssues: Issue[] = [];

const initialPrefs = {
  communication: "",
  reporting: "",
  escalation: "",
};

const initialKT = {
  tribal: "",
  watchouts: "",
  history: "",
};

type PlanItem = { done: boolean; title: string; detail: string };
const initialPlan: PlanItem[] = [];


// ---------- Component ----------
function Dashboard() {
  const [client, setClient] = useState(initialClient);
  const [stakeholders, setStakeholders] = useState<Stakeholder[]>(initialStakeholders);
  const [tasks, setTasks] = useState<RecurringTask[]>(initialTasks);
  const [platforms, setPlatforms] = useState(initialPlatforms);
  const [customPlatform, setCustomPlatform] = useState("");
  const [issues, setIssues] = useState<Issue[]>(initialIssues);
  const [prefs, setPrefs] = useState(initialPrefs);
  const [kt, setKT] = useState(initialKT);
  const [plan, setPlan] = useState<PlanItem[]>(initialPlan);

  const [uploadedFiles, setUploadedFiles] = useState<string[]>([]);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);


  const handleFiles = (files: FileList | null) => {
    if (!files) return;
    setUploadedFiles((p) => [...p, ...Array.from(files).map((f) => f.name)]);
  };

  const generateExcel = () => {
    // ----- Styling helpers -----
    const BRAND = "FF0F0F";
    const INK = "1A2233";
    const SUBINK = "3A4456";
    const SOFT = "F4F6FA";
    const BORDER_GRAY = "D8DEE7";
    const ZEBRA = "FAFBFD";

    const border = {
      top: { style: "thin", color: { rgb: BORDER_GRAY } },
      bottom: { style: "thin", color: { rgb: BORDER_GRAY } },
      left: { style: "thin", color: { rgb: BORDER_GRAY } },
      right: { style: "thin", color: { rgb: BORDER_GRAY } },
    };

    const titleStyle = {
      font: { name: "Calibri", sz: 18, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: INK } },
      alignment: { vertical: "center", horizontal: "left", indent: 1 },
    };
    const subtitleStyle = {
      font: { name: "Calibri", sz: 10, italic: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: SUBINK } },
      alignment: { vertical: "center", horizontal: "left", indent: 1 },
    };
    const sectionStyle = {
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: BRAND } },
      alignment: { vertical: "center", horizontal: "left", indent: 1 },
      border,
    };
    const headerStyle = {
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: INK } },
      alignment: { vertical: "center", horizontal: "left", indent: 1, wrapText: true },
      border,
    };
    const labelStyle = {
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: INK } },
      fill: { fgColor: { rgb: SOFT } },
      alignment: { vertical: "center", horizontal: "left", indent: 1, wrapText: true },
      border,
    };
    const valueStyle = (zebra = false) => ({
      font: { name: "Calibri", sz: 10, color: { rgb: SUBINK } },
      fill: { fgColor: { rgb: zebra ? ZEBRA : "FFFFFF" } },
      alignment: { vertical: "center", horizontal: "left", indent: 1, wrapText: true },
      border,
    });
    const metricLabelStyle = {
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: INK } },
      alignment: { vertical: "center", horizontal: "left", indent: 1, wrapText: true },
      border,
    };
    const metricValueStyle = {
      font: { name: "Calibri", sz: 16, bold: true, color: { rgb: BRAND } },
      fill: { fgColor: { rgb: "FFFFFF" } },
      alignment: { vertical: "center", horizontal: "left", indent: 1 },
      border,
    };
    const priorityStyle = (p: string) => {
      const map: Record<string, string> = {
        Critical: "C81E1E", High: "E0651A", Medium: "B58900", Low: "4A6B8A",
      };
      return {
        font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: map[p] || "4A6B8A" } },
        alignment: { vertical: "center", horizontal: "center" },
        border,
      };
    };
    const statusStyle = (s: string) => {
      const map: Record<string, string> = {
        Open: "EEF2F7", "In Progress": "DCEEFB", Escalated: "FDE2E2", Resolved: "D9F2E1",
      };
      const fg: Record<string, string> = {
        Open: SUBINK, "In Progress": "1E5A93", Escalated: "9B1C1C", Resolved: "1F6B3B",
      };
      return {
        font: { name: "Calibri", sz: 10, bold: true, color: { rgb: fg[s] || SUBINK } },
        fill: { fgColor: { rgb: map[s] || "EEF2F7" } },
        alignment: { vertical: "center", horizontal: "center" },
        border,
      };
    };
    const checkStyle = {
      font: { name: "Calibri", sz: 14, bold: true, color: { rgb: BRAND } },
      fill: { fgColor: { rgb: SOFT } },
      alignment: { vertical: "center", horizontal: "center" },
      border,
    };
    const footerStyle = {
      font: { name: "Calibri", sz: 9, italic: true, color: { rgb: "8892A6" } },
      alignment: { vertical: "center", horizontal: "left", indent: 1 },
    };

    // Build a sheet from AOA + per-cell style map
    type CellSpec = { v: any; s?: any; t?: string };
    const buildSheet = (
      rows: (CellSpec | string | number | null)[][],
      opts: {
        cols: { wch: number }[];
        rowHeights?: Record<number, number>;
        merges?: XLSX.Range[];
      }
    ): XLSX.WorkSheet => {
      const aoa = rows.map((r) => r.map((c) => (c && typeof c === "object" ? (c as CellSpec).v : c)));
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      rows.forEach((row, r) => {
        row.forEach((cell, c) => {
          if (cell && typeof cell === "object" && (cell as CellSpec).s) {
            const addr = XLSX.utils.encode_cell({ r, c });
            if (ws[addr]) ws[addr].s = (cell as CellSpec).s;
          }
        });
      });
      ws["!cols"] = opts.cols;
      if (opts.merges) ws["!merges"] = opts.merges;
      ws["!rows"] = Array.from({ length: rows.length }, (_, i) => ({
        hpt: opts.rowHeights?.[i] ?? 20,
      }));
      return ws;
    };

    const dateStr = new Date().toLocaleDateString("en-US", {
      year: "numeric", month: "long", day: "numeric",
    });

    const wb = XLSX.utils.book_new();
    wb.Props = {
      Title: `Client Handover — ${client.name}`,
      Subject: "B2B Client Handover & Account Transition",
      Author: "Handover OS",
      Company: "Handover OS",
      CreatedDate: new Date(),
    };

    // ============ 1. EXECUTIVE SUMMARY ============
    {
      const rows: (CellSpec | null)[][] = [
        [{ v: "B2B CLIENT HANDOVER  —  EXECUTIVE SUMMARY", s: titleStyle }, null, null, null],
        [{ v: `${client.name}  ·  Prepared ${dateStr}  ·  Confidential`, s: subtitleStyle }, null, null, null],
        [null, null, null, null],
        [{ v: "ACCOUNT SNAPSHOT", s: sectionStyle }, null, null, null],
        [{ v: "Client", s: labelStyle }, { v: client.name, s: valueStyle() }, { v: "Industry", s: labelStyle }, { v: client.industry, s: valueStyle() }],
        [{ v: "Region", s: labelStyle }, { v: client.region, s: valueStyle(true) }, { v: "Services", s: labelStyle }, { v: client.services, s: valueStyle(true) }],
        [null, null, null, null],
        [{ v: "EXECUTIVE METRICS", s: sectionStyle }, null, null, null],
        [{ v: "Time Saved per Handover", s: metricLabelStyle }, { v: "6.5 Hours", s: metricValueStyle }, { v: "Account Risk Mitigation Score", s: metricLabelStyle }, { v: "98%", s: metricValueStyle }],
        [null, null, null, null],
        [{ v: "30-DAY TRANSITION PLAN", s: sectionStyle }, null, null, null],
        [{ v: "Week", s: headerStyle }, { v: "Phase", s: headerStyle }, { v: "Key Deliverables", s: headerStyle }, { v: "Owner", s: headerStyle }],
        [{ v: "Week 1", s: { ...valueStyle(), font: { ...valueStyle().font, bold: true, color: { rgb: BRAND } } } }, { v: "Discovery & Access", s: valueStyle() }, { v: "Stakeholder intros, access provisioning, platform audit kickoff", s: valueStyle() }, { v: "Incoming Lead", s: valueStyle() }],
        [{ v: "Week 2", s: { ...valueStyle(true), font: { ...valueStyle(true).font, bold: true, color: { rgb: BRAND } } } }, { v: "Shadow & Validate", s: valueStyle(true) }, { v: "Shadow recurring tasks, validate SOPs, baseline reporting setup", s: valueStyle(true) }, { v: "Joint", s: valueStyle(true) }],
        [{ v: "Week 3", s: { ...valueStyle(), font: { ...valueStyle().font, bold: true, color: { rgb: BRAND } } } }, { v: "Co-Execute", s: valueStyle() }, { v: "Co-execute deliverables with outgoing lead, resolve top 3 open issues", s: valueStyle() }, { v: "Joint", s: valueStyle() }],
        [{ v: "Week 4", s: { ...valueStyle(true), font: { ...valueStyle(true).font, bold: true, color: { rgb: BRAND } } } }, { v: "Ownership Transfer", s: valueStyle(true) }, { v: "Full ownership transfer, executive QBR readout, 90-day roadmap delivery", s: valueStyle(true) }, { v: "Incoming Lead", s: valueStyle(true) }],
        [null, null, null, null],
        [{ v: "Generated by Handover OS  ·  handoveros.com  ·  Confidential & Proprietary", s: footerStyle }, null, null, null],
      ];
      const ws = buildSheet(rows, {
        cols: [{ wch: 26 }, { wch: 38 }, { wch: 26 }, { wch: 38 }],
        rowHeights: { 0: 36, 1: 22, 3: 24, 7: 24, 8: 38, 10: 24, 11: 24 },
        merges: [
          { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
          { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
          { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
          { s: { r: 7, c: 0 }, e: { r: 7, c: 3 } },
          { s: { r: 10, c: 0 }, e: { r: 10, c: 3 } },
          { s: { r: 16, c: 0 }, e: { r: 16, c: 3 } },
        ],
      });
      (ws as any)["!sheetView"] = [{ showGridLines: false }];
      XLSX.utils.book_append_sheet(wb, ws, "Executive Summary");
    }

    // Helper for standard styled table sheets
    const makeTableSheet = (
      title: string,
      headers: string[],
      dataRows: (string | { v: string; s: any })[][],
      cols: number[]
    ) => {
      const rows: (CellSpec | null)[][] = [
        [{ v: title.toUpperCase(), s: titleStyle }, ...Array(headers.length - 1).fill(null)],
        [{ v: `${client.name}  ·  ${dateStr}`, s: subtitleStyle }, ...Array(headers.length - 1).fill(null)],
        Array(headers.length).fill(null),
        headers.map((h) => ({ v: h, s: headerStyle })),
        ...dataRows.map((row, i) =>
          row.map((cell) =>
            typeof cell === "string" ? { v: cell, s: valueStyle(i % 2 === 1) } : cell
          )
        ),
      ];
      const ws = buildSheet(rows, {
        cols: cols.map((wch) => ({ wch })),
        rowHeights: { 0: 32, 1: 20, 3: 24 },
        merges: [
          { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
          { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
        ],
      });
      (ws as any)["!sheetView"] = [{ showGridLines: false }];
      ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: 3 + dataRows.length, c: headers.length - 1 } }) };
      ws["!freeze"] = { xSplit: 0, ySplit: 4 };
      return ws;
    };

    // 2. CLIENT OVERVIEW
    XLSX.utils.book_append_sheet(wb, makeTableSheet(
      "Client Overview", ["Field", "Value"],
      [
        ["Client Name", client.name],
        ["Industry", client.industry],
        ["Region", client.region],
        ["Services Delivered", client.services],
      ], [28, 60]
    ), "Client Overview");

    // 3. STAKEHOLDERS
    XLSX.utils.book_append_sheet(wb, makeTableSheet(
      "Stakeholders", ["Name", "Role", "Email", "Notes"],
      stakeholders.map((s) => [s.name, s.role, s.email, s.notes]),
      [24, 28, 32, 60]
    ), "Stakeholders");

    // 4. RECURRING TASKS
    XLSX.utils.book_append_sheet(wb, makeTableSheet(
      "Recurring Tasks", ["Task", "Frequency", "Instructions"],
      tasks.map((t) => [t.task, t.frequency, t.instructions]),
      [34, 26, 70]
    ), "Recurring Tasks");

    // 5. PLATFORMS
    XLSX.utils.book_append_sheet(wb, makeTableSheet(
      "Platforms", ["Platform", "Status"],
      Object.entries(platforms).map(([p, on]) => [
        p,
        {
          v: on ? "✓  Active" : "—  Not in scope",
          s: {
            font: { name: "Calibri", sz: 10, bold: true, color: { rgb: on ? "1F6B3B" : "8892A6" } },
            fill: { fgColor: { rgb: on ? "D9F2E1" : SOFT } },
            alignment: { vertical: "center", horizontal: "left", indent: 1 },
            border,
          },
        },
      ]), [36, 24]
    ), "Platforms");

    // 6. OPEN ISSUES (with colored priority/status pills)
    XLSX.utils.book_append_sheet(wb, makeTableSheet(
      "Open Issues", ["Issue", "Priority", "Status"],
      issues.map((it) => [
        it.issue,
        { v: it.priority.toUpperCase(), s: priorityStyle(it.priority) },
        { v: it.status, s: statusStyle(it.status) },
      ]), [70, 16, 18]
    ), "Open Issues");

    // 7. CLIENT PREFERENCES
    XLSX.utils.book_append_sheet(wb, makeTableSheet(
      "Client Preferences", ["Preference", "Detail"],
      [
        ["Communication Style", prefs.communication],
        ["Reporting Expectations", prefs.reporting],
        ["Escalation Path", prefs.escalation],
      ], [28, 80]
    ), "Client Preferences");

    // 8. KT CHECKLIST (with check column)
    {
      const ktData: (string | { v: string; s: any })[][] = [
        ["Tribal Knowledge", kt.tribal, { v: "☐", s: checkStyle }],
        ["Watch-outs", kt.watchouts, { v: "☐", s: checkStyle }],
        ["Historical Context", kt.history, { v: "☐", s: checkStyle }],
        ["Stakeholders introduced", "", { v: "☐", s: checkStyle }],
        ["Platform access verified", "", { v: "☐", s: checkStyle }],
        ["Recurring tasks shadowed", "", { v: "☐", s: checkStyle }],
        ["Open issues briefed", "", { v: "☐", s: checkStyle }],
        ["Executive QBR scheduled", "", { v: "☐", s: checkStyle }],
      ];
      XLSX.utils.book_append_sheet(wb, makeTableSheet(
        "KT Checklist", ["Category", "Notes", "Done"],
        ktData, [28, 70, 10]
      ), "KT Checklist");
    }

    const buf = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
    saveAs(new Blob([buf], { type: "application/octet-stream" }),
      `Handover_${client.name.replace(/\s+/g, "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-background text-foreground">
      {/* Top Banner */}
      <header className="border-b border-border bg-card">
        <div className="mx-auto max-w-[1600px] px-6 py-6">
          <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">
            <div className="flex items-start gap-5">
              <Logo />
              <div>
                <div className="flex items-center gap-2">
                  <span className="inline-flex items-center gap-1.5 rounded-full bg-brand/10 px-2.5 py-0.5 text-[10px] font-semibold uppercase tracking-wider text-brand">
                    <span className="h-1.5 w-1.5 rounded-full bg-brand" /> Handover OS
                  </span>
                  <span className="text-[11px] font-medium text-muted-foreground">v2.4 · Enterprise</span>
                </div>
                <h1 className="mt-1.5 text-2xl font-semibold tracking-tight text-slate-ink lg:text-[28px]">
                  B2B Client Handover & Account Transition
                </h1>
                <p className="mt-1 max-w-2xl text-sm text-muted-foreground">
                  Capture institutional knowledge, polish it with AI, and ship a board-ready Excel handover in minutes — not weeks.
                </p>
              </div>
            </div>

            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
              onClick={() => fileInputRef.current?.click()}
              className={`group flex cursor-pointer items-center gap-4 rounded-xl border-2 border-dashed px-5 py-4 transition-all lg:w-[460px] ${
                dragOver ? "border-brand bg-brand/5" : "border-border bg-muted/40 hover:border-brand/60 hover:bg-brand/5"
              }`}
            >
              <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-lg bg-slate-ink text-white">
                <Upload className="h-5 w-5" />
              </div>
              <div className="min-w-0 flex-1">
                <p className="text-sm font-semibold text-slate-ink">Upload Existing Documents</p>
                <p className="text-xs text-muted-foreground">Drag SOPs, notes, emails — or click to browse</p>
                {uploadedFiles.length > 0 && (
                  <p className="mt-1 truncate text-[11px] font-medium text-brand">
                    {uploadedFiles.length} file{uploadedFiles.length > 1 ? "s" : ""} attached · {uploadedFiles.slice(-1)[0]}
                  </p>
                )}
              </div>
              <input ref={fileInputRef} type="file" multiple className="hidden" onChange={(e) => handleFiles(e.target.files)} />
            </div>
          </div>
        </div>
      </header>

      {/* Main grid */}
      <main className="mx-auto max-w-[1600px] px-6 py-6">
        <div className="grid grid-cols-1 gap-6 xl:grid-cols-[minmax(0,1.15fr)_minmax(0,1fr)]">
          {/* LEFT PANEL */}
          <div className="space-y-5">
            <Section icon={<Briefcase className="h-4 w-4" />} title="Client Info" subtitle="Account fundamentals">
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                <Field label="Client Name"><input className="input-base" placeholder="e.g. Acme Corp" value={client.name} onChange={(e) => setClient({ ...client, name: e.target.value })} /></Field>
                <Field label="Industry"><input className="input-base" placeholder="e.g. SaaS / FinTech" value={client.industry} onChange={(e) => setClient({ ...client, industry: e.target.value })} /></Field>
                <Field label="Region"><input className="input-base" placeholder="e.g. North America" value={client.region} onChange={(e) => setClient({ ...client, region: e.target.value })} /></Field>
                <Field label="Services"><input className="input-base" placeholder="e.g. ABM, Demand Gen" value={client.services} onChange={(e) => setClient({ ...client, services: e.target.value })} /></Field>
              </div>
            </Section>

            <Section icon={<Users className="h-4 w-4" />} title="Stakeholders" subtitle="Key contacts & decision makers"
              action={<AddBtn onClick={() => setStakeholders([...stakeholders, { name: "", role: "", email: "", notes: "" }])} />}>
              <div className="space-y-3">
                {stakeholders.map((s, i) => (
                  <RowCard key={i} onRemove={() => setStakeholders(stakeholders.filter((_, x) => x !== i))}>
                    <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                      <input className="input-base" placeholder="Name" value={s.name} onChange={(e) => updateArr(setStakeholders, stakeholders, i, { ...s, name: e.target.value })} />
                      <input className="input-base" placeholder="Role" value={s.role} onChange={(e) => updateArr(setStakeholders, stakeholders, i, { ...s, role: e.target.value })} />
                      <input className="input-base md:col-span-2" placeholder="Email" value={s.email} onChange={(e) => updateArr(setStakeholders, stakeholders, i, { ...s, email: e.target.value })} />
                      <textarea className="input-base md:col-span-2" rows={2} placeholder="Notes" value={s.notes} onChange={(e) => updateArr(setStakeholders, stakeholders, i, { ...s, notes: e.target.value })} />
                    </div>
                  </RowCard>
                ))}
              </div>
            </Section>

            <Section icon={<ListChecks className="h-4 w-4" />} title="Recurring Tasks" subtitle="Operational cadence"
              action={<AddBtn onClick={() => setTasks([...tasks, { task: "", frequency: "", instructions: "" }])} />}>
              <div className="space-y-3">
                {tasks.map((t, i) => (
                  <RowCard key={i} onRemove={() => setTasks(tasks.filter((_, x) => x !== i))}>
                    <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                      <input className="input-base" placeholder="Task Name" value={t.task} onChange={(e) => updateArr(setTasks, tasks, i, { ...t, task: e.target.value })} />
                      <input className="input-base" placeholder="Frequency" value={t.frequency} onChange={(e) => updateArr(setTasks, tasks, i, { ...t, frequency: e.target.value })} />
                      <textarea className="input-base md:col-span-2" rows={2} placeholder="Instructions" value={t.instructions} onChange={(e) => updateArr(setTasks, tasks, i, { ...t, instructions: e.target.value })} />
                    </div>
                  </RowCard>
                ))}
              </div>
            </Section>

            <Section icon={<Layers className="h-4 w-4" />} title="Platforms Checklist" subtitle="Tech stack in scope">
              <div className="grid grid-cols-2 gap-2.5 md:grid-cols-3">
                {Object.entries(platforms).map(([name, on]) => (
                  <label key={name} className={`flex cursor-pointer items-center gap-2.5 rounded-lg border px-3 py-2.5 text-sm transition-all ${
                    on ? "border-brand/40 bg-brand/5 text-slate-ink" : "border-border bg-card text-muted-foreground hover:border-slate-soft/40"
                  }`}>
                    <input type="checkbox" checked={on} onChange={(e) => setPlatforms({ ...platforms, [name]: e.target.checked })}
                      className="h-4 w-4 accent-brand" />
                    <span className="font-medium">{name}</span>
                  </label>
                ))}
              </div>
              <div className="mt-3 flex gap-2">
                <input className="input-base" placeholder="Add custom platform…" value={customPlatform} onChange={(e) => setCustomPlatform(e.target.value)} />
                <button
                  onClick={() => { if (customPlatform.trim()) { setPlatforms({ ...platforms, [customPlatform.trim()]: true }); setCustomPlatform(""); }}}
                  className="shrink-0 rounded-lg bg-slate-ink px-4 text-sm font-medium text-white hover:bg-slate-ink/90">
                  Add
                </button>
              </div>
            </Section>

            <Section icon={<AlertTriangle className="h-4 w-4" />} title="Open Issues" subtitle="Active risks & blockers"
              action={<AddBtn onClick={() => setIssues([...issues, { issue: "", priority: "Medium", status: "Open" }])} />}>
              <div className="space-y-2">
                {issues.map((it, i) => (
                  <div key={i} className="grid grid-cols-12 gap-2">
                    <input className="input-base col-span-7" placeholder="Issue" value={it.issue} onChange={(e) => updateArr(setIssues, issues, i, { ...it, issue: e.target.value })} />
                    <select className="input-base col-span-2" value={it.priority} onChange={(e) => updateArr(setIssues, issues, i, { ...it, priority: e.target.value })}>
                      <option>Low</option><option>Medium</option><option>High</option><option>Critical</option>
                    </select>
                    <select className="input-base col-span-2" value={it.status} onChange={(e) => updateArr(setIssues, issues, i, { ...it, status: e.target.value })}>
                      <option>Open</option><option>In Progress</option><option>Escalated</option><option>Resolved</option>
                    </select>
                    <button onClick={() => setIssues(issues.filter((_, x) => x !== i))}
                      className="col-span-1 flex items-center justify-center rounded-lg border border-border text-muted-foreground hover:border-brand hover:text-brand">
                      <Trash2 className="h-4 w-4" />
                    </button>
                  </div>
                ))}
              </div>
            </Section>

            <Section icon={<MessageSquare className="h-4 w-4" />} title="Client Preferences" subtitle="How they like to work">
              <div className="space-y-4">
                <Field label="Communication Style"><textarea rows={2} className="input-base" placeholder="e.g. Weekly calls, Slack channel, email for formal updates" value={prefs.communication} onChange={(e) => setPrefs({ ...prefs, communication: e.target.value })} /></Field>
                <Field label="Reporting Expectations"><textarea rows={2} className="input-base" placeholder="e.g. Monthly dashboard, quarterly business review" value={prefs.reporting} onChange={(e) => setPrefs({ ...prefs, reporting: e.target.value })} /></Field>
                <Field label="Escalation Path"><textarea rows={2} className="input-base" placeholder="e.g. Account Lead → Director → VP" value={prefs.escalation} onChange={(e) => setPrefs({ ...prefs, escalation: e.target.value })} /></Field>
              </div>
            </Section>

            <Section icon={<BookOpen className="h-4 w-4" />} title="Knowledge Transfer Notes" subtitle="The things only the outgoing lead knows">
              <div className="space-y-4">
                <Field label="Tribal Knowledge"><textarea rows={3} className="input-base" placeholder="e.g. Unwritten rules, past decisions, or relationships that matter" value={kt.tribal} onChange={(e) => setKT({ ...kt, tribal: e.target.value })} /></Field>
                <Field label="Watch-outs"><textarea rows={3} className="input-base" placeholder="e.g. Sensitive dates, approval delays, or things to avoid" value={kt.watchouts} onChange={(e) => setKT({ ...kt, watchouts: e.target.value })} /></Field>
                <Field label="Historical Context"><textarea rows={3} className="input-base" placeholder="e.g. Account start date, past migrations, or team changes" value={kt.history} onChange={(e) => setKT({ ...kt, history: e.target.value })} /></Field>
              </div>
            </Section>
          </div>

          {/* RIGHT PANEL */}
          <div className="space-y-5 xl:sticky xl:top-6 xl:self-start">
            {/* 30-Day Transition Plan — editable form */}

            <Section
              icon={<Calendar className="h-4 w-4" />}
              title="30-Day Transition Plan"
              subtitle="Add the milestones, tasks, or deliverables for this handover"
              action={
                <button
                  onClick={() => setPlan([...plan, { done: false, title: "", detail: "" }])}
                  className="inline-flex items-center gap-1.5 rounded-md bg-slate-ink px-2.5 py-1.5 text-[11px] font-semibold text-white hover:bg-slate-ink/90"
                >
                  <Plus className="h-3.5 w-3.5" /> Add Item
                </button>
              }
            >
              {plan.length === 0 ? (
                <div className="rounded-lg border border-dashed border-border bg-muted/30 p-6 text-center">
                  <p className="text-[13px] text-muted-foreground">No plan items yet.</p>
                  <p className="mt-1 text-[11px] text-muted-foreground">Click "Add Item" to start building the transition plan.</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {plan.map((item, i) => (
                    <div key={i} className="rounded-lg border border-border bg-card p-3">
                      <div className="flex items-start gap-3">
                        <input
                          type="checkbox"
                          checked={item.done}
                          onChange={(e) => updateArr(setPlan, plan, i, { ...item, done: e.target.checked })}
                          className="mt-2 h-4 w-4 shrink-0 cursor-pointer accent-brand"
                        />
                        <div className="flex-1 space-y-2">
                          <input
                            className="input-base"
                            placeholder="e.g. Week 1 — Stakeholder intro calls"
                            value={item.title}
                            onChange={(e) => updateArr(setPlan, plan, i, { ...item, title: e.target.value })}
                          />
                          <textarea
                            rows={2}
                            className="input-base"
                            placeholder="Elaborate: deliverables, owners, dates, notes…"
                            value={item.detail}
                            onChange={(e) => updateArr(setPlan, plan, i, { ...item, detail: e.target.value })}
                          />
                        </div>
                        <button
                          onClick={() => setPlan(plan.filter((_, k) => k !== i))}
                          className="mt-1 rounded-md p-1.5 text-muted-foreground hover:bg-muted hover:text-destructive"
                          aria-label="Remove plan item"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Section>

            {/* Action */}
            <button
              onClick={generateExcel}
              className="group relative flex w-full items-center justify-center gap-3 overflow-hidden rounded-xl bg-brand px-6 py-5 text-base font-semibold text-brand-foreground shadow-brand transition-all hover:translate-y-[-1px] hover:bg-brand/95 active:translate-y-0"

            >
              <FileSpreadsheet className="h-5 w-5" />
              <span>Generate &amp; Download Excel</span>
              <Download className="h-5 w-5 transition-transform group-hover:translate-y-0.5" />
              <span className="ml-2 rounded-md bg-white/15 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wider">8 tabs · .xlsx</span>
            </button>

            <p className="text-center text-[11px] text-muted-foreground">
              Generates a fully-structured workbook from your live form data · No data leaves your browser
            </p>
          </div>
        </div>
      </main>

      <footer className="mt-8 border-t border-border bg-card">
        <div className="mx-auto flex max-w-[1600px] flex-col items-center justify-between gap-2 px-6 py-4 text-xs text-muted-foreground sm:flex-row">
          <span>© Handover OS — Enterprise Client Transition Platform</span>
          <span>SOC 2 Type II · ISO 27001 · GDPR-ready</span>
        </div>
      </footer>
    </div>
  );
}

// ---------- Sub-components ----------
function Logo() {
  return (
    <div className="flex shrink-0 items-center gap-3">
      <div className="relative flex h-14 w-14 items-center justify-center rounded-xl bg-slate-ink shadow-elevated">
        <div className="absolute inset-1 rounded-lg border border-white/10" />
        <span className="relative text-2xl font-black tracking-tight text-white">
          H<span className="text-brand">·</span>O
        </span>
      </div>
    </div>
  );
}

function Section({ icon, title, subtitle, action, children }: {
  icon: React.ReactNode; title: string; subtitle?: string; action?: React.ReactNode; children: React.ReactNode;
}) {
  return (
    <section className="overflow-hidden rounded-xl border border-border bg-card shadow-[0_1px_3px_rgb(15_23_42/0.04)]">
      <header className="flex items-center justify-between border-b border-border px-5 py-3.5">
        <div className="flex items-center gap-2.5">
          <span className="flex h-7 w-7 items-center justify-center rounded-md bg-slate-ink text-white">{icon}</span>
          <div>
            <h3 className="text-sm font-semibold text-slate-ink">{title}</h3>
            {subtitle && <p className="text-[11px] text-muted-foreground">{subtitle}</p>}
          </div>
        </div>
        {action}
      </header>
      <div className="p-5">{children}</div>
    </section>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="block"><span className="label-base">{label}</span>{children}</label>;
}

function RowCard({ children, onRemove }: { children: React.ReactNode; onRemove: () => void }) {
  return (
    <div className="group relative rounded-lg border border-border bg-muted/30 p-3 transition-colors hover:border-slate-soft/40">
      <button onClick={onRemove} className="absolute right-2 top-2 rounded-md p-1 text-muted-foreground opacity-0 transition-opacity hover:bg-brand/10 hover:text-brand group-hover:opacity-100">
        <Trash2 className="h-3.5 w-3.5" />
      </button>
      {children}
    </div>
  );
}

function AddBtn({ onClick }: { onClick: () => void }) {
  return (
    <button onClick={onClick} className="flex items-center gap-1 rounded-md border border-border bg-card px-2.5 py-1 text-[11px] font-semibold text-slate-ink transition-colors hover:border-brand hover:text-brand">
      <Plus className="h-3 w-3" /> Add
    </button>
  );
}

function MetricCard({ icon, value, label, trend, tone }: {
  icon: React.ReactNode; value: string; label: string; trend: string; tone: "dark" | "brand";
}) {
  const dark = tone === "dark";
  return (
    <div className={`relative overflow-hidden rounded-xl p-5 shadow-elevated ${dark ? "bg-slate-ink text-white" : "bg-brand text-brand-foreground"}`}>
      <div className="flex items-start justify-between">
        <span className={`flex h-9 w-9 items-center justify-center rounded-lg ${dark ? "bg-white/10" : "bg-white/15"}`}>{icon}</span>
        <span className={`text-[10px] font-semibold uppercase tracking-wider ${dark ? "text-white/60" : "text-white/80"}`}>Executive</span>
      </div>
      <div className="mt-5 text-[34px] font-bold leading-none tracking-tight">{value}</div>
      <div className={`mt-1.5 text-sm font-medium ${dark ? "text-white/85" : "text-white/95"}`}>{label}</div>
      <div className={`mt-3 text-[11px] font-medium ${dark ? "text-white/55" : "text-white/80"}`}>{trend}</div>
      <div className="absolute -right-8 -bottom-8 h-32 w-32 rounded-full bg-white/5" />
    </div>
  );
}

function PlanTab({ clientName }: { clientName: string }) {
  const weeks = [
    { w: "Week 1", title: "Discovery & Access", items: ["Stakeholder intro calls (4 contacts)", "Provision platform access (Marketo, SFDC, 6sense, ON24)", "Baseline audit of active programs"] },
    { w: "Week 2", title: "Shadow & Validate", items: ["Shadow weekly Marketo audit", "Validate SOPs against live workflows", "Stand up reporting baseline in Looker"] },
    { w: "Week 3", title: "Co-Execute", items: ["Co-own deliverables with outgoing lead", "Resolve top 3 open issues (SFDC sync, GDPR, intent flow)", "Run first ON24 webinar end-to-end"] },
    { w: "Week 4", title: "Ownership Transfer", items: ["Full handover sign-off with VP Demand Gen", "Executive QBR readout to CMO", "Deliver 90-day forward roadmap"] },
  ];
  return (
    <div>
      <div className="mb-4 rounded-lg border border-border bg-muted/40 p-4">
        <p className="text-[11px] font-semibold uppercase tracking-wider text-brand">Executive Summary</p>
        <p className="mt-1.5 text-sm leading-relaxed text-slate-ink">
          <strong>{clientName}</strong> transitions from outgoing lead to new account team across a 30-day structured handover.
          The plan prioritizes zero-disruption to in-flight campaigns, full institutional knowledge transfer, and executive-grade
          reporting continuity. Risk is mitigated through phased shadowing and co-execution before full ownership transfer.
        </p>
      </div>
      <div className="relative space-y-3 pl-6">
        <div className="absolute left-2 top-2 bottom-2 w-px bg-border" />
        {weeks.map((w, i) => (
          <div key={i} className="relative">
            <div className="absolute -left-[18px] top-1 flex h-3.5 w-3.5 items-center justify-center rounded-full bg-brand ring-4 ring-card" />
            <div className="rounded-lg border border-border bg-card p-3.5">
              <div className="flex items-center gap-2">
                <Calendar className="h-3.5 w-3.5 text-brand" />
                <span className="text-[11px] font-bold uppercase tracking-wider text-brand">{w.w}</span>
                <span className="text-sm font-semibold text-slate-ink">· {w.title}</span>
              </div>
              <ul className="mt-2 space-y-1">
                {w.items.map((it, k) => (
                  <li key={k} className="flex items-start gap-2 text-[13px] text-muted-foreground">
                    <CheckCircle2 className="mt-0.5 h-3.5 w-3.5 shrink-0 text-slate-ink/60" />
                    <span>{it}</span>
                  </li>
                ))}
              </ul>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function InsightsTab() {
  return (
    <div className="space-y-4">
      <div className="rounded-lg border border-dashed border-border bg-muted/40 p-4">
        <p className="text-[10px] font-bold uppercase tracking-wider text-muted-foreground">Raw Input</p>
        <p className="mt-1.5 text-[13px] italic leading-relaxed text-slate-soft">
          "margaret hates when ppl over-promise on attribution lol. last agency burned her. always show ranges, never single touch.
          also dont schedule launches during fiscal close (last week of mar/jun/sep/dec) — legal needs 5 biz days for gated stuff"
        </p>
      </div>
      <div className="flex justify-center">
        <div className="flex h-7 w-7 items-center justify-center rounded-full bg-brand text-white shadow-brand">
          <Sparkles className="h-3.5 w-3.5" />
        </div>
      </div>
      <div className="rounded-lg border border-brand/20 bg-gradient-to-br from-brand/5 to-transparent p-4">
        <p className="text-[10px] font-bold uppercase tracking-wider text-brand">AI-Polished · Compliance-Ready</p>
        <div className="mt-2 space-y-2.5 text-[13px] leading-relaxed text-slate-ink">
          <p>
            <strong>Executive Sponsor Sensitivity:</strong> The VP of Demand Generation has experienced material dissatisfaction
            with previous agency partners regarding marketing attribution methodology. All performance reporting must present
            results with conservative confidence ranges. Single-touch attribution models must not be presented without
            multi-touch contextualization.
          </p>
          <p>
            <strong>Operational Scheduling Constraints:</strong> Campaign launches must not be scheduled during the client's
            fiscal close periods, defined as the final calendar week of March, June, September, and December. Additionally,
            all gated content assets require a minimum five (5) business day legal review prior to publication.
          </p>
        </div>
      </div>
    </div>
  );
}

function StructureTab({ tabs }: { tabs: string[] }) {
  const meta = [
    "Board-ready overview · KPIs · 30-day plan",
    "Account fundamentals & service scope",
    "Contact directory with roles & notes",
    "Operational cadence & SOPs",
    "Tech stack inventory & access",
    "Active risks with priority & status",
    "Communication, reporting, escalation",
    "Sign-off checklist for handover completion",
  ];
  return (
    <div>
      <p className="mb-3 text-[13px] text-muted-foreground">
        Your downloadable workbook contains <strong className="text-slate-ink">8 fully-structured tabs</strong>, each
        formatted for direct use in QBRs, executive reviews, and account onboarding.
      </p>
      <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
        {tabs.map((t, i) => (
          <div key={t} className="flex items-center gap-3 rounded-lg border border-border bg-card p-3 transition-colors hover:border-brand/40">
            <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-md bg-slate-ink text-[11px] font-bold text-white">
              {String(i + 1).padStart(2, "0")}
            </span>
            <div className="min-w-0">
              <p className="truncate text-[13px] font-semibold text-slate-ink">{t}</p>
              <p className="truncate text-[11px] text-muted-foreground">{meta[i]}</p>
            </div>
            <FileText className="ml-auto h-4 w-4 shrink-0 text-muted-foreground" />
          </div>
        ))}
      </div>
    </div>
  );
}

function updateArr<T>(setter: (v: T[]) => void, arr: T[], i: number, v: T) {
  setter(arr.map((x, k) => (k === i ? v : x)));
}
