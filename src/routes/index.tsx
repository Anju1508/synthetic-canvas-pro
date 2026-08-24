import { createFileRoute } from "@tanstack/react-router";
import { useDeferredValue, useMemo, useRef, useState, type ChangeEvent, type ReactNode } from "react";
import ExcelJS from "exceljs";
import {
  ThumbsUp,
  ThumbsDown,
  MessageSquare,
  Trash2,
  Pencil,
  Plus,
  Upload,
  Download,
  FileJson,
  Check,
  X,
  Sparkles,
  Database,
  Search,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  LoaderCircle,
} from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "SynthLab — Review & Export LLM Fine-tuning Data" },
      {
        name: "description",
        content:
          "Upload synthetic JSON/JSONL datasets, review each sample with likes, dislikes and comments, then export a formatted XLSX for LLM fine-tuning workflows.",
      },
      { property: "og:title", content: "SynthLab — Review & Export LLM Fine-tuning Data" },
      {
        property: "og:description",
        content:
          "A clean workspace to inspect, edit, and curate synthetic LLM fine-tuning datasets with one-click XLSX export.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

type Feedback = "like" | "dislike" | null;

interface Row {
  id: string;
  data: Record<string, unknown>;
  feedback: Feedback;
  comment: string;
}

const PAGE_SIZE = 30;

function uid() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function parseInput(text: string): Record<string, unknown>[] {
  const trimmed = text.trim();
  if (!trimmed) return [];
  try {
    const parsed = JSON.parse(trimmed);
    if (Array.isArray(parsed)) {
      if (parsed.length > 0 && parsed.every(isChatMessage)) {
        return [{ messages: parsed }];
      }
      return parsed.map((r) => (isObj(r) ? r : { value: r }));
    }
    if (isObj(parsed)) {
      for (const k of ["data", "rows", "items", "samples", "examples"]) {
        const v = (parsed as Record<string, unknown>)[k];
        if (Array.isArray(v)) {
          if (v.length > 0 && v.every(isChatMessage)) return [{ messages: v }];
          return v.map((r) => (isObj(r) ? r : { value: r }));
        }
      }
      return [parsed];
    }
  } catch {
    /* fall through to JSONL */
  }
  const out: Record<string, unknown>[] = [];
  const lines = trimmed.split(/\r?\n/);
  for (const line of lines) {
    const l = line.trim();
    if (!l) continue;
    try {
      const p = JSON.parse(l);
      out.push(isObj(p) ? p : { value: p });
    } catch {
      // skip malformed line
    }
  }
  return out;
}

function isObj(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function isChatMessage(v: unknown): v is { role: string; content: string } {
  return (
    isObj(v) &&
    typeof (v as Record<string, unknown>).role === "string" &&
    typeof (v as Record<string, unknown>).content === "string"
  );
}

function extractThink(content: string): { think: string | null; body: string } {
  const m = content.match(/^\s*<think>([\s\S]*?)<\/think>\s*/i);
  if (!m) return { think: null, body: content };
  return { think: m[1].trim(), body: content.slice(m[0].length) };
}

function excelValue(value: ExcelJS.CellValue): unknown {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") {
    if ("result" in value) return value.result ?? "";
    if ("text" in value) return value.text;
    if ("richText" in value) return value.richText.map((part) => part.text).join("");
  }
  return value;
}

function parseCellValue(value: unknown): unknown {
  if (typeof value !== "string") return value;
  const trimmed = value.trim();
  if (!trimmed || (!trimmed.startsWith("{") && !trimmed.startsWith("["))) return value;
  try {
    return JSON.parse(trimmed);
  } catch {
    return value;
  }
}

async function parseWorkbook(file: File): Promise<Row[]> {
  const workbook = new ExcelJS.Workbook();
  // Some workbook producers omit optional sheet names; ExcelJS expects a string.
  workbook.addWorksheet("Import");
  await workbook.xlsx.load(await file.arrayBuffer());
  const worksheet = workbook.worksheets.find((sheet) => sheet.name !== "Import") ?? workbook.worksheets[0];
  if (!worksheet || worksheet.rowCount < 1) return [];

  const headers = worksheet.getRow(1).values as ExcelJS.CellValue[];
  const columns = Array.from(
    { length: worksheet.columnCount + 1 },
    (_, index) => String(excelValue(headers[index]) ?? "").trim(),
  );
  const feedbackIndex = columns.findIndex((value) => String(value ?? "").toLowerCase() === "feedback");
  const commentIndex = columns.findIndex((value) => String(value ?? "").toLowerCase() === "comment");
  const rows: Row[] = [];

  worksheet.eachRow((excelRow, rowNumber) => {
    if (rowNumber === 1) return;
    const data: Record<string, unknown> = {};
    let hasData = false;
    columns.forEach((header, index) => {
      if (!header || header === "#" || index === feedbackIndex || index === commentIndex) return;
      const value = excelValue(excelRow.getCell(index).value);
      data[header] = parseCellValue(value);
      if (value !== "" && value !== null && value !== undefined) hasData = true;
    });
    if (!hasData) return;
    const feedbackValue = feedbackIndex > 0
      ? String(excelValue(excelRow.getCell(feedbackIndex).value) ?? "").toLowerCase()
      : "";
    rows.push({
      id: uid(),
      data,
      feedback: feedbackValue === "like" || feedbackValue === "dislike" ? feedbackValue : null,
      comment: commentIndex > 0 ? String(excelValue(excelRow.getCell(commentIndex).value) ?? "") : "",
    });
  });
  return rows;
}

function Index() {
  const [rows, setRows] = useState<Row[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editDraft, setEditDraft] = useState<string>("");
  const [editError, setEditError] = useState<string>("");
  const [commentingId, setCommentingId] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<"all" | "like" | "dislike" | "none" | "commented">("all");
  const [page, setPage] = useState(1);
  const [busy, setBusy] = useState<"import" | "export" | null>(null);
  const [fileError, setFileError] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);
  const deferredSearch = useDeferredValue(search);

  const stats = useMemo(() => {
    let liked = 0;
    let disliked = 0;
    let commented = 0;
    for (const row of rows) {
      if (row.feedback === "like") liked += 1;
      if (row.feedback === "dislike") disliked += 1;
      if (row.comment.trim()) commented += 1;
    }
    return { total: rows.length, liked, disliked, commented, pending: rows.length - liked - disliked };
  }, [rows]);

  const rowIndexes = useMemo(() => new Map(rows.map((row, index) => [row.id, index])), [rows]);

  const visibleRows = useMemo(() => {
    let r = rows;
    if (filter === "like") r = r.filter((x) => x.feedback === "like");
    else if (filter === "dislike") r = r.filter((x) => x.feedback === "dislike");
    else if (filter === "none") r = r.filter((x) => x.feedback === null);
    else if (filter === "commented") r = r.filter((x) => x.comment.trim());
    if (deferredSearch.trim()) {
      const q = deferredSearch.toLowerCase();
      r = r.filter((x) => JSON.stringify(x.data).toLowerCase().includes(q));
    }
    return r;
  }, [rows, filter, deferredSearch]);

  const totalPages = Math.max(1, Math.ceil(visibleRows.length / PAGE_SIZE));
  const activePage = Math.min(page, totalPages);
  const pageRows = useMemo(
    () => visibleRows.slice((activePage - 1) * PAGE_SIZE, activePage * PAGE_SIZE),
    [visibleRows, activePage],
  );

  async function handleFile(e: ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0];
    if (!f) return;
    setBusy("import");
    setFileError("");
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    try {
      const importedRows = /\.xlsx$/i.test(f.name)
        ? await parseWorkbook(f)
        : parseInput(await f.text()).map((data) => ({ id: uid(), data, feedback: null, comment: "" }));
      if (!importedRows.length) throw new Error("No valid data rows were found in this file.");
      setRows(importedRows);
      setFileName(f.name);
      setPage(1);
    } catch (error) {
      setFileError(error instanceof Error ? error.message : "This file could not be imported.");
    } finally {
      setBusy(null);
    }
    e.target.value = "";
  }

  function loadSample() {
    const sample = [
      {
        instruction: "Explain photosynthesis in one sentence.",
        input: "",
        output:
          "Photosynthesis is the process by which green plants convert sunlight, water, and carbon dioxide into glucose and oxygen.",
      },
      {
        instruction: "Translate to French.",
        input: "Good morning, how are you?",
        output: "Bonjour, comment allez-vous ?",
      },
      {
        instruction: "Classify sentiment.",
        input: "The product broke after two days. Terrible experience.",
        output: "negative",
      },
      {
        instruction: "Summarize the paragraph.",
        input:
          "Large language models are neural networks trained on vast text corpora. They can generate coherent responses across many domains.",
        output:
          "LLMs are text-trained neural networks capable of generating coherent multi-domain responses.",
      },
    ];
    setRows(sample.map((d) => ({ id: uid(), data: d, feedback: null, comment: "" })));
    setFileName("sample-dataset.json");
  }

  function setFeedback(id: string, fb: Feedback) {
    setRows((rs) => rs.map((r) => (r.id === id ? { ...r, feedback: r.feedback === fb ? null : fb } : r)));
  }

  function updateComment(id: string, comment: string) {
    setRows((rs) => rs.map((r) => (r.id === id ? { ...r, comment } : r)));
  }

  function deleteRow(id: string) {
    setRows((rs) => rs.filter((r) => r.id !== id));
  }

  function startEdit(row: Row) {
    setEditingId(row.id);
    setEditDraft(JSON.stringify(row.data, null, 2));
    setEditError("");
  }

  function saveEdit() {
    try {
      const parsed = JSON.parse(editDraft);
      if (!isObj(parsed)) throw new Error("Must be a JSON object");
      setRows((rs) => rs.map((r) => (r.id === editingId ? { ...r, data: parsed } : r)));
      setEditingId(null);
      setEditDraft("");
      setEditError("");
    } catch (e) {
      setEditError(e instanceof Error ? e.message : "Invalid JSON");
    }
  }

  function addRow() {
    const template = rows[0]?.data
      ? Object.fromEntries(Object.keys(rows[0].data).map((k) => [k, ""]))
      : { instruction: "", input: "", output: "" };
    const newRow: Row = { id: uid(), data: template, feedback: null, comment: "" };
    setRows((rs) => [newRow, ...rs]);
    setEditingId(newRow.id);
    setEditDraft(JSON.stringify(template, null, 2));
    setEditError("");
  }

  async function exportXlsx() {
    if (!rows.length) return;
    setBusy("export");
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    const wb = new ExcelJS.Workbook();
    wb.creator = "SynthLab";
    wb.created = new Date();
    const ws = wb.addWorksheet("Dataset");

    // Collect all keys across rows to build columns
    const keySet = new Set<string>();
    rows.forEach((r) => Object.keys(r.data).forEach((k) => keySet.add(k)));
    const dataKeys = Array.from(keySet);
    const columns = [
      { header: "#", key: "__idx", width: 6 },
      ...dataKeys.map((k) => ({ header: k, key: k, width: 40 })),
      { header: "Feedback", key: "__feedback", width: 12 },
      { header: "Comment", key: "__comment", width: 40 },
    ];
    ws.columns = columns;

    // Header style
    const header = ws.getRow(1);
    header.font = { bold: true, color: { argb: "FF111827" } };
    header.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF3F4F6" },
    };
    header.alignment = { vertical: "middle", horizontal: "left" };
    header.height = 22;
    header.eachCell((c) => {
      c.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
    });

    rows.forEach((r, i) => {
      const rowData: Record<string, unknown> = {
        __idx: i + 1,
        __feedback: r.feedback ?? "",
        __comment: r.comment,
      };
      dataKeys.forEach((k) => {
        const v = r.data[k];
        rowData[k] =
          v == null
            ? ""
            : typeof v === "object"
              ? JSON.stringify(v)
              : (v as string | number | boolean);
      });
      const excelRow = ws.addRow(rowData);
      excelRow.alignment = { vertical: "top", wrapText: true };

      if (r.feedback === "dislike") {
        excelRow.eachCell((c) => {
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" },
          };
          c.font = { color: { argb: "FFFFFFFF" }, bold: true };
        });
      } else if (r.feedback === "like") {
        excelRow.eachCell((c) => {
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFECFDF5" },
          };
        });
      }
    });

    ws.views = [{ state: "frozen", ySplit: 1 }];

    try {
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const base = fileName.replace(/\.(jsonl?|txt|xlsx)$/i, "") || "dataset";
      a.download = `${base}-reviewed.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } finally {
      setBusy(null);
    }
  }

  return (
    <div className="min-h-screen bg-[#fafaf9]">
      {/* Header */}
      <header className="sticky top-0 z-20 border-b border-neutral-200 bg-white/80 backdrop-blur-md">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-neutral-900 text-white">
              <Sparkles size={18} />
            </div>
            <div>
              <h1 className="text-base font-semibold tracking-tight text-neutral-900">SynthLab</h1>
              <p className="text-xs text-neutral-500">Synthetic dataset review & export</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => fileRef.current?.click()}
              disabled={busy !== null}
              className="inline-flex items-center gap-2 rounded-lg border border-neutral-300 bg-white px-3 py-2 text-sm font-medium text-neutral-700 shadow-sm transition hover:border-neutral-400 hover:bg-neutral-50"
            >
              {busy === "import" ? <LoaderCircle className="animate-spin" size={14} /> : <Upload size={14} />} Upload
            </button>
            <input
              ref={fileRef}
              type="file"
              accept=".json,.jsonl,.txt,.xlsx,application/json,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              className="hidden"
              onChange={handleFile}
            />
            <button
              onClick={exportXlsx}
              disabled={!rows.length || busy !== null}
              className="inline-flex items-center gap-2 rounded-lg bg-neutral-900 px-3 py-2 text-sm font-medium text-white shadow-sm transition hover:bg-neutral-800 disabled:cursor-not-allowed disabled:opacity-40"
            >
              {busy === "export" ? <LoaderCircle className="animate-spin" size={14} /> : <Download size={14} />} {busy === "export" ? "Preparing…" : "Export XLSX"}
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl px-6 py-8">
        {fileError && (
          <div role="alert" className="mb-4 flex items-center justify-between rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-800">
            <span>{fileError}</span>
            <button type="button" onClick={() => setFileError("")} aria-label="Dismiss error"><X size={15} /></button>
          </div>
        )}
        {/* Empty state */}
        {rows.length === 0 ? (
          <EmptyState
            onPick={() => fileRef.current?.click()}
            onSample={loadSample}
            fileName={fileName}
          />
        ) : (
          <>
            {/* Stats */}
            <section className="mb-6 grid grid-cols-2 gap-3 md:grid-cols-5">
              <StatCard label="Total samples" value={stats.total} accent="neutral" icon={<Database size={14} />} />
              <StatCard label="Liked" value={stats.liked} accent="emerald" icon={<ThumbsUp size={14} />} />
              <StatCard label="Disliked" value={stats.disliked} accent="rose" icon={<ThumbsDown size={14} />} />
              <StatCard label="Commented" value={stats.commented} accent="amber" icon={<MessageSquare size={14} />} />
              <StatCard label="Pending" value={stats.pending} accent="slate" icon={<Sparkles size={14} />} />
            </section>

            {/* Toolbar */}
            <section className="mb-4 flex flex-wrap items-center justify-between gap-3 rounded-xl border border-neutral-200 bg-white p-3 shadow-sm">
              <div className="flex flex-1 items-center gap-2">
                <div className="relative flex-1 max-w-md">
                  <Search
                    size={14}
                    className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400"
                  />
                  <input
                    value={search}
                    onChange={(e) => { setSearch(e.target.value); setPage(1); }}
                    placeholder="Search across all fields..."
                    className="w-full rounded-lg border border-neutral-200 bg-neutral-50 py-2 pl-9 pr-3 text-sm text-neutral-800 placeholder:text-neutral-400 focus:border-neutral-400 focus:bg-white focus:outline-none"
                  />
                </div>
                <div className="flex items-center gap-1 rounded-lg border border-neutral-200 bg-neutral-50 p-1 text-xs">
                  {(
                    [
                      ["all", "All"],
                      ["like", "Liked"],
                      ["dislike", "Disliked"],
                      ["none", "Unrated"],
                      ["commented", "Commented"],
                    ] as const
                  ).map(([k, label]) => (
                    <button
                      key={k}
                      type="button"
                      onClick={() => { setFilter(k); setPage(1); }}
                      className={`rounded-md px-2.5 py-1.5 font-medium transition ${
                        filter === k
                          ? "bg-white text-neutral-900 shadow-sm"
                          : "text-neutral-500 hover:text-neutral-800"
                      }`}
                    >
                      {label}
                    </button>
                  ))}
                </div>
              </div>
              <div className="flex items-center gap-2">
                {fileName ? (
                  <span className="inline-flex items-center gap-1.5 rounded-md bg-neutral-100 px-2 py-1 text-xs text-neutral-600">
                    <FileJson size={12} /> {fileName}
                  </span>
                ) : null}
                <button
                  onClick={addRow}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-neutral-300 bg-white px-3 py-1.5 text-xs font-medium text-neutral-700 shadow-sm transition hover:border-neutral-400 hover:bg-neutral-50"
                >
                  <Plus size={14} /> Add sample
                </button>
              </div>
            </section>

            {/* Rows */}
            <section className="space-y-3">
              {visibleRows.length === 0 ? (
                <div className="rounded-xl border border-dashed border-neutral-300 bg-white p-10 text-center text-sm text-neutral-500">
                  No samples match this filter.
                </div>
              ) : (
                pageRows.map((row) => {
                  const realIdx = rowIndexes.get(row.id) ?? 0;
                  return (
                    <RowCard
                      key={row.id}
                      row={row}
                      index={realIdx}
                      isEditing={editingId === row.id}
                      editDraft={editDraft}
                      editError={editError}
                      showComment={commentingId === row.id}
                      onEdit={() => startEdit(row)}
                      onCancelEdit={() => {
                        setEditingId(null);
                        setEditError("");
                      }}
                      onSaveEdit={saveEdit}
                      onEditDraftChange={setEditDraft}
                      onDelete={() => deleteRow(row.id)}
                      onLike={() => setFeedback(row.id, "like")}
                      onDislike={() => setFeedback(row.id, "dislike")}
                      onToggleComment={() =>
                        setCommentingId((c) => (c === row.id ? null : row.id))
                      }
                      onCommentSave={(v) => updateComment(row.id, v)}
                    />
                  );
                })
              )}
            </section>
            {visibleRows.length > PAGE_SIZE && (
              <nav aria-label="Dataset pages" className="mt-5 flex items-center justify-between border-t border-neutral-200 pt-4">
                <p className="text-xs text-neutral-500">
                  Showing {(activePage - 1) * PAGE_SIZE + 1}–{Math.min(activePage * PAGE_SIZE, visibleRows.length)} of {visibleRows.length.toLocaleString()}
                </p>
                <div className="flex items-center gap-2">
                  <button type="button" aria-label="Previous page" disabled={activePage === 1} onClick={() => setPage((value) => Math.max(1, value - 1))} className="rounded-lg border border-neutral-300 bg-white p-2 text-neutral-700 disabled:opacity-40"><ChevronLeft size={15} /></button>
                  <span className="min-w-20 text-center text-xs font-medium text-neutral-700">{activePage} / {totalPages}</span>
                  <button type="button" aria-label="Next page" disabled={activePage === totalPages} onClick={() => setPage((value) => Math.min(totalPages, value + 1))} className="rounded-lg border border-neutral-300 bg-white p-2 text-neutral-700 disabled:opacity-40"><ChevronRight size={15} /></button>
                </div>
              </nav>
            )}
          </>
        )}
      </main>

      <footer className="mx-auto max-w-7xl px-6 py-10 text-center text-xs text-neutral-400">
        SynthLab · Curate. Review. Export. — built for LLM fine-tuning datasets.
      </footer>
    </div>
  );
}

function EmptyState({
  onPick,
  onSample,
  fileName,
}: {
  onPick: () => void;
  onSample: () => void;
  fileName: string;
}) {
  return (
    <div className="mx-auto mt-8 max-w-2xl rounded-2xl border border-neutral-200 bg-white p-10 text-center shadow-sm">
      <div className="mx-auto mb-5 flex h-14 w-14 items-center justify-center rounded-2xl bg-neutral-900 text-white">
        <FileJson size={26} />
      </div>
      <h2 className="text-xl font-semibold tracking-tight text-neutral-900">
        Upload your synthetic dataset
      </h2>
      <p className="mx-auto mt-2 max-w-md text-sm text-neutral-500">
        Drop in a <code className="rounded bg-neutral-100 px-1 py-0.5 text-xs">.json</code>,
        <code className="rounded bg-neutral-100 px-1 py-0.5 text-xs">.jsonl</code>, or
        <code className="rounded bg-emerald-50 px-1 py-0.5 text-xs text-emerald-700">.xlsx</code> file of samples.
        Review each row with likes, dislikes, and comments — then export a formatted XLSX for your
        fine-tuning pipeline.
      </p>
      <div className="mt-6 flex items-center justify-center gap-2">
        <button
          onClick={onPick}
          className="inline-flex items-center gap-2 rounded-lg bg-neutral-900 px-4 py-2.5 text-sm font-medium text-white shadow-sm transition hover:bg-neutral-800"
        >
          <FileSpreadsheet size={14} /> Choose JSON or Excel
        </button>
        <button
          onClick={onSample}
          className="inline-flex items-center gap-2 rounded-lg border border-neutral-300 bg-white px-4 py-2.5 text-sm font-medium text-neutral-700 shadow-sm transition hover:bg-neutral-50"
        >
          <Sparkles size={14} /> Try sample data
        </button>
      </div>
      {fileName && <p className="mt-4 text-xs text-neutral-400">Last file: {fileName}</p>}
    </div>
  );
}

function StatCard({
  label,
  value,
  accent,
  icon,
}: {
  label: string;
  value: number;
  accent: "neutral" | "emerald" | "rose" | "amber" | "slate";
  icon: React.ReactNode;
}) {
  const accents: Record<string, string> = {
    neutral: "bg-neutral-900 text-white",
    emerald: "bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200",
    rose: "bg-rose-50 text-rose-700 ring-1 ring-rose-200",
    amber: "bg-amber-50 text-amber-700 ring-1 ring-amber-200",
    slate: "bg-slate-50 text-slate-700 ring-1 ring-slate-200",
  };
  return (
    <div className="rounded-xl border border-neutral-200 bg-white p-4 shadow-sm">
      <div className="flex items-center justify-between">
        <span className="text-xs font-medium uppercase tracking-wider text-neutral-500">
          {label}
        </span>
        <span className={`inline-flex h-6 w-6 items-center justify-center rounded-md ${accents[accent]}`}>
          {icon}
        </span>
      </div>
      <div className="mt-2 text-2xl font-semibold tracking-tight text-neutral-900 tabular-nums">
        {value}
      </div>
    </div>
  );
}

function RowCard({
  row,
  index,
  isEditing,
  editDraft,
  editError,
  showComment,
  onEdit,
  onCancelEdit,
  onSaveEdit,
  onEditDraftChange,
  onDelete,
  onLike,
  onDislike,
  onToggleComment,
  onCommentSave,
}: {
  row: Row;
  index: number;
  isEditing: boolean;
  editDraft: string;
  editError: string;
  showComment: boolean;
  onEdit: () => void;
  onCancelEdit: () => void;
  onSaveEdit: () => void;
  onEditDraftChange: (v: string) => void;
  onDelete: () => void;
  onLike: () => void;
  onDislike: () => void;
  onToggleComment: () => void;
  onCommentSave: (v: string) => void;
}) {
  const [commentDraft, setCommentDraft] = useState(row.comment);
  const feedbackRing =
    row.feedback === "like"
      ? "ring-emerald-200 border-emerald-200"
      : row.feedback === "dislike"
        ? "ring-rose-200 border-rose-200"
        : "border-neutral-200";
  const stripe =
    row.feedback === "like"
      ? "bg-emerald-400"
      : row.feedback === "dislike"
        ? "bg-rose-400"
        : "bg-neutral-200";

  return (
    <article
      className={`relative overflow-hidden rounded-xl border bg-white shadow-sm transition ${feedbackRing}`}
    >
      <div className={`absolute inset-y-0 left-0 w-1 ${stripe}`} />
      <div className="grid grid-cols-1 md:grid-cols-[1fr_240px]">
        {/* Data side */}
        <div className="min-w-0 p-5 pl-6">
          <div className="mb-3 flex items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <span className="inline-flex h-6 min-w-[2rem] items-center justify-center rounded-md bg-neutral-100 px-2 text-xs font-semibold tabular-nums text-neutral-600">
                #{index + 1}
              </span>
              {row.feedback === "like" && (
                <span className="inline-flex items-center gap-1 rounded-md bg-emerald-50 px-2 py-0.5 text-xs font-medium text-emerald-700">
                  <ThumbsUp size={11} /> liked
                </span>
              )}
              {row.feedback === "dislike" && (
                <span className="inline-flex items-center gap-1 rounded-md bg-rose-50 px-2 py-0.5 text-xs font-medium text-rose-700">
                  <ThumbsDown size={11} /> disliked
                </span>
              )}
              {row.comment.trim() && (
                <span className="inline-flex items-center gap-1 rounded-md bg-amber-50 px-2 py-0.5 text-xs font-medium text-amber-700">
                  <MessageSquare size={11} /> note
                </span>
              )}
            </div>
            <div className="flex items-center gap-1">
              {!isEditing ? (
                <>
                  <button
                    onClick={onEdit}
                    className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs text-neutral-500 transition hover:bg-neutral-100 hover:text-neutral-800"
                  >
                    <Pencil size={12} /> Edit
                  </button>
                  <button
                    onClick={onDelete}
                    className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs text-neutral-500 transition hover:bg-rose-50 hover:text-rose-700"
                  >
                    <Trash2 size={12} /> Delete
                  </button>
                </>
              ) : (
                <>
                  <button
                    onClick={onSaveEdit}
                    className="inline-flex items-center gap-1 rounded-md bg-neutral-900 px-2 py-1 text-xs font-medium text-white transition hover:bg-neutral-800"
                  >
                    <Check size={12} /> Save
                  </button>
                  <button
                    onClick={onCancelEdit}
                    className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs text-neutral-500 transition hover:bg-neutral-100"
                  >
                    <X size={12} /> Cancel
                  </button>
                </>
              )}
            </div>
          </div>

          {isEditing ? (
            <div>
              <textarea
                value={editDraft}
                onChange={(e) => onEditDraftChange(e.target.value)}
                spellCheck={false}
                className="min-h-[220px] w-full resize-y rounded-lg border border-neutral-300 bg-neutral-50 p-3 font-mono text-xs text-neutral-800 focus:border-neutral-500 focus:bg-white focus:outline-none"
              />
              {editError && (
                <p className="mt-2 text-xs font-medium text-rose-600">{editError}</p>
              )}
            </div>
          ) : (
            <FieldGrid data={row.data} />
          )}
        </div>

        {/* Side panel */}
        <aside className="flex flex-col justify-between border-t border-neutral-100 bg-neutral-50/60 p-4 md:border-l md:border-t-0">
          <div className="space-y-2">
            <p className="text-[10px] font-semibold uppercase tracking-wider text-neutral-500">
              Review
            </p>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onLike}
                aria-pressed={row.feedback === "like"}
                className={`flex-1 inline-flex items-center justify-center gap-1.5 rounded-lg border px-2.5 py-2 text-xs font-medium transition ${
                  row.feedback === "like"
                    ? "border-emerald-500 bg-emerald-500 text-white shadow-sm"
                    : "border-neutral-200 bg-white text-neutral-700 hover:border-emerald-300 hover:text-emerald-700"
                }`}
              >
                <ThumbsUp size={13} /> Like
              </button>
              <button
                type="button"
                onClick={onDislike}
                aria-pressed={row.feedback === "dislike"}
                className={`flex-1 inline-flex items-center justify-center gap-1.5 rounded-lg border px-2.5 py-2 text-xs font-medium transition ${
                  row.feedback === "dislike"
                    ? "border-rose-500 bg-rose-500 text-white shadow-sm"
                    : "border-neutral-200 bg-white text-neutral-700 hover:border-rose-300 hover:text-rose-700"
                }`}
              >
                <ThumbsDown size={13} /> Dislike
              </button>
            </div>
            <button
              onClick={onToggleComment}
              className={`w-full inline-flex items-center justify-center gap-1.5 rounded-lg border px-2.5 py-2 text-xs font-medium transition ${
                showComment || row.comment
                  ? "border-amber-300 bg-amber-50 text-amber-800"
                  : "border-neutral-200 bg-white text-neutral-700 hover:border-amber-300 hover:text-amber-700"
              }`}
            >
              <MessageSquare size={13} /> {row.comment ? "Edit comment" : "Add comment"}
            </button>

            {(showComment || row.comment) && (
              <textarea
                value={commentDraft}
                onChange={(e) => setCommentDraft(e.target.value)}
                onBlur={() => {
                  if (commentDraft !== row.comment) onCommentSave(commentDraft);
                }}
                placeholder="Notes for reviewers, quality flags, edge cases..."
                className="mt-1 min-h-[80px] w-full resize-y rounded-lg border border-neutral-200 bg-white p-2 text-xs text-neutral-800 placeholder:text-neutral-400 focus:border-neutral-400 focus:outline-none"
              />
            )}
          </div>
        </aside>
      </div>
    </article>
  );
}

function FieldGrid({ data }: { data: Record<string, unknown> }) {
  const entries = Object.entries(data);
  if (!entries.length) {
    return <p className="text-xs italic text-neutral-400">Empty object</p>;
  }

  // Case 1: single {role, content} message
  if (entries.length <= 3 && isChatMessage(data)) {
    return <ChatBubble role={data.role} content={data.content} />;
  }

  // Case 2: object contains a messages array of chat messages
  const messagesEntry = entries.find(
    ([, v]) => Array.isArray(v) && v.length > 0 && (v as unknown[]).every(isChatMessage),
  );
  if (messagesEntry) {
    const [msgKey, msgs] = messagesEntry as [string, { role: string; content: string }[]];
    const others = entries.filter(([k]) => k !== msgKey);
    return (
      <div className="space-y-3">
        {others.length > 0 && (
          <div className="divide-y divide-neutral-100 rounded-lg border border-neutral-100 bg-neutral-50/40">
            {others.map(([k, v]) => (
              <FieldRow key={k} k={k} v={v} />
            ))}
          </div>
        )}
        <div className="space-y-2">
          <div className="flex items-center gap-2 px-1">
            <span className="text-[10px] font-semibold uppercase tracking-wider text-neutral-500">
              {msgKey}
            </span>
            <span className="text-[10px] text-neutral-400">
              {msgs.length} message{msgs.length === 1 ? "" : "s"}
            </span>
          </div>
          <div className="space-y-2">
            {msgs.map((m, i) => (
              <ChatBubble key={i} role={m.role} content={m.content} />
            ))}
          </div>
        </div>
      </div>
    );
  }

  // Case 3: QA fine-tuning record (question_text / answer / cot / chunk_content)
  if (typeof data.question_text === "string" || typeof data.answer === "string") {
    return <QARecordView data={data} />;
  }

  // Default: key/value grid
  return (
    <div className="divide-y divide-neutral-100 rounded-lg border border-neutral-100 bg-neutral-50/40">
      {entries.map(([k, v]) => (
        <FieldRow key={k} k={k} v={v} />
      ))}
    </div>
  );
}

const QA_MAIN_KEYS = ["question_text", "answer", "cot", "chunk_content"];
const QA_META_ORDER = [
  "questionId",
  "source_doc",
  "source_subject",
  "section",
  "chunkName",
  "chunk_page_range",
  "quality_tier",
  "quality_score",
  "citation_state",
  "cited_manual",
  "cited_page",
  "citation_in_range",
  "defect_class",
  "tags",
  "score",
  "provenance_exact",
  "held_out",
  "defunct",
  "abstention",
];

function parseTags(v: unknown): string[] {
  if (Array.isArray(v)) return v.map(String);
  if (typeof v === "string" && v.trim()) {
    try {
      const p = JSON.parse(v);
      if (Array.isArray(p)) return p.map(String);
    } catch {
      return v.split(",").map((s) => s.trim()).filter(Boolean);
    }
  }
  return [];
}

function Chip({
  label,
  value,
  tone = "neutral",
}: {
  label: string;
  value: string;
  tone?: "neutral" | "green" | "amber" | "rose" | "sky" | "violet";
}) {
  const tones = {
    neutral: "border-neutral-200 bg-white text-neutral-700",
    green: "border-emerald-200 bg-emerald-50 text-emerald-800",
    amber: "border-amber-200 bg-amber-50 text-amber-800",
    rose: "border-rose-200 bg-rose-50 text-rose-800",
    sky: "border-sky-200 bg-sky-50 text-sky-800",
    violet: "border-violet-200 bg-violet-50 text-violet-800",
  } as const;
  return (
    <span
      className={`inline-flex max-w-full items-center gap-1.5 rounded-md border px-2 py-0.5 text-[11px] ${tones[tone]}`}
      title={`${label}: ${value}`}
    >
      <span className="font-semibold uppercase tracking-wider opacity-60">{label}</span>
      <span className="truncate font-medium">{value}</span>
    </span>
  );
}

function Section({
  title,
  tone,
  children,
  meta,
  collapsible = false,
  defaultOpen = true,
}: {
  title: string;
  tone: "sky" | "emerald" | "violet" | "neutral";
  children: ReactNode;
  meta?: string;
  collapsible?: boolean;
  defaultOpen?: boolean;
}) {
  const [open, setOpen] = useState(defaultOpen);
  const tones = {
    sky: { wrap: "border-sky-200 bg-sky-50/50", bar: "bg-sky-400", badge: "bg-sky-600 text-white" },
    emerald: {
      wrap: "border-emerald-200 bg-emerald-50/40",
      bar: "bg-emerald-400",
      badge: "bg-emerald-600 text-white",
    },
    violet: {
      wrap: "border-violet-200 bg-violet-50/40",
      bar: "bg-violet-400",
      badge: "bg-violet-600 text-white",
    },
    neutral: {
      wrap: "border-neutral-200 bg-neutral-50",
      bar: "bg-neutral-400",
      badge: "bg-neutral-700 text-white",
    },
  }[tone];

  return (
    <div className={`relative overflow-hidden rounded-lg border ${tones.wrap}`}>
      <div className={`absolute inset-y-0 left-0 w-1 ${tones.bar}`} />
      <div className="pl-3">
        <div className="flex items-center justify-between gap-2 border-b border-black/5 px-3 py-1.5">
          <span
            className={`inline-flex items-center rounded-md px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider ${tones.badge}`}
          >
            {title}
          </span>
          <div className="flex items-center gap-2">
            {meta && <span className="text-[10px] text-neutral-500">{meta}</span>}
            {collapsible && (
              <button
                type="button"
                onClick={() => setOpen((o) => !o)}
                className="rounded-md border border-neutral-200 bg-white px-2 py-0.5 text-[10px] font-medium text-neutral-600 hover:text-neutral-900"
              >
                {open ? "Hide" : "Show"}
              </button>
            )}
          </div>
        </div>
        {(!collapsible || open) && <div className="px-3 py-3">{children}</div>}
      </div>
    </div>
  );
}

function HtmlTableView({ html }: { html: string }) {
  const tables = useMemo(() => {
    const out: { headers: string[]; rows: string[][] }[] = [];
    const decode = (s: string) =>
      s
        .replace(/<[^>]+>/g, "")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&nbsp;/g, " ")
        .replace(/&quot;/g, '"')
        .trim();
    for (const tableMatch of html.matchAll(/<table[\s\S]*?<\/table>/gi)) {
      const t = tableMatch[0];
      const rows: string[][] = [];
      let headers: string[] = [];
      for (const trMatch of t.matchAll(/<tr[\s\S]*?<\/tr>/gi)) {
        const tr = trMatch[0];
        const isHead = /<th[\s>]/i.test(tr);
        const cells = [...tr.matchAll(/<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi)].map((c) =>
          decode(c[1] ?? ""),
        );
        if (!cells.length) continue;
        if (isHead && !headers.length) headers = cells;
        else rows.push(cells);
      }
      if (rows.length || headers.length) out.push({ headers, rows });
    }
    return out;
  }, [html]);

  const plain = useMemo(() => html.replace(/<table[\s\S]*?<\/table>/gi, "").trim(), [html]);

  return (
    <div className="space-y-3">
      {plain && (
        <div className="whitespace-pre-wrap break-words text-[13px] leading-relaxed text-neutral-800">
          {plain}
        </div>
      )}
      {tables.map((t, ti) => (
        <div key={ti} className="overflow-x-auto rounded-lg border border-neutral-200 bg-white">
          <table className="w-full border-collapse text-[12.5px]">
            {t.headers.length > 0 && (
              <thead>
                <tr className="bg-neutral-100">
                  {t.headers.map((h, i) => (
                    <th
                      key={i}
                      className="border-b border-neutral-200 px-3 py-1.5 text-left font-semibold text-neutral-700"
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
            )}
            <tbody>
              {t.rows.map((r, ri) => (
                <tr key={ri} className={ri % 2 ? "bg-neutral-50/60" : "bg-white"}>
                  {r.map((c, ci) => (
                    <td
                      key={ci}
                      className="border-b border-neutral-100 px-3 py-1.5 align-top text-neutral-800"
                    >
                      {c}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}

function RichText({ text, className }: { text: string; className?: string }) {
  const blocks = useMemo(() => {
    const lines = text.replace(/\u00a0/g, " ").split(/\n/);
    const out: { type: "p" | "li"; content: string }[] = [];
    for (const raw of lines) {
      const line = raw.replace(/[\u2011\u2013\u2014]/g, "-").trimEnd();
      if (!line.trim()) continue;
      const bullet = line.match(/^\s*[-*•]\s+(.*)$/);
      if (bullet) out.push({ type: "li", content: bullet[1] ?? "" });
      else out.push({ type: "p", content: line.replace(/^#{1,6}\s*/, "") });
    }
    return out;
  }, [text]);

  const inline = (s: string, key: number) => (
    <span key={key}>
      {s.split(/(\*\*[^*]+\*\*)/g).map((part, i) =>
        part.startsWith("**") && part.endsWith("**") ? (
          <strong key={i} className="font-semibold text-neutral-900">
            {part.slice(2, -2)}
          </strong>
        ) : (
          <span key={i}>{part}</span>
        ),
      )}
    </span>
  );

  return (
    <div className={`space-y-1.5 ${className ?? ""}`}>
      {blocks.map((b, i) =>
        b.type === "li" ? (
          <div key={i} className="flex gap-2">
            <span className="select-none text-neutral-400">•</span>
            <span className="min-w-0 break-words">{inline(b.content, i)}</span>
          </div>
        ) : (
          <p key={i} className="break-words">
            {inline(b.content, i)}
          </p>
        ),
      )}
    </div>
  );
}

function QARecordView({ data }: { data: Record<string, unknown> }) {
  const question = typeof data.question_text === "string" ? data.question_text : "";
  const answer = typeof data.answer === "string" ? data.answer : "";
  const cot = typeof data.cot === "string" ? data.cot.trim() : "";
  const chunk = typeof data.chunk_content === "string" ? data.chunk_content : "";
  const tags = parseTags(data.tags);
  const tier = typeof data.quality_tier === "string" ? data.quality_tier : "";
  const defects = typeof data.defect_class === "string" && data.defect_class.trim()
    ? data.defect_class.split(",").map((d) => d.trim()).filter(Boolean)
    : [];
  const extras = Object.entries(data).filter(
    ([k]) => !QA_MAIN_KEYS.includes(k) && !QA_META_ORDER.includes(k),
  );
  const metaKeys = QA_META_ORDER.filter(
    (k) => k in data && !["tags", "defect_class", "quality_tier", "quality_score"].includes(k),
  );

  return (
    <div className="space-y-3">
      {/* Metadata header */}
      <div className="rounded-lg border border-neutral-200 bg-neutral-50/70 p-3">
        <div className="flex flex-wrap items-center gap-1.5">
          {tier && (
            <Chip
              label="tier"
              value={tier}
              tone={tier === "clean" ? "green" : tier === "warn" ? "amber" : "rose"}
            />
          )}
          {"quality_score" in data && (
            <Chip
              label="score"
              value={String(data.quality_score)}
              tone={Number(data.quality_score) >= 8 ? "green" : Number(data.quality_score) >= 5 ? "amber" : "rose"}
            />
          )}
          {typeof data.citation_state === "string" && (
            <Chip
              label="citation"
              value={data.citation_state}
              tone={data.citation_state === "ok" ? "green" : "rose"}
            />
          )}
          {tags.map((t) => (
            <span
              key={t}
              className="rounded-md border border-violet-200 bg-violet-50 px-2 py-0.5 text-[11px] font-medium text-violet-700"
            >
              #{t}
            </span>
          ))}
          {defects.map((d) => (
            <span
              key={d}
              className="rounded-md border border-rose-200 bg-rose-50 px-2 py-0.5 text-[11px] font-medium text-rose-700"
            >
              ⚠ {d}
            </span>
          ))}
        </div>
        <div className="mt-2 grid grid-cols-1 gap-x-4 gap-y-1 sm:grid-cols-2 lg:grid-cols-3">
          {metaKeys.map((k) => (
            <div key={k} className="flex min-w-0 items-baseline gap-2 text-[12px]">
              <span className="shrink-0 font-semibold uppercase tracking-wider text-neutral-400">
                {k.replace(/_/g, " ")}
              </span>
              <span className="truncate font-mono text-neutral-700" title={renderValue(data[k])}>
                {renderValue(data[k])}
              </span>
            </div>
          ))}
        </div>
      </div>

      <Section title="question" tone="sky" meta={`${question.length.toLocaleString()} chars`}>
        {question ? (
          <RichText text={question} className="text-[14px] font-medium leading-relaxed text-neutral-900" />
        ) : (
          <span className="text-sm italic text-neutral-400">(empty)</span>
        )}
      </Section>

      <Section title="answer" tone="emerald" meta={`${answer.length.toLocaleString()} chars`}>
        {answer ? (
          <RichText text={answer} className="text-[13.5px] leading-relaxed text-neutral-900" />
        ) : (
          <span className="text-sm italic text-neutral-400">(empty)</span>
        )}
      </Section>

      {cot && (
        <Section
          title="chain of thought"
          tone="violet"
          meta={`${cot.length.toLocaleString()} chars`}
          collapsible
          defaultOpen={false}
        >
          <RichText text={cot} className="text-[13px] leading-relaxed text-neutral-600" />
        </Section>
      )}

      {chunk && (
        <Section
          title="source chunk"
          tone="neutral"
          meta={`${chunk.length.toLocaleString()} chars`}
          collapsible
          defaultOpen={false}
        >
          <HtmlTableView html={chunk} />
        </Section>
      )}

      {extras.length > 0 && (
        <div className="divide-y divide-neutral-100 rounded-lg border border-neutral-100 bg-neutral-50/40">
          {extras.map(([k, v]) => (
            <FieldRow key={k} k={k} v={v} />
          ))}
        </div>
      )}
    </div>
  );
}


function FieldRow({ k, v }: { k: string; v: unknown }) {
  return (
    <div className="grid grid-cols-[140px_1fr] gap-3 p-3 text-sm">
      <div className="min-w-0">
        <span className="inline-block truncate rounded-md bg-white px-2 py-0.5 text-[11px] font-semibold uppercase tracking-wider text-neutral-500 ring-1 ring-neutral-200">
          {k}
        </span>
      </div>
      <div className="min-w-0 whitespace-pre-wrap break-words font-mono text-[13px] leading-relaxed text-neutral-800">
        {renderValue(v)}
      </div>
    </div>
  );
}

function ChatBubble({ role, content }: { role: string; content: string }) {
  const { think, body } = extractThink(content);
  const roleLower = role.toLowerCase();
  const styles =
    roleLower === "user"
      ? {
          wrap: "border-sky-200 bg-sky-50/60",
          badge: "bg-sky-600 text-white",
          accent: "bg-sky-400",
        }
      : roleLower === "assistant"
        ? {
            wrap: "border-violet-200 bg-violet-50/50",
            badge: "bg-violet-600 text-white",
            accent: "bg-violet-400",
          }
        : roleLower === "system"
          ? {
              wrap: "border-amber-200 bg-amber-50/60",
              badge: "bg-amber-600 text-white",
              accent: "bg-amber-400",
            }
          : {
              wrap: "border-neutral-200 bg-neutral-50",
              badge: "bg-neutral-700 text-white",
              accent: "bg-neutral-400",
            };

  return (
    <div className={`relative overflow-hidden rounded-lg border ${styles.wrap}`}>
      <div className={`absolute inset-y-0 left-0 w-1 ${styles.accent}`} />
      <div className="pl-3">
        <div className="flex items-center justify-between border-b border-black/5 px-3 py-1.5">
          <span
            className={`inline-flex items-center gap-1 rounded-md px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider ${styles.badge}`}
          >
            {role}
          </span>
          <span className="text-[10px] text-neutral-400">
            {content.length.toLocaleString()} chars
          </span>
        </div>
        <div className="space-y-2 px-3 py-3">
          {think !== null && (
            <details className="group rounded-md border border-dashed border-neutral-300 bg-white/70">
              <summary className="cursor-pointer select-none px-2.5 py-1.5 text-[11px] font-medium uppercase tracking-wider text-neutral-500 hover:text-neutral-700">
                Reasoning ({think.length.toLocaleString()} chars)
              </summary>
              <div className="whitespace-pre-wrap break-words border-t border-neutral-200 px-3 py-2 font-mono text-[12px] leading-relaxed text-neutral-500">
                {think}
              </div>
            </details>
          )}
          <div className="whitespace-pre-wrap break-words text-[13.5px] leading-relaxed text-neutral-900">
            {body || <span className="italic text-neutral-400">(empty)</span>}
          </div>
        </div>
      </div>
    </div>
  );
}

function renderValue(v: unknown): string {
  if (v === null || v === undefined) return "—";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  try {
    return JSON.stringify(v, null, 2);
  } catch {
    return String(v);
  }
}
