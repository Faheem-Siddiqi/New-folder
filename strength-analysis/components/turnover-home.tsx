"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ChevronRight, Download, FileSpreadsheet, Plus, Search, Upload, X } from "lucide-react";
import { Button } from "./ui/button";
import { DesignationEmployeesDialog } from "./designation-employees-dialog";
import { FeedbackDialog, type Feedback } from "./feedback-dialog";
import { ProgressDialog, type Progress } from "./progress-dialog";
import { loadReport, storeReport } from "@/lib/report-storage";
import { MAX_FILE_SIZE, TURNOVER_HEADERS, validateUpload, type TurnoverResult } from "@/lib/turnover-contracts";
import type { StrengthData } from "@/lib/strength-data";
import { cn } from "@/lib/utils";

type WorkerMessage = { type: "progress"; step: string; percent: number } | { type: "error"; message: string } | { type: "complete"; result?: TurnoverResult; buffer?: ArrayBuffer };

export function TurnoverHome({ template }: { template: StrengthData }) {
  const [employeeRowIndex, setEmployeeRowIndex] = useState<number | null>(null);
  const [report, setReport] = useState<TurnoverResult | null>(null);
  const [restoring, setRestoring] = useState(true);
  const [storageStatus, setStorageStatus] = useState("");
  const [category, setCategory] = useState(template.strengthStructure[0]?.category ?? "");
  const [search, setSearch] = useState("");
  const [progress, setProgress] = useState<Progress | null>(null);
  const [feedback, setFeedback] = useState<Feedback | null>(null);
  const [uploadOpen, setUploadOpen] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [uploadError, setUploadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [issuesOpen, setIssuesOpen] = useState(false);
  const [issuePage, setIssuePage] = useState(0);
  const [storedFile, setStoredFile] = useState<{ file: Blob; fileName: string } | null>(null);
  const uploadRef = useRef<HTMLDialogElement>(null);
  const issuesRef = useRef<HTMLDialogElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const activeJob = useRef<{ worker: Worker; cancel: () => void } | null>(null);
  const busyRef = useRef(false);
  const committing = useRef(false);
  const cancelled = useRef(false);

  const runWorker = useCallback((worker: Worker, payload: object, title: string, transfer: Transferable[] = []) => new Promise<WorkerMessage>((resolve, reject) => {
    const finish = () => { clearTimeout(timeout); worker.terminate(); activeJob.current = null; };
    const cancel = () => { finish(); reject(new DOMException("Processing cancelled", "AbortError")); };
    const timeout = setTimeout(() => { finish(); reject(new Error("Processing took too long. Try a smaller workbook or save a new Excel copy.")); }, 180_000);
    activeJob.current = { worker, cancel };
    worker.onmessage = (event: MessageEvent<WorkerMessage>) => {
      if (event.data.type === "progress") setProgress({ title, step: event.data.step, percent: event.data.percent });
      else if (event.data.type === "error") { finish(); reject(new Error(event.data.message)); }
      else { finish(); resolve(event.data); }
    };
    worker.onerror = () => { finish(); reject(new Error("The processing worker could not run. Refresh the page and try again.")); };
    worker.onmessageerror = () => { finish(); reject(new Error("The workbook results could not be read. Try a smaller file.")); };
    worker.postMessage(payload, transfer);
  }), []);

  const processFile = useCallback(async (source: Blob, fileName: string, restored = false) => {
    if (busyRef.current) return;
    busyRef.current = true;
    cancelled.current = false;
    setProgress({ title: "Scanning turnover report", step: "Loading uploaded workbook", percent: 2 });
    try {
      validateUpload(fileName, source.size);
      const buffer = await source.arrayBuffer();
      if (cancelled.current) throw new DOMException("Processing cancelled", "AbortError");
      const message = await runWorker(new Worker(new URL("../lib/turnover.worker.ts", import.meta.url)), { buffer, fileName, template }, "Scanning turnover report", [buffer]);
      if (message.type !== "complete" || !message.result) throw new Error("The scan did not return a report.");
      const next = message.result;
      committing.current = true;
      setProgress({ title: "Scanning turnover report", step: "Saving report for your next visit", percent: 98 });
      let persisted = true;
      try { await storeReport({ version: 1, file: source, fileName, result: next }); }
      catch { persisted = false; }
      setStoredFile({ file: source, fileName });
      setReport(next);
      setSearch("");
      setCategory(template.strengthStructure[0]?.category ?? "");
      setStorageStatus(persisted ? "Saved in this browser. Available after refresh or reopening." : "Browser storage failed. Download now; this new report may be lost after refresh.");
      const warnings = next.issues.length;
      setProgress({ title: "Scanning turnover report", step: "Scan complete", percent: 100 });
      await new Promise((resolve) => setTimeout(resolve, 200));
      setProgress(null);
      setFeedback({ title: restored ? "Saved upload reprocessed" : warnings ? "Report ready — review flagged records" : "Report ready", tone: !persisted || warnings ? "warning" : "success", message: `${next.matched.toLocaleString()} employees matched to the strength template. ${next.unmatched.toLocaleString()} employee records are unmatched or conflicting.\n\n${warnings ? `${warnings.toLocaleString()} validation notes are available under Review issues and in the Excel workbook. ` : ""}${persisted ? "Your report is saved in this browser. Review the category results, then download Excel." : "Browser storage is unavailable or full. Your previous saved report has not been replaced. Download this report before leaving."}` });
    } catch (error) {
      setProgress(null);
      if (!(error instanceof DOMException && error.name === "AbortError")) setFeedback({ title: "Report could not be processed", tone: "error", message: `${error instanceof Error ? error.message : "Please try uploading a new Excel copy."}\n\nYour previous report and Strength JSON template are unchanged.` });
    } finally { busyRef.current = false; committing.current = false; }
  }, [runWorker, template]);

  useEffect(() => {
    let disposed = false;
    loadReport().then((saved) => {
      if (disposed || !saved) return;
      setReport(saved.result);
      setStoredFile({ file: saved.file, fileName: saved.fileName });
      setCategory(saved.result.rows[0]?.category ?? "");
      setStorageStatus("Restored from this browser. Available after refresh or reopening.");
      if (saved.result.templateSignature !== JSON.stringify(template)) void processFile(saved.file, saved.fileName, true);
    }).catch((error) => {
      if (!disposed) { setStorageStatus("Browser storage could not be restored. Upload a workbook to continue."); setFeedback({ title: "Saved report unavailable", tone: "warning", message: `${error instanceof Error ? error.message : "Browser storage could not be accessed."} You can still upload, scan, and download a new report.` }); }
    }).finally(() => { if (!disposed) setRestoring(false); });
    return () => { disposed = true; };
  }, [processFile, template]);

  useEffect(() => () => activeJob.current?.cancel(), []);
  useEffect(() => {
    if (!uploadOpen && !issuesOpen) return;
    const dialog = uploadOpen ? uploadRef.current : issuesRef.current;
    dialog?.showModal();
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { dialog?.close(); document.body.style.overflow = overflow; };
  }, [uploadOpen, issuesOpen]);
  useEffect(() => {
    if (!progress) return;
    const warn = (event: BeforeUnloadEvent) => { event.preventDefault(); };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [progress]);

  const chooseFile = (files: FileList | File[]) => {
    setUploadError("");
    if (files.length !== 1) { setFile(null); setUploadError("Choose one workbook at a time."); setUploadOpen(true); return; }
    try { validateUpload(files[0].name, files[0].size); setFile(files[0]); }
    catch (error) { setFile(null); setUploadError(error instanceof Error ? error.message : "Unsupported file."); }
    setUploadOpen(true);
  };
  const openUpload = () => { setFile(null); setUploadError(""); setUploadOpen(true); };
  const cancelProcessing = () => {
    if (committing.current) return;
    cancelled.current = true;
    if (activeJob.current) activeJob.current.cancel();
    setProgress(null);
    setFeedback({ title: "Processing cancelled", tone: "warning", message: "Your previous report is still available. You can start another scan or download whenever you’re ready." });
  };
  const download = async () => {
    if (!report || busyRef.current) return;
    busyRef.current = true;
    cancelled.current = false;
    setProgress({ title: "Preparing Excel download", step: "Preparing worksheet layout", percent: 5 });
    try {
      let logo: string | undefined;
      try {
        const response = await fetch("/kohinoor-logo.png");
        if (response.ok) {
          const blob = await response.blob();
          logo = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result)); reader.onerror = reject; reader.readAsDataURL(blob); });
        }
      } catch { /* The workbook retains text branding if the logo is unavailable. */ }
      if (cancelled.current) throw new DOMException("Processing cancelled", "AbortError");
      const message = await runWorker(new Worker(new URL("../lib/export.worker.ts", import.meta.url)), { result: report, logo }, "Preparing Excel download");
      if (message.type !== "complete" || !message.buffer) throw new Error("The download did not return a valid workbook.");
      setProgress({ title: "Preparing Excel download", step: "Workbook ready", percent: 100 });
      await new Promise((resolve) => setTimeout(resolve, 200));
      const url = URL.createObjectURL(new Blob([message.buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `Kohinoor-Strength-Report-${report.generatedAt.slice(0, 10)}.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60_000);
      setProgress(null);
      setFeedback({ title: "Excel download started", tone: "success", message: "Your workbook includes a worksheet for each category, Strength Detail, Employees, Summary, and Validation. Check your browser’s downloads. Your scanned report remains saved here." });
    } catch (error) {
      setProgress(null);
      if (!(error instanceof DOMException && error.name === "AbortError")) setFeedback({ title: "Download could not be prepared", tone: "error", message: error instanceof Error ? error.message : "Try downloading again. Your scanned report is still available." });
    } finally { busyRef.current = false; }
  };

  const totals = useMemo(() => report?.rows.reduce((sum, row) => ({ approved: sum.approved + row.approvedStrength, onRoll: sum.onRoll + row.onRoll }), { approved: 0, onRoll: 0 }) ?? { approved: 0, onRoll: 0 }, [report]);
  const categories = report ? report.categories : template.strengthStructure.map((item) => item.category);
  const selectedRows = report?.rows.filter((row) => row.category === category) ?? [];
  const visibleRows = selectedRows.filter((row) => `${row.subcategory} ${row.designation} ${row.grade} ${row.cadre}`.toLowerCase().includes(search.toLowerCase()));
  const subcategories = [...new Set(visibleRows.map((row) => row.subcategory))];
  const outdated = report && report.templateSignature !== JSON.stringify(template);

  return <div className="config-page min-h-screen bg-white">
    <header className="border-b border-neutral-200"><div className="mx-auto flex max-w-[1440px] flex-wrap items-center justify-between gap-5 px-5 py-6 sm:px-8 lg:px-10"><div><h1 className="text-2xl font-semibold tracking-tight">Turnover Report</h1><p className="mt-1.5 text-sm text-neutral-500">Upload your employee report. Review workforce strength. Download Excel.</p></div><div className="flex flex-wrap gap-2"><Button variant="outline" onClick={openUpload} disabled={restoring || !!progress}><Upload size={16} />{report ? "Upload new report" : "Upload report"}</Button><Button onClick={download} disabled={!report || !!progress}><Download size={16} /> Download Excel</Button></div></div></header>
    <main className="mx-auto max-w-[1440px] px-5 py-7 sm:px-8 lg:px-10">
      <ol aria-label="How to generate a report" className="mb-7 grid gap-4 rounded-xl border border-neutral-200 p-5 sm:grid-cols-3">{[{ title: "Upload Excel", text: "Choose your turnover workbook (.xlsx or .xls)." }, { title: "Scan & review", text: "We detect headers and match employees to your strength template." }, { title: "Download report", text: "Get a formatted workbook with a worksheet for every category." }].map((step, index) => <li key={step.title} className="flex gap-3"><span className={cn("flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-semibold", report || index === 0 ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-500")}>{index + 1}</span><div><p className="text-sm font-medium">{step.title}</p><p className="mt-1 text-xs leading-5 text-neutral-500">{step.text}</p></div></li>)}</ol>
      {restoring ? <div role="status" className="rounded-xl bg-neutral-50 p-10 text-center text-sm text-neutral-500">Checking for your saved report…</div> : !report ? <section onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); chooseFile(event.dataTransfer.files); }} className={cn("rounded-xl border border-dashed px-5 py-14 text-center", dragging ? "border-neutral-900 bg-neutral-100" : "border-neutral-300 bg-neutral-50/60")}><FileSpreadsheet className="mx-auto h-9 w-9 text-neutral-400" /><h2 className="mt-4 text-lg font-semibold tracking-tight">Start with your turnover report</h2><p className="mx-auto mt-2 max-w-lg text-sm leading-6 text-neutral-500">Drop an Excel workbook here or choose a file. Columns can be in any order, and title rows are supported.</p><Button onClick={openUpload} className="mt-5"><Plus size={16} /> Choose workbook</Button><p className="mt-4 text-xs text-neutral-400">Excel .xlsx / .xls · Maximum 20 MB · One workbook at a time</p></section> : <>
        <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">{[{ label: "Approved strength", value: totals.approved, text: "From your strength template" }, { label: "On-roll employees", value: totals.onRoll, text: "Matched unique uploaded employees" }, { label: "Vacancies", value: totals.approved - totals.onRoll, text: "Approved less on-roll" }, { label: "Records to review", value: report.issues.length, text: "Unmatched, duplicate, or invalid data" }].map((metric) => <div key={metric.label} className="rounded-xl bg-neutral-50 p-4 sm:p-5"><p className="text-xs font-medium text-neutral-600">{metric.label}</p><p className="mt-3 text-3xl font-semibold tracking-tight tabular-nums">{metric.value.toLocaleString()}</p><p className="mt-1.5 text-xs text-neutral-400">{metric.text}</p></div>)}</div>
        <div className="mt-5 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-neutral-200 px-4 py-3"><div className="min-w-0"><p className="break-all text-sm font-medium">{report.fileName}</p><p className="mt-1 text-xs text-neutral-500">Scanned {new Date(report.generatedAt).toLocaleString("en-GB", { timeZone: "Asia/Karachi" })} PKT · {report.scannedRows.toLocaleString()} employee rows</p><p role="status" className="mt-1 text-xs text-neutral-500">{storageStatus}</p></div><Button variant="outline" size="sm" onClick={() => { setIssuePage(0); setIssuesOpen(true); }}>Review issues ({report.issues.length})</Button></div>
        {outdated && <div role="alert" className="mt-4 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900"><span>This report uses an earlier configuration. Reprocess the saved upload to use the current template.</span><Button size="sm" variant="outline" disabled={!!progress} onClick={() => storedFile && processFile(storedFile.file, storedFile.fileName, true)}>Reprocess saved upload</Button></div>}
        {!!report.unmatched && <p className="mt-4 rounded-lg bg-amber-50 px-4 py-3 text-sm text-amber-800">{report.unmatched.toLocaleString()} unmatched or conflicting employee records are excluded from on-roll counts. Review issues before using the totals.</p>}
        <div className="mt-8 grid items-start gap-6 lg:grid-cols-[240px_minmax(0,1fr)]"><aside className="min-w-0 rounded-xl border border-neutral-200 p-3 lg:sticky lg:top-6"><div className="flex justify-between px-2 pb-3 pt-1"><h2 className="text-sm font-semibold">Categories</h2><span className="text-xs text-neutral-400">{categories.length}</span></div><nav aria-label="Report categories" className="flex gap-1 overflow-x-auto lg:flex-col">{categories.map((item) => <button key={item} type="button" aria-current={item === category ? "page" : undefined} onClick={() => { setCategory(item); setSearch(""); }} className={cn("flex shrink-0 items-center justify-between gap-3 rounded-lg px-3 py-2.5 text-left text-sm lg:shrink", item === category ? "bg-neutral-100 font-semibold text-neutral-950" : "text-neutral-500 hover:bg-neutral-50")}><span className="break-words">{item}</span><ChevronRight size={14} className="shrink-0" /></button>)}</nav></aside>
        <div className="min-w-0"><h2 className="text-xl font-semibold tracking-tight">{category}</h2><p className="mt-1 text-sm text-neutral-500">{selectedRows.length} designations · {selectedRows.reduce((sum, row) => sum + row.onRoll, 0).toLocaleString()} on-roll employees</p><p className="mt-2 text-xs text-neutral-400">Select a designation to view its employees.</p><label className="relative mt-5 block"><Search className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-neutral-400" /><input aria-label="Search report designations" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search designations, grades, or cadres…" className="h-10 w-full rounded-lg border border-neutral-200 pl-10 pr-3 text-sm outline-none focus-visible:ring-2 focus-visible:ring-neutral-400" /></label><div className="mt-5 space-y-5">{subcategories.map((subcategory) => <section key={subcategory} className="overflow-hidden rounded-xl border border-neutral-200"><div className="flex items-center gap-3 px-5 py-4"><h3 className="break-words text-sm font-semibold">{subcategory}</h3><span className="rounded-md bg-neutral-100 px-2 py-0.5 text-xs text-neutral-500">{visibleRows.filter((row) => row.subcategory === subcategory).length}</span></div><table className="designation-table report-table w-full text-left text-sm"><thead className="border-y border-neutral-100 bg-neutral-50 text-xs text-neutral-500"><tr><th>Designation</th><th>Grade</th><th>Cadre</th><th className="text-right">Approved</th><th className="text-right">On-roll</th><th className="text-right">Vacancies</th></tr></thead><tbody>{visibleRows.filter((row) => row.subcategory === subcategory).map((row, index) => <tr key={`${row.designation}-${row.grade}-${index}`} className="cursor-pointer border-t border-neutral-100 transition-colors hover:bg-neutral-50" onClick={() => setEmployeeRowIndex(report.rows.indexOf(row))}><td data-label="Designation" className="font-medium"><button type="button" aria-label={`View employees for ${row.designation}, ${row.grade}, in ${row.subcategory}, ${row.category}`} aria-haspopup="dialog" onClick={(event) => { event.stopPropagation(); setEmployeeRowIndex(report.rows.indexOf(row)); }} className="group flex items-center gap-2 text-left font-medium text-neutral-900 hover:underline hover:underline-offset-4">{row.designation}<ChevronRight className="h-3.5 w-3.5 shrink-0 text-neutral-400 transition-transform group-hover:translate-x-0.5" /></button></td><td data-label="Grade">{row.grade}</td><td data-label="Cadre">{row.cadre}</td><td data-label="Approved" className="text-right tabular-nums">{row.approvedStrength}</td><td data-label="On-roll" className="text-right tabular-nums">{row.onRoll}</td><td data-label="Vacancies" className="text-right tabular-nums">{row.approvedStrength - row.onRoll}</td></tr>)}</tbody></table></section>)}{!subcategories.length && <p className="rounded-xl border border-dashed border-neutral-300 p-10 text-center text-sm text-neutral-500">No matching designations. Try a different search.</p>}</div></div></div>
      </>}
      {!report && storageStatus && <p className="mt-4 text-xs text-neutral-500">{storageStatus}</p>}
      <p className="mt-6 text-xs leading-5 text-neutral-400">Reports are stored on this browser and device. Clearing site data removes saved uploads. Scanning and downloading never change the Strength JSON.</p>
    </main>
    {uploadOpen && <dialog ref={uploadRef} className="config-dialog" aria-labelledby="upload-title" aria-describedby="upload-description" onCancel={() => setUploadOpen(false)}><div className="p-6"><div className="flex justify-between gap-4"><div><h2 id="upload-title" className="text-lg font-semibold tracking-tight">{report ? "Upload a new turnover report" : "Upload turnover report"}</h2><p id="upload-description" className="mt-2 text-sm leading-6 text-neutral-500">Choose a workbook, then start scanning. {report ? "Your current report is replaced only after a successful scan." : "We’ll detect the table and match it to your strength template."}</p></div><Button variant="ghost" size="icon" aria-label="Close upload" onClick={() => setUploadOpen(false)}><X size={16} /></Button></div><div onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); chooseFile(event.dataTransfer.files); }} className="mt-5 rounded-lg border border-dashed border-neutral-300 bg-neutral-50 px-4 py-7 text-center"><FileSpreadsheet className="mx-auto h-7 w-7 text-neutral-400" /><p className="mt-3 break-all text-sm font-medium">{file ? file.name : "Drop your Excel workbook here"}</p><p className="mt-1 text-xs text-neutral-500">{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB · Ready to scan` : `.xlsx or .xls · Up to ${MAX_FILE_SIZE / 1024 / 1024} MB`}</p><Button variant="outline" className="mt-4" onClick={() => inputRef.current?.click()}>{file ? "Choose another file" : "Browse files"}</Button><input ref={inputRef} aria-label="Turnover workbook" type="file" accept=".xlsx,.xls" className="sr-only" onChange={(event) => { if (event.target.files?.length) chooseFile(event.target.files); event.target.value = ""; }} /></div><details className="mt-5 text-sm"><summary className="cursor-pointer font-medium">Required column headers</summary><p className="mt-2 text-xs leading-6 text-neutral-500">{TURNOVER_HEADERS.join(" · ")}</p><p className="mt-2 text-xs leading-5 text-neutral-500">Column order and position can vary. Department or Unit must identify the template subcategory. Employee number + Unit identifies a unique employee.</p></details>{uploadError && <p role="alert" className="mt-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">{uploadError}</p>}<div className="mt-6 flex justify-end gap-2"><Button variant="outline" onClick={() => setUploadOpen(false)}>Cancel</Button><Button disabled={!file} onClick={() => { if (!file) return; setUploadOpen(false); void processFile(file, file.name); }}><Upload size={16} /> Start scanning</Button></div></div></dialog>}
    {issuesOpen && report && <dialog ref={issuesRef} className="config-dialog !w-[min(760px,calc(100%-32px))]" aria-labelledby="issues-title" onCancel={() => setIssuesOpen(false)}><div className="p-6"><div className="flex items-center justify-between gap-4"><h2 id="issues-title" className="text-lg font-semibold">Validation notes ({report.issues.length})</h2><Button variant="ghost" size="icon" aria-label="Close validation notes" onClick={() => setIssuesOpen(false)}><X size={16} /></Button></div><p className="mt-2 text-sm text-neutral-500">Correct source records and upload again to resolve issues. All notes are included in the Excel workbook.</p><div className="mt-5 space-y-3">{report.issues.slice(issuePage * 20, (issuePage + 1) * 20).map((issue, index) => <div key={index} className="rounded-lg border border-neutral-200 p-3"><p className="text-xs font-medium text-neutral-500">{issue.sheet}{issue.row ? ` · Row ${issue.row}` : ""}{issue.employeeId ? ` · Emp # ${issue.employeeId}` : ""}</p><p className="mt-1 text-sm">{issue.message}</p></div>)}{!report.issues.length && <p className="py-6 text-sm text-neutral-500">No validation issues found.</p>}</div><div className="mt-5 flex items-center justify-between gap-2"><Button variant="outline" size="sm" disabled={issuePage === 0} onClick={() => setIssuePage((page) => page - 1)}>Previous</Button><span className="text-xs text-neutral-500">Page {issuePage + 1} of {Math.max(1, Math.ceil(report.issues.length / 20))}</span><Button variant="outline" size="sm" disabled={(issuePage + 1) * 20 >= report.issues.length} onClick={() => setIssuePage((page) => page + 1)}>Next</Button></div></div></dialog>}
    {report && employeeRowIndex !== null && <DesignationEmployeesDialog key={employeeRowIndex} report={report} rowIndex={employeeRowIndex} onClose={() => setEmployeeRowIndex(null)} />}
    <ProgressDialog progress={progress} onCancel={cancelProcessing} />
    <FeedbackDialog feedback={feedback} onClose={() => setFeedback(null)} />
  </div>;
}
