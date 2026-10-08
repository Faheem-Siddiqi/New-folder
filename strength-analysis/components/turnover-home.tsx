"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Download, FileSpreadsheet, Plus, Upload, X } from "lucide-react";
import { getStrengthData, saveTurnoverResult, refreshTurnoverResult } from "@/lib/turnover-client";
import { getLatestReport } from "@/lib/latest-report-client";
import { confirmedAdjustments } from "@/lib/review-cases";
import { ReviewCases } from "./review-cases";
import { StrengthPanels } from "./strength-panels";
import { Button } from "./ui/button";
import { DesignationEmployeesDialog } from "./designation-employees-dialog";
import { FeedbackDialog, type Feedback } from "./feedback-dialog";
import { ProgressDialog, type Progress } from "./progress-dialog";
import { MATCHING_VERSION, MAX_FILE_SIZE, TURNOVER_HEADERS, validateUpload, type TurnoverResult } from "@/lib/turnover-contracts";
import type { StrengthData } from "@/lib/strength-data";
import { cn } from "@/lib/utils";

type WorkerMessage = { type: "progress"; step: string; percent: number } | { type: "error"; message: string } | { type: "complete"; result?: TurnoverResult; buffer?: ArrayBuffer };

async function loadCurrentTemplate() {
  try { return await getStrengthData(); }
  catch (error) { throw new Error(error instanceof Error ? error.message : "The latest strength configuration could not be loaded. Check your connection and try again."); }
}

function downloadWorkbook(buffer: ArrayBuffer, fileName: string) {
  const url = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

export function TurnoverHome({ template, initialReport, initialReportError = "" }: { template: StrengthData; initialReport: TurnoverResult | null; initialReportError?: string }) {
  const [employeeRowIndex, setEmployeeRowIndex] = useState<number | null>(null);
  const [report, setReport] = useState<TurnoverResult | null>(initialReport);
  const reportRef = useRef<TurnoverResult | null>(initialReport);
  const [reviewOpen, setReviewOpen] = useState(false);
  useEffect(() => { reportRef.current = report; }, [report]);
  const [storageStatus, setStorageStatus] = useState(initialReportError || (initialReport ? "Restored from the latest saved report." : ""));
  const [currentTemplate, setCurrentTemplate] = useState(template);
  const templateRef = useRef(template);
  const [progress, setProgress] = useState<Progress | null>(null);
  const [feedback, setFeedback] = useState<Feedback | null>(null);
  const [uploadOpen, setUploadOpen] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [uploadError, setUploadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [issuesOpen, setIssuesOpen] = useState(false);
  const [issuePage, setIssuePage] = useState(0);
  const uploadRef = useRef<HTMLDialogElement>(null);
  const issuesRef = useRef<HTMLDialogElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const activeJob = useRef<{ worker: Worker; cancel: () => void } | null>(null);
  const busyRef = useRef(false);
  const committing = useRef(false);
  const cancelled = useRef(false);
  const refreshSequence = useRef(0);
  const refreshSavedResult = useCallback(async () => {
    if (busyRef.current) return;
    const sequence = ++refreshSequence.current;
    try {
      const latest = await getLatestReport();
      if (sequence !== refreshSequence.current || busyRef.current) return;
      if (latest === reportRef.current) return;
      reportRef.current = latest;
      setReport(latest);
      setEmployeeRowIndex(null);
      setStorageStatus(latest ? "Loaded from the latest saved report." : "No saved report. Upload a turnover workbook to start.");
    } catch (error) {
      setStorageStatus(error instanceof Error ? error.message : "The latest saved report could not be loaded. Please retry.");
    }
  }, []);
  const previousHomeDialog = useRef(false);
  useEffect(() => {
    const open = uploadOpen || issuesOpen || employeeRowIndex !== null || feedback !== null;
    const closed = previousHomeDialog.current && !open;
    previousHomeDialog.current = open;
    if (closed) void refreshSavedResult();
  }, [uploadOpen, issuesOpen, employeeRowIndex, feedback, refreshSavedResult]);

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

  const processFile = useCallback(async (source: Blob, fileName: string) => {
    if (busyRef.current) return;
    busyRef.current = true;
    cancelled.current = false;
    setProgress({ title: "Scanning turnover report", step: "Loading uploaded workbook", percent: 2 });
    try {
      const latest = await loadCurrentTemplate();
      templateRef.current = latest;
      setCurrentTemplate(latest);
      validateUpload(fileName, source.size);
      const buffer = await source.arrayBuffer();
      if (cancelled.current) throw new DOMException("Processing cancelled", "AbortError");
      const message = await runWorker(new Worker(new URL("../lib/turnover.worker.ts", import.meta.url)), { buffer, fileName, template: latest }, "Scanning turnover report", [buffer]);
      if (message.type !== "complete" || !message.result) throw new Error("The scan did not return a report.");
      const scanned = message.result;
      if (cancelled.current) throw new DOMException("Processing cancelled", "AbortError");
      committing.current = true;
      setProgress({ title: "Scanning turnover report", step: "Saving the latest report and checking previous cases", percent: 96 });
      const response = await saveTurnoverResult(scanned);
      if (!response.ok) throw new Error(response.error);
      const next = response.result;
      setProgress({ title: "Scanning turnover report", step: "Saving report for your next visit", percent: 98 });
      reportRef.current = next;
      setReport(next);
      setEmployeeRowIndex(null);
      setStorageStatus("Latest report saved. Available after refresh or reopening.");
      const warnings = next.issues.filter((issue) => issue.severity !== "info").length;
      setProgress({ title: "Scanning turnover report", step: "Scan complete", percent: 100 });
      await new Promise((resolve) => setTimeout(resolve, 200));
      setProgress(null);
      setFeedback({ title: warnings ? "Report ready — review flagged records" : "Report ready", tone: warnings ? "warning" : "success", message: `${next.matched.toLocaleString()} employees matched to the strength template. ${next.unmatched.toLocaleString()} employee records are unmatched or conflicting.\n\n${warnings ? `${warnings.toLocaleString()} validation notes are available under Review issues. ` : ""}Your latest report and case changes are saved. Review the category results, then download Excel.` });
    } catch (error) {
      setProgress(null);
      if (!(error instanceof DOMException && error.name === "AbortError")) setFeedback({ title: "Report could not be processed", tone: "error", message: `${error instanceof Error ? error.message : "Please try uploading a new Excel copy."}\n\nYour previous report and Strength JSON template are unchanged.` });
    } finally { busyRef.current = false; committing.current = false; }
  }, [runWorker]);

  const refreshReportFromJson = async () => {
    const current = reportRef.current;
    if (busyRef.current || !current?.reportId) return;
    busyRef.current = true;
    committing.current = true;
    setProgress({ title: "Refreshing saved report", step: "Matching saved employee records to the current strength configuration", percent: 30 });
    try {
      const next = await refreshTurnoverResult(current.reportId);
      reportRef.current = next;
      setReport(next);
      setEmployeeRowIndex(null);
      const latest = JSON.parse(next.templateSignature) as StrengthData;
      templateRef.current = latest;
      setCurrentTemplate(latest);
      setStorageStatus("Latest report saved. Available after refresh or reopening.");
      setFeedback({ title: "Saved report refreshed", tone: "success", message: "The saved employee records have been matched to the current configuration and the latest-result JSON has been updated. Review any pending cases before downloading Excel." });
    } catch (error) {
      setFeedback({ title: "Saved report could not be refreshed", tone: "error", message: error instanceof Error ? error.message : "Please retry. Your previous saved report is preserved." });
    } finally { busyRef.current = false; committing.current = false; setProgress(null); }
  };

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
      const latest = await loadCurrentTemplate();
      templateRef.current = latest;
      setCurrentTemplate(latest);
      let exportReport = await getLatestReport();
      if (!exportReport) throw new Error("No saved report is available. Upload a turnover workbook first.");
      if (exportReport.templateSignature !== JSON.stringify(latest) || exportReport.matchingVersion !== MATCHING_VERSION) {
        if (!exportReport.reportId) throw new Error("This saved report needs a new turnover upload before it can be refreshed.");
        if (cancelled.current) throw new DOMException("Processing cancelled", "AbortError");
        committing.current = true;
        setProgress({ title: "Preparing Excel download", step: "Refreshing employee counts from the saved JSON", percent: 25 });
        try { exportReport = await refreshTurnoverResult(exportReport.reportId); }
        finally { committing.current = false; }
        setEmployeeRowIndex(null);
        setStorageStatus("Updated strength saved to the latest-result JSON.");
      }
      reportRef.current = exportReport;
      setReport(exportReport);
      const response = await fetch("/kohinoor-logo.png", { cache: "no-store" });
      if (!response.ok) throw new Error("The company logo could not be loaded. Please retry the download.");
      const logoBlob = await response.blob();
      const logo = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result));
        reader.onerror = () => reject(new Error("The company logo could not be read."));
        reader.readAsDataURL(logoBlob);
      });
      exportReport = { ...exportReport, generatedAt: new Date().toISOString() };
      if (cancelled.current) throw new DOMException("Processing cancelled", "AbortError");
      const message = await runWorker(new Worker(new URL("../lib/export.worker.ts", import.meta.url)), { result: exportReport, logo }, "Preparing Excel download");
      if (message.type !== "complete" || !message.buffer) throw new Error("The download did not return a valid workbook.");
      setProgress({ title: "Preparing Excel download", step: "Workbook ready", percent: 100 });
      await new Promise((resolve) => setTimeout(resolve, 200));
      downloadWorkbook(message.buffer, `Kohinoor-Strength-Report-${exportReport.generatedAt.slice(0, 10)}.xlsx`);
      setProgress(null);
      setFeedback({ title: "Excel download started", tone: "success", message: "Your workbook includes Summary and a worksheet for each category. Check your browser’s downloads. Your scanned report remains saved here." });
    } catch (error) {
      setProgress(null);
      if (!(error instanceof DOMException && error.name === "AbortError")) setFeedback({ title: "Download could not be prepared", tone: "error", message: error instanceof Error ? error.message : "Try downloading again. Your scanned report is still available." });
    } finally { busyRef.current = false; }
  };

  const totals = useMemo(() => report?.rows.reduce((sum, row) => ({ approved: sum.approved + row.approvedStrength, onRoll: sum.onRoll + row.onRoll }), { approved: 0, onRoll: 0 }) ?? { approved: 0, onRoll: 0 }, [report]);
  const outdated = report && (report.templateSignature !== JSON.stringify(currentTemplate) || report.matchingVersion !== MATCHING_VERSION);

  useEffect(() => {
    const refresh = async () => {
      if (document.visibilityState !== "visible" || busyRef.current || employeeRowIndex !== null || reviewOpen || uploadOpen || issuesOpen || feedback) return;
      try {
        await refreshSavedResult();
        const latest = await loadCurrentTemplate();
        const current = reportRef.current;
        if (busyRef.current || (JSON.stringify(latest) === JSON.stringify(templateRef.current) && (!current || (current.templateSignature === JSON.stringify(latest) && current.matchingVersion === MATCHING_VERSION)))) return;
        templateRef.current = latest;
        setCurrentTemplate(latest);
        // Configuration changes are shown immediately; reprocessing is explicit.
      } catch { /* Downloads independently verify freshness and report errors. */ }
    };
    void refresh();
    window.addEventListener("focus", refresh);
    document.addEventListener("visibilitychange", refresh);
    const timer = window.setInterval(refresh, 15000);
    return () => { window.removeEventListener("focus", refresh); document.removeEventListener("visibilitychange", refresh); window.clearInterval(timer); };
  }, [employeeRowIndex, reviewOpen, uploadOpen, issuesOpen, feedback, refreshSavedResult]);

  return <div className="config-page min-h-screen bg-white">
    <header className="border-b border-neutral-200"><div className="mx-auto flex max-w-[1440px] flex-wrap items-center justify-between gap-5 px-5 py-6 sm:px-8 lg:px-10"><div><h1 className="text-2xl font-semibold tracking-tight">Turnover Report</h1><p className="mt-1.5 text-sm text-neutral-500">Upload your employee report. Review workforce strength. Download Excel.</p></div><div className="flex flex-wrap gap-2"><Button variant="outline" onClick={openUpload} disabled={!!progress}><Upload size={16} />{report ? "Upload new report" : "Upload report"}</Button><Button onClick={download} disabled={!report || !!progress}><Download size={16} /> Download Excel</Button></div></div></header>
    <main className="mx-auto max-w-[1440px] px-5 py-7 sm:px-8 lg:px-10">
      <ol aria-label="How to generate a report" className="mb-7 grid gap-4 rounded-xl border border-neutral-200 p-5 sm:grid-cols-3">{[{ title: "Upload Excel", text: "Choose your turnover workbook (.xlsx or .xls)." }, { title: "Scan & review", text: "We detect headers and match employees to your strength template." }, { title: "Download report", text: "Get a formatted workbook with a worksheet for every category." }].map((step, index) => <li key={step.title} className="flex gap-3"><span className={cn("flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-semibold", report || index === 0 ? "bg-neutral-900 text-white" : "bg-neutral-100 text-neutral-500")}>{index + 1}</span><div><p className="text-sm font-medium">{step.title}</p><p className="mt-1 text-xs leading-5 text-neutral-500">{step.text}</p></div></li>)}</ol>
      {!report ? <section onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(event) => { event.preventDefault(); setDragging(false); chooseFile(event.dataTransfer.files); }} className={cn("rounded-xl border border-dashed px-5 py-14 text-center", dragging ? "border-neutral-900 bg-neutral-100" : "border-neutral-300 bg-neutral-50/60")}><FileSpreadsheet className="mx-auto h-9 w-9 text-neutral-400" /><h2 className="mt-4 text-lg font-semibold tracking-tight">Start with your turnover report</h2><p className="mx-auto mt-2 max-w-lg text-sm leading-6 text-neutral-500">Drop an Excel workbook here or choose a file. Columns can be in any order, and title rows are supported.</p><Button onClick={openUpload} className="mt-5"><Plus size={16} /> Choose workbook</Button><p className="mt-4 text-xs text-neutral-400">Excel .xlsx / .xls · Maximum 20 MB · One workbook at a time</p></section> : <>
        <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">{[{ label: "Approved strength", value: totals.approved, text: "From your strength template" }, { label: "On-roll employees", value: totals.onRoll, text: "Matched unique uploaded employees" }, { label: "Vacancies", value: totals.approved - totals.onRoll + confirmedAdjustments(report), text: "Approved less on-roll, plus saved cases" }, { label: "Excluded records", value: report.unmatched + report.skippedRows, text: "Unmatched, conflicting, or invalid data" }].map((metric) => <div key={metric.label} className="rounded-xl bg-neutral-50 p-4 sm:p-5"><p className="text-xs font-medium text-neutral-600">{metric.label}</p><p className={cn("mt-3 text-3xl font-semibold tracking-tight tabular-nums", metric.label === "Vacancies" && (metric.value > 0 ? "text-emerald-700" : metric.value < 0 ? "text-red-700" : "text-neutral-600"))}>{metric.value.toLocaleString()}</p><p className="mt-1.5 text-xs text-neutral-400">{metric.text}</p></div>)}</div>
        <div className="mt-5 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-neutral-200 px-4 py-3"><div className="min-w-0"><p className="break-all text-sm font-medium">{report.fileName}</p><p className="mt-1 text-xs text-neutral-500">Scanned {new Date(report.generatedAt).toLocaleString("en-GB", { timeZone: "Asia/Karachi" })} PKT · {report.scannedRows.toLocaleString()} employee rows</p><p role="status" className="mt-1 text-xs text-neutral-500">{storageStatus}</p></div><Button variant="outline" size="sm" onClick={() => { setIssuePage(0); setIssuesOpen(true); }}>Review issues ({report.issues.length})</Button></div>
        {outdated && <div role="alert" className="mt-4 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900"><span>Updated configuration or matching rules are available. Refresh the saved employee records to update the counts.</span><Button size="sm" variant="outline" disabled={!!progress} onClick={() => void refreshReportFromJson()}>Refresh saved report</Button></div>}
        {!!report.unmatched && <p className="mt-4 rounded-lg bg-amber-50 px-4 py-3 text-sm text-amber-800">{report.unmatched.toLocaleString()} unmatched or conflicting employee records are excluded from on-roll counts. Review issues before using the totals.</p>}

      </>}
      {report && <ReviewCases report={report} onInteraction={setReviewOpen} onClose={refreshSavedResult} onUpdate={async (updated) => {
        setReport(updated); reportRef.current = updated;
        setStorageStatus("Case changes saved to the latest-result JSON.");
      }} />}
      {report && <StrengthPanels data={report} onDesignation={setEmployeeRowIndex} />}
      {!report && storageStatus && <p className="mt-4 text-xs text-neutral-500">{storageStatus}</p>}
      <p className="mt-6 text-xs leading-5 text-neutral-400">Reports and approved cases are saved in the latest-result JSON. Configuration comes from the strength JSON.</p>
    </main>
    {uploadOpen && <dialog ref={uploadRef} className="config-dialog" aria-labelledby="upload-title" aria-describedby="upload-description" onCancel={() => setUploadOpen(false)}><div className="p-6"><div className="flex justify-between gap-4"><div><h2 id="upload-title" className="text-lg font-semibold tracking-tight">{report ? "Upload a new turnover report" : "Upload turnover report"}</h2><p id="upload-description" className="mt-2 text-sm leading-6 text-neutral-500">Choose a workbook, then start scanning. {report ? "Your current report is replaced only after a successful scan." : "We’ll detect the table and match it to your strength template."}</p></div><Button variant="ghost" size="icon" aria-label="Close upload" onClick={() => setUploadOpen(false)}><X size={16} /></Button></div><div onDragOver={(event) => event.preventDefault()} onDrop={(event) => { event.preventDefault(); chooseFile(event.dataTransfer.files); }} className="mt-5 rounded-lg border border-dashed border-neutral-300 bg-neutral-50 px-4 py-7 text-center"><FileSpreadsheet className="mx-auto h-7 w-7 text-neutral-400" /><p className="mt-3 break-all text-sm font-medium">{file ? file.name : "Drop your Excel workbook here"}</p><p className="mt-1 text-xs text-neutral-500">{file ? `${(file.size / 1024 / 1024).toFixed(2)} MB · Ready to scan` : `.xlsx or .xls · Up to ${MAX_FILE_SIZE / 1024 / 1024} MB`}</p><Button variant="outline" className="mt-4" onClick={() => inputRef.current?.click()}>{file ? "Choose another file" : "Browse files"}</Button><input ref={inputRef} aria-label="Turnover workbook" type="file" accept=".xlsx,.xls" className="sr-only" onChange={(event) => { if (event.target.files?.length) chooseFile(event.target.files); event.target.value = ""; }} /></div><details className="mt-5 text-sm"><summary className="cursor-pointer font-medium">Turnover column headers</summary><p className="mt-2 text-xs leading-6 text-neutral-500">{TURNOVER_HEADERS.join(" · ")}</p><p className="mt-2 text-xs leading-5 text-neutral-500">Column order and position can vary. Grade is optional and does not affect matching. Department or Unit must identify the template subcategory. Employee number + Unit identifies a unique employee.</p></details>{uploadError && <p role="alert" className="mt-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">{uploadError}</p>}<div className="mt-6 flex justify-end gap-2"><Button variant="outline" onClick={() => setUploadOpen(false)}>Cancel</Button><Button disabled={!file} onClick={() => { if (!file) return; setUploadOpen(false); void processFile(file, file.name); }}><Upload size={16} /> Start scanning</Button></div></div></dialog>}
    {issuesOpen && report && <dialog ref={issuesRef} className="config-dialog !w-[min(760px,calc(100%-32px))]" aria-labelledby="issues-title" onCancel={() => setIssuesOpen(false)}><div className="p-6"><div className="flex items-center justify-between gap-4"><h2 id="issues-title" className="text-lg font-semibold">Validation notes ({report.issues.length})</h2><Button variant="ghost" size="icon" aria-label="Close validation notes" onClick={() => setIssuesOpen(false)}><X size={16} /></Button></div><p className="mt-2 text-sm text-neutral-500">Notes marked Counted do not reduce on-roll strength. Excluded notes explain the missing or conflicting field. Correct those source or configuration values, then reprocess. Review the source workbook to correct excluded records.</p><div className="mt-5 space-y-3">{report.issues.slice(issuePage * 20, (issuePage + 1) * 20).map((issue, index) => <div key={index} className="rounded-lg border border-neutral-200 p-3"><p className="text-xs font-medium text-neutral-500">{issue.sheet}{issue.row ? ` · Row ${issue.row}` : ""}{issue.employeeId ? ` · Emp # ${issue.employeeId}` : ""}</p>{issue.outcome && <span className={cn("mt-2 inline-block rounded-md px-2 py-1 text-xs font-medium", issue.outcome === "counted" ? "bg-emerald-50 text-emerald-800" : issue.outcome === "excluded" ? "bg-amber-50 text-amber-800" : "bg-neutral-100 text-neutral-600")}>{issue.outcome === "counted" ? "Counted once" : issue.outcome === "excluded" ? "Excluded from on-roll" : "Information"}</span>}<p className="mt-2 text-sm">{issue.message}</p></div>)}{!report.issues.length && <p className="py-6 text-sm text-neutral-500">No validation issues found.</p>}</div><div className="mt-5 flex items-center justify-between gap-2"><Button variant="outline" size="sm" disabled={issuePage === 0} onClick={() => setIssuePage((page) => page - 1)}>Previous</Button><span className="text-xs text-neutral-500">Page {issuePage + 1} of {Math.max(1, Math.ceil(report.issues.length / 20))}</span><Button variant="outline" size="sm" disabled={(issuePage + 1) * 20 >= report.issues.length} onClick={() => setIssuePage((page) => page + 1)}>Next</Button></div></div></dialog>}
    {report && employeeRowIndex !== null && <DesignationEmployeesDialog key={employeeRowIndex} report={report} rowIndex={employeeRowIndex} onClose={() => setEmployeeRowIndex(null)} />}
    <ProgressDialog progress={progress} onCancel={cancelProcessing} />
    <FeedbackDialog feedback={feedback} onClose={() => setFeedback(null)} />
  </div>;
}
