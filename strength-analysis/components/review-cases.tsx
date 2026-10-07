"use client";

import { useEffect, useRef, useState } from "react";
import { X } from "lucide-react";
import { saveCaseResolution } from "@/app/actions";
import { againstPostOptions, caseIdentity, vacancyFor, resolutionLabel, reviewColors, type Resolution, type ReviewCase } from "@/lib/review-cases";
import type { TurnoverResult } from "@/lib/turnover-contracts";
import { Button } from "./ui/button";

export function ReviewCases({ report, onUpdate, onInteraction }: { report: TurnoverResult; onUpdate: (report: TurnoverResult) => Promise<void>; onInteraction: (open: boolean) => void }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(0);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");
  const ref = useRef<HTMLDialogElement>(null);
  const cases = report.reviewCases ?? [];
  const pending = cases.filter((entry) => !entry.resolution).length;
  const groups = [...new Set(cases.map((entry) => entry.rowIndex))].filter((index) => { const row = report.rows[index]; return `${row.category} ${row.subcategory} ${row.designation}`.toLowerCase().includes(query.toLowerCase().trim()); });
  useEffect(() => {
    onInteraction(open);
    if (!open) return;
    const dialog = ref.current;
    dialog?.showModal();
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { dialog?.close(); document.body.style.overflow = overflow; onInteraction(false); };
  }, [open, onInteraction]);
  async function confirm(entry: ReviewCase, target: string, reason = "") {
    if (saving || !target) return false;
    setSaving(true); setMessage("");
    try {
      const resolution: Resolution = { ...caseIdentity(report, entry.rowIndex, entry.caseNumber), id: crypto.randomUUID(), caseType: target === "__social_security__" ? "Social Security Leave" : target === "__other__" ? "Other" : "Against Post", assignedAgainstDesignation: target.startsWith("post:") ? target.slice(5) : null, ...(target === "__other__" ? { otherReason: reason.trim() } : {}), resolvedAt: new Date().toISOString() };
      const response = await saveCaseResolution(resolution);
      if (!response.ok) throw new Error(response.error);
      const saved = response.resolution;
      await onUpdate({ ...report, reviewCases: cases.map((item) => item.key === entry.key ? { ...item, resolution: saved, suggestion: saved } : item) });
      setMessage("Case saved. The designation?s vacancy has been adjusted by 1.");
      return true;
    } catch (error) { setMessage(error instanceof Error ? error.message : "Case could not be saved. Please retry."); return false; }
    finally { setSaving(false); }
  }
  return <section aria-label="Review Cases" className="mt-4 flex flex-wrap items-center justify-between gap-4 rounded-xl border border-neutral-200 bg-white p-4 sm:p-5">
    <div><h2 className="text-sm font-semibold">Review Cases <span className="ml-2 rounded-md bg-neutral-100 px-2 py-1 text-xs font-medium">{cases.length}</span></h2><p className="mt-2 text-xs leading-5 text-neutral-500">{pending} remaining ? {cases.length - pending} saved. Add cases to clear negative vacancies by designation.</p></div>
    <Button variant="outline" size="sm" onClick={() => { setOpen(true); setMessage(""); }}>Review cases</Button>
    {open && <dialog ref={ref} className="config-dialog !w-[min(800px,calc(100%-32px))]" aria-labelledby="review-cases-title" onCancel={(event) => { if (saving) event.preventDefault(); else setOpen(false); }}>
      <div className="flex max-h-[calc(100dvh-32px)] flex-col"><header className="border-b border-neutral-200 p-5 sm:p-6"><div className="flex justify-between gap-3"><h2 id="review-cases-title" className="text-lg font-semibold">Review Cases</h2><Button variant="ghost" size="icon" aria-label="Close review cases" disabled={saving} onClick={() => setOpen(false)}><X size={16}/></Button></div><p className="mt-2 text-sm leading-6 text-neutral-500">A vacancy of ?3 needs 3 cases. For each case, choose a designation, Social Security, or Other and type a reason. Save all 3 to bring the vacancy to 0.</p><input aria-label="Search review cases" placeholder="Search category, subcategory, or designation" value={query} onChange={(event) => { setQuery(event.target.value); setPage(0); }} className="mt-4 h-10 w-full rounded-lg border border-neutral-200 px-3 text-sm" /></header>
      <div className="min-h-0 flex-1 space-y-4 overflow-y-auto bg-neutral-50 p-5 sm:p-6">{groups.slice(page * 5, (page + 1) * 5).map((rowIndex) => { const row = report.rows[rowIndex]; const related = cases.filter((entry) => entry.rowIndex === rowIndex); const vacancy = vacancyFor(report, rowIndex); return <article key={rowIndex} className="rounded-xl border border-neutral-200 bg-white p-4 sm:p-5"><p className="text-xs leading-5 text-neutral-500">{row.category} / {row.subcategory}</p><div className="mt-2 flex flex-wrap items-center justify-between gap-2"><h3 className="text-sm font-semibold">{row.designation}</h3><span className={`rounded-md px-2 py-1 text-xs font-medium ${vacancy < 0 ? "bg-red-50 text-red-700" : "bg-emerald-50 text-emerald-700"}`}>Vacancy: {vacancy}</span></div><p className="mt-2 text-xs text-neutral-500">Approved {row.approvedStrength} ? On-roll {row.onRoll} ? {related.length} cases required ? {related.filter((entry) => entry.resolution).length} saved</p><div className="mt-4 divide-y divide-neutral-100">{related.map((entry) => <CaseEditor key={entry.key} report={report} entry={entry} disabled={saving} onConfirm={(target, reason) => confirm(entry, target, reason)} />)}</div></article>; })}{!groups.length && <p className="py-8 text-center text-sm text-neutral-500">{cases.length ? "No designations match your search." : "No negative vacancies require review in this upload."}</p>}</div>
      <footer className="border-t border-neutral-200 p-5">{message && <p role="status" className="mb-4 rounded-lg bg-neutral-100 p-3 text-sm">{message}</p>}<div className="flex items-center justify-between gap-3"><Button variant="outline" size="sm" disabled={saving || page === 0} onClick={() => setPage(page - 1)}>Previous</Button><span className="text-xs text-neutral-500">{pending} remaining ? Page {page + 1} / {Math.max(1, Math.ceil(groups.length / 5))}</span><Button variant="outline" size="sm" disabled={saving || (page + 1) * 5 >= groups.length} onClick={() => setPage(page + 1)}>Next</Button></div></footer></div>
    </dialog>}
  </section>;
}

function targetValue(resolution?: Resolution) { return resolution?.caseType === "Social Security Leave" ? "__social_security__" : resolution?.caseType === "Other" ? "__other__" : resolution?.assignedAgainstDesignation ? `post:${resolution.assignedAgainstDesignation}` : ""; }
function CaseEditor({ report, entry, disabled, onConfirm }: { report: TurnoverResult; entry: ReviewCase; disabled: boolean; onConfirm: (target: string, reason?: string) => Promise<boolean> }) {
  const [target, setTarget] = useState(targetValue(entry.resolution));
  const [reason, setReason] = useState(entry.resolution?.otherReason ?? "");
  const [changing, setChanging] = useState(false);
  const [changingSuggestion, setChangingSuggestion] = useState(false);
  const options = againstPostOptions(report.rows, report.rows[entry.rowIndex]);
  const suggestion = entry.suggestion;
  return <div data-testid="designation-case" className="py-4 first:pt-0 last:pb-0"><p className="mb-2 text-xs font-semibold">Case {entry.caseNumber}</p>
    {entry.resolution && !changing ? <div className="flex flex-wrap items-center justify-between gap-2"><p className={`rounded-lg border px-3 py-2 text-sm ${reviewColors[entry.resolution.caseType].className}`}>{resolutionLabel(entry.resolution)} ? Saved</p><Button variant="outline" size="sm" disabled={disabled} onClick={() => setChanging(true)}>Change case</Button></div> : <>
      {!entry.resolution && suggestion && !changingSuggestion && <div className="mb-3 rounded-lg border border-blue-200 bg-blue-50 p-3"><p className="text-xs font-medium text-blue-900">Previous Match Found</p><p className="mt-1 text-xs text-blue-800">{resolutionLabel(suggestion)}. Apply this previous case?</p><div className="mt-3 flex gap-2"><Button size="sm" disabled={disabled} onClick={() => onConfirm(targetValue(suggestion), suggestion.otherReason)}>Apply</Button><Button variant="outline" size="sm" disabled={disabled} onClick={() => setChangingSuggestion(true)}>Change</Button></div></div>}
      <div className="flex flex-col gap-3 sm:flex-row sm:items-end"><label className="block min-w-0 flex-1 text-xs font-medium">Case type<select aria-label={`Case ${entry.caseNumber} for ${report.rows[entry.rowIndex].designation}`} disabled={disabled} value={target} onChange={(event) => setTarget(event.target.value)} className="mt-2 h-10 w-full rounded-lg border border-neutral-200 bg-white px-3 text-sm"><option value="">Select an option</option><option value="__social_security__">Social Security</option><optgroup label="Against Post ? Designations">{options.map((designation) => <option key={designation} value={`post:${designation}`}>{designation}</option>)}</optgroup><option value="__other__">Other (type text)</option></select></label><Button size="sm" disabled={disabled || !target || target === "__other__" && !reason.trim()} onClick={async () => { if (await onConfirm(target, reason)) setChanging(false); }}>{disabled ? "Saving?" : "Save case"}</Button>{changing && <Button variant="outline" size="sm" disabled={disabled} onClick={() => setChanging(false)}>Cancel change</Button>}</div>
      {target === "__other__" && <label className="mt-3 block text-xs font-medium">Other reason<input aria-label={`Other reason for case ${entry.caseNumber} of ${report.rows[entry.rowIndex].designation}`} value={reason} maxLength={200} disabled={disabled} onChange={(event) => setReason(event.target.value)} placeholder="Type the reason for this case" className="mt-2 h-10 w-full rounded-lg border border-neutral-200 px-3 text-sm" /></label>}
    </>}
  </div>;
}
