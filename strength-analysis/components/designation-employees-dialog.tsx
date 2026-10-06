"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { Search, Users, X } from "lucide-react";
import { Button } from "./ui/button";
import { TURNOVER_HEADERS, type TurnoverResult } from "@/lib/turnover-contracts";

export function DesignationEmployeesDialog({ report, rowIndex, onClose }: { report: TurnoverResult; rowIndex: number; onClose: () => void }) {
  const ref = useRef<HTMLDialogElement>(null);
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(0);
  const row = report.rows[rowIndex];
  const employees = useMemo(() => report.employees.filter((record) => record.match === rowIndex), [report, rowIndex]);
  const filtered = useMemo(() => {
    const terms = query.normalize("NFKC").trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
    return employees.filter((record) => {
      const text = TURNOVER_HEADERS.map((header) => record.employee[header]).join(" ").normalize("NFKC").toLocaleLowerCase();
      return terms.every((term) => text.includes(term));
    });
  }, [employees, query]);
  const pageSize = 10;
  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
  const fields = TURNOVER_HEADERS.filter((header) => header !== "Employee Name" && header !== "Emp #");

  useEffect(() => {
    const dialog = ref.current;
    const trigger = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    dialog?.showModal();
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { dialog?.close(); document.body.style.overflow = overflow; trigger?.focus({ preventScroll: true }); };
  }, []);

  return <dialog ref={ref} className="config-dialog employee-dialog" aria-labelledby="employee-dialog-title" aria-describedby="employee-dialog-context" onCancel={onClose}>
    <div className="flex max-h-[calc(100dvh-32px)] flex-col">
      <header className="shrink-0 border-b border-neutral-200 p-5 sm:p-6">
        <div className="flex items-start justify-between gap-4"><div className="min-w-0"><p id="employee-dialog-context" className="break-words text-xs leading-5 text-neutral-500">{row.category} / {row.subcategory}</p><h2 id="employee-dialog-title" className="mt-1 break-words text-lg font-semibold tracking-tight sm:text-xl">{row.designation}</h2><div className="mt-2 flex flex-wrap items-center gap-2 text-xs text-neutral-500"><span className="rounded-md bg-neutral-100 px-2 py-1">Grade {row.grade}</span><span className="rounded-md bg-neutral-100 px-2 py-1">{row.cadre}</span><span>{employees.length.toLocaleString()} on-roll {employees.length === 1 ? "employee" : "employees"}</span></div></div><Button variant="ghost" size="icon" className="shrink-0" aria-label="Close employee details" onClick={onClose}><X size={16} /></Button></div>
        <label className="relative mt-5 block"><Search className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-neutral-400" /><input autoFocus aria-label="Search designation employees" value={query} onChange={(event) => { setQuery(event.target.value); setPage(0); }} placeholder="Search name, employee number, shift, or any field…" className="h-10 w-full border border-neutral-200 bg-white pl-10 pr-10 text-sm outline-none focus-visible:ring-2 focus-visible:ring-neutral-400" />{query && <button type="button" aria-label="Clear employee search" onClick={() => { setQuery(""); setPage(0); }} className="absolute right-1 top-1 flex h-8 w-8 items-center justify-center rounded-md text-neutral-500 hover:bg-neutral-100"><X size={14} /></button>}</label>
        <p role="status" aria-live="polite" className="mt-2 text-xs text-neutral-500">{query ? `${filtered.length.toLocaleString()} of ${employees.length.toLocaleString()} employees match your search` : "Employee information from your uploaded turnover report."}</p>
      </header>
      <div className="min-h-0 flex-1 overflow-y-auto bg-neutral-50/60 p-4 sm:p-6">
        {filtered.length ? <div className="space-y-4">{filtered.slice(page * pageSize, (page + 1) * pageSize).map((record) => <article key={`${record.sheet}-${record.row}`} className="rounded-xl border border-neutral-200 bg-white">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-neutral-100 px-4 py-3"><h3 className="min-w-0 break-words text-sm font-semibold">{record.employee["Employee Name"] || "Name not provided"}</h3><span className="rounded-md bg-neutral-100 px-2 py-1 text-xs font-medium text-neutral-600">Emp # {record.employee["Emp #"]}</span></div>
          <dl className="grid grid-cols-2 gap-x-5 gap-y-4 p-4 sm:grid-cols-3 lg:grid-cols-4">{fields.map((field) => <div key={field} className="min-w-0"><dt className="text-xs text-neutral-500">{field}</dt><dd className="mt-1 break-words text-sm text-neutral-900">{record.employee[field] || <span className="text-neutral-400">Not provided</span>}</dd></div>)}</dl>
        </article>)}</div> : <div className="flex min-h-52 flex-col items-center justify-center px-4 py-8 text-center"><Users className="h-7 w-7 text-neutral-400" /><h3 className="mt-3 text-sm font-semibold">{employees.length ? "No employees match your search" : "No on-roll employees in this designation"}</h3><p className="mt-2 max-w-md text-sm leading-6 text-neutral-500">{employees.length ? "Try a different name, employee number, shift, or department." : "The uploaded report has no employees matched to this category, subcategory, designation, and grade."}</p>{query && <Button variant="outline" size="sm" className="mt-4" onClick={() => { setQuery(""); setPage(0); }}>Clear search</Button>}</div>}
      </div>
      <footer className="flex shrink-0 flex-wrap items-center justify-between gap-3 border-t border-neutral-200 bg-white px-5 py-4 sm:px-6"><p className="text-xs text-neutral-500">{filtered.length ? `${page * pageSize + 1}–${Math.min((page + 1) * pageSize, filtered.length)} of ${filtered.length.toLocaleString()} employees` : "0 employees"}</p><div className="flex items-center gap-2"><Button variant="outline" size="sm" disabled={page === 0} onClick={() => setPage((current) => current - 1)}>Previous</Button><span className="text-xs text-neutral-500">{page + 1} / {pageCount}</span><Button variant="outline" size="sm" disabled={page + 1 >= pageCount} onClick={() => setPage((current) => current + 1)}>Next</Button></div></footer>
    </div>
  </dialog>;
}
