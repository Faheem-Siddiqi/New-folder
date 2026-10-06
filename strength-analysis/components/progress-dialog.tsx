"use client";

import { useEffect, useRef } from "react";
import { LoaderCircle } from "lucide-react";
import { Button } from "./ui/button";

export type Progress = { title: string; step: string; percent: number };
export function ProgressDialog({ progress, onCancel }: { progress: Progress | null; onCancel: () => void }) {
  const ref = useRef<HTMLDialogElement>(null);
  const active = Boolean(progress);
  useEffect(() => {
    if (!active) return;
    const dialog = ref.current;
    dialog?.showModal();
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { dialog?.close(); document.body.style.overflow = overflow; };
  }, [active]);
  if (!progress) return null;
  return <dialog ref={ref} className="config-dialog" aria-labelledby="progress-title" aria-describedby="progress-step" onCancel={(event) => { event.preventDefault(); if (progress.percent < 98) onCancel(); }}>
    <div className="p-6"><LoaderCircle className="h-6 w-6 animate-spin text-neutral-600" /><h2 id="progress-title" className="mt-4 text-lg font-semibold">{progress.title}</h2><p id="progress-step" role="status" aria-live="polite" className="mt-2 text-sm text-neutral-500">{progress.step}</p><div role="progressbar" aria-label={progress.title} aria-valuemin={0} aria-valuemax={100} aria-valuenow={progress.percent} className="mt-5 h-2 overflow-hidden rounded-full bg-neutral-100"><div className="h-full rounded-full bg-neutral-900 transition-[width]" style={{ width: `${progress.percent}%` }} /></div><div className="mt-2 flex justify-between text-xs text-neutral-500"><span>Step progress</span><span>{progress.percent}%</span></div><p className="mt-4 text-xs leading-5 text-neutral-400">Keep this page open until processing finishes. Your previous report stays available if this operation fails or is cancelled.</p><div className="mt-5 flex justify-end"><Button variant="outline" disabled={progress.percent >= 98} onClick={onCancel}>{progress.percent >= 98 ? "Finishing…" : "Cancel processing"}</Button></div></div>
  </dialog>;
}
