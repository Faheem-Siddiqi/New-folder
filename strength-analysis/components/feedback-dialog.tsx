"use client";

import { useEffect, useId, useRef } from "react";
import { Check, AlertCircle, X } from "lucide-react";
import { Button } from "./ui/button";

export type Feedback = { title: string; message: string; tone: "success" | "error" | "warning" };
export function FeedbackDialog({ feedback, onClose }: { feedback: Feedback | null; onClose: () => void }) {
  const ref = useRef<HTMLDialogElement>(null);
  const id = useId();
  useEffect(() => {
    if (!feedback) return;
    const dialog = ref.current;
    dialog?.showModal();
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { dialog?.close(); document.body.style.overflow = overflow; };
  }, [feedback]);
  if (!feedback) return null;
  return <dialog ref={ref} className="config-dialog" aria-labelledby={`${id}-title`} aria-describedby={`${id}-description`} onCancel={onClose}>
    <div className="p-6"><div className="flex items-start justify-between"><div className={`flex h-10 w-10 items-center justify-center rounded-full ${feedback.tone === "success" ? "bg-emerald-50 text-emerald-700" : feedback.tone === "error" ? "bg-red-50 text-red-600" : "bg-amber-50 text-amber-700"}`}>{feedback.tone === "success" ? <Check size={20} /> : <AlertCircle size={20} />}</div><Button variant="ghost" size="icon" aria-label="Close message" onClick={onClose}><X size={16} /></Button></div><h2 id={`${id}-title`} className="mt-4 text-lg font-semibold tracking-tight">{feedback.title}</h2><p id={`${id}-description`} className="mt-2 whitespace-pre-line text-sm leading-6 text-neutral-500">{feedback.message}</p><div className="mt-6 flex justify-end"><Button onClick={onClose}>Got it</Button></div></div>
  </dialog>;
}
