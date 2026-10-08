"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { ChevronRight, Pencil, Plus, Search, X } from "lucide-react";
import { saveStrengthData } from "@/app/actions";
import {
  addCategory,
  addDesignation,
  addSubcategory,
  configurationNameKey,
  validateStrengthData,
  emptyDesignationForm,
  normalizeData,
  toggleDesignationStatus,
  type Designation,
  type StrengthData,
  updateDesignation,
} from "@/lib/strength-data";
import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import { FeedbackDialog, type Feedback } from "@/components/feedback-dialog";

type ModalMode = "category" | "subcategory" | "designation" | "edit" | null;

type FormState = {
  category: string;
  subcategory: string;
  designation: Designation;
};

const initialForm = (): FormState => ({
  category: "",
  subcategory: "",
  designation: emptyDesignationForm(),
});

export function ConfigEditor({ initialData }: { initialData: StrengthData }) {
  const [data, setData] = useState(() => normalizeData(initialData));
  const [selectedCategory, setSelectedCategory] = useState<string>(() => initialData.strengthStructure[0]?.category ?? "");
  const [form, setForm] = useState<FormState>(() => initialForm());
  const [modalMode, setModalMode] = useState<ModalMode>(null);
  const [error, setError] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);
  const [feedback, setFeedback] = useState<Feedback | null>(null);
  const [savedData, setSavedData] = useState(() => JSON.stringify(normalizeData(initialData)));
  const [savedAt, setSavedAt] = useState("");
  const [search, setSearch] = useState("");
  const [originalName, setOriginalName] = useState("");
  const dialogRef = useRef<HTMLDialogElement>(null);
  const savingRef = useRef(false);
  const dirty = JSON.stringify(data) !== savedData;

  useEffect(() => {
    if (!modalMode) return;
    const dialog = dialogRef.current;
    dialog?.showModal();
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { dialog?.close(); document.body.style.overflow = previousOverflow; };
  }, [modalMode]);

  const categories = data.strengthStructure;
  const selectedCategoryData = categories.find((category) => category.category === selectedCategory) ?? categories[0];

  const totals = useMemo(() => {
    const approved = categories.reduce(
      (sum, category) =>
        sum +
        category.subcategories.reduce(
          (subcategorySum, subcategory) =>
            subcategorySum + subcategory.designations.reduce((designationSum, designation) => designationSum + Number(designation.approvedStrength), 0),
          0,
        ),
      0,
    );
    const onRoll = categories.reduce(
      (sum, category) =>
        sum +
        category.subcategories.reduce(
          (subcategorySum, subcategory) =>
            subcategorySum + subcategory.designations.reduce((designationSum, designation) => designationSum + Number(designation.onRoll), 0),
          0,
        ),
      0,
    );
    const active = categories.reduce(
      (sum, category) =>
        sum +
        category.subcategories.reduce(
          (subcategorySum, subcategory) => subcategorySum + subcategory.designations.filter((designation) => designation.status === "Active").length,
          0,
        ),
      0,
    );
    return { approved, onRoll, active, vacancies: approved - onRoll };
  }, [categories]);

  const openAddCategory = () => {
    setModalMode("category");
    setForm(initialForm());
    setError("");
  };

  const openAddSubcategory = () => {
    if (!selectedCategoryData) return;
    setForm({ ...initialForm(), category: selectedCategoryData.category });
    setError("");
    setModalMode("subcategory");
  };

  const openAddDesignation = (subcategory: string) => {
    if (!selectedCategoryData) return;
    setForm({ ...initialForm(), category: selectedCategoryData.category, subcategory });
    setError("");
    setModalMode("designation");
  };

  const openEditDesignation = (subcategory: string, designation: Designation) => {
    if (!selectedCategoryData) return;
    setOriginalName(designation.designation);
    setForm({ category: selectedCategoryData.category, subcategory, designation: { ...designation } });
    setError("");
    setModalMode("edit");
  };

  const closeModal = () => {
    setModalMode(null);
    setError("");
    setForm(initialForm());
  };

  const submitModal = () => {
    if (savingRef.current || !modalMode) return;
    if (modalMode === "category") {
      const name = form.category.trim();
      if (!name) return setError("Category name is required.");
      if (name.length > 500) return setError("Category name must be 500 characters or fewer.");
      if (categories.some((item) => configurationNameKey(item.category) === configurationNameKey(name))) return setError("This category already exists.");
      setData(addCategory(data, name));
      setSelectedCategory(name);
      setSearch("");
    } else if (modalMode === "subcategory") {
      const name = form.subcategory.trim();
      if (!name) return setError("Subcategory name is required.");
      if (name.length > 500) return setError("Subcategory name must be 500 characters or fewer.");
      if (!selectedCategoryData) return setError("Select a category first.");
      if (selectedCategoryData.subcategories.some((item) => configurationNameKey(item.subcategory) === configurationNameKey(name))) return setError("This subcategory already exists.");
      setData(addSubcategory(data, form.category, name));
      setSearch("");
    } else if (modalMode === "designation" || modalMode === "edit") {
      const designation = { ...form.designation, designation: form.designation.designation.trim(), grade: form.designation.grade.trim() };
      if (!designation.designation || !designation.grade || !designation.cadre) return setError("Designation, grade, and cadre are required.");
      if (![designation.approvedStrength, designation.onRoll].every((value) => Number.isSafeInteger(value) && value >= 0)) return setError("Strength values must be non-negative whole numbers.");
      const group = data.strengthStructure.find((item) => item.category === form.category)?.subcategories.find((item) => item.subcategory === form.subcategory);
      if (!group || modalMode === "edit" && !group.designations.some((item) => item.designation === originalName)) return setError("This designation group changed. Close the dialog and try again.");
      const siblings = group.designations;
      if (siblings.some((item) => configurationNameKey(item.designation) === configurationNameKey(designation.designation) && (modalMode !== "edit" || item.designation !== originalName))) return setError("This designation already exists in this group.");
      setData(modalMode === "edit" ? updateDesignation(data, form.category, form.subcategory, originalName, designation) : addDesignation(data, form.category, form.subcategory, designation));
      setSearch("");
    }
    const action = modalMode === "edit" ? "Designation updated" : modalMode === "category" ? "Category added" : modalMode === "subcategory" ? "Subcategory added" : "Designation added";
    closeModal();
    setFeedback({ title: action, tone: "success", message: "Your change is ready. Select Save changes at the top of the page to save it to the strength configuration." });
  };

  const handleSave = async () => {
    if (savingRef.current || !dirty || modalMode) return;
    savingRef.current = true;
    setIsSaving(true);
    setError("");
    try {
      validateStrengthData(data);
      const result = await saveStrengthData(data, savedData);
      setData(result);
      setSavedData(JSON.stringify(result));
      setSavedAt(new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", timeZone: "Asia/Karachi" }));
      setFeedback({ title: "Configuration saved", tone: "success", message: "Your strength configuration has been saved successfully. The Home Page will use this template when processing reports." });
    } catch (error) {
      setError("The configuration could not be saved. Please try again.");
      setFeedback({ title: "Configuration could not be saved", tone: "error", message: `${error instanceof Error ? error.message : "Please retry."} Your changes remain on this page. Try Save changes again before leaving or refreshing.` });
    } finally {
      setIsSaving(false);
      savingRef.current = false;
    }
  };

  const toggleStatus = (subcategory: string, designation: string) => {
    if (!selectedCategoryData) return;
    setData(toggleDesignationStatus(data, selectedCategoryData.category, subcategory, designation));
  };

  const filteredSubcategories = selectedCategoryData?.subcategories.map((subcategory) => ({
    ...subcategory,
    designations: subcategory.designations.filter((designation) =>
      `${designation.designation} ${designation.grade} ${designation.cadre} ${subcategory.subcategory}`.toLowerCase().includes(search.toLowerCase()),
    ),
  })).filter((subcategory) => !search || subcategory.designations.length > 0) ?? [];

  return (
    <div className="config-page min-h-screen bg-white">
      <header className="border-b border-neutral-200">
        <div className="mx-auto flex max-w-[1440px] flex-wrap items-center justify-between gap-5 px-5 py-6 sm:px-8 lg:px-10">
          <div>
            <h1 className="text-2xl font-semibold tracking-tight text-neutral-950">Strength Configuration</h1>
            <p className="mt-1.5 text-sm text-neutral-500">Manage categories, designations, and workforce strength.</p>
            <p className="mt-2 text-xs text-neutral-400">1. Select a category · 2. Add or edit designations · 3. Save changes</p>
          </div>
          <div className="flex items-center gap-4">
            <span role="status" className="text-xs text-neutral-500">{dirty ? "Unsaved changes" : savedAt ? `Saved at ${savedAt}` : "All changes saved"}</span>
            <Button onClick={handleSave} disabled={isSaving || !dirty}>{isSaving ? "Saving..." : "Save changes"}</Button>
          </div>
        </div>
      </header>
      <main className="mx-auto max-w-[1440px] px-5 py-7 sm:px-8 lg:px-10"><fieldset disabled={isSaving} className="min-w-0">
        <section aria-label="Workforce overview" className="grid grid-cols-2 gap-3 lg:grid-cols-4">
          <MetricCard label="Approved strength" value={totals.approved.toLocaleString()} description="Total sanctioned positions" />
          <MetricCard label="On-roll" value={totals.onRoll.toLocaleString()} description="Current workforce" />
          <MetricCard label="Vacancies" value={totals.vacancies.toLocaleString()} description="Approved less on-roll" />
          <MetricCard label="Active designations" value={totals.active.toLocaleString()} description="Across all categories" />
        </section>
        {error && !modalMode && <div role="alert" className="mt-5 rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">{error}</div>}
        <div className="mt-8 grid items-start gap-6 lg:grid-cols-[240px_minmax(0,1fr)]">
          <aside className="min-w-0 rounded-xl border border-neutral-200 p-3 lg:sticky lg:top-6">
            <div className="flex items-center justify-between px-2 pb-3 pt-1">
              <h2 className="text-sm font-semibold">Categories</h2>
              <span className="text-xs text-neutral-400">{categories.length}</span>
            </div>
            <nav aria-label="Categories" className="flex gap-1 overflow-x-auto lg:flex-col">
              {categories.map((category) => (
                <button key={category.category} type="button" aria-current={selectedCategoryData?.category === category.category ? "page" : undefined}
                  onClick={() => { setSelectedCategory(category.category); setSearch(""); }}
                  className={cn("flex shrink-0 items-center justify-between gap-3 rounded-lg px-3 py-2.5 text-left text-sm transition-colors lg:shrink", selectedCategoryData?.category === category.category ? "bg-neutral-100 font-semibold text-neutral-950" : "text-neutral-500 hover:bg-neutral-50 hover:text-neutral-900")}>
                  <span className="break-words">{category.category}</span><ChevronRight className="h-3.5 w-3.5 shrink-0" />
                </button>
              ))}
            </nav>
            <Button variant="outline" onClick={openAddCategory} className="mt-3 w-full"><Plus className="h-4 w-4" /> Add category</Button>
          </aside>
          <div className="min-w-0">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-semibold tracking-tight">{selectedCategoryData?.category ?? "Create your first category"}</h2>
                <p className="mt-1 text-sm text-neutral-500">{selectedCategoryData ? `${selectedCategoryData.subcategories.length} subcategories / ${selectedCategoryData.subcategories.reduce((sum, item) => sum + item.designations.length, 0)} designations` : "Organize your workforce to get started."}</p>
              </div>
              {selectedCategoryData && <Button onClick={openAddSubcategory}><Plus className="h-4 w-4" /> Add subcategory</Button>}
            </div>
            {selectedCategoryData && <label className="relative mt-5 block">
              <Search className="pointer-events-none absolute left-3 top-3 h-4 w-4 text-neutral-400" />
              <input aria-label="Search designations" placeholder="Search designations, grades, or cadres..." value={search} onChange={(event) => setSearch(event.target.value)} className="h-10 w-full rounded-lg border border-neutral-200 bg-white pl-10 pr-3 text-sm outline-none focus-visible:ring-2 focus-visible:ring-neutral-400" />
            </label>}
            <div className="mt-5 space-y-5">
              {filteredSubcategories.map((subcategory) => (
                <section key={subcategory.subcategory} className="overflow-hidden rounded-xl border border-neutral-200 bg-white">
                  <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-4">
                    <div className="flex min-w-0 items-center gap-2.5"><h3 className="break-words text-sm font-semibold">{subcategory.subcategory}</h3><span className="rounded-md bg-neutral-100 px-2 py-0.5 text-xs text-neutral-500">{subcategory.designations.length}</span></div>
                    <Button variant="outline" size="sm" onClick={() => openAddDesignation(subcategory.subcategory)}><Plus className="h-3.5 w-3.5" /> Add designation</Button>
                  </div>
                  {subcategory.designations.length === 0 ? (
                    <div className="border-t border-neutral-100 px-5 py-9 text-center"><p className="text-sm font-medium">No designations yet</p><p className="mt-1 text-xs text-neutral-500">Add a designation to configure strength for this group.</p></div>
                  ) : (
                    <table className="designation-table w-full text-left text-sm">
                      <thead className="border-y border-neutral-100 bg-neutral-50 text-xs text-neutral-500"><tr>
                        <th>Designation</th><th>Grade</th><th>Cadre</th><th className="text-right">Approved</th><th className="text-right">On-roll</th><th>Status</th><th><span className="sr-only">Edit designation</span></th>
                      </tr></thead>
                      <tbody>{subcategory.designations.map((designation) => (
                        <tr key={designation.designation} className="border-t border-neutral-100 hover:bg-neutral-50/70">
                          <td data-label="Designation" className="font-medium text-neutral-900">{designation.designation}</td>
                          <td data-label="Grade">{designation.grade}</td><td data-label="Cadre">{designation.cadre}</td>
                          <td data-label="Approved" className="text-right tabular-nums">{designation.approvedStrength}</td><td data-label="On-roll" className="text-right tabular-nums">{designation.onRoll}</td>
                          <td data-label="Status"><button type="button" aria-label={`Change status of ${designation.designation}, currently ${designation.status}`} onClick={() => toggleStatus(subcategory.subcategory, designation.designation)} className={cn("inline-flex items-center gap-1.5 rounded-md px-2 py-1 text-xs font-medium focus-visible:outline-2 focus-visible:outline-offset-2", designation.status === "Active" ? "bg-emerald-50 text-emerald-700" : "bg-neutral-100 text-neutral-500")}><span className={cn("h-1.5 w-1.5 rounded-full", designation.status === "Active" ? "bg-emerald-500" : "bg-neutral-400")} />{designation.status}</button></td>
                          <td data-label="Actions" className="text-right"><Button variant="ghost" size="icon" aria-label={`Edit ${designation.designation}`} onClick={() => openEditDesignation(subcategory.subcategory, designation)}><Pencil className="h-3.5 w-3.5" /></Button></td>
                        </tr>
                      ))}</tbody>
                    </table>
                  )}
                </section>
              ))}
              {filteredSubcategories.length === 0 && <div className="rounded-xl border border-dashed border-neutral-300 px-6 py-12 text-center"><h3 className="text-sm font-medium">{search ? "No matching designations" : selectedCategoryData ? "Add your first subcategory" : "Your configuration starts here"}</h3><p className="mt-2 text-sm text-neutral-500">{search ? "Try a different designation, grade, or cadre." : "Use the add button to build your workforce structure."}</p></div>}
            </div>
          </div>
        </div>
      </fieldset></main>

      {modalMode && (
        <dialog ref={dialogRef} aria-labelledby="config-dialog-title" aria-describedby="config-dialog-description" onCancel={closeModal} className="config-dialog" onClick={(event) => { if (event.target === event.currentTarget) closeModal(); }}>
          <form onSubmit={(event) => { event.preventDefault(); submitModal(); }} className="p-5 sm:p-6">
            <div className="mb-5 flex items-start justify-between gap-4">
              <div>
                
                <h3 id="config-dialog-title" className="text-lg font-semibold tracking-tight text-neutral-950">
                  {modalMode === "category" ? "Add category" : modalMode === "subcategory" ? "Add subcategory" : modalMode === "edit" ? "Edit designation" : "Add designation"}
                </h3>
                <p id="config-dialog-description" className="mt-1.5 text-sm text-neutral-500">{modalMode === "category" ? "Create a category to organize your workforce." : modalMode === "subcategory" ? `Create a group within ${form.category}.` : `${form.category} / ${form.subcategory}`}</p>
              </div>
              <Button type="button" variant="ghost" size="icon" aria-label="Close dialog" onClick={closeModal}><X className="h-4 w-4" /></Button>
            </div>

            {modalMode === "category" && (
              <div className="space-y-3">
                <label className="block text-sm font-medium text-slate-700">Category name<input autoFocus required value={form.category} onChange={(event) => setForm({ ...form, category: event.target.value })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none ring-0 transition focus:border-neutral-400" placeholder="e.g. FINISHING" /></label>
              </div>
            )}

            {modalMode === "subcategory" && (
              <div className="space-y-3">
                <label className="block text-sm font-medium text-slate-700">Subcategory name<input autoFocus required value={form.subcategory} onChange={(event) => setForm({ ...form, subcategory: event.target.value })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400" placeholder="e.g. Pre-spinning" /></label>
              </div>
            )}

            {(modalMode === "designation" || modalMode === "edit") && (
              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <label className="block text-sm font-medium text-slate-700 sm:col-span-2">Designation<input autoFocus required value={form.designation.designation} onChange={(event) => setForm({ ...form, designation: { ...form.designation, designation: event.target.value } })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400" placeholder="e.g. Operator" /></label>
                <label className="block text-sm font-medium text-slate-700">Grade<input required value={form.designation.grade} onChange={(event) => setForm({ ...form, designation: { ...form.designation, grade: event.target.value } })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400" placeholder="E-03" /></label>
                <label className="block text-sm font-medium text-slate-700">Cadre<select value={form.designation.cadre} onChange={(event) => setForm({ ...form, designation: { ...form.designation, cadre: event.target.value } })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400"><option value="" disabled>Select cadre</option><option>Staff</option><option>Worker</option></select></label>
                <label className="block text-sm font-medium text-slate-700">Approved strength<input type="number" required min="0" step="1" max={Number.MAX_SAFE_INTEGER} value={form.designation.approvedStrength} onChange={(event) => setForm({ ...form, designation: { ...form.designation, approvedStrength: Number(event.target.value) } })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400" /></label>
                <label className="block text-sm font-medium text-slate-700">On-roll<input type="number" min="0" value={form.designation.onRoll} onChange={(event) => setForm({ ...form, designation: { ...form.designation, onRoll: Number(event.target.value) } })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400" /></label>
                <label className="block text-sm font-medium text-slate-700 sm:col-span-2">Status<select value={form.designation.status} onChange={(event) => setForm({ ...form, designation: { ...form.designation, status: event.target.value as Designation["status"] } })} className="mt-1 w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 outline-none transition focus:border-neutral-400"><option value="Active">Active</option><option value="Inactive">Inactive</option></select></label>
              </div>
            )}

            {error && <p role="alert" className="mt-4 rounded-lg bg-red-50 p-3 text-sm text-red-700">{error}</p>}
            <div className="-mx-5 -mb-5 mt-6 flex justify-end gap-2 border-t border-neutral-200 bg-neutral-50 px-5 py-4 sm:-mx-6 sm:-mb-6 sm:px-6">
              <Button type="button" variant="outline" onClick={closeModal}>Cancel</Button>
              <Button type="submit">{modalMode === "edit" ? "Update designation" : modalMode === "category" ? "Add category" : modalMode === "subcategory" ? "Add subcategory" : "Add designation"}</Button>
            </div>
          </form>
        </dialog>
      )}
      <FeedbackDialog feedback={feedback} onClose={() => setFeedback(null)} />
    </div>
  );
}

function MetricCard({ label, value, description }: { label: string; value: string; description: string }) {
  return (
    <div className="rounded-xl bg-neutral-50 p-4 sm:p-5">
      <p className="text-xs font-medium text-neutral-600">{label}</p>
      <p className="mt-3 text-3xl font-semibold tracking-tight text-neutral-950 tabular-nums">{value}</p>
      <p className="mt-1.5 text-xs text-neutral-400">{description}</p>
    </div>
  );
}
