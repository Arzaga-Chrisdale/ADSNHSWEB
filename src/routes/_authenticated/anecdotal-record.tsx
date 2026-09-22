import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useMemo, useState } from "react";
import { ArrowLeft, Plus, Trash2, GraduationCap, Users } from "lucide-react";
import type { ClassRow, StudentRow } from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import { StudentMultiSelect } from "@/components/StudentMultiSelect";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/anecdotal-record")({
  component: AnecdotalRecordPage,
});

type Row = { date: string; incident: string; action: string; remarks: string };

function AnecdotalRecordPage() {
  const [classId] = useState<string>(readSchoolFormsClassId);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [rowsByStudent, setRowsByStudent] = useState<Record<string, Row[]>>({});

  const { data: profile } = useQuery({
    queryKey: ["profile"],
    queryFn: async () => (await supabase.from("profiles").select("*").maybeSingle()).data as any,
  });
  const { data: classes = [] } = useQuery({
    queryKey: ["classes"],
    queryFn: async () => (await supabase.from("classes").select("*").order("created_at", { ascending: false })).data as ClassRow[],
  });

  const selectedSchoolFormsClass = classes.find((item) => item.id === classId);
  const active = selectedSchoolFormsClass?.id || classes[0]?.id;
  const klass = classes.find((c) => c.id === active);

  const { data: students = [] } = useQuery({
    enabled: !!active,
    queryKey: ["students", active],
    queryFn: async () =>
      (await supabase.from("students").select("*").eq("class_id", active!).order("last_name")).data as StudentRow[],
  });

  const targets = useMemo(
    () => (selected.size ? students.filter((s) => selected.has(s.id)) : students.slice(0, 1)),
    [students, selected],
  );

  const rowsFor = (id: string) => rowsByStudent[id] ?? [{ date: "", incident: "", action: "", remarks: "" }];

  const setRows = (id: string, next: Row[]) =>
    setRowsByStudent((prev) => ({ ...prev, [id]: next }));

  return (
    <div className="space-y-4 pb-8">
      <div className="flex flex-wrap items-center gap-3">
        <Link
          to="/school-forms"
          className="inline-flex items-center gap-1.5 rounded-lg border-2 px-3 py-1.5 text-sm font-medium text-white shadow-sm hover:opacity-90"
          style={{ backgroundColor: DEPED_BLUE, borderColor: DEPED_BLUE }}
        >
          <ArrowLeft className="size-4" /> Back to Forms
        </Link>
        <div>
          <div className="flex items-center gap-2 text-lg font-semibold">📕 Anecdotal Record</div>
          <div className="text-xs text-muted-foreground">
            Grade {klass?.grade_level || "—"} - {klass?.section || "—"} · {selected.size} selected
          </div>
        </div>
      </div>

      <div
        className="relative overflow-hidden rounded-2xl border-2 px-5 py-3 shadow-sm sm:px-6 sm:py-0.5"
        style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}
      >
        <div className="text-sm font-semibold text-amber-900">
          Form Configuration
        </div>

        <div className="mx-auto mt-1 flex max-w-2xl flex-col items-center">
          <div className="flex w-full items-center justify-center gap-3 sm:gap-5">
            <div className="hidden h-px w-20 bg-amber-400 sm:block" />
            <div
              className="grid size-9 shrink-0 place-items-center rounded-full"
              style={{ backgroundColor: "#FEF0B6" }}
            >
              <GraduationCap className="size-5 text-amber-950" />
            </div>
            <div className="hidden h-px w-20 bg-amber-400 sm:block" />
          </div>

          <div className="mt-1 text-center text-xs font-bold uppercase tracking-[0.32em] text-amber-950">
            Class
          </div>

          <div className="mt-1 flex w-full max-w-[300px] items-center gap-1.5 rounded-lg border border-amber-300 bg-white/75 px-2 py-2 shadow-sm">
            <div className="grid size-5 shrink-0 place-items-center rounded-md bg-amber-50 text-amber-700">
              <Users className="size-3" />
            </div>

            <div className="min-w-0 flex-1 text-center">
              <div className="truncate text-[11px] font-bold text-amber-950">
                {klass
                  ? `${klass.grade_level || "—"} · ${klass.subject || "—"} · ${klass.section || "—"}`
                  : "No class selected"}
              </div>
            </div>
          </div>

          <p className="mt-0.5 text-center text-[9px] text-amber-900/60">
            Selected class for this form
          </p>
        </div>
      </div>

      <div className="rounded-2xl border bg-card p-4">
        <StudentMultiSelect students={students} selected={selected} onChange={setSelected} color={DEPED_BLUE} />
        {students.length === 0 && <div className="mt-2 text-xs text-muted-foreground">No students in this class.</div>}
      </div>

      <PdfPreviewShell
        fileName={`Anecdotal_${klass?.section || "class"}.pdf`}
        docBaseName={`Anecdotal_${klass?.section || "class"}`}
        printTargetId="anecdotal-doc"
      >
        <div id="anecdotal-doc" className="p-6 text-[11px] text-black" style={{ width: 1000, fontFamily: "Arial, Helvetica, sans-serif" }}>
          <div className="space-y-8">
            {targets.map((st) => {
              const rows = rowsFor(st.id);
              return (
                <div key={st.id} className="border-2 border-black p-4">
                  <div className="text-center">
                    <div className="text-[13px] font-bold">ANECDOTAL RECORD</div>
                    <div className="text-[11px]">{profile?.school_name || "School Name"}</div>
                  </div>
                  <div className="mt-2 text-[10px]">
                    <b>Name:</b> {st.last_name}, {st.first_name} | <b>Grade & Section:</b> Grade {klass?.grade_level || ""} - {klass?.section || ""} | <b>SY:</b> <u>{klass?.school_year || "\u00A0\u00A0\u00A0\u00A0\u00A0"}</u>
                  </div>
                  <table className="mt-2 w-full border border-black border-collapse text-[10px]">
                    <thead>
                      <tr className="text-white" style={{ backgroundColor: DEPED_BLUE }}>
                        <th className="border border-black px-2 py-1 w-28">Date</th>
                        <th className="border border-black px-2 py-1">Incident / Observation</th>
                        <th className="border border-black px-2 py-1">Action Taken</th>
                        <th className="border border-black px-2 py-1">Remarks</th>
                        <th className="border border-black px-2 py-1 w-8 print:hidden">×</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r, i) => (
                        <tr key={i}>
                          <td className="border border-black px-2 py-1">
                            <input
                              type="date"
                              value={r.date}
                              onChange={(e) => setRows(st.id, rows.map((x, j) => (j === i ? { ...x, date: e.target.value } : x)))}
                              className="w-full bg-transparent outline-none"
                            />
                          </td>
                          <td className="border border-black px-2 py-1">
                            <textarea
                              value={r.incident}
                              onChange={(e) => setRows(st.id, rows.map((x, j) => (j === i ? { ...x, incident: e.target.value } : x)))}
                              className="min-h-[40px] w-full resize-y bg-transparent outline-none"
                            />
                          </td>
                          <td className="border border-black px-2 py-1">
                            <textarea
                              value={r.action}
                              onChange={(e) => setRows(st.id, rows.map((x, j) => (j === i ? { ...x, action: e.target.value } : x)))}
                              className="min-h-[40px] w-full resize-y bg-transparent outline-none"
                            />
                          </td>
                          <td className="border border-black px-2 py-1">
                            <textarea
                              value={r.remarks}
                              onChange={(e) => setRows(st.id, rows.map((x, j) => (j === i ? { ...x, remarks: e.target.value } : x)))}
                              className="min-h-[40px] w-full resize-y bg-transparent outline-none"
                            />
                          </td>
                          <td className="border border-black px-1 py-1 text-center print:hidden">
                            <button
                              onClick={() => setRows(st.id, rows.filter((_, j) => j !== i))}
                              className="text-red-600 hover:text-red-800"
                              aria-label="Remove row"
                            >
                              <Trash2 className="mx-auto size-3.5" />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <div className="mt-2 print:hidden">
                    <button
                      onClick={() => setRows(st.id, [...rows, { date: "", incident: "", action: "", remarks: "" }])}
                      className="inline-flex items-center gap-1 rounded-md border-2 px-3 py-1 text-[11px] font-semibold"
                      style={{ borderColor: DEPED_BLUE, color: DEPED_BLUE, backgroundColor: "#EEF4FF" }}
                    >
                      <Plus className="size-3" /> Add Row
                    </button>
                  </div>
                  <div className="mt-6 grid grid-cols-2 gap-6 text-[9px]">
                    <div className="text-center">
                      <div className="mx-auto w-40 border-t border-black">&nbsp;</div>
                      Class Adviser
                    </div>
                    <div className="text-center">
                      <div className="mx-auto w-40 border-t border-black">&nbsp;</div>
                      School Head
                    </div>
                  </div>
                </div>
              );
            })}
            {targets.length === 0 && <div className="p-10 text-center text-slate-500">Select at least one student to preview.</div>}
          </div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}
