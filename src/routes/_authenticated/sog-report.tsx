import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useMemo, useState, Fragment } from "react";
import { ArrowLeft, FileText, GraduationCap, Users } from "lucide-react";
import { SUBJECTS, computeAverage, type ClassRow, type GradeRow, type StudentRow } from "@/lib/data";
import { PdfPreviewShell } from "@/components/PdfPreviewShell";
import { DepEdHeader, DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/sog-report")({
  component: SOGReport,
});

function descriptor(avg: number | null) {
  if (avg == null) return "";
  if (avg >= 90) return "Outstanding";
  if (avg >= 85) return "Very Satisfactory";
  if (avg >= 80) return "Satisfactory";
  if (avg >= 75) return "Fairly Satisfactory";
  return "Did Not Meet";
}

function SOGReport() {
  const [classId] = useState<string>(readSchoolFormsClassId);
  const [length, setLength] = useState<"short" | "full">("full");

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
    queryFn: async () => (await supabase.from("students").select("*").eq("class_id", active!).order("last_name")).data as StudentRow[],
  });
  const { data: grades = [] } = useQuery({
    enabled: !!active,
    queryKey: ["grades-all", active],
    queryFn: async () => (await supabase.from("grades").select("*").eq("class_id", active!)).data as GradeRow[],
  });

  const rows = useMemo(() => students.map((st) => {
    const cells = SUBJECTS.map((s) => {
      const t1 = grades.find((g) => g.student_id === st.id && g.subject === s && g.term === "1")?.score ?? null;
      const t2 = grades.find((g) => g.student_id === st.id && g.subject === s && g.term === "2")?.score ?? null;
      const t3 = grades.find((g) => g.student_id === st.id && g.subject === s && g.term === "3")?.score ?? null;
      const f = computeAverage([t1, t2, t3]);
      return { t1, t2, t3, f };
    });
    const avg = computeAverage(cells.map((c) => c.f));
    return { st, cells, avg };
  }), [students, grades]);

  const withRank = useMemo(() => {
    const sorted = [...rows].filter((r) => r.avg != null).sort((a, b) => (b.avg! - a.avg!));
    const rank = new Map<string, number>();
    sorted.forEach((r, i) => rank.set(r.st.id, i + 1));
    return rank;
  }, [rows]);

  const male = rows.filter((r) => r.st.sex === "male");
  const female = rows.filter((r) => r.st.sex === "female");
  const other = rows.filter((r) => !r.st.sex);
  const fmt = (n: number | null) => (n == null ? "" : Math.round(n));
  const showAllTerms = length === "full";

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
        <div className="flex items-center gap-2 text-lg font-semibold">
          <FileText className="size-5" style={{ color: DEPED_BLUE }} />
          Summary of Grades Per Term
        </div>
      </div>

      <div
        className="relative overflow-hidden rounded-2xl border-2 px-5 py-3 shadow-sm sm:px-6 sm:py-0.5"
        style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}
      >
        <div className="text-sm font-semibold text-amber-900">
          Report Configuration
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

      <PdfPreviewShell
        fileName={`SOG_${klass?.section || "class"}.pdf`}
        docBaseName={`SOG_${klass?.section || "class"}`}
        printTargetId="sog-doc"
        length={length}
        onLengthChange={setLength}
        pageLabel="Page 1 of 1"
      >
        <div id="sog-doc" className="p-8 text-[10px] text-black" style={{ width: 1200, fontFamily: "Arial, Helvetica, sans-serif" }}>
          <DepEdHeader
            region={profile?.region}
            division={profile?.division}
            district={profile?.district}
            schoolName={profile?.school_name}
            title="SUMMARY OF GRADES PER TERM"
            subtitle={`SY ${klass?.school_year || ""} · Grade ${klass?.grade_level || ""} — ${klass?.section || ""}`}
          />

          <table className="mt-3 w-full border border-black border-collapse">
            <thead className="text-white text-[9px]" style={{ backgroundColor: DEPED_BLUE }}>
              <tr>
                <th rowSpan={2} className="border border-black px-1 py-1 w-6">#</th>
                <th rowSpan={2} className="border border-black px-1 py-1 text-left">LEARNER'S NAME</th>
                <th rowSpan={2} className="border border-black px-1 py-1 w-20">LRN</th>
                {SUBJECTS.map((s) => (
                  <th key={s} colSpan={showAllTerms ? 4 : 1} className="border border-black px-1 py-1">{s}</th>
                ))}
                <th rowSpan={2} className="border border-black px-1 py-1 w-12">FINAL AVG</th>
                <th rowSpan={2} className="border border-black px-1 py-1 w-10">RANK</th>
                <th rowSpan={2} className="border border-black px-1 py-1 w-24">DESCRIPTOR</th>
              </tr>
              <tr>
                {SUBJECTS.flatMap((s) => (
                  showAllTerms
                    ? ["T1", "T2", "T3", "F"].map((t) => (
                        <th key={s + t} className="border border-black px-1 py-0.5 w-8" style={{ backgroundColor: t === "F" ? DEPED_YELLOW : undefined, color: t === "F" ? "#111" : undefined }}>{t}</th>
                      ))
                    : [<th key={s + "F"} className="border border-black px-1 py-0.5 w-10" style={{ backgroundColor: DEPED_YELLOW, color: "#111" }}>Final</th>]
                ))}
              </tr>
            </thead>
            <tbody>
              {[
                { label: "MALE", list: male },
                { label: "FEMALE", list: female },
                ...(other.length ? [{ label: "OTHER", list: other }] : []),
              ].map((g) => (
                <Fragment key={g.label}>
                  <tr style={{ backgroundColor: "#FFF7D6" }} className="font-bold">
                    <td className="border border-black px-1 py-0.5" colSpan={3 + SUBJECTS.length * (showAllTerms ? 4 : 1) + 3}>{g.label}</td>
                  </tr>
                  {g.list.map((r, i) => (
                    <tr key={r.st.id}>
                      <td className="border border-black px-1 py-0.5 text-center">{i + 1}</td>
                      <td className="border border-black px-1 py-0.5">{r.st.last_name}, {r.st.first_name} {r.st.middle_name || ""}</td>
                      <td className="border border-black px-1 py-0.5 text-center">{r.st.lrn || ""}</td>
                      {r.cells.map((c, ci) => (
                        <Fragment key={ci}>
                          {showAllTerms && <td className="border border-black px-1 py-0.5 text-center">{fmt(c.t1)}</td>}
                          {showAllTerms && <td className="border border-black px-1 py-0.5 text-center">{fmt(c.t2)}</td>}
                          {showAllTerms && <td className="border border-black px-1 py-0.5 text-center">{fmt(c.t3)}</td>}
                          <td className="border border-black px-1 py-0.5 text-center font-semibold" style={{ backgroundColor: "#FFFBEB" }}>{fmt(c.f)}</td>
                        </Fragment>
                      ))}
                      <td className="border border-black px-1 py-0.5 text-center font-bold">{fmt(r.avg)}</td>
                      <td className="border border-black px-1 py-0.5 text-center">{withRank.get(r.st.id) ?? ""}</td>
                      <td className="border border-black px-1 py-0.5 text-center">{descriptor(r.avg)}</td>
                    </tr>
                  ))}
                </Fragment>
              ))}
              {rows.length === 0 && (
                <tr><td className="border border-black p-6 text-center text-slate-500" colSpan={3 + SUBJECTS.length * (showAllTerms ? 4 : 1) + 3}>No learners in this class.</td></tr>
              )}
            </tbody>
          </table>

          <div className="mt-8 grid grid-cols-2 gap-8 text-[11px]">
            <div>
              <div>Prepared by:</div>
              <div className="mt-8 border-t border-black pt-1 text-center font-semibold">{klass?.teacher_name || profile?.full_name || ""}</div>
              <div className="text-center">Class Adviser</div>
            </div>
            <div>
              <div>Noted by:</div>
              <div className="mt-8 border-t border-black pt-1 text-center font-semibold">{profile?.principal || ""}</div>
              <div className="text-center">School Head</div>
            </div>
          </div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}
