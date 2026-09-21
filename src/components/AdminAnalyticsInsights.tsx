import { useEffect, useMemo, useState, type ReactNode } from "react";
import { useQuery } from "@tanstack/react-query";
import {
  BarChart3,
  BookOpen,
  Crown,
  GraduationCap,
  Search,
  ShieldCheck,
  UserRound,
} from "lucide-react";
import { AnalyticsInsightsPanel } from "@/components/AnalyticsInsightsPanel";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import { supabase } from "@/integrations/supabase/client";
import type { StudentRow } from "@/lib/data";

type AdminAnalyticsClass = {
  id: string;
  teacher_id: string;
  subject: string | null;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
  created_at?: string | null;
};

type AdminAnalyticsProfile = {
  id: string;
  full_name: string | null;
  email: string | null;
  teacher_type?: string | null;
};

type TeacherTypeFilter = "all" | "class_adviser" | "subject_teacher";

function normalize(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .trim();
}

function classLabel(row: AdminAnalyticsClass) {
  const grade = row.grade_level?.trim() || "Grade";
  const section = row.section?.trim();
  return section ? `${grade} · ${section}` : grade;
}

function teacherTypeLabel(value: string | null | undefined) {
  const normalized = normalize(value);

  if (normalized === "class_adviser") return "Class Adviser";
  if (normalized === "subject_teacher") return "Subject Teacher";
  return "Teacher";
}

function sortGradeLevels(values: string[]) {
  return [...values].sort((left, right) => {
    const leftNumber = Number(left.match(/\d+/)?.[0] ?? Number.MAX_SAFE_INTEGER);
    const rightNumber = Number(
      right.match(/\d+/)?.[0] ?? Number.MAX_SAFE_INTEGER,
    );

    if (leftNumber !== rightNumber) return leftNumber - rightNumber;
    return left.localeCompare(right);
  });
}

export function AdminAnalyticsInsights() {
  const [schoolYear, setSchoolYear] = useState("all");
  const [gradeLevel, setGradeLevel] = useState("all");
  const [section, setSection] = useState("all");
  const [teacherType, setTeacherType] =
    useState<TeacherTypeFilter>("all");
  const [subject, setSubject] = useState("all");
  const [search, setSearch] = useState("");
  const [selectedClassId, setSelectedClassId] = useState<string | null>(null);

  const {
    data: directory,
    isLoading: directoryLoading,
    error: directoryError,
  } = useQuery({
    queryKey: ["admin-analytics-directory"],
    queryFn: async () => {
      const [classesResponse, profilesResponse] = await Promise.all([
        supabase
          .from("classes")
          .select(
            "id, teacher_id, subject, grade_level, section, school_year, created_at",
          )
          .order("school_year", { ascending: false })
          .order("grade_level", { ascending: true })
          .order("section", { ascending: true }),
        supabase
          .from("profiles")
          .select("id, full_name, email, teacher_type"),
      ]);

      if (classesResponse.error) throw classesResponse.error;
      if (profilesResponse.error) throw profilesResponse.error;

      return {
        classes: (classesResponse.data ?? []) as AdminAnalyticsClass[],
        profiles: (profilesResponse.data ?? []) as AdminAnalyticsProfile[],
      };
    },
  });

  const classes = directory?.classes ?? [];
  const profiles = directory?.profiles ?? [];

  const profileById = useMemo(
    () => new Map(profiles.map((profile) => [profile.id, profile])),
    [profiles],
  );

  const schoolYears = useMemo(
    () =>
      [...new Set(classes.map((row) => row.school_year).filter(Boolean) as string[])]
        .sort()
        .reverse(),
    [classes],
  );

  const gradeLevels = useMemo(
    () =>
      sortGradeLevels(
        [...new Set(classes.map((row) => row.grade_level).filter(Boolean) as string[])],
      ),
    [classes],
  );

  const sections = useMemo(
    () =>
      [...new Set(classes.map((row) => row.section).filter(Boolean) as string[])].sort(
        (a, b) => a.localeCompare(b),
      ),
    [classes],
  );

  const subjects = useMemo(
    () =>
      [...new Set(classes.map((row) => row.subject).filter(Boolean) as string[])].sort(
        (a, b) => a.localeCompare(b),
      ),
    [classes],
  );

  const filteredClasses = useMemo(() => {
    const searchTerm = normalize(search);

    return classes.filter((row) => {
      const teacher = profileById.get(row.teacher_id);
      const rowTeacherType = normalize(teacher?.teacher_type);

      if (schoolYear !== "all" && row.school_year !== schoolYear) return false;
      if (gradeLevel !== "all" && row.grade_level !== gradeLevel) return false;
      if (section !== "all" && row.section !== section) return false;
      if (subject !== "all" && row.subject !== subject) return false;
      if (teacherType !== "all" && rowTeacherType !== teacherType) return false;

      if (!searchTerm) return true;

      return [
        row.grade_level,
        row.section,
        row.subject,
        row.school_year,
        teacher?.full_name,
        teacher?.email,
        teacherTypeLabel(teacher?.teacher_type),
      ].some((value) => normalize(value).includes(searchTerm));
    });
  }, [
    classes,
    gradeLevel,
    profileById,
    schoolYear,
    search,
    section,
    subject,
    teacherType,
  ]);

  useEffect(() => {
    if (filteredClasses.length === 0) {
      setSelectedClassId(null);
      return;
    }

    if (
      !selectedClassId ||
      !filteredClasses.some((row) => row.id === selectedClassId)
    ) {
      setSelectedClassId(filteredClasses[0].id);
    }
  }, [filteredClasses, selectedClassId]);

  const selectedClass =
    filteredClasses.find((row) => row.id === selectedClassId) ??
    classes.find((row) => row.id === selectedClassId) ??
    null;

  const selectedTeacher = selectedClass
    ? profileById.get(selectedClass.teacher_id) ?? null
    : null;

  const {
    data: selectedStudents = [],
    isLoading: studentsLoading,
    error: studentsError,
  } = useQuery<StudentRow[]>({
    queryKey: ["admin-analytics-students", selectedClass?.id],
    enabled: Boolean(selectedClass?.id),
    queryFn: async () => {
      if (!selectedClass?.id) return [];

      const { data, error } = await supabase
        .from("students")
        .select("*")
        .eq("class_id", selectedClass.id)
        .order("sex", { ascending: false })
        .order("last_name", { ascending: true })
        .order("first_name", { ascending: true });

      if (error) throw error;
      return (data ?? []) as StudentRow[];
    },
  });

  const clearFilters = () => {
    setSchoolYear("all");
    setGradeLevel("all");
    setSection("all");
    setTeacherType("all");
    setSubject("all");
    setSearch("");
  };

  if (directoryLoading) {
    return <AdminAnalyticsInsightsSkeleton />;
  }

  if (directoryError) {
    return (
      <section className="rounded-2xl border border-rose-200 bg-rose-50 p-6 text-sm text-rose-700">
        {directoryError instanceof Error
          ? directoryError.message
          : "Unable to load Admin Analytics & Insights."}
      </section>
    );
  }

  return (
    <section className="space-y-5">
      <div className="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
        <div className="flex min-w-0 items-start gap-3">
          <div className="grid size-12 shrink-0 place-items-center rounded-2xl bg-rose-50 text-[var(--admin-primary)]">
            <BarChart3 className="size-6" />
          </div>
          <div className="min-w-0">
            <h1 className="text-2xl font-bold text-[var(--admin-text-heading)] sm:text-3xl">
              Analytics &amp; Insights
            </h1>
            <p className="mt-1 max-w-3xl text-sm text-[var(--admin-text-muted)]">
              View academic analytics across all classes created by Class
              Advisers and Subject Teachers.
            </p>
          </div>
        </div>

        <div className="flex max-w-sm items-start gap-3 rounded-2xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-emerald-800">
          <Crown className="mt-0.5 size-5 shrink-0" />
          <div>
            <div className="text-sm font-bold">Admin Access: All Classes</div>
            <div className="mt-0.5 text-xs leading-5 text-emerald-700">
              You can view analytics for all classes across all grade levels and
              sections.
            </div>
          </div>
        </div>
      </div>

      <div className="rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-4 shadow-[var(--admin-shadow-panel)]">
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-6">
          <FilterField label="School Year">
            <select
              value={schoolYear}
              onChange={(event) => setSchoolYear(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All School Years</option>
              {schoolYears.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Grade Level">
            <select
              value={gradeLevel}
              onChange={(event) => setGradeLevel(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Grade Levels</option>
              {gradeLevels.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Section">
            <select
              value={section}
              onChange={(event) => setSection(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Sections</option>
              {sections.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Teacher Type">
            <select
              value={teacherType}
              onChange={(event) =>
                setTeacherType(event.target.value as TeacherTypeFilter)
              }
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Teachers</option>
              <option value="class_adviser">Class Adviser</option>
              <option value="subject_teacher">Subject Teacher</option>
            </select>
          </FilterField>

          <FilterField label="Subject">
            <select
              value={subject}
              onChange={(event) => setSubject(event.target.value)}
              className="h-10 w-full rounded-xl border border-[var(--admin-border)] bg-[var(--admin-white)] px-3 text-xs text-[var(--admin-text-heading)] outline-none transition focus:border-[var(--admin-primary)]"
            >
              <option value="all">All Subjects</option>
              {subjects.map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </FilterField>

          <FilterField label="Search">
            <div className="relative">
              <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-[var(--admin-text-muted)]" />
              <Input
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                placeholder="Search class or teacher..."
                className="h-10 rounded-xl pl-9 text-xs"
              />
            </div>
          </FilterField>
        </div>

        {(schoolYear !== "all" ||
          gradeLevel !== "all" ||
          section !== "all" ||
          teacherType !== "all" ||
          subject !== "all" ||
          search.trim()) && (
          <div className="mt-3 flex justify-end">
            <Button
              type="button"
              variant="outline"
              onClick={clearFilters}
              className="h-8 rounded-lg px-3 text-xs"
            >
              Clear filters
            </Button>
          </div>
        )}
      </div>

      <div className="grid gap-4 2xl:grid-cols-[minmax(420px,0.9fr)_minmax(0,1.35fr)]">
        <div className="overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-[var(--admin-shadow-panel)]">
          <div className="flex items-start justify-between gap-3 border-b border-[var(--admin-divider)] px-4 py-4">
            <div className="flex items-start gap-3">
              <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-rose-50 text-[var(--admin-primary)]">
                <BookOpen className="size-4" />
              </div>
              <div>
                <h2 className="font-bold text-[var(--admin-text-heading)]">
                  Classes (All Classes)
                </h2>
                <p className="mt-0.5 text-xs text-[var(--admin-text-muted)]">
                  List of all classes created by Class Advisers and Subject
                  Teachers.
                </p>
              </div>
            </div>
            <span className="shrink-0 rounded-full bg-rose-50 px-2.5 py-1 text-[10px] font-bold text-[var(--admin-primary)]">
              {filteredClasses.length} class
              {filteredClasses.length === 1 ? "" : "es"}
            </span>
          </div>

          {filteredClasses.length === 0 ? (
            <div className="px-6 py-16 text-center">
              <GraduationCap className="mx-auto size-8 text-[var(--admin-text-muted)]" />
              <p className="mt-3 text-sm font-semibold text-[var(--admin-text-heading)]">
                No classes found
              </p>
              <p className="mt-1 text-xs text-[var(--admin-text-muted)]">
                Change the filters or search term.
              </p>
            </div>
          ) : (
            <div className="max-h-[860px] overflow-auto">
              <table className="w-full min-w-[680px] border-collapse text-left">
                <thead className="sticky top-0 z-10 bg-[var(--admin-surface-soft)]">
                  <tr className="border-b border-[var(--admin-divider)] text-[9px] font-bold uppercase tracking-wide text-[var(--admin-text-muted)]">
                    <th className="px-3 py-3">#</th>
                    <th className="px-3 py-3">Class</th>
                    <th className="px-3 py-3">Teacher</th>
                    <th className="px-3 py-3">Subject</th>
                    <th className="px-3 py-3">Type</th>
                    <th className="px-3 py-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredClasses.map((row, index) => {
                    const teacher = profileById.get(row.teacher_id);
                    const selected = row.id === selectedClassId;
                    const type = teacherTypeLabel(teacher?.teacher_type);

                    return (
                      <tr
                        key={row.id}
                        className={`border-b border-[var(--admin-divider)] text-xs transition last:border-b-0 ${
                          selected
                            ? "bg-rose-50/60"
                            : "hover:bg-[var(--admin-nav-hover)]"
                        }`}
                      >
                        <td className="px-3 py-3 text-[var(--admin-text-muted)]">
                          {index + 1}
                        </td>
                        <td className="px-3 py-3">
                          <div className="font-bold text-[var(--admin-text-heading)]">
                            {classLabel(row)}
                          </div>
                          <div className="mt-0.5 text-[10px] text-[var(--admin-text-muted)]">
                            SY {row.school_year || "—"}
                          </div>
                        </td>
                        <td className="px-3 py-3">
                          <div className="max-w-[150px] truncate font-medium text-[var(--admin-text-heading)]">
                            {teacher?.full_name || "Unknown Teacher"}
                          </div>
                        </td>
                        <td className="px-3 py-3 text-[var(--admin-color-735e55)]">
                          {row.subject || "—"}
                        </td>
                        <td className="px-3 py-3">
                          <span
                            className={`inline-flex rounded-full px-2.5 py-1 text-[9px] font-semibold ${
                              normalize(teacher?.teacher_type) === "class_adviser"
                                ? "bg-emerald-100 text-emerald-700"
                                : normalize(teacher?.teacher_type) ===
                                    "subject_teacher"
                                  ? "bg-sky-100 text-sky-700"
                                  : "bg-slate-100 text-slate-700"
                            }`}
                          >
                            {type}
                          </span>
                        </td>
                        <td className="px-3 py-3 text-right">
                          <Button
                            type="button"
                            size="sm"
                            onClick={() => setSelectedClassId(row.id)}
                            className="h-8 gap-1.5 rounded-lg bg-[var(--admin-primary)] px-3 text-[10px] hover:bg-[var(--admin-primary-hover)]"
                          >
                            <BarChart3 className="size-3.5" />
                            {selected ? "Viewing" : "Open Analytics"}
                          </Button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        <div className="min-w-0 overflow-hidden rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] shadow-[var(--admin-shadow-panel)]">
          <div className="flex flex-col gap-3 border-b border-[var(--admin-divider)] px-4 py-4 sm:flex-row sm:items-start sm:justify-between">
            <div className="flex min-w-0 items-start gap-3">
              <div className="grid size-9 shrink-0 place-items-center rounded-xl bg-rose-50 text-[var(--admin-primary)]">
                <BarChart3 className="size-4" />
              </div>
              <div className="min-w-0">
                <h2 className="font-bold text-[var(--admin-text-heading)]">
                  Class Analytics Preview
                </h2>
                <p className="mt-0.5 text-xs text-[var(--admin-text-muted)]">
                  Select a class and review the same analytics available to its
                  Class Adviser or Subject Teacher.
                </p>
              </div>
            </div>

            {selectedClass ? (
              <div className="shrink-0 rounded-xl border border-emerald-200 bg-emerald-50 px-3 py-2 text-right">
                <div className="text-xs font-bold text-emerald-800">
                  {classLabel(selectedClass)}
                </div>
                <div className="mt-0.5 text-[10px] text-emerald-700">
                  {selectedClass.subject || "No Subject"}
                </div>
              </div>
            ) : null}
          </div>

          <div className="p-3 sm:p-4">
            {!selectedClass ? (
              <div className="grid min-h-[420px] place-items-center rounded-xl border border-dashed border-[var(--admin-border)] bg-[var(--admin-surface-soft)] p-8 text-center">
                <div>
                  <ShieldCheck className="mx-auto size-9 text-[var(--admin-text-muted)]" />
                  <p className="mt-3 text-sm font-semibold text-[var(--admin-text-heading)]">
                    Select a class
                  </p>
                  <p className="mt-1 max-w-sm text-xs leading-5 text-[var(--admin-text-muted)]">
                    Choose any class from the list to view its term analytics,
                    proficiency distribution, and forecast panel.
                  </p>
                </div>
              </div>
            ) : studentsError ? (
              <div className="rounded-xl border border-rose-200 bg-rose-50 p-5 text-sm text-rose-700">
                {studentsError instanceof Error
                  ? studentsError.message
                  : "Unable to load learners for this class."}
              </div>
            ) : studentsLoading ? (
              <div className="space-y-3">
                <Skeleton className="h-16 w-full rounded-xl" />
                <Skeleton className="h-40 w-full rounded-xl" />
                <Skeleton className="h-72 w-full rounded-xl" />
              </div>
            ) : (
              <AnalyticsInsightsPanel
                classId={selectedClass.id}
                students={selectedStudents}
                subject={selectedClass.subject || ""}
                initialTerm="1"
                forecastData={null}
                adminView
                classLabel={classLabel(selectedClass)}
                teacherLabel={
                  selectedTeacher?.full_name ||
                  selectedTeacher?.email ||
                  "Unknown Teacher"
                }
              />
            )}
          </div>
        </div>
      </div>

      <div className="flex items-start gap-2 rounded-xl border border-sky-200 bg-sky-50 px-4 py-3 text-xs leading-5 text-sky-800">
        <UserRound className="mt-0.5 size-4 shrink-0" />
        <span>
          Admin Analytics is read-only. It reuses the existing class analytics
          calculations so the Admin sees the same actual grades, proficiency
          distribution, and model forecast output as the teacher.
        </span>
      </div>
    </section>
  );
}

function FilterField({
  label,
  children,
}: {
  label: string;
  children: ReactNode;
}) {
  return (
    <label className="block">
      <span className="mb-1.5 block text-[10px] font-bold uppercase tracking-wide text-[var(--admin-text-muted)]">
        {label}
      </span>
      {children}
    </label>
  );
}

function AdminAnalyticsInsightsSkeleton() {
  return (
    <section className="space-y-5">
      <div className="flex items-start gap-3">
        <Skeleton className="size-12 rounded-2xl" />
        <div className="space-y-2">
          <Skeleton className="h-7 w-64" />
          <Skeleton className="h-4 w-[34rem] max-w-full" />
        </div>
      </div>

      <div className="grid gap-3 rounded-2xl border border-[var(--admin-border)] bg-[var(--admin-white)] p-4 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-6">
        {Array.from({ length: 6 }).map((_, index) => (
          <Skeleton key={index} className="h-10 rounded-xl" />
        ))}
      </div>

      <div className="grid gap-4 2xl:grid-cols-[minmax(420px,0.9fr)_minmax(0,1.35fr)]">
        <Skeleton className="h-[620px] rounded-2xl" />
        <Skeleton className="h-[620px] rounded-2xl" />
      </div>
    </section>
  );
}
