import { createFileRoute } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import {
  ArrowLeftRight,
  BookOpen,
  CalendarDays,
  ChevronLeft, 
  ChevronRight,
  Eye,
  GraduationCap,
  Lightbulb,
  LogIn,
  LogOut,
  Search,
  School,
  UserRound,
} from "lucide-react";
import { useEffect, useMemo, useState, type ReactNode } from "react";

import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Skeleton } from "@/components/ui/skeleton";

export const Route = createFileRoute("/_authenticated/transfers")({
  component: TransfersPage,
});

type TransferType = "in" | "out";
type TransferTab = "all" | TransferType;

type TransferRecord = {
  id: string;
  student_id: string;
  transfer_type: TransferType;
  from_class_id: string | null;
  to_class_id: string | null;
  school_year: string | null;
  effective_term: string | null;
  transfer_date: string;
  previous_school: string | null;
  destination_school: string | null;
  reason: string | null;
  remarks: string | null;
  created_by: string | null;
  created_at: string;
};

type TransferStudent = {
  id: string;
  lrn: string | null;
  last_name: string | null;
  first_name: string | null;
  middle_name: string | null;
  sex: string | null;
  enrollment_status?: string | null;
  is_active?: boolean | null;
};

type TransferClass = {
  id: string;
  school_name: string | null;
  school_year: string | null;
  grade_level: string | null;
  section: string | null;
  subject: string | null;
  teacher_id: string | null;
};

type TransferProfile = {
  id: string;
  full_name: string | null;
  email: string | null;
  teacher_type?: string | null;
};

type TransferViewRow = {
  record: TransferRecord;
  student: TransferStudent | null;
  klass: TransferClass | null;
  creator: TransferProfile | null;
  requestNo: string;
  learnerName: string;
  lrn: string;
  gradeSection: string;
  subject: string;
  schoolYear: string;
  counterpartSchool: string;
  classSchool: string;
  recordedBy: string;
};

function normalize(value: unknown) {
  return String(value ?? "").trim().toLowerCase();
}

function displayLearnerName(student: TransferStudent | null) {
  if (!student) return "Unknown learner";

  const givenNames = [student.first_name, student.middle_name]
    .filter(Boolean)
    .join(" ");

  return [student.last_name, givenNames].filter(Boolean).join(", ") || "Unknown learner";
}

function formatDate(value: string | null | undefined) {
  if (!value) return "—";

  const date = new Date(`${value.slice(0, 10)}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;

  return new Intl.DateTimeFormat("en-PH", {
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(date);
}

function buildRequestNo(record: TransferRecord) {
  const prefix = record.transfer_type === "in" ? "TIN" : "TOUT";
  const year = record.transfer_date?.slice(0, 4) || "----";
  return `${prefix}-${year}-${record.id.slice(0, 6).toUpperCase()}`;
}

function uniqueSorted(values: Array<string | null | undefined>) {
  return Array.from(
    new Set(values.map((value) => String(value ?? "").trim()).filter(Boolean)),
  ).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function TransfersPage() {
  const [activeTab, setActiveTab] = useState<TransferTab>("all");
  const [schoolYear, setSchoolYear] = useState("all");
  const [gradeLevel, setGradeLevel] = useState("all");
  const [section, setSection] = useState("all");
  const [subject, setSubject] = useState("all");
  const [search, setSearch] = useState("");
  const [selectedTransfer, setSelectedTransfer] =
    useState<TransferViewRow | null>(null);
  const [page, setPage] = useState(1);

  const { data: viewer, isLoading: viewerLoading } = useQuery({
    queryKey: ["transfer-page-viewer"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) return null;

      const { data: profile, error: profileError } = await (supabase as any)
        .from("profiles")
        .select("id, full_name, email, teacher_type")
        .eq("id", user.id)
        .maybeSingle();

      if (profileError) throw profileError;

      return {
        id: user.id,
        full_name: profile?.full_name ?? null,
        email: profile?.email ?? user.email ?? null,
        // Preserve compatibility with accounts created before teacher_type
        // existed: only an explicit subject_teacher is treated as non-adviser.
        teacher_type:
          profile?.teacher_type === "subject_teacher"
            ? "subject_teacher"
            : "class_adviser",
      } as TransferProfile;
    },
  });

  const {
    data: transferRows = [],
    isLoading: transfersLoading,
    error: transfersError,
  } = useQuery({
    queryKey: ["student-transfers"],
    enabled: viewer?.teacher_type === "class_adviser",
    queryFn: async () => {
      const { data: transfers, error: transferError } = await (supabase as any)
        .from("student_transfers")
        .select("*")
        .order("transfer_date", { ascending: false })
        .order("created_at", { ascending: false });

      if (transferError) throw transferError;

      const records = (transfers ?? []) as TransferRecord[];

      const studentIds = Array.from(
        new Set(records.map((row) => row.student_id).filter(Boolean)),
      );
      const classIds = Array.from(
        new Set(
          records
            .flatMap((row) => [row.from_class_id, row.to_class_id])
            .filter((value): value is string => Boolean(value)),
        ),
      );
      const creatorIds = Array.from(
        new Set(
          records
            .map((row) => row.created_by)
            .filter((value): value is string => Boolean(value)),
        ),
      );

      const [
        { data: studentRows, error: studentError },
        { data: classRows, error: classError },
        { data: creatorRows, error: creatorError },
      ] = await Promise.all([
        studentIds.length
          ? (supabase as any)
              .from("students")
              .select(
                "id, lrn, last_name, first_name, middle_name, sex, enrollment_status, is_active",
              )
              .in("id", studentIds)
          : Promise.resolve({ data: [], error: null }),
        classIds.length
          ? (supabase as any)
              .from("classes")
              .select(
                "id, school_name, school_year, grade_level, section, subject, teacher_id",
              )
              .in("id", classIds)
          : Promise.resolve({ data: [], error: null }),
        creatorIds.length
          ? (supabase as any)
              .from("profiles")
              .select("id, full_name, email, teacher_type")
              .in("id", creatorIds)
          : Promise.resolve({ data: [], error: null }),
      ]);

      if (studentError) throw studentError;
      if (classError) throw classError;
      if (creatorError) throw creatorError;

      const studentById = new Map<string, TransferStudent>(
        ((studentRows ?? []) as TransferStudent[]).map((row) => [row.id, row]),
      );

      const classById = new Map<string, TransferClass>(
        ((classRows ?? []) as TransferClass[]).map((row) => [row.id, row]),
      );

      const creatorById = new Map<string, TransferProfile>(
        ((creatorRows ?? []) as TransferProfile[]).map((row) => [row.id, row]),
      );

      return records.map((record): TransferViewRow => {
        const student = studentById.get(record.student_id) ?? null;

        const classId =
          record.transfer_type === "in"
            ? record.to_class_id ?? record.from_class_id
            : record.from_class_id ?? record.to_class_id;

        const klass = classId ? classById.get(classId) ?? null : null;
        const creator = record.created_by
          ? creatorById.get(record.created_by) ?? null
          : null;

        const counterpartSchool =
          record.transfer_type === "in"
            ? record.previous_school || "Previous school not specified"
            : record.destination_school || "Receiving school not specified";

        return {
          record,
          student,
          klass,
          creator,
          requestNo: buildRequestNo(record),
          learnerName: displayLearnerName(student),
          lrn: student?.lrn || "—",
          gradeSection:
            [klass?.grade_level, klass?.section].filter(Boolean).join(" - ") ||
            "—",
          subject: klass?.subject || "—",
          schoolYear:
            record.school_year || klass?.school_year || "No school year",
          counterpartSchool,
          classSchool: klass?.school_name || "Agusan del Sur NHS",
          recordedBy:
            creator?.full_name || creator?.email || "Class Adviser",
        };
      });
    },
    refetchOnWindowFocus: true,
  });

  const schoolYears = useMemo(
    () => uniqueSorted(transferRows.map((row) => row.schoolYear)).reverse(),
    [transferRows],
  );

  const gradeLevels = useMemo(
    () => uniqueSorted(transferRows.map((row) => row.klass?.grade_level)),
    [transferRows],
  );

  const sections = useMemo(
    () => uniqueSorted(transferRows.map((row) => row.klass?.section)),
    [transferRows],
  );

  const subjects = useMemo(
    () => uniqueSorted(transferRows.map((row) => row.subject)),
    [transferRows],
  );

  const filteredRows = useMemo(() => {
    const query = normalize(search);

    return transferRows.filter((row) => {
      const matchesTab =
        activeTab === "all" || row.record.transfer_type === activeTab;

      const matchesSchoolYear =
        schoolYear === "all" || row.schoolYear === schoolYear;

      const matchesGrade =
        gradeLevel === "all" || row.klass?.grade_level === gradeLevel;

      const matchesSection =
        section === "all" || row.klass?.section === section;

      const matchesSubject =
        subject === "all" || row.subject === subject;

      const matchesSearch =
        !query ||
        [
          row.requestNo,
          row.learnerName,
          row.lrn,
          row.gradeSection,
          row.subject,
          row.counterpartSchool,
          row.recordedBy,
        ].some((value) => normalize(value).includes(query));

      return (
        matchesTab &&
        matchesSchoolYear &&
        matchesGrade &&
        matchesSection &&
        matchesSubject &&
        matchesSearch
      );
    });
  }, [
    activeTab,
    gradeLevel,
    schoolYear,
    search,
    section,
    subject,
    transferRows,
  ]);

  useEffect(() => {
    setPage(1);
  }, [activeTab, schoolYear, gradeLevel, section, subject, search]);

  const pageSize = 10;
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const pageRows = filteredRows.slice(
    (safePage - 1) * pageSize,
    safePage * pageSize,
  );

  if (viewerLoading) {
    return <TransfersSkeleton />;
  }

  if (!viewer || viewer.teacher_type !== "class_adviser") {
    return (
      <div className="mx-auto max-w-3xl py-14">
        <div className="rounded-3xl border bg-card p-8 text-center shadow-sm">
          <div className="mx-auto grid size-14 place-items-center rounded-2xl bg-amber-50 text-amber-700">
            <ArrowLeftRight className="size-7" />
          </div>
          <h1 className="mt-4 text-2xl font-bold">
            Class Adviser access only
          </h1>
          <p className="mt-2 text-sm text-muted-foreground">
            The Transfer In &amp; Transfer Out history page is available to
            Class Adviser accounts.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-4 pb-8">
      <section className="grid gap-4 xl:grid-cols-[1fr_330px]">
        <div className="relative overflow-hidden rounded-3xl border border-amber-200/80 bg-gradient-to-r from-[#fffdf9] via-[#fffaf1] to-[#fff4e6] px-5 py-5 shadow-sm">
          <div className="pointer-events-none absolute -right-10 -top-12 size-40 rounded-full border-[18px] border-amber-400/10" />
          <div className="pointer-events-none absolute right-20 top-4 rotate-[-7deg] font-serif text-sm font-semibold italic text-rose-400/45">
            Together
            <br />
            for Greater Learners
          </div>

          <div className="relative flex items-center gap-4">
            <div className="grid size-12 shrink-0 place-items-center rounded-2xl border border-rose-200 bg-white/80 text-rose-600 shadow-sm">
              <ArrowLeftRight className="size-6" />
            </div>

            <div>
              <h1 className="text-2xl font-bold tracking-tight">
                Transfer In &amp; Transfer Out
              </h1>
              <p className="mt-1 text-sm text-muted-foreground">
                View incoming and outgoing learner transfer records for your
                classes.
              </p>
            </div>
          </div>
        </div>

        <div className="flex items-start gap-3 rounded-2xl border border-emerald-200 bg-emerald-50/70 px-4 py-4 text-sm text-emerald-900">
          <Lightbulb className="mt-0.5 size-5 shrink-0 text-emerald-600" />
          <p className="leading-6">
            Class Advisers can review transfer history and monitor learner
            movement between classes or schools. This page is view-only.
          </p>
        </div>
      </section>

      <section className="flex flex-wrap gap-2">
        <TransferTabButton
          active={activeTab === "all"}
          label="All Requests"
          icon={<BookOpen className="size-4" />}
          onClick={() => setActiveTab("all")}
        />
        <TransferTabButton
          active={activeTab === "in"}
          label="Transfer In"
          icon={<LogIn className="size-4" />}
          onClick={() => setActiveTab("in")}
          activeClass="bg-emerald-600 text-white border-emerald-600"
        />
        <TransferTabButton
          active={activeTab === "out"}
          label="Transfer Out"
          icon={<LogOut className="size-4" />}
          onClick={() => setActiveTab("out")}
          activeClass="bg-red-600 text-white border-red-600"
        />
      </section>

      <section className="rounded-2xl border bg-card p-3 shadow-sm">
        <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[1fr_1fr_1fr_1.15fr_1.55fr]">
          <FilterField label="School Year">
            <Select value={schoolYear} onValueChange={setSchoolYear}>
              <SelectTrigger className="h-10">
                <SelectValue placeholder="All School Years" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All School Years</SelectItem>
                {schoolYears.map((value) => (
                  <SelectItem key={value} value={value}>
                    {value}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </FilterField>

          <FilterField label="Grade Level">
            <Select value={gradeLevel} onValueChange={setGradeLevel}>
              <SelectTrigger className="h-10">
                <SelectValue placeholder="All Grade Levels" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Grade Levels</SelectItem>
                {gradeLevels.map((value) => (
                  <SelectItem key={value} value={value}>
                    {value}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </FilterField>

          <FilterField label="Section">
            <Select value={section} onValueChange={setSection}>
              <SelectTrigger className="h-10">
                <SelectValue placeholder="All Sections" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Sections</SelectItem>
                {sections.map((value) => (
                  <SelectItem key={value} value={value}>
                    {value}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </FilterField>

          <FilterField label="Subject">
            <Select value={subject} onValueChange={setSubject}>
              <SelectTrigger className="h-10">
                <SelectValue placeholder="All Subjects" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Subjects</SelectItem>
                {subjects.map((value) => (
                  <SelectItem key={value} value={value}>
                    {value}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </FilterField>

          <FilterField label="Search">
            <div className="relative">
              <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
              <Input
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                placeholder="Search learner name, LRN, request no..."
                className="h-10 pl-9"
              />
            </div>
          </FilterField>
        </div>
      </section>

      <section className="overflow-hidden rounded-2xl border bg-card shadow-sm">
        {transfersError ? (
          <div className="p-8 text-center">
            <div className="font-semibold text-destructive">
              Unable to load transfer records.
            </div>
            <p className="mt-2 text-sm text-muted-foreground">
              {transfersError instanceof Error
                ? transfersError.message
                : "Check that the student transfer migration has been applied."}
            </p>
          </div>
        ) : transfersLoading ? (
          <div className="space-y-2 p-4">
            {Array.from({ length: 8 }).map((_, index) => (
              <Skeleton key={index} className="h-11 w-full" />
            ))}
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="min-w-[1320px] w-full text-sm">
                <thead className="border-b bg-[#fffaf4]">
                  <tr className="text-left text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
                    <th className="px-4 py-3">#</th>
                    <th className="px-4 py-3">Request No.</th>
                    <th className="px-4 py-3">Learner Name</th>
                    <th className="px-4 py-3">LRN</th>
                    <th className="px-4 py-3">Grade &amp; Section</th>
                    <th className="px-4 py-3">Subject</th>
                    <th className="px-4 py-3">Transfer Type</th>
                    <th className="px-4 py-3">From / To School</th>
                    <th className="px-4 py-3">Recorded By</th>
                    <th className="px-4 py-3">Date</th>
                    <th className="px-4 py-3">Status</th>
                    <th className="px-4 py-3 text-center">Action</th>
                  </tr>
                </thead>

                <tbody>
                  {pageRows.map((row, index) => (
                    <tr
                      key={row.record.id}
                      className="border-b last:border-b-0 hover:bg-muted/20"
                    >
                      <td className="px-4 py-3 text-muted-foreground">
                        {(safePage - 1) * pageSize + index + 1}
                      </td>
                      <td className="px-4 py-3 font-medium text-slate-600">
                        {row.requestNo}
                      </td>
                      <td className="px-4 py-3 font-medium">
                        {row.learnerName}
                      </td>
                      <td className="px-4 py-3 text-muted-foreground">
                        {row.lrn}
                      </td>
                      <td className="px-4 py-3">{row.gradeSection}</td>
                      <td className="px-4 py-3">{row.subject}</td>
                      <td className="px-4 py-3">
                        <TransferTypeBadge type={row.record.transfer_type} />
                      </td>
                      <td className="px-4 py-3">
                        <div className="max-w-[220px] truncate font-medium">
                          {row.counterpartSchool}
                        </div>
                        <div className="text-[11px] text-muted-foreground">
                          {row.record.transfer_type === "in"
                            ? "Previous school"
                            : "Receiving school"}
                        </div>
                      </td>
                      <td className="px-4 py-3">{row.recordedBy}</td>
                      <td className="px-4 py-3">
                        {formatDate(row.record.transfer_date)}
                      </td>
                      <td className="px-4 py-3">
                        <span className="inline-flex rounded-full bg-blue-50 px-2.5 py-1 text-xs font-semibold text-blue-700">
                          Recorded
                        </span>
                      </td>
                      <td className="px-4 py-3 text-center">
                        <Button
                          type="button"
                          size="sm"
                          onClick={() => setSelectedTransfer(row)}
                          className="bg-[#8f2222] text-white hover:bg-[#761b1b]"
                        >
                          <Eye className="mr-1.5 size-4" />
                          View
                        </Button>
                      </td>
                    </tr>
                  ))}

                  {!pageRows.length && (
                    <tr>
                      <td
                        colSpan={12}
                        className="px-6 py-16 text-center text-sm text-muted-foreground"
                      >
                        No matching transfer records found.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="flex flex-wrap items-center justify-between gap-3 border-t px-4 py-4">
              <div className="text-sm text-muted-foreground">
                {filteredRows.length === 0
                  ? "Showing 0 transfer records"
                  : `Showing ${(safePage - 1) * pageSize + 1} to ${Math.min(
                      safePage * pageSize,
                      filteredRows.length,
                    )} of ${filteredRows.length} transfer records`}
              </div>

              <div className="flex items-center gap-2">
                <Button
                  type="button"
                  variant="outline"
                  size="icon"
                  disabled={safePage <= 1}
                  onClick={() => setPage((value) => Math.max(1, value - 1))}
                >
                  <ChevronLeft className="size-4" />
                </Button>

                <span className="grid min-w-9 place-items-center rounded-lg bg-primary px-3 py-2 text-xs font-semibold text-primary-foreground">
                  {safePage}
                </span>

                <span className="text-xs text-muted-foreground">
                  of {totalPages}
                </span>

                <Button
                  type="button"
                  variant="outline"
                  size="icon"
                  disabled={safePage >= totalPages}
                  onClick={() =>
                    setPage((value) => Math.min(totalPages, value + 1))
                  }
                >
                  <ChevronRight className="size-4" />
                </Button>
              </div>
            </div>
          </>
        )}
      </section>

      <TransferDetailsDialog
        transfer={selectedTransfer}
        onClose={() => setSelectedTransfer(null)}
      />
    </div>
  );
}

function TransferTabButton({
  active,
  label,
  icon,
  onClick,
  activeClass = "bg-[#8f2222] text-white border-[#8f2222]",
}: {
  active: boolean;
  label: string;
  icon: ReactNode;
  onClick: () => void;
  activeClass?: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`inline-flex min-w-44 items-center justify-center gap-2 rounded-xl border px-5 py-2.5 text-sm font-semibold shadow-sm transition ${
        active
          ? activeClass
          : "bg-card text-foreground hover:bg-muted/40"
      }`}
    >
      {icon}
      {label}
    </button>
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
    <div>
      <div className="mb-1.5 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
        {label}
      </div>
      {children}
    </div>
  );
}

function TransferTypeBadge({ type }: { type: TransferType }) {
  if (type === "in") {
    return (
      <span className="inline-flex items-center gap-1 rounded-full bg-emerald-50 px-2.5 py-1 text-xs font-semibold text-emerald-700">
        <LogIn className="size-3.5" />
        Transfer In
      </span>
    );
  }

  return (
    <span className="inline-flex items-center gap-1 rounded-full bg-orange-50 px-2.5 py-1 text-xs font-semibold text-orange-700">
      <LogOut className="size-3.5" />
      Transfer Out
    </span>
  );
}

function TransferDetailsDialog({
  transfer,
  onClose,
}: {
  transfer: TransferViewRow | null;
  onClose: () => void;
}) {
  return (
    <Dialog open={Boolean(transfer)} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="sm:max-w-[650px]">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2">
            <Eye className="size-5 text-primary" />
            Transfer Details
          </DialogTitle>
        </DialogHeader>

        {transfer && (
          <div className="space-y-5">
            <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border bg-muted/20 p-4">
              <div>
                <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                  Request No.
                </div>
                <div className="mt-1 font-bold">{transfer.requestNo}</div>
              </div>
              <TransferTypeBadge type={transfer.record.transfer_type} />
            </div>

            <div className="grid gap-3 sm:grid-cols-2">
              <DetailCard
                icon={<UserRound className="size-4" />}
                label="Learner"
                value={transfer.learnerName}
              />
              <DetailCard
                icon={<GraduationCap className="size-4" />}
                label="LRN"
                value={transfer.lrn}
              />
              <DetailCard
                icon={<BookOpen className="size-4" />}
                label="Grade & Section"
                value={transfer.gradeSection}
              />
              <DetailCard
                icon={<BookOpen className="size-4" />}
                label="Subject"
                value={transfer.subject}
              />
              <DetailCard
                icon={<CalendarDays className="size-4" />}
                label="School Year"
                value={transfer.schoolYear}
              />
              <DetailCard
                icon={<CalendarDays className="size-4" />}
                label="Effective Term"
                value={transfer.record.effective_term || "—"}
              />
              <DetailCard
                icon={<CalendarDays className="size-4" />}
                label="Transfer Date"
                value={formatDate(transfer.record.transfer_date)}
              />
              <DetailCard
                icon={<UserRound className="size-4" />}
                label="Recorded By"
                value={transfer.recordedBy}
              />
            </div>

            <div className="rounded-2xl border p-4">
              <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                <School className="size-4" />
                {transfer.record.transfer_type === "in"
                  ? "Previous School"
                  : "Receiving School"}
              </div>
              <div className="mt-2 font-semibold">
                {transfer.counterpartSchool}
              </div>
            </div>

            <div className="rounded-2xl border p-4">
              <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                Reason / Remarks
              </div>
              <div className="mt-2 text-sm leading-6">
                {transfer.record.reason ||
                  transfer.record.remarks ||
                  "No additional remarks."}
              </div>
            </div>

            <div className="rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-800">
              This page is read-only. Transfer In and Transfer Out actions are
              performed from the learner&apos;s E-Class Record.
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}

function DetailCard({
  icon,
  label,
  value,
}: {
  icon: ReactNode;
  label: string;
  value: string;
}) {
  return (
    <div className="rounded-2xl border bg-muted/10 p-4">
      <div className="flex items-center gap-2 text-[11px] font-semibold uppercase tracking-wide text-muted-foreground">
        {icon}
        {label}
      </div>
      <div className="mt-2 font-medium">{value || "—"}</div>
    </div>
  );
}

function TransfersSkeleton() {
  return (
    <div className="space-y-4">
      <Skeleton className="h-24 w-full rounded-3xl" />
      <div className="flex gap-2">
        <Skeleton className="h-10 w-44 rounded-xl" />
        <Skeleton className="h-10 w-44 rounded-xl" />
        <Skeleton className="h-10 w-44 rounded-xl" />
      </div>
      <Skeleton className="h-24 w-full rounded-2xl" />
      <Skeleton className="h-[520px] w-full rounded-2xl" />
    </div>
  );
}
