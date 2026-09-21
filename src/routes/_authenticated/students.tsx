import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useEffect, useMemo, useState, Fragment } from "react";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { ArrowLeft, SquarePen } from "lucide-react";
import { toast } from "sonner";
import type { ClassRow, StudentRow } from "@/lib/data";

export const Route = createFileRoute("/_authenticated/students")({
  component: StudentsPage,
});

type SF1StudentRow = StudentRow & {
  mother_tongue?: string | null;
  ip_ethnic_group?: string | null;
  ip_ethnic?: string | null;
  religion?: string | null;
  house_street?: string | null;
  municipality_city?: string | null;
  municipality?: string | null;
  province?: string | null;
  guardian_relationship?: string | null;
  learning_modality?: string | null;
  remarks?: string | null;
};

type EditableField =
  | "lrn"
  | "last_name"
  | "first_name"
  | "middle_name"
  | "sex"
  | "birthdate"
  | "mother_tongue"
  | "ip_ethnic_group"
  | "religion"
  | "house_street"
  | "municipality_city"
  | "province"
  | "father_name"
  | "mother_name"
  | "guardian"
  | "contact_number"
  | "learning_modality"
  | "remarks";

function firstFridayOfJune(schoolYear?: string | null) {
  const matchedYear = schoolYear?.match(/\d{4}/)?.[0];
  const year = matchedYear ? Number(matchedYear) : new Date().getFullYear();
  const date = new Date(Date.UTC(year, 5, 1));
  const daysUntilFriday = (5 - date.getUTCDay() + 7) % 7;
  date.setUTCDate(date.getUTCDate() + daysUntilFriday);
  return date;
}

function ageOnFirstFridayOfJune(
  birthdate?: string | null,
  schoolYear?: string | null,
) {
  if (!birthdate) return "";

  const birth = new Date(`${birthdate.slice(0, 10)}T00:00:00Z`);
  if (Number.isNaN(birth.getTime())) return "";

  const reference = firstFridayOfJune(schoolYear);
  let age = reference.getUTCFullYear() - birth.getUTCFullYear();

  const beforeBirthday =
    reference.getUTCMonth() < birth.getUTCMonth() ||
    (reference.getUTCMonth() === birth.getUTCMonth() &&
      reference.getUTCDate() < birth.getUTCDate());

  if (beforeBirthday) age -= 1;

  return String(Math.max(0, age));
}

function StudentsPage() {
  const qc = useQueryClient();
  const navigate = useNavigate();

  const { data: teacherType, isLoading: accessLoading } = useQuery({
    queryKey: ["students-page-teacher-type"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) return null;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: profile, error } = await (supabase as any)
        .from("profiles")
        .select("teacher_type")
        .eq("id", user.id)
        .maybeSingle();

      if (error) throw error;

      return profile?.teacher_type ?? null;
    },
    staleTime: 5 * 60 * 1000,
  });

  const canManageStudentInfo = teacherType === "class_adviser";

  useEffect(() => {
    if (!accessLoading && teacherType === "subject_teacher") {
      void navigate({
        to: "/subject-teacher-dashboard",
        replace: true,
      });
    }
  }, [accessLoading, navigate, teacherType]);

  const { data: classes = [], isLoading: classesLoading } = useQuery({
    queryKey: ["classes"],
    enabled: canManageStudentInfo,
    queryFn: async () =>
      (
        await supabase
          .from("classes")
          .select("*")
          .order("created_at", { ascending: false })
      ).data as ClassRow[],
  });

  const [classId, setClassId] = useState("");
  const activeClassId = classId || classes[0]?.id || "";
  const [editMode, setEditMode] = useState(false);

  const { data: students = [], isLoading: studentsLoading } = useQuery({
    queryKey: ["students", activeClassId],
    enabled: canManageStudentInfo && !!activeClassId,
    queryFn: async () =>
      (
        await supabase
          .from("students")
          .select("*")
          .eq("class_id", activeClassId)
          .order("last_name")
          .order("first_name")
      ).data as SF1StudentRow[],
  });

  const activeClass = classes.find((c) => c.id === activeClassId);

  const { male, female, unspecified } = useMemo(() => {
    const male = students.filter((s) => s.sex?.toLowerCase() === "male");
    const female = students.filter((s) => s.sex?.toLowerCase() === "female");
    const unspecified = students.filter(
      (s) =>
        s.sex?.toLowerCase() !== "male" &&
        s.sex?.toLowerCase() !== "female",
    );

    return { male, female, unspecified };
  }, [students]);

  const updateStudent = useMutation({
    mutationFn: async ({
      id,
      field,
      value,
    }: {
      id: string;
      field: EditableField;
      value: string;
    }) => {
      const patch = { [field]: value || null } as Record<
        string,
        string | null
      >;

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { error } = await supabase
        .from("students")
        .update(patch as any)
        .eq("id", id);

      if (error) throw error;
    },
    onSuccess: () =>
      qc.invalidateQueries({ queryKey: ["students", activeClassId] }),
    onError: (e: Error) => toast.error(e.message),
  });

  const pageLoading =
    accessLoading ||
    (canManageStudentInfo && classesLoading) ||
    (canManageStudentInfo && !!activeClassId && studentsLoading);

  if (pageLoading) {
    return <StudentsPageSkeleton />;
  }

  if (!canManageStudentInfo) {
    return null;
  }

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border bg-card px-4 py-3 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex min-w-0 items-center gap-3">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => void navigate({ to: "/school-forms" })}
              className="h-9 shrink-0 gap-1.5 rounded-xl px-3"
            >
              <ArrowLeft className="size-4" />
              Back
            </Button>

            <div className="min-w-0">
              <h1 className="truncate text-xl font-bold">My Students</h1>
              <p className="text-sm text-muted-foreground">
                {activeClass
                  ? `${activeClass.grade_level || "Grade"} · ${
                      activeClass.section || "No section"
                    } · `
                  : ""}
                {students.length} student{students.length === 1 ? "" : "s"}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <Button
              type="button"
              variant={editMode ? "default" : "outline"}
              size="sm"
              onClick={() => setEditMode((v) => !v)}
              className="gap-1.5"
            >
              <SquarePen className="size-4" />
              {editMode ? "Editing" : "Edit"}
            </Button>

            <Select value={activeClassId} onValueChange={setClassId}>
              <SelectTrigger className="w-[240px] bg-background">
                <SelectValue placeholder="Select class" />
              </SelectTrigger>
              <SelectContent>
                {classes.map((c) => (
                  <SelectItem key={c.id} value={c.id}>
                    {c.subject} {c.section ? `( ${c.section} )` : ""}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>
      </div>

      <p className="text-xs text-muted-foreground">
        Fill in your students' information at your own pace. This data is
        optional — school forms will use whatever you've entered.
      </p>

      <div className="overflow-x-auto rounded-lg border bg-card shadow-sm">
        <table className="w-full min-w-[3300px] border-collapse text-xs">
          <thead className="bg-secondary/60 text-[11px] font-semibold uppercase text-muted-foreground">
            <tr>
              <Th>#</Th>
              <Th>LRN</Th>
              <Th>Last Name</Th>
              <Th>First Name</Th>
              <Th>Middle Name</Th>
              <Th>Sex</Th>
              <Th>Birth Date</Th>
              <Th>
                Age as of
                <br />
                1st Friday June
              </Th>
              <Th>
                Mother Tongue
                <br />
                (Grade 1 to 3 Only)
              </Th>
              <Th>
                IP
                <br />
                (Ethnic Group)
              </Th>
              <Th>Religion</Th>
              <Th>
                House #/Street/
                <br />
                Sitio/Purok
              </Th>
              <Th>Municipality/City</Th>
              <Th>Province</Th>
              <Th>
                Father's Name
                <br />
                (Last Name, First Name, Middle Name)
              </Th>
              <Th>
                Mother's Maiden Name
                <br />
                (Last Name, First Name, Middle Name)
              </Th>
              <Th>Parent/Guardian</Th>
              <Th>
                Contact Number of
                <br />
                Parent or Guardian
              </Th>
              <Th>Learning Modality</Th>
              <Th>Remarks</Th>
            </tr>
          </thead>

          <tbody>
            {[
              { key: "male", label: "MALE", group: male },
              { key: "female", label: "FEMALE", group: female },
              {
                key: "unspecified",
                label: "UNSPECIFIED",
                group: unspecified,
              },
            ].map(({ key, label, group }) => (
              <Fragment key={key}>
                <tr className="bg-accent/50">
                  <td
                    colSpan={20}
                    className="px-3 py-1.5 text-[11px] font-bold uppercase tracking-wide text-primary"
                  >
                    {label}
                  </td>
                </tr>

                {group.length === 0 && (
                  <tr>
                    <td
                      colSpan={20}
                      className="px-3 py-3 text-center italic text-muted-foreground"
                    >
                      No {label.toLowerCase()} learners in this class.
                    </td>
                  </tr>
                )}

                {group.map((s, i) => (
                  <tr
                    key={s.id}
                    className="border hover:bg-secondary/20 transition-colors"
                  >
                    <Td className="text-center">{i + 1}</Td>

                    <EditCell
                      student={s}
                      field="lrn"
                      placeholder="Enter LRN"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="last_name"
                      placeholder="Last name"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="first_name"
                      placeholder="First name"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="middle_name"
                      placeholder="Middle name"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <SexCell
                      student={s}
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="birthdate"
                      type="date"
                      placeholder="mm/dd/yyyy"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <Td className="text-center font-medium">
                      {ageOnFirstFridayOfJune(
                        s.birthdate,
                        activeClass?.school_year,
                      )}
                    </Td>

                    <EditCell
                      student={s}
                      field="mother_tongue"
                      placeholder="Mother tongue"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="ip_ethnic_group"
                      fallbackValue={s.ip_ethnic ?? ""}
                      placeholder="Ethnic group"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="religion"
                      placeholder="Religion"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="house_street"
                      fallbackValue={s.address ?? ""}
                      placeholder="House #/Street/Sitio/Purok"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="municipality_city"
                      fallbackValue={s.municipality ?? ""}
                      placeholder="Municipality/City"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="province"
                      placeholder="Province"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="father_name"
                      placeholder="Father's Name"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="mother_name"
                      placeholder="Mother's Maiden Name"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="guardian"
                      placeholder="Parent/Guardian"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="contact_number"
                      placeholder="Contact number"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="learning_modality"
                      placeholder="Learning modality"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />

                    <EditCell
                      student={s}
                      field="remarks"
                      placeholder="Remarks"
                      onSave={updateStudent.mutate}
                      disabled={!editMode}
                    />
                  </tr>
                ))}
              </Fragment>
            ))}

            {students.length === 0 && activeClass && (
              <tr>
                <td
                  colSpan={20}
                  className="p-6 text-center text-muted-foreground"
                >
                  No learners in <b>{activeClass.subject}</b> yet. Open the class
                  and use the Roster tab to add students.
                </td>
              </tr>
            )}

            {classes.length === 0 && (
              <tr>
                <td
                  colSpan={20}
                  className="p-6 text-center text-muted-foreground"
                >
                  No classes yet. Create a class from the dashboard first.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}


function StudentsPageSkeleton() {
  const headerColumns = [
    "w-10",
    "w-24",
    "w-40",
    "w-40",
    "w-40",
    "w-24",
    "w-36",
    "w-36",
    "w-44",
    "w-36",
  ];

  return (
    <div
      className="space-y-4"
      aria-busy="true"
      aria-label="Loading students"
    >
      {/* Header skeleton */}
      <div className="rounded-2xl border bg-card px-4 py-3 shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex min-w-0 items-center gap-3">
            <SkeletonBlock className="h-9 w-[92px] rounded-xl" />

            <div className="space-y-2">
              <SkeletonBlock className="h-6 w-36 rounded-md" />
              <SkeletonBlock className="h-4 w-48 rounded-md" />
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <SkeletonBlock className="h-9 w-24 rounded-lg" />
            <SkeletonBlock className="h-9 w-[240px] rounded-lg" />
          </div>
        </div>
      </div>

      {/* Description skeleton */}
      <SkeletonBlock className="h-4 w-full max-w-2xl rounded-md" />

      {/* Student table skeleton */}
      <div className="overflow-hidden rounded-lg border bg-card shadow-sm">
        <div className="overflow-x-auto">
          <div className="min-w-[1500px]">
            <div className="flex h-[58px] items-center border-b bg-secondary/60 px-3">
              {headerColumns.map((width, index) => (
                <div
                  key={index}
                  className="flex h-full shrink-0 items-center justify-center border-r px-3 last:border-r-0"
                >
                  <SkeletonBlock className={`h-4 ${width} rounded-md`} />
                </div>
              ))}
            </div>

            <SkeletonStudentGroup labelWidth="w-16" rowCount={5} />
            <SkeletonStudentGroup labelWidth="w-20" rowCount={3} />
            <SkeletonStudentGroup labelWidth="w-28" rowCount={1} />
          </div>
        </div>

        <div className="border-t px-3 py-2">
          <SkeletonBlock className="h-2 w-1/3 rounded-full" />
        </div>
      </div>
    </div>
  );
}

function SkeletonStudentGroup({
  labelWidth,
  rowCount,
}: {
  labelWidth: string;
  rowCount: number;
}) {
  return (
    <>
      <div className="flex h-9 items-center border-b bg-accent/50 px-4">
        <SkeletonBlock className={`h-3 ${labelWidth} rounded-md`} />
      </div>

      {Array.from({ length: rowCount }).map((_, rowIndex) => (
        <div
          key={rowIndex}
          className="flex h-[48px] items-center gap-3 border-b px-3"
        >
          <SkeletonBlock className="h-4 w-6 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-28 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-36 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-36 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-36 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-24 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-36 shrink-0 rounded-md" />
          <SkeletonBlock className="h-5 w-16 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-40 shrink-0 rounded-md" />
          <SkeletonBlock className="h-8 w-36 shrink-0 rounded-md" />
        </div>
      ))}
    </>
  );
}

function SkeletonBlock({ className = "" }: { className?: string }) {
  return (
    <div
      className={`animate-pulse bg-muted/80 ${className}`}
      aria-hidden="true"
    />
  );
}

function Th({ children }: { children: React.ReactNode }) {
  return (
    <th className="border px-3 py-2 text-center align-middle whitespace-nowrap">
      {children}
    </th>
  );
}

function Td({
  children,
  className = "",
}: {
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <td className={`border px-3 py-2 whitespace-nowrap ${className}`}>
      {children}
    </td>
  );
}

function SexCell({
  student,
  onSave,
  disabled = false,
}: {
  student: SF1StudentRow;
  onSave: (v: { id: string; field: EditableField; value: string }) => void;
  disabled?: boolean;
}) {
  const initial =
    student.sex?.toLowerCase() === "male"
      ? "male"
      : student.sex?.toLowerCase() === "female"
        ? "female"
        : "";

  const [value, setValue] = useState(initial);

  if (disabled) {
    return (
      <Td className="text-center uppercase">
        {value === "male" ? "M" : value === "female" ? "F" : ""}
      </Td>
    );
  }

  return (
    <td className="border px-2 py-1">
      <Select
        value={value}
        onValueChange={(nextValue) => {
          setValue(nextValue);
          if (nextValue !== initial) {
            onSave({
              id: student.id,
              field: "sex",
              value: nextValue,
            });
          }
        }}
      >
        <SelectTrigger className="h-9 min-w-[95px] rounded-sm">
          <SelectValue placeholder="Sex" />
        </SelectTrigger>
        <SelectContent>
          <SelectItem value="male">Male</SelectItem>
          <SelectItem value="female">Female</SelectItem>
        </SelectContent>
      </Select>
    </td>
  );
}

function EditCell({
  student,
  field,
  placeholder,
  type = "text",
  onSave,
  disabled = false,
  fallbackValue = "",
}: {
  student: SF1StudentRow;
  field: EditableField;
  placeholder: string;
  type?: string;
  onSave: (v: { id: string; field: EditableField; value: string }) => void;
  disabled?: boolean;
  fallbackValue?: string;
}) {
  const storedValue = student[field as keyof SF1StudentRow];
  const initial =
    typeof storedValue === "string" && storedValue.length > 0
      ? storedValue
      : fallbackValue;

  const [val, setVal] = useState(initial);

  return (
    <td className="border px-2 py-1">
      <Input
        type={type}
        value={val}
        placeholder={placeholder}
        disabled={disabled}
        onChange={(e) => {
          const { value } = e.target;

          if (field === "lrn") {
            if (/^\d{0,12}$/.test(value)) {
              setVal(value);
            }
          } else {
            setVal(value);
          }
        }}
        onBlur={() => {
          if (val !== initial) {
            onSave({
              id: student.id,
              field,
              value: val,
            });
          }
        }}
        className="h-9 w-full min-w-[145px] rounded-sm border bg-background px-2 text-xs shadow-none focus:ring-1 focus:ring-primary disabled:cursor-not-allowed disabled:opacity-60"
        maxLength={field === "lrn" ? 12 : undefined}
      />
    </td>
  );
}
