import { createFileRoute } from "@tanstack/react-router";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { useState, type ChangeEvent } from "react";
import {
  BookOpen,
  ClipboardList,
  ClipboardPenLine,
  GraduationCap,
  Loader2,
  Mail,
  School,
  Trash2,
  Upload,
  UserRound,
  X,
} from "lucide-react";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";

export const Route = createFileRoute("/_authenticated/profile")({
  component: ProfilePage,
});

type TeacherType = "class_adviser" | "subject_teacher";

type ProfileRow = {
  id: string;
  full_name: string | null;
  email: string | null;
  teacher_type: TeacherType | string | null;
  school_name: string | null;
  avatar_url: string | null;
};

type ClassRow = {
  id: string;
  teacher_id: string;
  subject: string | null;
  grade_level: string | null;
  section: string | null;
  school_year: string | null;
};

type StudentRow = {
  id: string;
  class_id: string;
  teacher_id: string;
};

function normalize(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .trim();
}

function roleLabel(value: TeacherType | string | null | undefined) {
  return value === "subject_teacher" ? "Subject Teacher" : "Class Adviser";
}

function displayNameFromEmail(email: string) {
  const name = email.split("@")[0] || "Teacher";
  return name
    .split(/[._-]/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function getProfileImagePath(publicUrl: string | null | undefined) {
  if (!publicUrl) return null;

  const marker = "/profile-images/";
  const markerIndex = publicUrl.indexOf(marker);

  if (markerIndex === -1) return null;

  return decodeURIComponent(publicUrl.slice(markerIndex + marker.length));
}

function ProfilePageSkeleton() {
  return (
    <div className="space-y-6 pb-24 md:pb-6" role="status" aria-live="polite" aria-busy="true">
      <span className="sr-only">Loading teacher profile...</span>

      <section className="overflow-hidden rounded-3xl border border-[#ead8b8] bg-[#fffdf9] shadow-[0_14px_35px_rgba(102,63,38,0.10)]">
        <div className="relative overflow-hidden bg-gradient-to-r from-[#f5e5e2] via-[#fff8ef] to-[#fffdf8] p-6">
          <div className="pointer-events-none absolute -left-20 -bottom-28 size-52 rounded-full border border-[#efc9bf]/70" />
          <div className="pointer-events-none absolute -right-16 -top-20 size-48 rounded-full bg-[#f5cfc2]/35" />

          <div className="relative z-10 flex flex-col gap-4 sm:flex-row sm:items-center">
            <Skeleton className="size-24 shrink-0 rounded-3xl" />

            <div className="min-w-0 flex-1 space-y-2">
              <Skeleton className="h-8 w-full max-w-64" />
              <Skeleton className="h-4 w-32" />
              <Skeleton className="h-4 w-full max-w-sm" />

              <div className="flex flex-wrap gap-2 pt-2">
                <Skeleton className="h-9 w-40 rounded-lg" />
                <Skeleton className="h-9 w-32 rounded-lg" />
              </div>
            </div>
          </div>
        </div>

        <div className="grid gap-4 border-t border-[#f0dfc3] bg-[#fffdf9] p-6 sm:grid-cols-2 lg:grid-cols-4">
          {Array.from({ length: 4 }, (_, index) => (
            <div
              key={index}
              className="rounded-2xl border border-[#ead8b8] bg-gradient-to-br from-[#fffaf0] to-[#fffdf9] p-4"
            >
              <div className="flex items-center gap-2">
                <Skeleton className="size-4 rounded-full" />
                <Skeleton className="h-3 w-24" />
              </div>
              <Skeleton className={`mt-3 h-4 ${index === 1 ? "w-40" : "w-28"}`} />
            </div>
          ))}
        </div>
      </section>

      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        {Array.from({ length: 4 }, (_, index) => (
          <div key={index} className="rounded-3xl border bg-card p-5 shadow-sm">
            <div className="flex items-center justify-between">
              <div className="space-y-3">
                <Skeleton className="h-4 w-28" />
                <Skeleton className="h-9 w-14" />
              </div>
              <Skeleton className="size-11 rounded-full" />
            </div>
          </div>
        ))}
      </div>

      <section className="rounded-3xl border bg-card p-5 shadow-sm">
        <Skeleton className="h-6 w-44" />

        <div className="mt-4 overflow-hidden rounded-xl border">
          <div className="grid grid-cols-4 gap-px border-b bg-border">
            {Array.from({ length: 4 }, (_, index) => (
              <div key={index} className="bg-muted/40 px-3 py-3">
                <Skeleton className="h-3 w-20" />
              </div>
            ))}
          </div>

          <div className="divide-y">
            {Array.from({ length: 5 }, (_, rowIndex) => (
              <div key={rowIndex} className="grid grid-cols-4 gap-px bg-border">
                {Array.from({ length: 4 }, (_, columnIndex) => (
                  <div key={columnIndex} className="bg-card px-3 py-4">
                    <Skeleton
                      className={`h-4 ${columnIndex === 2 || rowIndex % 2 === 0 ? "w-28" : "w-20"}`}
                    />
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      </section>
    </div>
  );
}

function ProfilePage() {
  const queryClient = useQueryClient();
  const [zoomImageOpen, setZoomImageOpen] = useState(false);
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);

  const { data, isLoading, error } = useQuery({
    queryKey: ["profile-page-data"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) throw new Error("You must be signed in.");

      const [profileRes, classesRes, studentsRes, credentialRequestsRes] = await Promise.all([
        (supabase as any)
          .from("profiles")
          .select("id, full_name, email, teacher_type, school_name, avatar_url")
          .eq("id", user.id)
          .maybeSingle(),
        supabase.from("classes").select("*").eq("teacher_id", user.id),
        supabase.from("students").select("id, class_id, teacher_id").eq("teacher_id", user.id),
        (supabase as any)
          .from("student_credential_requests")
          .select("id, status, requester_id")
          .eq("requester_id", user.id),
      ]);

      if (profileRes.error) throw profileRes.error;
      if (classesRes.error) throw classesRes.error;
      if (studentsRes.error) throw studentsRes.error;

      const profile = (profileRes.data ?? {
        id: user.id,
        full_name:
          user.user_metadata?.full_name ||
          (user.email ? displayNameFromEmail(user.email) : "Teacher"),
        email: user.email,
        teacher_type: user.user_metadata?.teacher_type || "class_adviser",
        school_name: "Agusan del Sur National Science High School",
        avatar_url: null,
      }) as ProfileRow;

      return {
        userId: user.id,
        profile,
        classes: (classesRes.data ?? []) as ClassRow[],
        students: (studentsRes.data ?? []) as StudentRow[],
        credentialRequests: credentialRequestsRes.error
          ? []
          : ((credentialRequestsRes.data ?? []) as Array<{ id: string; status: string | null }>),
      };
    },
  });

  const uploadProfileImage = useMutation({
    mutationFn: async (file: File) => {
      if (!data?.userId) throw new Error("You must be signed in.");

      if (!file.type.startsWith("image/")) {
        throw new Error("Please select an image file.");
      }

      if (file.size > 3 * 1024 * 1024) {
        throw new Error("Image must be 3MB or smaller.");
      }

      const extension = file.name.split(".").pop()?.toLowerCase() || "jpg";
      const filePath = `${data.userId}/profile-${Date.now()}.${extension}`;

      const { error: uploadError } = await supabase.storage
        .from("profile-images")
        .upload(filePath, file, {
          cacheControl: "3600",
          upsert: true,
        });

      if (uploadError) throw uploadError;

      const {
        data: { publicUrl },
      } = supabase.storage.from("profile-images").getPublicUrl(filePath);

      const { error: updateError } = await (supabase as any)
        .from("profiles")
        .update({ avatar_url: publicUrl })
        .eq("id", data.userId);

      if (updateError) throw updateError;

      return publicUrl;
    },
    onSuccess: async () => {
      await Promise.all([
        queryClient.invalidateQueries({ queryKey: ["profile-page-data"] }),
        queryClient.invalidateQueries({ queryKey: ["current-user-shell-profile"] }),
      ]);
    },
  });

  const deleteProfileImage = useMutation({
    mutationFn: async () => {
      if (!data?.userId) throw new Error("You must be signed in.");

      const currentImageUrl = data.profile.avatar_url;
      const imagePath = getProfileImagePath(currentImageUrl);

      if (imagePath) {
        const { error: removeError } = await supabase.storage
          .from("profile-images")
          .remove([imagePath]);

        if (removeError) throw removeError;
      }

      const { error: updateError } = await (supabase as any)
        .from("profiles")
        .update({ avatar_url: null })
        .eq("id", data.userId);

      if (updateError) throw updateError;
    },
    onSuccess: async () => {
      setZoomImageOpen(false);
      setDeleteConfirmOpen(false);
      await Promise.all([
        queryClient.invalidateQueries({ queryKey: ["profile-page-data"] }),
        queryClient.invalidateQueries({ queryKey: ["current-user-shell-profile"] }),
      ]);
    },
  });

  const handleImageChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";

    if (file) {
      uploadProfileImage.mutate(file);
    }
  };

  if (isLoading) {
    return <ProfilePageSkeleton />;
  }

  if (error) {
    return (
      <div className="rounded-2xl border border-red-200 bg-red-50 p-5 text-sm text-red-700">
        {(error as Error).message}
      </div>
    );
  }

  const profile = data?.profile;
  const teacherType =
    profile?.teacher_type === "subject_teacher" ? "subject_teacher" : "class_adviser";
  const classes = data?.classes ?? [];
  const students = data?.students ?? [];
  const credentialRequests = data?.credentialRequests ?? [];
  const pendingRequests = credentialRequests.filter(
    (item) => normalize(item.status) === "pending",
  ).length;
  const completedRequests = credentialRequests.filter(
    (item) => normalize(item.status) === "completed",
  ).length;

  const subjects = Array.from(new Set(classes.map((item) => item.subject).filter(Boolean)));
  const gradeLevels = Array.from(new Set(classes.map((item) => item.grade_level).filter(Boolean)));
  const sections = Array.from(new Set(classes.map((item) => item.section).filter(Boolean)));

  const isImageBusy = uploadProfileImage.isPending || deleteProfileImage.isPending;

  return (
    <div className="space-y-6 pb-24 md:pb-6">
      <section className="overflow-hidden rounded-3xl border border-[#ead8b8] bg-[#fffdf9] shadow-[0_14px_35px_rgba(102,63,38,0.10)]">
        <div className="relative isolate overflow-hidden bg-gradient-to-r from-[#f5e5e2] via-[#fff8ef] to-[#fffdf8] p-6 sm:p-7">
          {/* Soft decorative profile background */}
          <div className="pointer-events-none absolute -left-20 -bottom-28 size-56 rounded-full border border-[#e8bdb4]/70" />
          <div className="pointer-events-none absolute -left-10 -bottom-20 size-40 rounded-full bg-[#f5c8c2]/30" />
          <div className="pointer-events-none absolute -right-20 -top-24 size-60 rounded-full border border-[#e6c37f]/55" />
          <div className="pointer-events-none absolute -right-12 -top-16 size-44 rounded-full bg-[#f5c7b8]/35" />
          <div className="pointer-events-none absolute right-8 bottom-5 grid grid-cols-6 gap-2 opacity-35">
            {Array.from({ length: 24 }).map((_, index) => (
              <span
                key={`profile-dot-${index}`}
                className="size-1.5 rounded-full bg-[#d98f79]"
              />
            ))}
          </div>

          <div className="relative z-10 flex flex-col gap-4 sm:flex-row sm:items-center">
            <div className="relative size-24 shrink-0">
              {profile?.avatar_url ? (
                <button
                  type="button"
                  onClick={() => setZoomImageOpen(true)}
                  className="block rounded-3xl focus:outline-none focus:ring-2 focus:ring-primary focus:ring-offset-2"
                  title="Click to zoom profile image"
                >
                  <img
                    src={profile.avatar_url}
                    alt={profile.full_name || "Profile image"}
                    className="size-24 rounded-3xl object-cover shadow-sm"
                  />
                </button>
              ) : (
                <div className="grid size-24 place-items-center rounded-3xl bg-[#8f2928] text-white shadow-[0_10px_24px_rgba(143,41,40,0.22)]">
                  <UserRound className="size-11" />
                </div>
              )}
            </div>

            <div className="min-w-0 flex-1">
              <h1 className="text-2xl font-bold">{profile?.full_name || "Teacher Profile"}</h1>
              <p className="text-sm font-semibold text-[#8f2928]">{roleLabel(teacherType)}</p>
              <p className="mt-1 text-sm text-muted-foreground">
                {profile?.school_name || "Agusan del Sur National Science High School"}
              </p>

              <div className="mt-4 flex flex-wrap items-center gap-2">
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  disabled={isImageBusy}
                  onClick={() => document.getElementById("profile-image-picker")?.click()}
                  className="border-[#ead8b8] bg-[#fffaf0]/90 text-[#3c241b] shadow-sm hover:bg-white"
                >
                  {uploadProfileImage.isPending ? (
                    <Loader2 className="mr-2 size-4 animate-spin" />
                  ) : (
                    <Upload className="mr-2 size-4" />
                  )}
                  Edit Profile Image
                </Button>

                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  disabled={isImageBusy || !profile?.avatar_url}
                  onClick={() => setDeleteConfirmOpen(true)}
                  className="border-[#f2cbc5] bg-[#fff8f6]/85 text-[#e25c59] shadow-sm hover:bg-[#fff1ee] disabled:cursor-not-allowed disabled:opacity-50"
                  title={
                    profile?.avatar_url ? "Delete profile image" : "No profile image to delete"
                  }
                >
                  {deleteProfileImage.isPending ? (
                    <Loader2 className="mr-2 size-4 animate-spin" />
                  ) : (
                    <Trash2 className="mr-2 size-4" />
                  )}
                  Delete Image
                </Button>

                <input
                  id="profile-image-picker"
                  type="file"
                  accept="image/*"
                  className="hidden"
                  onChange={handleImageChange}
                  disabled={isImageBusy}
                />
              </div>

              {(uploadProfileImage.error || deleteProfileImage.error) && (
                <p className="mt-2 text-xs text-red-600">
                  {((uploadProfileImage.error || deleteProfileImage.error) as Error).message}
                </p>
              )}
            </div>
          </div>
        </div>

        <div className="grid gap-4 p-6 sm:grid-cols-2 lg:grid-cols-4">
          <InfoCard icon={UserRound} label="Full Name" value={profile?.full_name || "-"} />
          <InfoCard icon={Mail} label="Email" value={profile?.email || "-"} />
          <InfoCard icon={School} label="Teacher Type" value={roleLabel(teacherType)} />
          <InfoCard icon={BookOpen} label="Assigned Classes" value={classes.length} />
        </div>
      </section>

      {teacherType === "class_adviser" ? (
        <>
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <MetricCard icon={BookOpen} label="My Classes" value={classes.length} />
            <MetricCard icon={GraduationCap} label="My Students" value={students.length} />
            <MetricCard icon={ClipboardList} label="Pending Requests" value={pendingRequests} />
            <MetricCard icon={ClipboardList} label="Completed Requests" value={completedRequests} />
          </div>

          <section className="rounded-3xl border bg-card p-5 shadow-sm">
            <h2 className="text-lg font-semibold">Advisory Information</h2>
            <div className="mt-4 overflow-x-auto">
              <table className="w-full min-w-[700px] text-sm">
                <thead>
                  <tr className="border-b text-left text-xs uppercase text-muted-foreground">
                    <th className="px-3 py-2">Grade Level</th>
                    <th className="px-3 py-2">Section</th>
                    <th className="px-3 py-2">Subject</th>
                    <th className="px-3 py-2">School Year</th>
                  </tr>
                </thead>
                <tbody>
                  {classes.map((item) => (
                    <tr key={item.id} className="border-b last:border-0">
                      <td className="px-3 py-3">{item.grade_level || "-"}</td>
                      <td className="px-3 py-3">{item.section || "-"}</td>
                      <td className="px-3 py-3">{item.subject || "-"}</td>
                      <td className="px-3 py-3">{item.school_year || "-"}</td>
                    </tr>
                  ))}
                  {classes.length === 0 && (
                    <tr>
                      <td colSpan={4} className="px-3 py-8 text-center text-muted-foreground">
                        No advisory class records found.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </>
      ) : (
        <>
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            <MetricCard icon={ClipboardPenLine} label="Subjects Handled" value={subjects.length} />
            <MetricCard icon={GraduationCap} label="Grade Levels" value={gradeLevels.length} />
            <MetricCard icon={BookOpen} label="Sections" value={sections.length} />
            <MetricCard icon={GraduationCap} label="Classes" value={classes.length} />
          </div>

          <section className="rounded-3xl border bg-card p-5 shadow-sm">
            <h2 className="text-lg font-semibold">Teaching Load</h2>
            <div className="mt-4 overflow-x-auto">
              <table className="w-full min-w-[700px] text-sm">
                <thead>
                  <tr className="border-b text-left text-xs uppercase text-muted-foreground">
                    <th className="px-3 py-2">Subject</th>
                    <th className="px-3 py-2">Grade Level</th>
                    <th className="px-3 py-2">Section</th>
                    <th className="px-3 py-2">School Year</th>
                  </tr>
                </thead>
                <tbody>
                  {classes.map((item) => (
                    <tr key={item.id} className="border-b last:border-0">
                      <td className="px-3 py-3">{item.subject || "-"}</td>
                      <td className="px-3 py-3">{item.grade_level || "-"}</td>
                      <td className="px-3 py-3">{item.section || "-"}</td>
                      <td className="px-3 py-3">{item.school_year || "-"}</td>
                    </tr>
                  ))}
                  {classes.length === 0 && (
                    <tr>
                      <td colSpan={4} className="px-3 py-8 text-center text-muted-foreground">
                        No subject load records found.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}

      {deleteConfirmOpen && profile?.avatar_url && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/40 p-4">
          <button
            type="button"
            className="absolute inset-0"
            aria-label="Cancel delete image"
            onClick={() => {
              if (!deleteProfileImage.isPending) setDeleteConfirmOpen(false);
            }}
          />

          <div className="relative z-10 w-full max-w-md rounded-3xl border bg-card p-6 shadow-2xl">
            <div className="flex items-start gap-4">
              <div className="grid size-12 shrink-0 place-items-center rounded-2xl bg-red-50 text-red-700">
                <Trash2 className="size-5" />
              </div>

              <div className="min-w-0 flex-1">
                <h2 className="text-lg font-bold">Delete profile image</h2>
                <p className="mt-1 text-sm text-muted-foreground">
                  This will remove your current profile picture from your account.
                </p>

                {deleteProfileImage.error && (
                  <p className="mt-3 text-xs text-red-600">
                    {(deleteProfileImage.error as Error).message}
                  </p>
                )}

                <div className="mt-5 flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
                  <Button
                    type="button"
                    variant="outline"
                    disabled={deleteProfileImage.isPending}
                    onClick={() => setDeleteConfirmOpen(false)}
                  >
                    Cancel
                  </Button>

                  <Button
                    type="button"
                    disabled={deleteProfileImage.isPending}
                    onClick={() => deleteProfileImage.mutate()}
                    className="bg-red-700 text-white hover:bg-red-800"
                  >
                    {deleteProfileImage.isPending ? (
                      <Loader2 className="mr-2 size-4 animate-spin" />
                    ) : (
                      <Trash2 className="mr-2 size-4" />
                    )}
                    Delete Image
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {zoomImageOpen && profile?.avatar_url && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/75 p-4">
          <button
            type="button"
            className="absolute inset-0"
            aria-label="Close zoomed image"
            onClick={() => setZoomImageOpen(false)}
          />

          <div className="relative z-10 max-h-[90vh] max-w-[90vw]">
            <button
              type="button"
              onClick={() => setZoomImageOpen(false)}
              className="absolute -right-3 -top-3 grid size-9 place-items-center rounded-full bg-background text-foreground shadow-lg"
              aria-label="Close"
            >
              <X className="size-4" />
            </button>

            <img
              src={profile.avatar_url}
              alt={profile.full_name || "Profile image"}
              className="max-h-[90vh] max-w-[90vw] rounded-3xl object-contain shadow-2xl"
            />
          </div>
        </div>
      )}
    </div>
  );
}

function InfoCard({
  icon: Icon,
  label,
  value,
}: {
  icon: typeof UserRound;
  label: string;
  value: number | string;
}) {
  return (
    <div className="rounded-2xl border border-[#ead8b8] bg-gradient-to-br from-[#fff9ee] to-[#fffdf8] p-4 shadow-[0_5px_14px_rgba(116,77,47,0.05)]">
      <div className="flex items-center gap-2 text-xs font-medium text-[#7d6657]">
        <Icon className="size-4 text-[#9b7b68]" />
        {label}
      </div>
      <div className="mt-2 truncate text-sm font-semibold text-[#2f1c15]">{value}</div>
    </div>
  );
}

function MetricCard({
  icon: Icon,
  label,
  value,
}: {
  icon: typeof UserRound;
  label: string;
  value: number | string;
}) {
  return (
    <div className="rounded-3xl border bg-card p-5 shadow-sm">
      <div className="flex items-center justify-between">
        <div>
          <p className="text-sm text-muted-foreground">{label}</p>
          <p className="mt-2 text-3xl font-bold">{value}</p>
        </div>
        <div className="grid size-11 place-items-center rounded-full bg-primary/10 text-primary">
          <Icon className="size-5" />
        </div>
      </div>
    </div>
  );
}