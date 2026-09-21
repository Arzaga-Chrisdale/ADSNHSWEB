import { useEffect, useMemo, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  ArrowLeft,
  Bell,
  Check,
  CheckCheck,
  FileCheck2,
  GraduationCap,
  Loader2,
  Megaphone,
  Search,
  ShieldCheck,
  UserCog,
  Users,
} from "lucide-react";
import { format, formatDistanceToNow } from "date-fns";
import { toast } from "sonner";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/_authenticated/notifications")({
  component: NotificationsPage,
});

type NotificationItem = {
  id: string;
  recipient_id: string;
  title: string;
  message: string;
  category: string;
  related_path: string | null;
  read_at: string | null;
  created_at: string;
};

type FilterValue = "all" | "unread" | "forms" | "grades" | "other";

const QUERY_KEY = "current-user-notifications";

const filters: Array<{ value: FilterValue; label: string }> = [
  { value: "all", label: "All" },
  { value: "unread", label: "Unread" },
  { value: "forms", label: "School Forms" },
  { value: "grades", label: "Grades" },
  { value: "other", label: "Other" },
];

function categoryDetails(category: string) {
  switch (category) {
    case "form_submission":
    case "form_review":
      return {
        label: "School Form",
        icon: FileCheck2,
        tone: "bg-violet-100 text-violet-700",
      };
    case "grade_submission":
    case "grade_review":
      return {
        label: "Grade",
        icon: GraduationCap,
        tone: "bg-emerald-100 text-emerald-700",
      };
    case "assignment":
      return {
        label: "Assignment",
        icon: UserCog,
        tone: "bg-amber-100 text-amber-700",
      };
    case "learner":
      return {
        label: "Learner",
        icon: Users,
        tone: "bg-sky-100 text-sky-700",
      };
    case "announcement":
      return {
        label: "Announcement",
        icon: Megaphone,
        tone: "bg-orange-100 text-orange-700",
      };
    case "account":
    case "security":
      return {
        label: category === "account" ? "Account" : "Security",
        icon: ShieldCheck,
        tone: "bg-rose-100 text-rose-700",
      };
    default:
      return {
        label: "System",
        icon: Bell,
        tone: "bg-slate-100 text-slate-700",
      };
  }
}

function isInFilter(item: NotificationItem, filter: FilterValue) {
  if (filter === "all") return true;
  if (filter === "unread") return !item.read_at;
  if (filter === "forms") {
    return item.category === "form_submission" || item.category === "form_review";
  }
  if (filter === "grades") {
    return item.category === "grade_submission" || item.category === "grade_review";
  }
  return ![
    "form_submission",
    "form_review",
    "grade_submission",
    "grade_review",
  ].includes(item.category);
}

function NotificationsLoadingSkeleton() {
  return (
    <div
      className="divide-y"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading notifications...</span>

      {Array.from({ length: 5 }).map((_, index) => (
        <article
          key={index}
          className={`flex gap-3 p-4 sm:gap-4 sm:p-5 ${
            index < 2 ? "bg-blue-50/55" : "bg-white"
          }`}
          aria-hidden="true"
        >
          <Skeleton className="size-10 shrink-0 rounded-xl sm:size-11" />

          <div className="min-w-0 flex-1">
            <div className="flex items-start justify-between gap-4">
              <div className="flex min-w-0 flex-1 items-center gap-2">
                {index < 2 && (
                  <Skeleton className="size-2 shrink-0 rounded-full" />
                )}
                <Skeleton
                  className={`h-5 ${index % 2 === 0 ? "w-48" : "w-36"}`}
                />
              </div>
              <Skeleton className="h-3 w-20 shrink-0" />
            </div>

            <div className="mt-2 space-y-2">
              <Skeleton className="h-3.5 w-full max-w-2xl" />
              <Skeleton
                className={`h-3.5 ${index % 2 === 0 ? "w-3/4" : "w-1/2"}`}
              />
            </div>

            <div className="mt-3 flex items-center justify-between gap-2">
              <Skeleton className="h-6 w-24 rounded-full" />
              {index < 2 && <Skeleton className="h-8 w-28 rounded-md" />}
            </div>
          </div>
        </article>
      ))}
    </div>
  );
}

function NotificationsPage() {
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const [filter, setFilter] = useState<FilterValue>("all");
  const [search, setSearch] = useState("");
  const [busyId, setBusyId] = useState<string | null>(null);
  const [markingAll, setMarkingAll] = useState(false);

  const { data, isLoading, isError, error, refetch } = useQuery({
    queryKey: [QUERY_KEY, "history"],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) return { userId: "", notifications: [] as NotificationItem[] };

      const { data: notifications, error: notificationsError } = await (
        supabase as any
      )
        .from("notifications")
        .select(
          "id, recipient_id, title, message, category, related_path, read_at, created_at",
        )
        .eq("recipient_id", user.id)
        .order("created_at", { ascending: false });

      if (notificationsError) throw notificationsError;

      return {
        userId: user.id,
        notifications: (notifications ?? []) as NotificationItem[],
      };
    },
    staleTime: 15_000,
    refetchOnWindowFocus: true,
  });

  const notifications = data?.notifications ?? [];
  const unreadCount = notifications.filter((item) => !item.read_at).length;

  useEffect(() => {
    if (!data?.userId) return;

    const channel = supabase
      .channel(`notification-history:${data.userId}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "notifications",
          filter: `recipient_id=eq.${data.userId}`,
        },
        () => {
          void queryClient.invalidateQueries({ queryKey: [QUERY_KEY] });
        },
      )
      .subscribe();

    return () => {
      void supabase.removeChannel(channel);
    };
  }, [data?.userId, queryClient]);

  const visibleNotifications = useMemo(() => {
    const needle = search.trim().toLowerCase();
    return notifications.filter((item) => {
      if (!isInFilter(item, filter)) return false;
      if (!needle) return true;
      return `${item.title} ${item.message} ${item.category}`
        .toLowerCase()
        .includes(needle);
    });
  }, [filter, notifications, search]);

  const markOneAsRead = async (item: NotificationItem, openRelated = false) => {
    setBusyId(item.id);
    try {
      if (!item.read_at) {
        const { error: updateError } = await (supabase as any)
          .from("notifications")
          .update({ read_at: new Date().toISOString() })
          .eq("id", item.id)
          .eq("recipient_id", data?.userId ?? "");

        if (updateError) throw updateError;
        await queryClient.invalidateQueries({ queryKey: [QUERY_KEY] });
      }

      if (openRelated && item.related_path) {
        navigate({ to: item.related_path as any });
      }
    } catch (updateError) {
      console.error("Could not update notification:", updateError);
      toast.error("The notification could not be marked as read.");
    } finally {
      setBusyId(null);
    }
  };

  const markAllAsRead = async () => {
    if (!data?.userId || unreadCount === 0) return;
    setMarkingAll(true);
    try {
      const { error: updateError } = await (supabase as any)
        .from("notifications")
        .update({ read_at: new Date().toISOString() })
        .eq("recipient_id", data.userId)
        .is("read_at", null);

      if (updateError) throw updateError;
      await queryClient.invalidateQueries({ queryKey: [QUERY_KEY] });
      toast.success("All notifications were marked as read.");
    } catch (updateError) {
      console.error("Could not update notifications:", updateError);
      toast.error("Notifications could not be marked as read.");
    } finally {
      setMarkingAll(false);
    }
  };

  return (
    <main className="min-h-[calc(100vh-80px)] bg-[#fbfaf7] px-4 py-6 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-5xl space-y-6">
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-start gap-3">
            <Button
              type="button"
              variant="outline"
              size="icon"
              className="mt-1 shrink-0 rounded-xl"
              onClick={() => window.history.back()}
              aria-label="Go back"
            >
              <ArrowLeft className="size-4" />
            </Button>
            <div>
              <h1 className="text-2xl font-bold tracking-tight text-slate-900">
                Notifications
              </h1>
              <p className="mt-1 text-sm text-slate-500">
                Requests, grade updates, school-form reviews, and system activity.
              </p>
            </div>
          </div>

          <Button
            type="button"
            variant="outline"
            onClick={() => void markAllAsRead()}
            disabled={unreadCount === 0 || markingAll}
            className="rounded-xl"
          >
            {markingAll ? (
              <Loader2 className="mr-2 size-4 animate-spin" />
            ) : (
              <CheckCheck className="mr-2 size-4" />
            )}
            Mark all as read
          </Button>
        </div>

        <section className="grid gap-3 sm:grid-cols-2">
          <div className="rounded-2xl border bg-white p-4 shadow-sm">
            <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
              Total notifications
            </p>
            <div className="mt-2 text-3xl font-bold text-slate-900">
              {isLoading ? (
                <Skeleton className="h-9 w-12" aria-label="Loading total" />
              ) : (
                notifications.length
              )}
            </div>
          </div>
          <div className="rounded-2xl border bg-white p-4 shadow-sm">
            <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">
              Unread
            </p>
            <div className="mt-2 text-3xl font-bold text-red-600">
              {isLoading ? (
                <Skeleton className="h-9 w-12" aria-label="Loading unread" />
              ) : (
                unreadCount
              )}
            </div>
          </div>
        </section>

        <section className="overflow-hidden rounded-2xl border bg-white shadow-sm">
          <div className="space-y-4 border-b p-4 sm:p-5">
            <div className="relative">
              <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-slate-400" />
              <Input
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                placeholder="Search notifications..."
                className="h-10 rounded-xl pl-9"
              />
            </div>

            <div className="flex gap-2 overflow-x-auto pb-1">
              {filters.map((item) => (
                <Button
                  key={item.value}
                  type="button"
                  size="sm"
                  variant={filter === item.value ? "default" : "outline"}
                  onClick={() => setFilter(item.value)}
                  className="shrink-0 rounded-full"
                >
                  {item.label}
                  {item.value === "unread" && unreadCount > 0 && (
                    <span className="ml-2 rounded-full bg-red-600 px-1.5 py-0.5 text-[10px] text-white">
                      {unreadCount}
                    </span>
                  )}
                </Button>
              ))}
            </div>
          </div>

          {isLoading ? (
            <NotificationsLoadingSkeleton />
          ) : isError ? (
            <div className="flex min-h-72 flex-col items-center justify-center gap-3 px-6 text-center">
              <Bell className="size-9 text-red-500" />
              <div>
                <p className="font-semibold text-slate-900">Could not load notifications</p>
                <p className="mt-1 text-sm text-slate-500">
                  {error instanceof Error ? error.message : "Please try again."}
                </p>
              </div>
              <Button variant="outline" onClick={() => void refetch()}>
                Try again
              </Button>
            </div>
          ) : visibleNotifications.length === 0 ? (
            <div className="flex min-h-72 flex-col items-center justify-center px-6 text-center">
              <div className="grid size-14 place-items-center rounded-full bg-slate-100">
                <Bell className="size-6 text-slate-500" />
              </div>
              <p className="mt-4 font-semibold text-slate-900">
                {notifications.length === 0
                  ? "No notifications yet"
                  : "No matching notifications"}
              </p>
              <p className="mt-1 max-w-sm text-sm text-slate-500">
                {notifications.length === 0
                  ? "New requests, updates, assignments, and review results will appear here."
                  : "Try another filter or search term."}
              </p>
            </div>
          ) : (
            <div className="divide-y">
              {visibleNotifications.map((item) => {
                const details = categoryDetails(item.category);
                const Icon = details.icon;
                const isUnread = !item.read_at;
                const isBusy = busyId === item.id;

                return (
                  <article
                    key={item.id}
                    className={cn(
                      "group flex gap-3 p-4 transition sm:gap-4 sm:p-5",
                      isUnread ? "bg-blue-50/55" : "bg-white",
                      item.related_path && "cursor-pointer hover:bg-slate-50",
                    )}
                    onClick={() => void markOneAsRead(item, true)}
                  >
                    <div
                      className={cn(
                        "grid size-10 shrink-0 place-items-center rounded-xl sm:size-11",
                        details.tone,
                      )}
                    >
                      <Icon className="size-5" />
                    </div>

                    <div className="min-w-0 flex-1">
                      <div className="flex flex-wrap items-start justify-between gap-x-4 gap-y-1">
                        <div className="flex min-w-0 items-center gap-2">
                          {isUnread && (
                            <span className="size-2 shrink-0 rounded-full bg-blue-600" />
                          )}
                          <h2 className="font-semibold text-slate-900">{item.title}</h2>
                        </div>
                        <time
                          dateTime={item.created_at}
                          title={format(new Date(item.created_at), "PPpp")}
                          className="shrink-0 text-xs text-slate-500"
                        >
                          {formatDistanceToNow(new Date(item.created_at), {
                            addSuffix: true,
                          })}
                        </time>
                      </div>

                      <p className="mt-1 text-sm leading-6 text-slate-600">{item.message}</p>

                      <div className="mt-3 flex flex-wrap items-center justify-between gap-2">
                        <span className="rounded-full bg-slate-100 px-2.5 py-1 text-[11px] font-semibold text-slate-600">
                          {details.label}
                        </span>

                        {isUnread && (
                          <Button
                            type="button"
                            variant="ghost"
                            size="sm"
                            disabled={isBusy}
                            className="h-8 text-xs"
                            onClick={(event) => {
                              event.stopPropagation();
                              void markOneAsRead(item);
                            }}
                          >
                            {isBusy ? (
                              <Loader2 className="mr-1.5 size-3.5 animate-spin" />
                            ) : (
                              <Check className="mr-1.5 size-3.5" />
                            )}
                            Mark as read
                          </Button>
                        )}
                      </div>
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </section>
      </div>
    </main>
  );
}