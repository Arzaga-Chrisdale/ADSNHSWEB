import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "@tanstack/react-router";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { Bell, BellRing, CheckCheck, Loader2 } from "lucide-react";
import { formatDistanceToNow } from "date-fns";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import {
  Popover,
  PopoverContent,
  PopoverTrigger,
} from "@/components/ui/popover";
import { ScrollArea } from "@/components/ui/scroll-area";
import { cn } from "@/lib/utils";

export type NotificationCategory =
  | "account"
  | "assignment"
  | "learner"
  | "grade_submission"
  | "grade_review"
  | "form_submission"
  | "form_review"
  | "announcement"
  | "security"
  | "system";

export type NotificationItem = {
  id: string;
  recipient_id: string;
  title: string;
  message: string;
  category: NotificationCategory | string;
  related_path: string | null;
  read_at: string | null;
  created_at: string;
  metadata?: Record<string, unknown> | null;
};

type NotificationBellProps = {
  className?: string;
  historyPath?: string;
  limit?: number;
  onNotificationClick?: (
    notification: NotificationItem,
  ) => void | Promise<void>;
  onOpenHistory?: () => void;
};

const NOTIFICATIONS_QUERY_KEY = "current-user-notifications";

function categoryDot(category: string) {
  switch (category) {
    case "form_submission":
    case "form_review":
      return "bg-violet-500";
    case "grade_submission":
    case "grade_review":
      return "bg-emerald-500";
    case "account":
    case "security":
      return "bg-rose-500";
    case "assignment":
      return "bg-amber-500";
    case "learner":
      return "bg-sky-500";
    case "announcement":
      return "bg-orange-500";
    default:
      return "bg-slate-400";
  }
}

function relativeTime(value: string) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Just now";
  return formatDistanceToNow(date, { addSuffix: true });
}

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";
const SCHOOL_FORMS_NOTIFICATION_CLASS_EVENT =
  "school-forms-notification-class-change";

function notificationClassId(notification: NotificationItem) {
  const metadata = notification.metadata;
  if (!metadata || typeof metadata !== "object") return "";

  const candidate =
    metadata.class_id ??
    metadata.advisory_class_id ??
    metadata.subject_class_id;

  return typeof candidate === "string" ? candidate : "";
}

function prepareSchoolFormsNotification(notification: NotificationItem) {
  if (typeof window === "undefined") return;

  const path = notification.related_path ?? "";
  if (!path.startsWith("/school-forms")) return;

  const classId = notificationClassId(notification);
  if (!classId) return;

  window.sessionStorage.setItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY, classId);

  window.dispatchEvent(
    new CustomEvent(SCHOOL_FORMS_NOTIFICATION_CLASS_EVENT, {
      detail: { classId },
    }),
  );
}

export function NotificationBell({
  className,
  historyPath = "/notifications",
  limit = 20,
  onNotificationClick,
  onOpenHistory,
}: NotificationBellProps) {
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const [open, setOpen] = useState(false);

  const { data, isLoading } = useQuery({
    queryKey: [NOTIFICATIONS_QUERY_KEY, limit],
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;
      if (!user) return { userId: "", notifications: [] as NotificationItem[] };

      const { data: notifications, error } = await (supabase as any)
        .from("notifications")
        .select(
          "id, recipient_id, title, message, category, related_path, read_at, created_at, metadata",
        )
        .eq("recipient_id", user.id)
        .order("created_at", { ascending: false })
        .limit(limit);

      if (error) throw error;

      return {
        userId: user.id,
        notifications: (notifications ?? []) as NotificationItem[],
      };
    },
    staleTime: 15_000,
    refetchOnWindowFocus: true,
  });

  const notifications = data?.notifications ?? [];
  const unreadCount = useMemo(
    () => notifications.filter((item) => !item.read_at).length,
    [notifications],
  );

  useEffect(() => {
    if (!data?.userId) return;

    const channel = supabase
      .channel(`notifications:${data.userId}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "notifications",
          filter: `recipient_id=eq.${data.userId}`,
        },
        () => {
          void queryClient.invalidateQueries({
            queryKey: [NOTIFICATIONS_QUERY_KEY],
          });
        },
      )
      .subscribe();

    return () => {
      void supabase.removeChannel(channel);
    };
  }, [data?.userId, queryClient]);

  const markAsRead = async (notification: NotificationItem) => {
    if (!notification.read_at) {
      const { error } = await (supabase as any)
        .from("notifications")
        .update({ read_at: new Date().toISOString() })
        .eq("id", notification.id)
        .eq("recipient_id", data?.userId ?? "");

      if (error) throw error;

      await queryClient.invalidateQueries({
        queryKey: [NOTIFICATIONS_QUERY_KEY],
      });
    }

    setOpen(false);

    // When a School Forms notification contains a class id, remember that
    // exact class before navigating so School Forms opens the correct
    // Grade Level / Section instead of the last dropdown selection.
    prepareSchoolFormsNotification(notification);

    // Allow pages with internal navigation (such as the Admin dashboard)
    // to decide exactly which section a notification should open.
    if (onNotificationClick) {
      await onNotificationClick(notification);
      return;
    }

    // Preserve the original route-based behavior everywhere else.
    if (notification.related_path) {
      navigate({ to: notification.related_path as any });
    }
  };

  const markAllAsRead = async () => {
    if (!data?.userId || unreadCount === 0) return;

    const { error } = await (supabase as any)
      .from("notifications")
      .update({ read_at: new Date().toISOString() })
      .eq("recipient_id", data.userId)
      .is("read_at", null);

    if (error) throw error;

    await queryClient.invalidateQueries({
      queryKey: [NOTIFICATIONS_QUERY_KEY],
    });
  };

  const openHistory = () => {
    setOpen(false);

    if (onOpenHistory) {
      onOpenHistory();
      return;
    }

    navigate({ to: historyPath as any });
  };

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <button
          type="button"
          className={cn(
            "relative grid size-10 place-items-center rounded-xl border bg-background text-foreground shadow-sm transition hover:bg-accent focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring",
            className,
          )}
          aria-label={
            unreadCount > 0
              ? `Notifications, ${unreadCount} unread`
              : "Notifications"
          }
        >
          {unreadCount > 0 ? (
            <BellRing className="size-5" />
          ) : (
            <Bell className="size-5" />
          )}

          {unreadCount > 0 && (
            <span className="absolute -right-1.5 -top-1.5 grid min-h-5 min-w-5 place-items-center rounded-full bg-red-600 px-1 text-[10px] font-bold leading-none text-white ring-2 ring-background">
              {unreadCount > 99 ? "99+" : unreadCount}
            </span>
          )}
        </button>
      </PopoverTrigger>

      <PopoverContent align="end" className="w-[min(92vw,390px)] p-0">
        <div className="flex items-center justify-between gap-3 border-b px-4 py-3">
          <div>
            <h2 className="text-sm font-semibold">Notifications</h2>
            <p className="text-xs text-muted-foreground">
              {unreadCount === 0
                ? "You are all caught up"
                : `${unreadCount} unread notification${unreadCount === 1 ? "" : "s"}`}
            </p>
          </div>

          <Button
            type="button"
            variant="ghost"
            size="sm"
            onClick={() => void markAllAsRead()}
            disabled={unreadCount === 0}
            className="h-8 px-2 text-xs"
          >
            <CheckCheck className="mr-1 size-4" />
            Mark all read
          </Button>
        </div>

        <ScrollArea className="h-[min(60vh,420px)]">
          {isLoading ? (
            <div className="grid min-h-40 place-items-center text-muted-foreground">
              <Loader2 className="size-5 animate-spin" />
            </div>
          ) : notifications.length === 0 ? (
            <div className="flex min-h-52 flex-col items-center justify-center gap-2 px-6 text-center">
              <div className="grid size-11 place-items-center rounded-full bg-muted">
                <Bell className="size-5 text-muted-foreground" />
              </div>
              <p className="text-sm font-medium">No notifications yet</p>
              <p className="text-xs text-muted-foreground">
                New requests, grade updates, form reviews, and assignments will
                appear here.
              </p>
            </div>
          ) : (
            <div className="divide-y">
              {notifications.map((notification) => (
                <button
                  key={notification.id}
                  type="button"
                  onClick={() => void markAsRead(notification)}
                  className={cn(
                    "flex w-full gap-3 px-4 py-3 text-left transition hover:bg-accent",
                    !notification.read_at && "bg-primary/[0.06]",
                  )}
                >
                  <span
                    className={cn(
                      "mt-1.5 size-2.5 shrink-0 rounded-full",
                      notification.read_at
                        ? "bg-muted-foreground/25"
                        : categoryDot(notification.category),
                    )}
                  />

                  <span className="min-w-0 flex-1">
                    <span className="flex items-start justify-between gap-2">
                      <span
                        className={cn(
                          "text-sm",
                          notification.read_at ? "font-medium" : "font-semibold",
                        )}
                      >
                        {notification.title}
                      </span>
                      {!notification.read_at && (
                        <span className="mt-1 size-2 shrink-0 rounded-full bg-primary" />
                      )}
                    </span>

                    <span className="mt-0.5 line-clamp-2 block text-xs leading-5 text-muted-foreground">
                      {notification.message}
                    </span>
                    <span className="mt-1 block text-[11px] text-muted-foreground/80">
                      {relativeTime(notification.created_at)}
                    </span>
                  </span>
                </button>
              ))}
            </div>
          )}
        </ScrollArea>

        <div className="border-t p-2">
          <Button
            type="button"
            variant="ghost"
            onClick={openHistory}
            className="w-full text-xs font-semibold"
          >
            View all notifications
          </Button>
        </div>
      </PopoverContent>
    </Popover>
  );
}

export default NotificationBell;
