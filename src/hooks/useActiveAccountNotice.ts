import { useCallback, useEffect, useRef, useState } from "react";
import { toast } from "sonner";

import { supabase } from "../integrations/supabase/client";

export type ActiveAccountSession = {
  id: string;
  user_id: string;
  device_id: string;
  device_name: string | null;
  browser_name: string | null;
  last_seen_at: string;
  created_at: string;
};

type TeacherType =
  | "class_adviser"
  | "subject_teacher"
  | "admin"
  | string
  | null;

const DEVICE_ID_KEY = "sigla.active-account-device-id";
const HEARTBEAT_MS = 30_000;
const ACTIVE_WINDOW_MS = 90_000;
const STALE_CLEANUP_MS = 24 * 60 * 60 * 1000;

// The generated Supabase TypeScript types may not yet contain the new
// active_account_sessions table. This cast is intentionally local to this
// feature so the rest of your Supabase client stays fully typed.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const sessionDb = supabase as any;

function getOrCreateDeviceId() {
  if (typeof window === "undefined") {
    return "server";
  }

  const existing =
    window.localStorage.getItem(DEVICE_ID_KEY);

  if (existing) {
    return existing;
  }

  const created =
    typeof crypto !== "undefined" &&
    typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `device-${Date.now()}-${Math.random()
          .toString(36)
          .slice(2)}`;

  window.localStorage.setItem(DEVICE_ID_KEY, created);

  return created;
}

function detectBrowserName() {
  if (typeof navigator === "undefined") {
    return "Browser";
  }

  const userAgent = navigator.userAgent;

  if (/Edg\//i.test(userAgent)) return "Microsoft Edge";
  if (/OPR\//i.test(userAgent)) return "Opera";
  if (/Chrome\//i.test(userAgent)) return "Chrome";
  if (/Firefox\//i.test(userAgent)) return "Firefox";
  if (/Safari\//i.test(userAgent)) return "Safari";

  return "Browser";
}

function detectDeviceName() {
  if (typeof navigator === "undefined") {
    return "Device";
  }

  const userAgent = navigator.userAgent;

  if (/iPad/i.test(userAgent)) return "iPad";
  if (/iPhone/i.test(userAgent)) return "iPhone";
  if (/Android/i.test(userAgent)) return "Android Device";
  if (/Windows/i.test(userAgent)) return "Windows Device";
  if (/Macintosh|Mac OS X/i.test(userAgent)) return "Mac";
  if (/Linux/i.test(userAgent)) return "Linux Computer";

  return "Device";
}

function activeSinceIso() {
  return new Date(
    Date.now() - ACTIVE_WINDOW_MS,
  ).toISOString();
}

function staleBeforeIso() {
  return new Date(
    Date.now() - STALE_CLEANUP_MS,
  ).toISOString();
}

async function getTeacherType(
  userId: string,
): Promise<TeacherType> {
  const { data, error } = await supabase
    .from("profiles")
    .select("teacher_type")
    .eq("id", userId)
    .maybeSingle();

  if (error) {
    throw error;
  }

  return data?.teacher_type ?? null;
}

function isTeacherNoticeEnabled(
  teacherType: TeacherType,
) {
  return (
    teacherType === "class_adviser" ||
    teacherType === "subject_teacher"
  );
}

async function registerCurrentDevice(userId: string) {
  const deviceId = getOrCreateDeviceId();
  const now = new Date().toISOString();

  const { error } = await sessionDb
    .from("active_account_sessions")
    .upsert(
      {
        user_id: userId,
        device_id: deviceId,
        device_name: detectDeviceName(),
        browser_name: detectBrowserName(),
        last_seen_at: now,
      },
      {
        onConflict: "user_id,device_id",
      },
    );

  if (error) {
    throw error;
  }

  // Clean only stale rows belonging to the same account.
  // This does not log out or remove another currently active device.
  const { error: cleanupError } = await sessionDb
    .from("active_account_sessions")
    .delete()
    .eq("user_id", userId)
    .lt("last_seen_at", staleBeforeIso());

  if (cleanupError) {
    console.warn(
      "Unable to remove stale active-session rows:",
      cleanupError,
    );
  }

  return deviceId;
}

async function listOtherActiveSessions(
  userId: string,
  deviceId: string,
): Promise<ActiveAccountSession[]> {
  const { data, error } = await sessionDb
    .from("active_account_sessions")
    .select(
      [
        "id",
        "user_id",
        "device_id",
        "device_name",
        "browser_name",
        "last_seen_at",
        "created_at",
      ].join(", "),
    )
    .eq("user_id", userId)
    .neq("device_id", deviceId)
    .gte("last_seen_at", activeSinceIso())
    .order("last_seen_at", {
      ascending: false,
    });

  if (error) {
    throw error;
  }

  return (data ?? []) as ActiveAccountSession[];
}

export function useActiveAccountNotice() {
  const [enabled, setEnabled] = useState(false);
  const [open, setOpen] = useState(false);

  const [
    detectedSession,
    setDetectedSession,
  ] = useState<ActiveAccountSession | null>(null);

  const [
    activeOtherSessions,
    setActiveOtherSessions,
  ] = useState<ActiveAccountSession[]>([]);

  const mountedRef = useRef(true);
  const userIdRef = useRef("");
  const deviceIdRef = useRef("");

  const dismissedSessionIdsRef =
    useRef<Set<string>>(new Set());

  const refreshOtherSessions =
    useCallback(async () => {
      const userId = userIdRef.current;
      const deviceId = deviceIdRef.current;

      if (!userId || !deviceId) {
        return;
      }

      try {
        const sessions =
          await listOtherActiveSessions(
            userId,
            deviceId,
          );

        if (!mountedRef.current) {
          return;
        }

        setActiveOtherSessions(sessions);

        const nextUndismissed =
          sessions.find(
            (session) =>
              !dismissedSessionIdsRef.current.has(
                session.id,
              ),
          );

        if (nextUndismissed) {
          setDetectedSession((current) => {
            if (
              current?.id !==
              nextUndismissed.id
            ) {
              toast.warning(
                "This account is also active on another device.",
              );
            }

            return nextUndismissed;
          });

          setOpen(true);
          return;
        }

        if (sessions.length === 0) {
          setDetectedSession(null);
          setOpen(false);
        }
      } catch (error) {
        console.warn(
          "Unable to check other active sessions:",
          error,
        );
      }
    }, []);

  const dismissNotice = useCallback(() => {
    if (detectedSession) {
      dismissedSessionIdsRef.current.add(
        detectedSession.id,
      );
    }

    setOpen(false);
  }, [detectedSession]);

  useEffect(() => {
    mountedRef.current = true;

    let heartbeatTimer: number | null = null;

    let realtimeChannel:
      | ReturnType<typeof supabase.channel>
      | null = null;

    const start = async () => {
      try {
        const { data, error } =
          await supabase.auth.getUser();

        if (
          error ||
          !data.user ||
          !mountedRef.current
        ) {
          return;
        }

        const teacherType =
          await getTeacherType(data.user.id);

        if (
          !mountedRef.current ||
          !isTeacherNoticeEnabled(teacherType)
        ) {
          setEnabled(false);
          return;
        }

        setEnabled(true);

        userIdRef.current = data.user.id;
        deviceIdRef.current =
          await registerCurrentDevice(
            data.user.id,
          );

        await refreshOtherSessions();

        realtimeChannel = supabase
          .channel(
            `active-account-notice:${data.user.id}`,
          )
          .on(
            "postgres_changes",
            {
              event: "*",
              schema: "public",
              table: "active_account_sessions",
              filter: `user_id=eq.${data.user.id}`,
            },
            () => {
              void refreshOtherSessions();
            },
          )
          .subscribe();

        heartbeatTimer =
          window.setInterval(() => {
            void (async () => {
              try {
                await registerCurrentDevice(
                  data.user.id,
                );

                await refreshOtherSessions();
              } catch (heartbeatError) {
                console.warn(
                  "Unable to refresh active-account session:",
                  heartbeatError,
                );
              }
            })();
          }, HEARTBEAT_MS);
      } catch (error) {
        console.warn(
          "Unable to start active-account notice:",
          error,
        );
      }
    };

    void start();

    const handleVisibilityChange = () => {
      if (
        document.visibilityState === "visible"
      ) {
        void refreshOtherSessions();
      }
    };

    document.addEventListener(
      "visibilitychange",
      handleVisibilityChange,
    );

    return () => {
      mountedRef.current = false;

      document.removeEventListener(
        "visibilitychange",
        handleVisibilityChange,
      );

      if (heartbeatTimer !== null) {
        window.clearInterval(heartbeatTimer);
      }

      if (realtimeChannel) {
        void supabase.removeChannel(
          realtimeChannel,
        );
      }
    };
  }, [refreshOtherSessions]);

  return {
    enabled,
    open,
    detectedSession,
    activeOtherSessions,
    dismissNotice,
  };
}
