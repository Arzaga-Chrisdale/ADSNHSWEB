import { useEffect, useMemo, useState } from "react";
import {
  AlertTriangle,
  Check,
  Laptop,
  MonitorSmartphone,
  UsersRound,
  X,
} from "lucide-react";

import { Button } from "./ui/button";
import type { ActiveAccountSession } from "../hooks/useActiveAccountNotice";

function relativeTime(value: string) {
  const timestamp = new Date(value).getTime();

  if (Number.isNaN(timestamp)) return "just now";

  const seconds = Math.max(
    0,
    Math.floor((Date.now() - timestamp) / 1000),
  );

  if (seconds < 60) return "just now";

  const minutes = Math.floor(seconds / 60);

  if (minutes === 1) return "1 minute ago";
  if (minutes < 60) return `${minutes} minutes ago`;

  const hours = Math.floor(minutes / 60);

  if (hours === 1) return "1 hour ago";

  return `${hours} hours ago`;
}

function sessionLabel(session: ActiveAccountSession) {
  const device = session.device_name?.trim() || "Another device";
  const browser = session.browser_name?.trim() || "Browser";

  return `${device} • ${browser}`;
}

export function ActiveAccountNotice({
  open,
  detectedSession,
  activeOtherSessions,
  onClose,
}: {
  open: boolean;
  detectedSession: ActiveAccountSession | null;
  activeOtherSessions: ActiveAccountSession[];
  onClose: () => void;
}) {
  const [showAllSessions, setShowAllSessions] = useState(false);

  useEffect(() => {
    if (open) {
      setShowAllSessions(false);
    }
  }, [open, detectedSession?.id]);

  const sessionsToShow = useMemo(() => {
    if (!showAllSessions) {
      return detectedSession ? [detectedSession] : [];
    }

    return activeOtherSessions;
  }, [showAllSessions, detectedSession, activeOtherSessions]);

  if (!open || !detectedSession) {
    return null;
  }

  return (
    <div
      className="fixed inset-0 z-[100] flex items-center justify-center bg-black/50 px-4 py-8 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
      aria-labelledby="active-account-notice-title"
    >
      <div className="relative w-full max-w-[680px] overflow-hidden rounded-[22px] border border-[#ead7c0] bg-[#fffdf9] shadow-[0_30px_90px_rgba(60,35,25,0.28)]">
        <button
          type="button"
          onClick={onClose}
          className="absolute right-5 top-5 grid size-9 place-items-center rounded-full text-[#5f342b] transition hover:bg-[#f7eee6]"
          aria-label="Close active account notice"
        >
          <X className="size-5" />
        </button>

        <div className="px-8 pb-8 pt-9 sm:px-10">
          <div className="flex items-start gap-5 pr-10">
            <div className="grid size-16 shrink-0 place-items-center rounded-2xl bg-[#fff2db] text-[#f4a62a]">
              <AlertTriangle className="size-10" strokeWidth={2.2} />
            </div>

            <div className="min-w-0">
              <h2
                id="active-account-notice-title"
                className="text-2xl font-black text-[#7f1d1d] sm:text-[30px]"
              >
                Account Active on Another Device
              </h2>

              <p className="mt-2 max-w-[540px] text-[15px] leading-6 text-[#76564e]">
                This account has been detected on another device or browser.
                This is only a notification to keep you informed.
                <span className="font-semibold text-[#5f342b]">
                  {" "}
                  You will remain signed in and no automatic logout will happen.
                </span>
              </p>
            </div>
          </div>

          <div className="mt-7 rounded-[18px] border border-[#ead7c0] bg-[#fffaf2] p-5">
            <div className="mb-4 flex items-center gap-2 text-sm font-bold text-[#5f342b]">
              <MonitorSmartphone className="size-4" />
              {showAllSessions
                ? "Other Active Sessions"
                : "Detected Session"}
            </div>

            <div className="space-y-3">
              {sessionsToShow.map((session) => (
                <div
                  key={session.id}
                  className="flex items-center gap-4 rounded-2xl border border-[#efdfcc] bg-white/80 p-4"
                >
                  <div className="grid size-12 shrink-0 place-items-center rounded-full bg-[#f8efe4] text-[#7f3028]">
                    <Laptop className="size-6" />
                  </div>

                  <div className="min-w-0 flex-1">
                    <div className="truncate text-base font-bold text-[#3d211b]">
                      {sessionLabel(session)}
                    </div>

                    <div className="mt-1 text-sm text-[#8a685f]">
                      Active {relativeTime(session.last_seen_at)}
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {showAllSessions && activeOtherSessions.length === 0 && (
              <p className="text-sm text-[#8a685f]">
                No other active sessions are currently detected.
              </p>
            )}
          </div>

          <div className="mt-7 grid gap-3 sm:grid-cols-2">
            <Button
              type="button"
              onClick={onClose}
              className="h-12 rounded-xl bg-[#8d2727] text-base font-bold text-white shadow-md shadow-red-950/10 hover:bg-[#742020]"
            >
              <Check className="mr-2 size-5" />
              OK
            </Button>

            <Button
              type="button"
              variant="outline"
              onClick={() => setShowAllSessions((current) => !current)}
              className="h-12 rounded-xl border-[#e5cdb4] bg-[#fffaf2] text-base font-semibold text-[#4c2a23] hover:bg-[#f9efe4]"
            >
              <UsersRound className="mr-2 size-5" />
              {showAllSessions
                ? "Show Detected Session"
                : "Review Active Sessions"}
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}
