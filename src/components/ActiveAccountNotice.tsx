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
      className="fixed inset-0 z-[100] overflow-y-auto bg-black/50 px-3 py-4 backdrop-blur-[1px] sm:px-4 sm:py-8"
      role="dialog"
      aria-modal="true"
      aria-labelledby="active-account-notice-title"
    >
      <div className="flex min-h-full items-center justify-center">
        <div
          className="
            relative
            w-full
            max-w-[680px]
            max-h-[calc(100dvh-2rem)]
            overflow-y-auto
            overflow-x-hidden
            rounded-[22px]
            border
            border-[#ead7c0]
            bg-[#fffdf9]
            shadow-[0_30px_90px_rgba(60,35,25,0.28)]
            sm:max-h-[calc(100dvh-4rem)]
          "
        >
          <button
            type="button"
            onClick={onClose}
            className="
              absolute
              right-3
              top-3
              z-10
              grid
              size-9
              shrink-0
              place-items-center
              rounded-full
              text-[#5f342b]
              transition
              hover:bg-[#f7eee6]
              sm:right-5
              sm:top-5
            "
            aria-label="Close active account notice"
          >
            <X className="size-5" />
          </button>

          <div className="px-4 pb-5 pt-6 sm:px-8 sm:pb-8 sm:pt-9 lg:px-10">
            <div className="flex min-w-0 flex-col gap-4 pr-9 sm:flex-row sm:items-start sm:gap-5 sm:pr-10">
              <div
                className="
                  grid
                  size-12
                  shrink-0
                  place-items-center
                  rounded-2xl
                  bg-[#fff2db]
                  text-[#f4a62a]
                  sm:size-16
                "
              >
                <AlertTriangle
                  className="size-7 sm:size-10"
                  strokeWidth={2.2}
                />
              </div>

              <div className="min-w-0 flex-1">
                <h2
                  id="active-account-notice-title"
                  className="
                    break-words
                    text-xl
                    font-black
                    leading-tight
                    text-[#7f1d1d]
                    sm:text-2xl
                    lg:text-[30px]
                  "
                >
                  Account Active on Another Device
                </h2>

                <p className="mt-2 max-w-[540px] break-words text-sm leading-6 text-[#76564e] sm:text-[15px]">
                  This account has been detected on another device or browser.
                  This is only a notification to keep you informed.
                  <span className="font-semibold text-[#5f342b]">
                    {" "}
                    You will remain signed in and no automatic logout will happen.
                  </span>
                </p>
              </div>
            </div>

            <div className="mt-5 rounded-[18px] border border-[#ead7c0] bg-[#fffaf2] p-3 sm:mt-7 sm:p-5">
              <div className="mb-3 flex min-w-0 items-center gap-2 text-sm font-bold text-[#5f342b] sm:mb-4">
                <MonitorSmartphone className="size-4 shrink-0" />
                <span className="min-w-0 break-words">
                  {showAllSessions
                    ? "Other Active Sessions"
                    : "Detected Session"}
                </span>
              </div>

              <div className="max-h-[280px] space-y-3 overflow-y-auto pr-1">
                {sessionsToShow.map((session) => (
                  <div
                    key={session.id}
                    className="
                      flex
                      min-w-0
                      items-start
                      gap-3
                      rounded-2xl
                      border
                      border-[#efdfcc]
                      bg-white/80
                      p-3
                      sm:items-center
                      sm:gap-4
                      sm:p-4
                    "
                  >
                    <div className="grid size-10 shrink-0 place-items-center rounded-full bg-[#f8efe4] text-[#7f3028] sm:size-12">
                      <Laptop className="size-5 sm:size-6" />
                    </div>

                    <div className="min-w-0 flex-1">
                      <div className="break-words text-sm font-bold leading-5 text-[#3d211b] sm:text-base">
                        {sessionLabel(session)}
                      </div>

                      <div className="mt-1 break-words text-xs text-[#8a685f] sm:text-sm">
                        Active {relativeTime(session.last_seen_at)}
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {showAllSessions && activeOtherSessions.length === 0 && (
                <p className="text-sm leading-5 text-[#8a685f]">
                  No other active sessions are currently detected.
                </p>
              )}
            </div>

            <div className="mt-5 grid grid-cols-1 gap-3 sm:mt-7 sm:grid-cols-2">
              <Button
                type="button"
                onClick={onClose}
                className="
                  min-h-12
                  h-auto
                  w-full
                  whitespace-normal
                  rounded-xl
                  bg-[#8d2727]
                  px-4
                  py-3
                  text-sm
                  font-bold
                  leading-5
                  text-white
                  shadow-md
                  shadow-red-950/10
                  hover:bg-[#742020]
                  sm:text-base
                "
              >
                <Check className="mr-2 size-5 shrink-0" />
                <span className="break-words">OK</span>
              </Button>

              <Button
                type="button"
                variant="outline"
                onClick={() => setShowAllSessions((current) => !current)}
                className="
                  min-h-12
                  h-auto
                  w-full
                  whitespace-normal
                  rounded-xl
                  border-[#e5cdb4]
                  bg-[#fffaf2]
                  px-4
                  py-3
                  text-sm
                  font-semibold
                  leading-5
                  text-[#4c2a23]
                  hover:bg-[#f9efe4]
                  sm:text-base
                "
              >
                <UsersRound className="mr-2 size-5 shrink-0" />
                <span className="break-words text-center">
                  {showAllSessions
                    ? "Show Detected Session"
                    : "Review Active Sessions"}
                </span>
              </Button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
