import { useEffect } from "react";

import { supabase } from "../integrations/supabase/client";
import {
  claimAdminAccess,
  rememberAdminAccessLimitNotice,
} from "../lib/admin-session";

const ADMIN_HEARTBEAT_MS = 30_000;

export function useAdminSessionPresence() {
  useEffect(() => {
    let mounted = true;
    let heartbeatTimer: number | null = null;

    const stopHeartbeat = () => {
      if (heartbeatTimer !== null) {
        window.clearInterval(heartbeatTimer);
        heartbeatTimer = null;
      }
    };

    const redirectBecauseLimitWasReached = async (
      onlineCount: number,
      maxAdmins: number,
    ) => {
      rememberAdminAccessLimitNotice(onlineCount, maxAdmins);

      stopHeartbeat();

      try {
        await supabase.auth.signOut();
      } finally {
        if (typeof window !== "undefined") {
          window.location.replace("/auth");
        }
      }
    };

    const refreshAdminPresence = async () => {
      try {
        const { data: userData, error: userError } =
          await supabase.auth.getUser();

        if (userError || !userData.user || !mounted) {
          return;
        }

        const { data: isAdmin, error: adminError } =
          await supabase.rpc("is_admin");

        if (adminError || isAdmin !== true || !mounted) {
          return;
        }

        const claim = await claimAdminAccess();

        if (!claim.allowed && mounted) {
          await redirectBecauseLimitWasReached(
            claim.onlineCount,
            claim.maxAdmins,
          );
        }
      } catch (error) {
        console.warn("Unable to refresh Admin online presence:", error);
      }
    };

    void refreshAdminPresence();

    heartbeatTimer = window.setInterval(() => {
      void refreshAdminPresence();
    }, ADMIN_HEARTBEAT_MS);

    const handleVisibilityChange = () => {
      if (document.visibilityState === "visible") {
        void refreshAdminPresence();
      }
    };

    document.addEventListener(
      "visibilitychange",
      handleVisibilityChange,
    );

    return () => {
      mounted = false;
      stopHeartbeat();

      document.removeEventListener(
        "visibilitychange",
        handleVisibilityChange,
      );
    };
  }, []);
}
