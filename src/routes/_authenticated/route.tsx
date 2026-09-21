import { createFileRoute, Outlet, redirect } from "@tanstack/react-router";
import { supabase } from "../../integrations/supabase/client";
import { AppShell } from "../../components/AppShell";
import { ActiveAccountNotice } from "../../components/ActiveAccountNotice";
import { useActiveAccountNotice } from "../../hooks/useActiveAccountNotice";
import { useAdminSessionPresence } from "@/hooks/useAdminSessionPresence";

export const Route = createFileRoute("/_authenticated")({
  ssr: false,
  beforeLoad: async () => {
    const { data, error } = await supabase.auth.getUser();

    if (error || !data.user) {
      throw redirect({ to: "/auth" });
    }

    return { user: data.user };
  },
  component: AuthenticatedLayout,
});

function AuthenticatedLayout() {
  // Class Adviser / Subject Teacher:
  // notification only when the same account is active on another device.
  const {
    enabled,
    open,
    detectedSession,
    activeOtherSessions,
    dismissNotice,
  } = useActiveAccountNotice();

  // Admin:
  // keeps the current Admin presence alive and enforces the 3-admin limit.
  useAdminSessionPresence();

  return (
    <>
      <AppShell>
        <Outlet />
      </AppShell>

      {enabled && (
        <ActiveAccountNotice
          open={open}
          detectedSession={detectedSession}
          activeOtherSessions={activeOtherSessions}
          onClose={dismissNotice}
        />
      )}
    </>
  );
}
