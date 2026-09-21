import { useMemo } from "react";
import { useQuery } from "@tanstack/react-query";

import { supabase } from "../integrations/supabase/client";

export type OnlineAdmin = {
  user_id: string;
  full_name: string | null;
  email: string | null;
  avatar_url: string | null;
  last_seen_at: string;
};

const MAX_ADMINS = 3;

export function useAdminOnlineStatus() {
  const query = useQuery<OnlineAdmin[]>({
    queryKey: ["admin-online-status"],
    queryFn: async () => {
      const { data, error } = await (supabase as any).rpc(
        "get_online_admins",
      );

      if (error) {
        throw error;
      }

      return (data ?? []) as OnlineAdmin[];
    },
    refetchInterval: 5_000,
    refetchOnWindowFocus: true,
    staleTime: 2_000,
  });

  const onlineAdmins = query.data ?? [];

  const onlineAdminIds = useMemo(
    () => new Set(onlineAdmins.map((admin) => admin.user_id)),
    [onlineAdmins],
  );

  return {
    ...query,
    onlineAdmins,
    onlineAdminIds,
    onlineCount: onlineAdmins.length,
    maxAdmins: MAX_ADMINS,
  };
}
