import { supabase } from "../integrations/supabase/client";

const ADMIN_DEVICE_ID_KEY = "sigla.admin-device-id";
const ADMIN_LIMIT_NOTICE_KEY = "sigla.admin-access-limit";

export type AdminAccessClaim = {
  allowed: boolean;
  onlineCount: number;
  maxAdmins: number;
};

type AdminLimitNotice = {
  onlineCount: number;
  maxAdmins: number;
};

function getOrCreateAdminDeviceId() {
  if (typeof window === "undefined") {
    return "server";
  }

  const existing = window.localStorage.getItem(ADMIN_DEVICE_ID_KEY);

  if (existing) {
    return existing;
  }

  const created =
    typeof crypto !== "undefined" &&
    typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `admin-device-${Date.now()}-${Math.random().toString(36).slice(2)}`;

  window.localStorage.setItem(ADMIN_DEVICE_ID_KEY, created);

  return created;
}

function detectBrowserName() {
  if (typeof navigator === "undefined") return "Browser";

  const userAgent = navigator.userAgent;

  if (/Edg\//i.test(userAgent)) return "Microsoft Edge";
  if (/OPR\//i.test(userAgent)) return "Opera";
  if (/Chrome\//i.test(userAgent)) return "Chrome";
  if (/Firefox\//i.test(userAgent)) return "Firefox";
  if (/Safari\//i.test(userAgent)) return "Safari";

  return "Browser";
}

function detectDeviceName() {
  if (typeof navigator === "undefined") return "Device";

  const userAgent = navigator.userAgent;

  if (/iPad/i.test(userAgent)) return "iPad";
  if (/iPhone/i.test(userAgent)) return "iPhone";
  if (/Android/i.test(userAgent)) return "Android Device";
  if (/Windows/i.test(userAgent)) return "Windows Device";
  if (/Macintosh|Mac OS X/i.test(userAgent)) return "Mac";
  if (/Linux/i.test(userAgent)) return "Linux Computer";

  return "Device";
}

export async function claimAdminAccess(): Promise<AdminAccessClaim> {
  const deviceId = getOrCreateAdminDeviceId();

  const { data, error } = await (supabase as any).rpc(
    "claim_admin_session",
    {
      p_device_id: deviceId,
      p_device_name: detectDeviceName(),
      p_browser_name: detectBrowserName(),
    },
  );

  if (error) {
    throw error;
  }

  const row = Array.isArray(data) ? data[0] : data;

  return {
    allowed: Boolean(row?.allowed),
    onlineCount: Number(row?.online_count ?? 0),
    maxAdmins: Number(row?.max_admins ?? 3),
  };
}

export async function releaseAdminAccess() {
  const deviceId =
    typeof window === "undefined"
      ? "server"
      : window.localStorage.getItem(ADMIN_DEVICE_ID_KEY);

  if (!deviceId) return;

  const { error } = await (supabase as any).rpc(
    "release_admin_session",
    {
      p_device_id: deviceId,
    },
  );

  if (error) {
    throw error;
  }
}

export function rememberAdminAccessLimitNotice(
  onlineCount: number,
  maxAdmins: number,
) {
  if (typeof window === "undefined") return;

  const notice: AdminLimitNotice = {
    onlineCount,
    maxAdmins,
  };

  window.sessionStorage.setItem(
    ADMIN_LIMIT_NOTICE_KEY,
    JSON.stringify(notice),
  );
}

export function consumeAdminAccessLimitNotice(): AdminLimitNotice | null {
  if (typeof window === "undefined") return null;

  const raw = window.sessionStorage.getItem(ADMIN_LIMIT_NOTICE_KEY);

  if (!raw) return null;

  window.sessionStorage.removeItem(ADMIN_LIMIT_NOTICE_KEY);

  try {
    const parsed = JSON.parse(raw) as Partial<AdminLimitNotice>;

    return {
      onlineCount: Number(parsed.onlineCount ?? 0),
      maxAdmins: Number(parsed.maxAdmins ?? 3),
    };
  } catch {
    return {
      onlineCount: 3,
      maxAdmins: 3,
    };
  }
}
