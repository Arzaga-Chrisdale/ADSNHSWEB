import { createFileRoute, useNavigate } from "@tanstack/react-router";
import emailjs from "@emailjs/browser";
import {
  useEffect,
  useState,
  type ChangeEvent,
  type FormEvent,
  type ComponentType,
} from "react";
import { supabase } from "@/integrations/supabase/client";
import {
  claimAdminAccess,
  consumeAdminAccessLimitNotice,
} from "../lib/admin-session";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Tabs, TabsContent } from "@/components/ui/tabs";
import { Checkbox } from "@/components/ui/checkbox";
import { Skeleton } from "@/components/ui/skeleton";
import logo from "@/assets/ASNSHS Logo.png";
import {
  Mail,
  Lock,
  Shield,
  AlertCircle,
  CheckCircle2,
  Loader2,
  Eye,
  EyeOff,
} from "lucide-react";

export const Route = createFileRoute("/auth")({
  component: AuthPage,
});

type FeedbackState = {
  open: boolean;
  variant: "loading" | "error" | "success";
  title: string;
  message: string;
  buttonText?: string;
};

type BrowserNetworkInformation = {
  effectiveType?: "slow-2g" | "2g" | "3g" | "4g" | string;
  downlink?: number;
  rtt?: number;
  saveData?: boolean;
  addEventListener?: (type: "change", listener: () => void) => void;
  removeEventListener?: (type: "change", listener: () => void) => void;
};

type NetworkState = {
  online: boolean;
  effectiveType: string;
  downlink?: number;
  rtt?: number;
  slow: boolean;
};

function readNetworkState(): NetworkState {
  if (typeof navigator === "undefined") {
    return {
      online: true,
      effectiveType: "unknown",
      slow: false,
    };
  }

  const connection = (
    navigator as Navigator & {
      connection?: BrowserNetworkInformation;
      mozConnection?: BrowserNetworkInformation;
      webkitConnection?: BrowserNetworkInformation;
    }
  ).connection ??
    (
      navigator as Navigator & {
        mozConnection?: BrowserNetworkInformation;
      }
    ).mozConnection ??
    (
      navigator as Navigator & {
        webkitConnection?: BrowserNetworkInformation;
      }
    ).webkitConnection;

  const effectiveType = connection?.effectiveType ?? "unknown";
  const downlink = connection?.downlink;
  const rtt = connection?.rtt;

  const slowByType =
    effectiveType === "slow-2g" ||
    effectiveType === "2g" ||
    effectiveType === "3g";

  const slowByMetrics =
    (typeof downlink === "number" && downlink > 0 && downlink < 1.5) ||
    (typeof rtt === "number" && rtt >= 500);

  return {
    online: navigator.onLine,
    effectiveType,
    downlink,
    rtt,
    slow: slowByType || slowByMetrics,
  };
}

function getSignInLoadingMessage(network: NetworkState) {
  if (!network.online) {
    return "No internet connection.";
  }

  if (network.slow) {
    return "Slow network ,Please wait.";
  }

  return "Please wait.";
}

const REMEMBER_ME_KEY = "sigla.rememberMe";
const REMEMBERED_EMAIL_KEY = "sigla.rememberedEmail";

function getInitialRememberMe() {
  if (typeof window === "undefined") return false;
  return window.localStorage.getItem(REMEMBER_ME_KEY) === "true";
}

function getInitialRememberedEmail() {
  if (typeof window === "undefined") return "";

  const shouldRemember =
    window.localStorage.getItem(REMEMBER_ME_KEY) === "true";

  return shouldRemember
    ? window.localStorage.getItem(REMEMBERED_EMAIL_KEY) ?? ""
    : "";
}


type BrowserPasswordCredential = Credential & {
  id: string;
  password?: string;
};

type PasswordCredentialConstructor = new (data: {
  id: string;
  password: string;
  name?: string;
}) => BrowserPasswordCredential;

/**
 * Uses the browser's credential/password manager when it is supported.
 * The raw password is NOT written to localStorage or sessionStorage.
 */
async function saveCredentialToBrowser(email: string, password: string) {
  if (typeof window === "undefined" || !navigator.credentials) return;

  const PasswordCredentialCtor = (
    window as typeof window & {
      PasswordCredential?: PasswordCredentialConstructor;
    }
  ).PasswordCredential;

  if (!PasswordCredentialCtor) return;

  try {
    const credential = new PasswordCredentialCtor({
      id: email,
      password,
      name: email,
    });

    await navigator.credentials.store(credential);
  } catch {
    // Some browsers/password managers do not support programmatic storage.
    // Normal autocomplete="username/current-password" still works as fallback.
  }
}

async function getCredentialFromBrowser() {
  if (typeof window === "undefined" || !navigator.credentials) return null;

  try {
    const credentials = navigator.credentials as CredentialsContainer & {
      get: (
        options?: CredentialRequestOptions & {
          password?: boolean;
          mediation?: "silent" | "optional" | "required" | "conditional";
        },
      ) => Promise<Credential | null>;
    };

    const credential = (await credentials.get({
      password: true,
      mediation: "optional",
    })) as BrowserPasswordCredential | null;

    if (
      credential &&
      typeof credential.id === "string" &&
      typeof credential.password === "string"
    ) {
      return {
        email: credential.id,
        password: credential.password,
      };
    }
  } catch {
    // Ignore unsupported credential-manager behavior.
  }

  return null;
}

type DashboardNavigationResult =
  | {
      status: "navigated";
      target: "admin" | "dashboard";
    }
  | {
      status: "admin_limit";
      onlineCount: number;
      maxAdmins: number;
    };

async function goToCorrectDashboard(
  navigate: ReturnType<typeof useNavigate>,
): Promise<DashboardNavigationResult> {
  const { data, error } = await supabase.rpc("is_admin");

  if (!error && data === true) {
    const claim = await claimAdminAccess();

    if (!claim.allowed) {
      return {
        status: "admin_limit",
        onlineCount: claim.onlineCount,
        maxAdmins: claim.maxAdmins,
      };
    }

    await navigate({ to: "/admin", replace: true });

    return {
      status: "navigated",
      target: "admin",
    };
  }

  await navigate({ to: "/dashboard", replace: true });

  return {
    status: "navigated",
    target: "dashboard",
  };
}

function AuthPageSkeleton() {
  return (
    <div
      className="relative min-h-screen overflow-hidden bg-[var(--login-bg)] text-[var(--login-brown)]"
      role="status"
      aria-live="polite"
      aria-busy="true"
    >
      <span className="sr-only">Loading the school portal...</span>

      <div className="pointer-events-none absolute left-[-140px] top-[-90px] h-[430px] w-[430px] rounded-full border border-[var(--login-gold)] opacity-25" />
      <div className="pointer-events-none absolute right-[-160px] bottom-[-130px] h-[520px] w-[520px] rounded-full border border-[var(--login-gold)] opacity-20" />
      <div className="pointer-events-none absolute right-[-40px] top-[-70px] h-[220px] w-[560px] rounded-bl-[100%] bg-[var(--login-red)] opacity-75" />
      <div className="pointer-events-none absolute bottom-[170px] left-[140px] h-[420px] w-[420px] rounded-full bg-[var(--login-soft-circle)] opacity-60" />

      <main className="relative z-10 flex min-h-screen flex-col items-center justify-center px-6 py-10">
        <div className="mb-8 flex w-full max-w-4xl flex-col items-center">
          <Skeleton className="mb-4 h-24 w-24 rounded-full bg-[#eadfd4]" />
          <Skeleton className="h-12 w-full max-w-[440px] rounded-xl bg-[#eadfd4]" />
          <Skeleton className="mt-3 h-7 w-full max-w-[560px] rounded-lg bg-[#eadfd4]/90" />
          <Skeleton className="mt-6 h-10 w-36 rounded-lg bg-[#eadfd4]" />

          <div className="mt-4 flex w-full max-w-[460px] items-center gap-4">
            <Skeleton className="h-px flex-1 rounded-none bg-[var(--login-gold)]/50" />
            <Skeleton className="h-5 w-44 bg-[#eadfd4]" />
            <Skeleton className="h-px flex-1 rounded-none bg-[var(--login-gold)]/50" />
          </div>
        </div>

        <div className="w-full max-w-[720px] overflow-hidden rounded-[28px] border border-[var(--login-border)] bg-white/90 shadow-[0_25px_70px_rgba(124,74,47,0.12)] backdrop-blur">
          <div className="flex flex-col items-center border-b border-[var(--login-border)] bg-white/70 px-8 py-6">
            <Skeleton className="h-9 w-36 bg-[#eadfd4]" />
            <Skeleton className="mt-3 h-4 w-64 bg-[#f0e8e1]" />
          </div>

          <div className="space-y-6 px-11 py-9">
            {["email", "password"].map((field) => (
              <div key={field} className="space-y-2">
                <Skeleton className="h-5 w-24 bg-[#eadfd4]" />
                <div className="flex h-12 items-center gap-3 rounded-[14px] border border-[var(--login-border)] bg-white/70 px-4">
                  <Skeleton className="h-5 w-5 shrink-0 rounded-full bg-[#dccdc2]" />
                  <Skeleton className="h-4 w-2/3 bg-[#eee4dc]" />
                </div>
              </div>
            ))}

            <div className="flex items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <Skeleton className="h-5 w-5 bg-[#eadfd4]" />
                <Skeleton className="h-5 w-28 bg-[#eadfd4]" />
              </div>
              <Skeleton className="h-5 w-32 bg-[#eadfd4]" />
            </div>

            <Skeleton className="h-[62px] w-full rounded-xl bg-[var(--login-red)]/25" />
          </div>
        </div>

        <footer className="mt-7 flex flex-col items-center">
          <Skeleton className="mb-3 h-11 w-11 rounded-full bg-[#eadfd4]" />
          <Skeleton className="h-5 w-80 max-w-[80vw] bg-[#eadfd4]" />
          <Skeleton className="mt-3 h-5 w-44 bg-[#eadfd4]" />
        </footer>
      </main>
    </div>
  );
}

function AuthField({
  label,
  type = "text",
  value,
  onChange,
  placeholder,
  icon: Icon,
  name,
  autoComplete,
}: {
  label: string;
  type?: string;
  value: string;
  onChange: (event: ChangeEvent<HTMLInputElement>) => void;
  placeholder?: string;
  icon: ComponentType<{ className?: string }>;
  name?: string;
  autoComplete?: string;
}) {
  return (
    <div className="space-y-2">
      <Label className="text-[15px] font-medium text-[var(--login-brown)]">
        {label}
      </Label>

      <div className="relative">
        <Icon className="absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-[#9b7c67]" />

        <Input
          type={type}
          name={name}
          autoComplete={autoComplete}
          value={value}
          onChange={onChange}
          placeholder={placeholder}
          className="h-12 rounded-[14px] border border-[var(--login-border)] bg-white/70 pl-11 pr-4 text-[15px] text-[var(--login-brown)] shadow-sm outline-none focus-visible:border-[var(--login-border)] focus-visible:ring-0"
        />
      </div>
    </div>
  );
}

function FeedbackModal({
  feedback,
  onClose,
}: {
  feedback: FeedbackState | null;
  onClose: () => void;
}) {
  if (!feedback?.open) return null;

  const isLoading = feedback.variant === "loading";
  const isError = feedback.variant === "error";

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#efe9e2]/70 px-4 backdrop-blur-[2px]">
      <div className="w-full max-w-[390px] rounded-[24px] border border-[var(--login-border)] bg-white px-8 py-8 text-center shadow-[0_20px_60px_rgba(115,67,45,0.22)]">
        <div className="mb-5 flex justify-center">
          {isLoading ? (
            <div className="grid h-20 w-20 place-items-center rounded-full bg-[#fdf6f4]">
              <Loader2 className="h-10 w-10 animate-spin text-[var(--login-red)]" />
            </div>
          ) : (
            <div
              className={`grid h-20 w-20 place-items-center rounded-full border-4 ${
                isError
                  ? "border-[#f5d7d6] bg-[#fff8f8]"
                  : "border-[#dcedd9] bg-[#f8fff8]"
              }`}
            >
              {isError ? (
                <AlertCircle className="h-10 w-10 text-[var(--login-red)]" />
              ) : (
                <CheckCircle2 className="h-10 w-10 text-green-600" />
              )}
            </div>
          )}
        </div>

        <h3 className="text-[30px] font-black leading-tight text-[var(--login-red)]">
          {feedback.title}
        </h3>

        <p className="mt-4 text-[17px] leading-7 text-[var(--login-brown)]">
          {feedback.message}
        </p>

        {!isLoading && (
          <Button
            type="button"
            onClick={onClose}
            className="mt-7 h-12 w-full rounded-xl bg-[var(--login-red)] text-base font-bold text-white shadow-lg shadow-red-900/15 hover:bg-[var(--login-red-dark)]"
          >
            {feedback.buttonText ?? "OK"}
          </Button>
        )}
      </div>
    </div>
  );
}

/** Generates a cryptographically-random 6-digit OTP string. */
function generateOtp(): string {
  const arr = new Uint8Array(6);
  crypto.getRandomValues(arr);
  return Array.from(arr, (b) => b % 10).join("");
}

function AuthPage() {
  const navigate = useNavigate();

  const [checkingSession, setCheckingSession] = useState(true);
  const [loading, setLoading] = useState(false);
  const [networkState, setNetworkState] = useState<NetworkState>(
    readNetworkState,
  );
  const [rememberMe, setRememberMe] = useState(getInitialRememberMe);

  const [signInEmail, setSignInEmail] = useState(getInitialRememberedEmail);
  const [signInPassword, setSignInPassword] = useState("");
  const [forgotEmail, setForgotEmail] = useState("");

  const [tab, setTab] = useState<"signin" | "forgot">("signin");
  const [feedback, setFeedback] = useState<FeedbackState | null>(null);

  // ── Forgot-password OTP flow ─────────────────────────────────────────────
  const [forgotStep, setForgotStep] = useState<"email" | "verify" | "newpassword">("email");
  const [forgotOtpInput, setForgotOtpInput] = useState("");
  const [generatedOtp, setGeneratedOtp] = useState("");
  const [otpExpiry, setOtpExpiry] = useState<Date | null>(null);
  const [pendingResetEmail, setPendingResetEmail] = useState("");

  // ── Step 3: New password fields ───────────────────────────────────────────
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showNewPw, setShowNewPw] = useState(false);
  const [showConfirmPw, setShowConfirmPw] = useState(false);



  const showLoadingModal = (title: string, message: string) => {
    setFeedback({
      open: true,
      variant: "loading",
      title,
      message,
    });
  };

  const showErrorModal = (title: string, message: string) => {
    setFeedback({
      open: true,
      variant: "error",
      title,
      message,
      buttonText: "OK",
    });
  };

  const showSuccessModal = (title: string, message: string) => {
    setFeedback({
      open: true,
      variant: "success",
      title,
      message,
      buttonText: "OK",
    });
  };

  const closeFeedbackModal = () => {
    if (feedback?.variant === "loading") return;
    setFeedback(null);
  };

  // If an authenticated Admin was redirected here because the maximum
  // number of online administrators was reached, show the message once.
  useEffect(() => {
    const limitNotice = consumeAdminAccessLimitNotice();

    if (!limitNotice) return;

    showErrorModal(
      "Admin access limit reached",
      `${limitNotice.onlineCount} administrators are currently online. Only ${limitNotice.maxAdmins} administrators can use the Admin panel at the same time.`,
    );
    // This notice is intentionally consumed only once when /auth mounts.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // If "Remember me" was previously enabled, ask the browser/password
  // manager for a saved login. The browser remains in control of access.
  useEffect(() => {
    let active = true;

    if (!rememberMe) {
      return () => {
        active = false;
      };
    }

    void (async () => {
      const savedCredential = await getCredentialFromBrowser();

      if (!active || !savedCredential) return;

      // Do not replace a different email the user has already typed.
      if (
        !signInEmail ||
        signInEmail.toLowerCase() === savedCredential.email.toLowerCase()
      ) {
        setSignInEmail(savedCredential.email);
        setSignInPassword(savedCredential.password);
      }
    })();

    return () => {
      active = false;
    };
    // Run once when the auth page mounts.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Remember only the email and checkbox preference in app storage.
  // The raw password is never stored in localStorage/sessionStorage.
  useEffect(() => {
    if (typeof window === "undefined") return;

    if (rememberMe) {
      window.localStorage.setItem(REMEMBER_ME_KEY, "true");
      window.localStorage.setItem(
        REMEMBERED_EMAIL_KEY,
        signInEmail.trim(),
      );
    } else {
      window.localStorage.removeItem(REMEMBER_ME_KEY);
      window.localStorage.removeItem(REMEMBERED_EMAIL_KEY);
    }
  }, [rememberMe, signInEmail]);

  // Keep the sign-in experience aware of the browser's current network.
  // Network Information API is optional, so navigator.onLine remains the fallback.
  useEffect(() => {
    if (typeof window === "undefined" || typeof navigator === "undefined") {
      return;
    }

    const updateNetworkState = () => {
      setNetworkState(readNetworkState());
    };

    const connection = (
      navigator as Navigator & {
        connection?: BrowserNetworkInformation;
        mozConnection?: BrowserNetworkInformation;
        webkitConnection?: BrowserNetworkInformation;
      }
    ).connection ??
      (
        navigator as Navigator & {
          mozConnection?: BrowserNetworkInformation;
        }
      ).mozConnection ??
      (
        navigator as Navigator & {
          webkitConnection?: BrowserNetworkInformation;
        }
      ).webkitConnection;

    updateNetworkState();

    window.addEventListener("online", updateNetworkState);
    window.addEventListener("offline", updateNetworkState);
    connection?.addEventListener?.("change", updateNetworkState);

    return () => {
      window.removeEventListener("online", updateNetworkState);
      window.removeEventListener("offline", updateNetworkState);
      connection?.removeEventListener?.("change", updateNetworkState);
    };
  }, []);

  // Do not fake a loading duration. The modal stays open for the real request.
  // If the network/server response is taking longer, update the message.
  useEffect(() => {
    if (!loading) return;

    const slowTimer = window.setTimeout(() => {
      setFeedback((current) => {
        if (!current?.open || current.variant !== "loading") {
          return current;
        }

        return {
          ...current,
          title: networkState.slow
            ? "Slow connection..."
            : "Still signing in...",
          message: networkState.slow
            ? "Your connection is slow. We are still securely verifying your account."
            : "The server is taking a little longer than usual. Please keep this page open.",
        };
      });
    }, 3500);

    const verySlowTimer = window.setTimeout(() => {
      setFeedback((current) => {
        if (!current?.open || current.variant !== "loading") {
          return current;
        }

        return {
          ...current,
          title: "Still working...",
          message:
            "Your sign-in request is still being processed. This can happen on a slow or unstable network.",
        };
      });
    }, 8000);

    return () => {
      window.clearTimeout(slowTimer);
      window.clearTimeout(verySlowTimer);
    };
  }, [loading, networkState.slow]);

  useEffect(() => {
    let mounted = true;

    const checkSession = async () => {
      try {
        const { data } = await supabase.auth.getSession();

        if (!mounted) return;

        if (data.session) {
          const navigationResult = await goToCorrectDashboard(navigate);

          if (navigationResult.status === "admin_limit") {
            await supabase.auth.signOut();

            if (mounted) {
              setCheckingSession(false);
              showErrorModal(
                "Admin access limit reached",
                `${navigationResult.onlineCount} administrators are currently online. Only ${navigationResult.maxAdmins} administrators can use the Admin panel at the same time.`,
              );
            }

            return;
          }

          return;
        }
      } catch {
        // If session restoration fails, keep the login page available.
      }

      if (mounted) {
        setCheckingSession(false);
      }
    };

    void checkSession();

    return () => {
      mounted = false;
    };
  }, [navigate]);

  const signIn = async () => {
    const trimmedEmail = signInEmail.trim();
    const currentNetwork = readNetworkState();

    setNetworkState(currentNetwork);

    if (!trimmedEmail || !signInPassword) {
      showErrorModal(
        "Enter email and password",
        "Please complete both fields before signing in.",
      );
      return;
    }

    if (!currentNetwork.online) {
      showErrorModal(
        "No internet connection",
        "Connect to the internet and try signing in again.",
      );
      return;
    }

    setLoading(true);
    showLoadingModal(
      "Signing in...",
      getSignInLoadingMessage(currentNetwork),
    );

    try {
      const { data, error } = await supabase.auth.signInWithPassword({
        email: trimmedEmail,
        password: signInPassword,
      });

      if (error) {
        if (error.message.toLowerCase().includes("invalid")) {
          showErrorModal(
            "Wrong email or password",
            "Please check your credentials or use Forgot password if needed.",
          );
          return;
        }

        showErrorModal("Sign in failed", error.message);
        return;
      }

      if (!data.session) {
        showErrorModal(
          "Unable to sign in",
          "Unable to create a login session.",
        );
        return;
      }

      const { data: profile, error: profileError } = await supabase
        .from("profiles")
        .select("id")
        .eq("id", data.session.user.id)
        .maybeSingle();

      if (profileError) {
        showErrorModal("Profile check failed", profileError.message);
        return;
      }

      if (!profile) {
        await supabase.auth.signOut();
        showErrorModal(
          "Account not registered",
          "This account is not registered in the school portal. Contact the administrator.",
        );
        return;
      }

      if (rememberMe) {
        // Let the browser/password manager remember the credential securely.
        // Depending on the browser, the user may see a Save/Update Password prompt.
        await saveCredentialToBrowser(trimmedEmail, signInPassword);
      }

      const navigationResult = await goToCorrectDashboard(navigate);

      if (navigationResult.status === "admin_limit") {
        await supabase.auth.signOut();

        showErrorModal(
          "Admin access limit reached",
          `${navigationResult.onlineCount} administrators are currently online. Only ${navigationResult.maxAdmins} administrators can use the Admin panel at the same time.`,
        );

        return;
      }
    } catch (error) {
      showErrorModal(
        "Unable to sign in",
        error instanceof Error
          ? error.message
          : "Something went wrong during sign in.",
      );
    } finally {
      setLoading(false);
    }
  };

  const handleSignInSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    if (loading) return;

    void signIn();
  };

  const forgot = async () => {
    const trimmedEmail = forgotEmail.trim();
    const currentNetwork = readNetworkState();

    setNetworkState(currentNetwork);

    if (!trimmedEmail) {
      showErrorModal(
        "Enter your email",
        "Please enter your registered email to continue.",
      );
      return;
    }

    if (!currentNetwork.online) {
      showErrorModal(
        "No internet connection",
        "Connect to the internet and try again.",
      );
      return;
    }

    setLoading(true);
    showLoadingModal(
      "Sending verification code...",
      currentNetwork.slow
        ? "Slow network detected. Please wait."
        : "Sending a 6-digit code to your email...",
    );

    try {
      // ── Step 1: Get role label via SECURITY DEFINER RPC ───────────────────
      // This bypasses RLS so unauthenticated users on the forgot-password
      // page can read the role without being logged in.
      const { data: roleLabel, error: rpcError } = await (
        supabase as unknown as {
          rpc: (
            functionName: string,
            args: Record<string, unknown>,
          ) => Promise<{
            data: string | null;
            error: { message: string } | null;
          }>;
        }
      ).rpc("get_role_label_for_reset", {
        p_email: trimmedEmail,
      });

      if (rpcError) {
        showErrorModal("Unable to send code", rpcError.message);
        return;
      }

      // RPC returns NULL when the email isn't registered
      if (roleLabel === null) {
        showErrorModal(
          "Email not found",
          "This email is not registered in the school portal. Please check and try again.",
        );
        return;
      }

      // Use the email as the display name (full_name is behind RLS for anon)
      // The OTP email will show the email address as the greeting fallback.
      const userName = trimmedEmail;

      // ── Step 2: Generate 6-digit OTP with 15-min expiry ───────────────────
      const otp = generateOtp();
      const expiry = new Date(Date.now() + 15 * 60 * 1000);
      const timeStr = expiry.toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
      });

      // ── Step 3: Send OTP via EmailJS ──────────────────────────────────────
      const EMAILJS_SERVICE_ID =
        import.meta.env.VITE_EMAILJS_SERVICE_ID ?? "service_li6vpin";
      const EMAILJS_TEMPLATE_ID =
        import.meta.env.VITE_EMAILJS_TEMPLATE_ID ?? "template_rwau5mc";
      const EMAILJS_PUBLIC_KEY =
        import.meta.env.VITE_EMAILJS_PUBLIC_KEY ?? "d5q8DJK9m_ChusXXo";

      await emailjs.send(
        EMAILJS_SERVICE_ID,
        EMAILJS_TEMPLATE_ID,
        {
          email: trimmedEmail,
          to_email: trimmedEmail,
          user_name: userName,
          user_role: roleLabel as string,
          passcode: otp,
          time: timeStr,
        },
        { publicKey: EMAILJS_PUBLIC_KEY },
      );

      // ── Step 4: Store OTP state and switch to verify step ─────────────────
      setGeneratedOtp(otp);
      setOtpExpiry(expiry);
      setPendingResetEmail(trimmedEmail);
      setForgotOtpInput("");
      setFeedback(null);
      setForgotStep("verify");
    } catch (error) {
      showErrorModal(
        "Unable to send code",
        error instanceof Error
          ? error.message
          : "Something went wrong while sending the verification code.",
      );
    } finally {
      setLoading(false);
    }
  };


  const verifyOtp = async () => {
    const enteredCode = forgotOtpInput.trim();
    const currentNetwork = readNetworkState();

    setNetworkState(currentNetwork);

    if (!enteredCode) {
      showErrorModal(
        "Enter the code",
        "Please enter the 6-digit code sent to your email.",
      );
      return;
    }

    if (!currentNetwork.online) {
      showErrorModal(
        "No internet connection",
        "Connect to the internet and try again.",
      );
      return;
    }

    // Check expiry first
    if (otpExpiry && new Date() > otpExpiry) {
      showErrorModal(
        "Code expired",
        "Your verification code has expired. Please request a new one.",
      );
      setForgotStep("email");
      setGeneratedOtp("");
      setOtpExpiry(null);
      setForgotOtpInput("");
      return;
    }

    // Verify code matches
    if (enteredCode !== generatedOtp) {
      showErrorModal(
        "Incorrect code",
        "The code you entered doesn't match. Please check your email and try again.",
      );
      return;
    }

    // ✅ OTP is correct — advance to Step 3 (new password form)
    setNewPassword("");
    setConfirmPassword("");
    setFeedback(null);
    setForgotStep("newpassword");
  };

  // ── Step 3: Save the new password via Edge Function ───────────────────────
  const saveNewPassword = async () => {
    if (newPassword.length < 6) {
      showErrorModal(
        "Password too short",
        "Your new password must be at least 6 characters long.",
      );
      return;
    }

    if (newPassword !== confirmPassword) {
      showErrorModal(
        "Passwords don't match",
        "Please make sure both password fields are identical.",
      );
      return;
    }

    const resetEmail = pendingResetEmail.trim().toLowerCase();

    if (!resetEmail) {
      showErrorModal(
        "Reset session expired",
        "Your reset email is missing. Please start the forgot-password process again.",
      );
      setForgotStep("email");
      return;
    }

    const currentNetwork = readNetworkState();
    setNetworkState(currentNetwork);

    if (!currentNetwork.online) {
      showErrorModal(
        "No internet connection",
        "Connect to the internet and try again.",
      );
      return;
    }

    setLoading(true);
    showLoadingModal(
      "Updating password...",
      "Please wait while we securely update your password.",
    );

    try {
      // Re-check the account before changing its Supabase Auth password.
      // Only Admin, Class Adviser, and Subject Teacher accounts are allowed.
      const { data: roleLabel, error: roleError } = await (
        supabase as unknown as {
          rpc: (
            functionName: string,
            args: Record<string, unknown>,
          ) => Promise<{
            data: string | null;
            error: { message: string } | null;
          }>;
        }
      ).rpc("get_role_label_for_reset", {
        p_email: resetEmail,
      });

      if (roleError) {
        showErrorModal("Update failed", roleError.message);
        return;
      }

      const allowedRoles = new Set([
        "Admin",
        "Class Adviser",
        "Subject Teacher",
      ]);

      if (!roleLabel || !allowedRoles.has(roleLabel)) {
        showErrorModal(
          "Update failed",
          "Only Admin, Class Adviser, and Subject Teacher accounts can reset a password here.",
        );
        return;
      }

      // Keep the current EmailJS OTP and UI exactly as they are.
      // The actual Supabase Auth password update is performed on the server.
      const { data, error: functionError } =
        await supabase.functions.invoke("reset-user-password", {
          body: {
            email: resetEmail,
            new_password: newPassword,
          },
        });

      if (functionError) {
        let message =
          functionError.message || "Unable to update the password.";

        const response = (
          functionError as unknown as {
            context?: Response;
          }
        ).context;

        if (response) {
          try {
            const payload = (await response.clone().json()) as {
              error?: string;
              message?: string;
            };

            message = payload.error ?? payload.message ?? message;
          } catch {
            // Keep the original function error when the response is not JSON.
          }
        }

        if (
          message.toLowerCase().includes("failed to send a request") ||
          message.toLowerCase().includes("failed to fetch")
        ) {
          message =
            "The Supabase reset-user-password Edge Function could not be reached. Deploy that function to the same Supabase project used by this app, then try again.";
        }

        showErrorModal("Update failed", message);
        return;
      }

      const result = (data ?? {}) as {
        success?: boolean;
        error?: string;
        email?: string;
      };

      if (result.error || result.success !== true) {
        showErrorModal(
          "Update failed",
          result.error ?? "The password was not updated. Please try again.",
        );
        return;
      }

      showSuccessModal(
        "Password updated!",
        `The password for ${resetEmail} has been changed successfully. You can now sign in with the new password.`,
      );

      setForgotEmail("");
      setForgotOtpInput("");
      setGeneratedOtp("");
      setOtpExpiry(null);
      setPendingResetEmail("");
      setNewPassword("");
      setConfirmPassword("");
      setShowNewPw(false);
      setShowConfirmPw(false);
      setForgotStep("email");
      setTab("signin");
    } catch (error) {
      showErrorModal(
        "Update failed",
        error instanceof Error
          ? error.message
          : "Something went wrong while updating the password.",
      );
    } finally {
      setLoading(false);
    }
  };


  // Pressing Enter anywhere on the auth page triggers the active action.

  // On the Sign in tab, Enter always uses the same signIn() function as the button,
  // so keyboard and mouse sign-ins both show the same loading screen and validation.
  useEffect(() => {
    const handleEnterKey = (event: KeyboardEvent) => {
      if (event.key !== "Enter" || event.repeat) return;

      // Do not allow another action while an async request is already running.
      if (loading || checkingSession) {
        event.preventDefault();
        return;
      }

      // When a feedback message is open, Enter acts like the OK button.
      if (feedback?.open) {
        event.preventDefault();

        if (feedback.variant !== "loading") {
          setFeedback(null);
        }

        return;
      }

      event.preventDefault();

      if (tab === "signin") {
        void signIn();
        return;
      }

      if (tab === "forgot") {
        if (forgotStep === "verify") {
          void verifyOtp();
        } else if (forgotStep === "newpassword") {
          void saveNewPassword();
        } else {
          void forgot();
        }
      }
    };

    window.addEventListener("keydown", handleEnterKey, true);

    return () => {
      window.removeEventListener("keydown", handleEnterKey, true);
    };
  }, [
    tab,
    forgotStep,
    loading,
    checkingSession,
    feedback,
    signInEmail,
    signInPassword,
    forgotEmail,
    forgotOtpInput,
    generatedOtp,
    otpExpiry,
    newPassword,
    confirmPassword,
    pendingResetEmail,
  ]);


  if (checkingSession) {
    return <AuthPageSkeleton />;
  }

  return (
    <div className="relative min-h-screen overflow-hidden bg-[var(--login-bg)] text-[var(--login-brown)]">
      <div className="pointer-events-none absolute left-[-140px] top-[-90px] h-[430px] w-[430px] rounded-full border border-[var(--login-gold)] opacity-45" />
      <div className="pointer-events-none absolute right-[-160px] bottom-[-130px] h-[520px] w-[520px] rounded-full border border-[var(--login-gold)] opacity-35" />
      <div className="pointer-events-none absolute right-[-40px] top-[-70px] h-[220px] w-[560px] rounded-bl-[100%] bg-[var(--login-red)]" />
      <div className="pointer-events-none absolute right-[-20px] top-[-42px] h-[170px] w-[500px] rounded-bl-[100%] border-b-2 border-[var(--login-gold)] opacity-70" />

      <div className="pointer-events-none absolute left-[70px] top-[95px] grid grid-cols-7 gap-4 opacity-30">
        {Array.from({ length: 49 }).map((_, index) => (
          <span
            key={index}
            className="h-1.5 w-1.5 rounded-full bg-[var(--login-dot)]"
          />
        ))}
      </div>

      <div className="pointer-events-none absolute right-[280px] top-[140px] h-[270px] w-[270px] rounded-full bg-[var(--login-soft-circle)]" />
      <div className="pointer-events-none absolute bottom-[170px] left-[140px] h-[420px] w-[420px] rounded-full bg-[var(--login-soft-circle)]" />

      <main className="relative z-10 flex min-h-screen flex-col items-center justify-center px-6 py-10">
        <div className="mb-8 flex flex-col items-center text-center">
          <img
            src={logo}
            alt="Agusan del Sur National Science High School"
            className="mb-4 h-24 w-24 object-contain"
          />

          <h1 className="font-serif text-5xl font-black tracking-[0.03em] text-[var(--login-red)]">
            AGUSAN DEL SUR
          </h1>

          <h2 className="mt-1 text-2xl font-semibold tracking-[0.34em] text-[var(--login-red)]">
            NATIONAL SCIENCE HIGH SCHOOL
          </h2>

          <div className="mt-5 flex flex-col items-center text-center">
            <h3 className="font-serif text-4xl font-black tracking-[0.18em] text-[var(--login-red)]">
              SIGLA
            </h3>

            <div className="mt-3 flex items-center justify-center gap-4 text-xl text-[var(--login-muted)]">
              <span className="h-px w-32 bg-[var(--login-gold)]" />
              <span>School ID No. 304704</span>
              <span className="h-px w-32 bg-[var(--login-gold)]" />
            </div>
          </div>
        </div>

        <div className="w-full max-w-[720px] overflow-hidden rounded-[28px] border border-[var(--login-border)] bg-white/90 shadow-[0_25px_70px_rgba(124,74,47,0.18)] backdrop-blur">
          <div className="border-b border-[var(--login-border)] bg-white/70 px-8 py-6 text-center">
            <h2 className="text-3xl font-black text-[var(--login-red)]">
              {tab === "forgot"
                ? forgotStep === "verify"
                  ? "Enter your code"
                  : forgotStep === "newpassword"
                    ? "Set new password"
                    : "Reset your password"
                : "Sign in"}
            </h2>

            <p className="mt-1 text-sm text-[var(--login-muted)]">
              {tab === "forgot"
                ? forgotStep === "verify"
                  ? `A 6-digit code was sent to ${pendingResetEmail}.`
                  : forgotStep === "newpassword"
                    ? "Choose a strong password for your school portal account."
                    : "Enter your registered email to receive a 6-digit code."
                : ""}
            </p>
          </div>

          <Tabs
            value={tab}
            onValueChange={(value) => {
              setTab(value as "signin" | "forgot");
              // Reset OTP flow when switching away from forgot tab
              if (value !== "forgot") {
                setForgotStep("email");
                setForgotOtpInput("");
                setGeneratedOtp("");
                setOtpExpiry(null);
                setPendingResetEmail("");
                setNewPassword("");
                setConfirmPassword("");
                setShowNewPw(false);
                setShowConfirmPw(false);
              }
            }}
          >
            <div className="px-11 py-9">
              <TabsContent value="signin" className="mt-0">
                <form
                  onSubmit={handleSignInSubmit}
                  className="space-y-6"
                  autoComplete="on"
                >
                  <AuthField
                    label="Email"
                    type="email"
                    name="username"
                    autoComplete="username"
                    value={signInEmail}
                    onChange={(event) => setSignInEmail(event.target.value)}
                    placeholder=""
                    icon={Mail}
                  />

                  <AuthField
                    label="Password"
                    type="password"
                    name="password"
                    autoComplete="current-password"
                    value={signInPassword}
                    onChange={(event) => setSignInPassword(event.target.value)}
                    placeholder=""
                    icon={Lock}
                  />

                  <div className="flex items-center justify-between">
                    <label className="flex cursor-pointer items-center gap-3 text-base text-[var(--login-brown)]">
                      <Checkbox
                        checked={rememberMe}
                        onCheckedChange={(value) =>
                          setRememberMe(value === true)
                        }
                        className="h-5 w-5 border-2 border-[var(--login-red)]"
                      />
                      Remember me
                    </label>

                    <button
                      type="button"
                      onClick={() => {
                        setForgotEmail(signInEmail);
                        setForgotStep("email");
                        setTab("forgot");
                      }}
                      className="text-base font-semibold text-[var(--login-red)] hover:underline"
                    >
                      Forgot password?
                    </button>
                  </div>

                  <Button
                    type="submit"
                    className="h-[62px] w-full rounded-xl bg-[var(--login-red)] text-xl font-black text-white shadow-lg shadow-red-900/20 hover:bg-[var(--login-red-dark)]"
                    disabled={loading}
                  >
                    <Lock className="mr-2 h-5 w-5" />
                    {loading ? "Please wait..." : "Sign in"}
                  </Button>
                </form>
              </TabsContent>

              <TabsContent value="forgot" className="mt-0 space-y-6">
                {forgotStep === "email" ? (
                  <>
                    {/* ─────────────────────────────────────────────
                        STEP 1: ENTER REGISTERED EMAIL
                    ───────────────────────────────────────────── */}

                    <AuthField
                      label="Email"
                      type="email"
                      name="forgot-email"
                      autoComplete="email"
                      value={forgotEmail}
                      onChange={(event) => setForgotEmail(event.target.value)}
                      placeholder=""
                      icon={Mail}
                    />

                    <Button
                      type="button"
                      className="h-[62px] w-full rounded-xl bg-[var(--login-red)] text-xl font-black text-white shadow-lg shadow-red-900/20 hover:bg-[var(--login-red-dark)]"
                      onClick={() => void forgot()}
                      disabled={loading}
                    >
                      {loading ? "Please wait..." : "Send verification code"}
                    </Button>

                    <button
                      type="button"
                      onClick={() => {
                        setForgotStep("email");
                        setForgotOtpInput("");
                        setGeneratedOtp("");
                        setOtpExpiry(null);
                        setPendingResetEmail("");
                        setNewPassword("");
                        setConfirmPassword("");
                        setShowNewPw(false);
                        setShowConfirmPw(false);
                        setTab("signin");
                      }}
                      className="w-full text-center text-base font-semibold text-[var(--login-red)] hover:underline"
                    >
                      Back to sign in
                    </button>
                  </>
                ) : forgotStep === "verify" ? (
                  <>
                    {/* ─────────────────────────────────────────────
                        STEP 2: ENTER THE 6-DIGIT OTP
                    ───────────────────────────────────────────── */}

                    <div className="space-y-2">
                      <Label className="text-[15px] font-medium text-[var(--login-brown)]">
                        Verification Code
                      </Label>

                      <div className="relative">
                        <Shield className="absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-[#9b7c67]" />

                        <Input
                          type="text"
                          inputMode="numeric"
                          pattern="[0-9]*"
                          maxLength={6}
                          value={forgotOtpInput}
                          onChange={(event) =>
                            setForgotOtpInput(
                              event.target.value.replace(/\D/g, "").slice(0, 6),
                            )
                          }
                          placeholder="000000"
                          className="h-12 rounded-[14px] border border-[var(--login-border)] bg-white/70 pl-11 pr-4 text-center text-[22px] font-bold tracking-[0.4em] text-[var(--login-brown)] shadow-sm outline-none focus-visible:border-[var(--login-border)] focus-visible:ring-0"
                          autoFocus
                        />
                      </div>

                      <p className="text-xs text-[var(--login-muted)]">
                        Enter the 6-digit code from your email. Valid for 15
                        minutes.
                      </p>
                    </div>

                    <Button
                      type="button"
                      className="h-[62px] w-full rounded-xl bg-[var(--login-red)] text-xl font-black text-white shadow-lg shadow-red-900/20 hover:bg-[var(--login-red-dark)]"
                      onClick={() => void verifyOtp()}
                      disabled={loading || forgotOtpInput.length < 6}
                    >
                      {loading ? "Please wait..." : "Verify & Continue"}
                    </Button>

                    <div className="flex items-center justify-between">
                      <button
                        type="button"
                        onClick={() => {
                          setForgotStep("email");
                          setForgotOtpInput("");
                          setGeneratedOtp("");
                          setOtpExpiry(null);
                        }}
                        className="text-base font-semibold text-[var(--login-red)] hover:underline"
                      >
                        ← Back
                      </button>

                      <button
                        type="button"
                        onClick={() => void forgot()}
                        disabled={loading}
                        className="text-base font-semibold text-[var(--login-red)] hover:underline disabled:opacity-40"
                      >
                        Resend code
                      </button>
                    </div>
                  </>
                ) : (
                  <>
                    {/* ─────────────────────────────────────────────
                        STEP 3: CREATE AND SAVE THE NEW PASSWORD
                    ───────────────────────────────────────────── */}

                    <div className="space-y-2">
                      <Label className="text-[15px] font-medium text-[var(--login-brown)]">
                        New Password
                      </Label>

                      <div className="relative">
                        <Lock className="absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-[#9b7c67]" />

                        <Input
                          type={showNewPw ? "text" : "password"}
                          name="new-password"
                          autoComplete="new-password"
                          value={newPassword}
                          onChange={(event) =>
                            setNewPassword(event.target.value)
                          }
                          className="h-12 rounded-[14px] border border-[var(--login-border)] bg-white/70 pl-11 pr-12 text-[15px] text-[var(--login-brown)] shadow-sm outline-none focus-visible:border-[var(--login-border)] focus-visible:ring-0"
                          autoFocus
                        />

                        <button
                          type="button"
                          onClick={() => setShowNewPw((current) => !current)}
                          className="absolute right-4 top-1/2 -translate-y-1/2 text-[#9b7c67] transition hover:text-[var(--login-red)]"
                          aria-label={
                            showNewPw
                              ? "Hide new password"
                              : "Show new password"
                          }
                        >
                          {showNewPw ? (
                            <EyeOff className="h-5 w-5" />
                          ) : (
                            <Eye className="h-5 w-5" />
                          )}
                        </button>
                      </div>
                    </div>

                    <div className="space-y-2">
                      <Label className="text-[15px] font-medium text-[var(--login-brown)]">
                        Confirm New Password
                      </Label>

                      <div className="relative">
                        <Lock className="absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-[#9b7c67]" />

                        <Input
                          type={showConfirmPw ? "text" : "password"}
                          name="confirm-new-password"
                          autoComplete="new-password"
                          value={confirmPassword}
                          onChange={(event) =>
                            setConfirmPassword(event.target.value)
                          }
                          className="h-12 rounded-[14px] border border-[var(--login-border)] bg-white/70 pl-11 pr-12 text-[15px] text-[var(--login-brown)] shadow-sm outline-none focus-visible:border-[var(--login-border)] focus-visible:ring-0"
                        />

                        <button
                          type="button"
                          onClick={() =>
                            setShowConfirmPw((current) => !current)
                          }
                          className="absolute right-4 top-1/2 -translate-y-1/2 text-[#9b7c67] transition hover:text-[var(--login-red)]"
                          aria-label={
                            showConfirmPw
                              ? "Hide confirmed password"
                              : "Show confirmed password"
                          }
                        >
                          {showConfirmPw ? (
                            <EyeOff className="h-5 w-5" />
                          ) : (
                            <Eye className="h-5 w-5" />
                          )}
                        </button>
                      </div>
                    </div>

                    <div className="rounded-[14px] border border-[var(--login-border)] bg-[#fffaf5] px-4 py-3">
                      <div className="flex items-start gap-3">
                        <Shield className="mt-0.5 h-5 w-5 shrink-0 text-[var(--login-red)]" />

                        <div>
                          <p className="text-sm font-semibold text-[var(--login-brown)]">
                            Password requirements
                          </p>

                          <p className="mt-1 text-xs leading-5 text-[var(--login-muted)]">
                            Use at least 6 characters and make sure both
                            password fields match.
                          </p>
                        </div>
                      </div>
                    </div>

                    {confirmPassword.length > 0 &&
                      newPassword !== confirmPassword && (
                        <div className="flex items-center gap-2 rounded-[12px] border border-red-200 bg-red-50 px-4 py-3 text-sm font-medium text-red-700">
                          <AlertCircle className="h-4 w-4 shrink-0" />
                          Passwords do not match yet.
                        </div>
                      )}

                    {confirmPassword.length > 0 &&
                      newPassword.length >= 6 &&
                      newPassword === confirmPassword && (
                        <div className="flex items-center gap-2 rounded-[12px] border border-green-200 bg-green-50 px-4 py-3 text-sm font-medium text-green-700">
                          <CheckCircle2 className="h-4 w-4 shrink-0" />
                          Passwords match.
                        </div>
                      )}

                    <Button
                      type="button"
                      className="h-[62px] w-full rounded-xl bg-[var(--login-red)] text-xl font-black text-white shadow-lg shadow-red-900/20 hover:bg-[var(--login-red-dark)]"
                      onClick={() => void saveNewPassword()}
                      disabled={
                        loading ||
                        newPassword.length < 6 ||
                        confirmPassword.length < 6
                      }
                    >
                      <Lock className="mr-2 h-5 w-5" />
                      {loading
                        ? "Updating password..."
                        : "Save New Password"}
                    </Button>

                    <button
                      type="button"
                      onClick={() => {
                        setForgotStep("verify");
                        setNewPassword("");
                        setConfirmPassword("");
                        setShowNewPw(false);
                        setShowConfirmPw(false);
                      }}
                      className="w-full text-center text-base font-semibold text-[var(--login-red)] hover:underline"
                    >
                      ← Back to verification
                    </button>
                  </>
                )}
              </TabsContent>
            </div>
          </Tabs>

        </div>

        <footer className="mt-7 flex flex-col items-center text-center">
          <div className="mb-3 grid h-11 w-11 place-items-center rounded-full bg-[var(--login-icon-bg)] text-[var(--login-red)]">
            <Shield className="h-6 w-6" />
          </div>

          <p className="text-base text-[var(--login-brown)]">
            © 2026 Agusan del Sur National Science High School
          </p>

          <div className="mt-2 flex items-center gap-4 text-base font-semibold text-[var(--login-red)]">
            <span>Privacy Policy</span>
            <span>•</span>
            <span>Support</span>
          </div>
        </footer>
      </main>

      <FeedbackModal feedback={feedback} onClose={closeFeedbackModal} />
    </div>
  );
}
