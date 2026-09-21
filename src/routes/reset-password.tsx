import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useEffect, useState, type ChangeEvent } from "react";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import logo from "@/assets/ASNSHS Logo.png";
import {
  Lock,
  Eye,
  EyeOff,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Shield,
} from "lucide-react";

export const Route = createFileRoute("/reset-password")({
  component: ResetPassword,
});

// ── helpers ──────────────────────────────────────────────────────────────────

type Strength = "weak" | "fair" | "strong";

function getStrength(pw: string): Strength {
  if (pw.length < 6) return "weak";
  const hasUpper = /[A-Z]/.test(pw);
  const hasLower = /[a-z]/.test(pw);
  const hasNumber = /[0-9]/.test(pw);
  const hasSpecial = /[^A-Za-z0-9]/.test(pw);
  const score = [hasUpper, hasLower, hasNumber, hasSpecial].filter(Boolean).length;
  if (pw.length >= 10 && score >= 3) return "strong";
  if (pw.length >= 6 && score >= 2) return "fair";
  return "weak";
}

const strengthConfig: Record<
  Strength,
  { label: string; color: string; bars: number }
> = {
  weak:   { label: "Weak",   color: "#dc2626", bars: 1 },
  fair:   { label: "Fair",   color: "#d97706", bars: 2 },
  strong: { label: "Strong", color: "#16a34a", bars: 3 },
};

// ── types ─────────────────────────────────────────────────────────────────────

type FeedbackState = {
  open: boolean;
  variant: "loading" | "error" | "success";
  title: string;
  message: string;
};

// ── FeedbackModal (same style as auth page) ───────────────────────────────────

function FeedbackModal({
  feedback,
  onClose,
}: {
  feedback: FeedbackState | null;
  onClose: () => void;
}) {
  if (!feedback?.open) return null;
  const isLoading = feedback.variant === "loading";
  const isError   = feedback.variant === "error";

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
            OK
          </Button>
        )}
      </div>
    </div>
  );
}

// ── PasswordField ─────────────────────────────────────────────────────────────

function PasswordField({
  label,
  value,
  onChange,
  placeholder,
  name,
  autoComplete,
}: {
  label: string;
  value: string;
  onChange: (e: ChangeEvent<HTMLInputElement>) => void;
  placeholder?: string;
  name?: string;
  autoComplete?: string;
}) {
  const [show, setShow] = useState(false);

  return (
    <div className="space-y-2">
      <Label className="text-[15px] font-medium text-[var(--login-brown)]">
        {label}
      </Label>
      <div className="relative">
        <Lock className="absolute left-4 top-1/2 h-5 w-5 -translate-y-1/2 text-[#9b7c67]" />
        <Input
          type={show ? "text" : "password"}
          name={name}
          autoComplete={autoComplete}
          value={value}
          onChange={onChange}
          placeholder={placeholder}
          className="h-12 rounded-[14px] border border-[var(--login-border)] bg-white/70 pl-11 pr-12 text-[15px] text-[var(--login-brown)] shadow-sm outline-none focus-visible:border-[var(--login-border)] focus-visible:ring-0"
        />
        <button
          type="button"
          tabIndex={-1}
          onClick={() => setShow((s) => !s)}
          className="absolute right-4 top-1/2 -translate-y-1/2 text-[#9b7c67] hover:text-[var(--login-brown)]"
          aria-label={show ? "Hide password" : "Show password"}
        >
          {show ? (
            <EyeOff className="h-5 w-5" />
          ) : (
            <Eye className="h-5 w-5" />
          )}
        </button>
      </div>
    </div>
  );
}

// ── Main component ────────────────────────────────────────────────────────────

function ResetPassword() {
  const navigate = useNavigate();

  const [ready, setReady]             = useState(false);
  const [checking, setChecking]       = useState(true);
  const [password, setPassword]       = useState("");
  const [confirm, setConfirm]         = useState("");
  const [loading, setLoading]         = useState(false);
  const [feedback, setFeedback]       = useState<FeedbackState | null>(null);

  // ── Wait for Supabase to exchange the recovery token ─────────────────────
  useEffect(() => {
    // Supabase places the recovery token in the URL hash and sets up a
    // PASSWORD_RECOVERY session automatically when the page loads.
    const { data: sub } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY" || event === "SIGNED_IN") {
        setReady(true);
        setChecking(false);
      }
    });

    supabase.auth.getSession().then(({ data }) => {
      if (data.session) {
        setReady(true);
        setChecking(false);
      } else {
        // Give the hash-based token a moment to be exchanged
        setTimeout(() => setChecking(false), 2500);
      }
    });

    return () => sub.subscription.unsubscribe();
  }, []);

  const strength = password ? getStrength(password) : null;
  const cfg      = strength ? strengthConfig[strength] : null;

  const passwordsMatch = confirm.length > 0 && password === confirm;
  const passwordsDiff  = confirm.length > 0 && password !== confirm;

  // ── Submit ────────────────────────────────────────────────────────────────
  const handleUpdate = async () => {
    if (password.length < 6) {
      setFeedback({
        open: true,
        variant: "error",
        title: "Password too short",
        message: "Your password must be at least 6 characters long.",
      });
      return;
    }

    if (password !== confirm) {
      setFeedback({
        open: true,
        variant: "error",
        title: "Passwords don't match",
        message: "Please make sure both password fields are identical.",
      });
      return;
    }

    setLoading(true);
    setFeedback({
      open: true,
      variant: "loading",
      title: "Updating password...",
      message: "Please wait while we securely update your password.",
    });

    const { error } = await supabase.auth.updateUser({ password });
    setLoading(false);

    if (error) {
      setFeedback({
        open: true,
        variant: "error",
        title: "Update failed",
        message: error.message,
      });
      return;
    }

    setFeedback({
      open: true,
      variant: "success",
      title: "Password updated!",
      message:
        "Your password has been changed successfully. You are now signed in.",
    });
  };

  const closeFeedback = () => {
    if (feedback?.variant === "loading") return;

    // After success, navigate to dashboard
    if (feedback?.variant === "success") {
      setFeedback(null);
      void navigate({ to: "/dashboard" });
      return;
    }

    setFeedback(null);
  };

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div className="relative min-h-screen overflow-hidden bg-[var(--login-bg)] text-[var(--login-brown)]">
      {/* Decorative background blobs */}
      <div className="pointer-events-none absolute left-[-140px] top-[-90px] h-[430px] w-[430px] rounded-full border border-[var(--login-gold)] opacity-45" />
      <div className="pointer-events-none absolute right-[-160px] bottom-[-130px] h-[520px] w-[520px] rounded-full border border-[var(--login-gold)] opacity-35" />
      <div className="pointer-events-none absolute right-[-40px] top-[-70px] h-[220px] w-[560px] rounded-bl-[100%] bg-[var(--login-red)]" />
      <div className="pointer-events-none absolute right-[-20px] top-[-42px] h-[170px] w-[500px] rounded-bl-[100%] border-b-2 border-[var(--login-gold)] opacity-70" />
      <div className="pointer-events-none absolute right-[280px] top-[140px] h-[270px] w-[270px] rounded-full bg-[var(--login-soft-circle)]" />
      <div className="pointer-events-none absolute bottom-[170px] left-[140px] h-[420px] w-[420px] rounded-full bg-[var(--login-soft-circle)]" />

      <main className="relative z-10 flex min-h-screen flex-col items-center justify-center px-6 py-10">
        {/* School branding */}
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

        {/* Card */}
        <div className="w-full max-w-[520px] overflow-hidden rounded-[28px] border border-[var(--login-border)] bg-white/90 shadow-[0_25px_70px_rgba(124,74,47,0.18)] backdrop-blur">
          {/* Card header */}
          <div className="border-b border-[var(--login-border)] bg-white/70 px-8 py-6 text-center">
            <h2 className="text-3xl font-black text-[var(--login-red)]">
              Set new password
            </h2>
            <p className="mt-1 text-sm text-[var(--login-muted)]">
              Choose a strong password for your school portal account.
            </p>
          </div>

          <div className="px-11 py-9 space-y-6">
            {checking ? (
              /* Verifying the reset token */
              <div className="flex flex-col items-center gap-4 py-6">
                <Loader2 className="h-10 w-10 animate-spin text-[var(--login-red)]" />
                <p className="text-sm text-[var(--login-muted)]">
                  Verifying your reset link…
                </p>
              </div>
            ) : !ready ? (
              /* Invalid / expired link */
              <div className="flex flex-col items-center gap-4 py-6 text-center">
                <div className="grid h-16 w-16 place-items-center rounded-full border-4 border-[#f5d7d6] bg-[#fff8f8]">
                  <AlertCircle className="h-8 w-8 text-[var(--login-red)]" />
                </div>
                <div>
                  <p className="text-lg font-bold text-[var(--login-red)]">
                    Link invalid or expired
                  </p>
                  <p className="mt-1 text-sm text-[var(--login-muted)]">
                    This reset link is no longer valid. Please request a new one.
                  </p>
                </div>
                <Button
                  type="button"
                  onClick={() => void navigate({ to: "/auth" })}
                  className="h-12 rounded-xl bg-[var(--login-red)] px-8 font-bold text-white hover:bg-[var(--login-red-dark)]"
                >
                  Back to sign in
                </Button>
              </div>
            ) : (
              /* New password form */
              <>
                <PasswordField
                  label="New password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Enter new password"
                  name="new-password"
                  autoComplete="new-password"
                />

                {/* Password strength bar */}
                {password.length > 0 && cfg && (
                  <div className="space-y-1 -mt-3">
                    <div className="flex gap-1.5">
                      {[1, 2, 3].map((bar) => (
                        <div
                          key={bar}
                          className="h-1.5 flex-1 rounded-full transition-colors duration-300"
                          style={{
                            backgroundColor:
                              bar <= cfg.bars ? cfg.color : "#e5e7eb",
                          }}
                        />
                      ))}
                    </div>
                    <p
                      className="text-xs font-semibold"
                      style={{ color: cfg.color }}
                    >
                      {cfg.label} password
                    </p>
                  </div>
                )}

                <PasswordField
                  label="Confirm new password"
                  value={confirm}
                  onChange={(e) => setConfirm(e.target.value)}
                  placeholder="Re-enter new password"
                  name="confirm-password"
                  autoComplete="new-password"
                />

                {/* Match indicator */}
                {confirm.length > 0 && (
                  <p
                    className={`-mt-3 flex items-center gap-1.5 text-xs font-semibold ${
                      passwordsMatch ? "text-green-600" : "text-[var(--login-red)]"
                    }`}
                  >
                    {passwordsMatch ? (
                      <CheckCircle2 className="h-4 w-4" />
                    ) : (
                      <AlertCircle className="h-4 w-4" />
                    )}
                    {passwordsMatch
                      ? "Passwords match"
                      : "Passwords do not match"}
                  </p>
                )}

                {/* Requirements hint */}
                <ul className="space-y-1 text-xs text-[var(--login-muted)]">
                  <li className={`flex items-center gap-1.5 ${password.length >= 6 ? "text-green-600" : ""}`}>
                    <CheckCircle2 className="h-3.5 w-3.5 shrink-0" />
                    At least 6 characters
                  </li>
                  <li className={`flex items-center gap-1.5 ${/[A-Z]/.test(password) ? "text-green-600" : ""}`}>
                    <CheckCircle2 className="h-3.5 w-3.5 shrink-0" />
                    At least one uppercase letter
                  </li>
                  <li className={`flex items-center gap-1.5 ${/[0-9]/.test(password) ? "text-green-600" : ""}`}>
                    <CheckCircle2 className="h-3.5 w-3.5 shrink-0" />
                    At least one number
                  </li>
                </ul>

                <Button
                  type="button"
                  onClick={() => void handleUpdate()}
                  disabled={
                    loading ||
                    password.length < 6 ||
                    passwordsDiff
                  }
                  className="h-[62px] w-full rounded-xl bg-[var(--login-red)] text-xl font-black text-white shadow-lg shadow-red-900/20 hover:bg-[var(--login-red-dark)] disabled:opacity-60"
                >
                  <Lock className="mr-2 h-5 w-5" />
                  {loading ? "Updating…" : "Update password"}
                </Button>

                <button
                  type="button"
                  onClick={() => void navigate({ to: "/auth" })}
                  className="w-full text-center text-base font-semibold text-[var(--login-red)] hover:underline"
                >
                  Back to sign in
                </button>
              </>
            )}
          </div>
        </div>

        {/* Footer */}
        <footer className="mt-7 flex flex-col items-center text-center">
          <div className="mb-3 grid h-11 w-11 place-items-center rounded-full bg-[var(--login-icon-bg)] text-[var(--login-red)]">
            <Shield className="h-6 w-6" />
          </div>
          <p className="text-base text-[var(--login-brown)]">
            © 2026 Agusan del Sur National Science High School
          </p>
        </footer>
      </main>

      <FeedbackModal feedback={feedback} onClose={closeFeedback} />
    </div>
  );
}
