type AppErrorOptions = {
  mechanism?: "manual" | "onerror" | "unhandledrejection" | "react_error_boundary";
  handled?: boolean;
  severity?: "error" | "warning" | "info";
};

export function reportAppError(
  error: unknown,
  context: Record<string, unknown> = {},
  _options?: AppErrorOptions,
) {
  // Generic application error reporting.
  // Keeps errors visible in development without relying on Lovable.
  if (typeof console !== "undefined") {
    console.error("Application error:", error, context);
  }
}
