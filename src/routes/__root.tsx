import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import {
  Outlet,
  Link,
  createRootRouteWithContext,
  useRouter,
  HeadContent,
  Scripts,
} from "@tanstack/react-router";
import { useEffect, type ReactNode } from "react";

import appCss from "../styles.css?url";
import { reportAppError } from "@/lib/error-reporting";
import { supabase } from "@/integrations/supabase/client";
import { Toaster } from "@/components/ui/sonner";

function NotFoundComponent() {
  return (
    <div className="flex min-h-screen items-center justify-center bg-background px-4">
      <div className="max-w-md text-center">
        <h1 className="text-7xl font-bold text-foreground">404</h1>

        <h2 className="mt-4 text-xl font-semibold">
          Page not found
        </h2>

        <p className="mt-2 text-sm text-muted-foreground">
          The page you're looking for doesn't exist.
        </p>

        <Link
          to="/"
          className="mt-6 inline-flex rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:opacity-90"
        >
          Go home
        </Link>
      </div>
    </div>
  );
}

function ErrorComponent({
  error,
  reset,
}: {
  error: Error;
  reset: () => void;
}) {
  const router = useRouter();

  useEffect(() => {
    reportAppError(error, {
      boundary: "root",
    });
  }, [error]);

  return (
    <div className="flex min-h-screen items-center justify-center bg-background px-4">
      <div className="max-w-md text-center">
        <h1 className="text-xl font-semibold">
          This page didn't load
        </h1>

        <p className="mt-2 text-sm text-muted-foreground">
          {error.message}
        </p>

        <div className="mt-6 flex justify-center gap-2">
          <button
            onClick={() => {
              router.invalidate();
              reset();
            }}
            className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground"
          >
            Try again
          </button>

          <a
            href="/"
            className="rounded-md border px-4 py-2 text-sm"
          >
            Go home
          </a>
        </div>
      </div>
    </div>
  );
}

export const Route =
  createRootRouteWithContext<{
    queryClient: QueryClient;
  }>()({
    head: () => ({
      meta: [
        {
          charSet: "utf-8",
        },
        {
          name: "viewport",
          content: "width=device-width, initial-scale=1",
        },
        {
          title:
            "Agusan del Sur National Science High School — Class Record",
        },
        {
          name: "description",
          content:
            "School Learner Attendance & Grade Monitoring for Agusan del Sur National Science High School.",
        },
        {
          property: "og:title",
          content:
            "Agusan del Sur National Science High School — Class Record",
        },
        {
          property: "og:description",
          content:
            "School Learner Attendance & Grade Monitoring for Agusan del Sur National Science High School.",
        },
        {
          property: "og:type",
          content: "website",
        },
        {
          name: "twitter:card",
          content: "summary",
        },
        {
          name: "twitter:title",
          content:
            "Agusan del Sur National Science High School — Class Record",
        },
        {
          name: "twitter:description",
          content:
            "School Learner Attendance & Grade Monitoring for Agusan del Sur National Science High School.",
        },
      ],

      links: [
        {
          rel: "stylesheet",
          href: appCss,
        },
        {
          rel: "icon",
          href: "/favicon.png",
          type: "image/png",
        },
      ],
    }),

    shellComponent: RootShell,
    component: RootComponent,
    notFoundComponent: NotFoundComponent,
    errorComponent: ErrorComponent,
  });

function RootShell({
  children,
}: {
  children: ReactNode;
}) {
  return (
    <html lang="en">
      <head>
        <HeadContent />
      </head>

      <body>
        {children}
        <Scripts />
      </body>
    </html>
  );
}

function RootComponent() {
  const { queryClient } = Route.useRouteContext();
  const router = useRouter();

  useEffect(() => {
    const { data: sub } =
      supabase.auth.onAuthStateChange((event) => {
        if (
          event !== "SIGNED_IN" &&
          event !== "SIGNED_OUT" &&
          event !== "USER_UPDATED"
        ) {
          return;
        }

        router.invalidate();

        if (event !== "SIGNED_OUT") {
          queryClient.invalidateQueries();
        }
      });

    return () => {
      sub.subscription.unsubscribe();
    };
  }, [router, queryClient]);

  return (
    <QueryClientProvider client={queryClient}>
      <Outlet />

      <Toaster
        richColors
        position="top-right"
      />
    </QueryClientProvider>
  );
}
