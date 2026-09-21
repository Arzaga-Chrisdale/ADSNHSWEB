import { Link, useNavigate, useRouterState } from "@tanstack/react-router";
import {
  ArrowLeftRight,
  ChevronDown,
  ClipboardList,
  ClipboardPenLine,
  FileText,
  GraduationCap,
  Inbox,
  LayoutDashboard,
  LogOut,
  Menu,
  UserRound,
  X,
} from "lucide-react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { useState, type ComponentType, type ReactNode } from "react";
import { supabase } from "@/integrations/supabase/client";
import { Button } from "@/components/ui/button";
import { NotificationBell } from "@/components/NotificationBell";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import logo from "@/assets/ASNSHS Logo.png";

type TeacherType = "class_adviser" | "subject_teacher";

type NavItem = {
  to: string;
  label: string;
  icon: ComponentType<{ className?: string }>;
  tone: string;
};

type CurrentUserProfile = {
  id: string;
  full_name: string;
  email: string;
  teacher_type: TeacherType;
  avatar_url: string | null;
};

const classAdviserNav: NavItem[] = [
  {
    to: "/dashboard",
    label: "Dashboard",
    icon: LayoutDashboard,
    tone: "bg-rose-50 text-rose-700 border-rose-200",
  },
  {
    to: "/summary-of-grades",
    label: "Summary of Grades",
    icon: GraduationCap,
    tone: "bg-emerald-50 text-emerald-700 border-emerald-200",
  },
  {
    to: "/school-forms",
    label: "School Forms",
    icon: FileText,
    tone: "bg-violet-50 text-violet-700 border-violet-200",
  },
  {
    to: "/request-form",
    label: "Request Form",
    icon: ClipboardPenLine,
    tone: "bg-orange-50 text-orange-700 border-orange-200",
  },
  {
    to: "/received-requests",
    label: "Receive Form",
    icon: Inbox,
    tone: "bg-orange-50 text-orange-700 border-orange-200",
  },
  {
    to: "/credential-requests",
    label: "Credential Request",
    icon: ClipboardList,
    tone: "bg-orange-50 text-orange-700 border-orange-200",
  },
  {
    to: "/transfers",
    label: "Transfer In / Out",
    icon: ArrowLeftRight,
    tone: "bg-orange-50 text-orange-700 border-orange-200",
  },
];

const subjectTeacherNav: NavItem[] = [
  {
    to: "/subject-teacher-dashboard",
    label: "Dashboard",
    icon: LayoutDashboard,
    tone: "bg-rose-50 text-rose-700 border-rose-200",
  },
  {
    to: "/summary-of-grades",
    label: "Summary of Grades",
    icon: GraduationCap,
    tone: "bg-emerald-50 text-emerald-700 border-emerald-200",
  },
  {
    to: "/request-form",
    label: "Request Form",
    icon: ClipboardPenLine,
    tone: "bg-orange-50 text-orange-700 border-orange-200",
  },
  {
    to: "/received-requests",
    label: "Receive Form",
    icon: Inbox,
    tone: "bg-orange-50 text-orange-700 border-orange-200",
  },
];

function isTeacherType(value: unknown): value is TeacherType {
  return value === "class_adviser" || value === "subject_teacher";
}

function roleLabel(value: TeacherType) {
  return value === "subject_teacher" ? "Subject Teacher" : "Class Adviser";
}

function displayNameFromEmail(email: string) {
  const name = email.split("@")[0] || "Teacher";
  return name
    .split(/[._-]/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function isNavigationItemActive(pathname: string, destination: string) {
  return (
    pathname === destination ||
    (destination !== "/dashboard" &&
      destination !== "/subject-teacher-dashboard" &&
      pathname.startsWith(destination))
  );
}

export function AppShell({ children }: { children: ReactNode }) {
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const pathname = useRouterState({
    select: (state) => state.location.pathname,
  });
  const [mobileNavigationOpen, setMobileNavigationOpen] = useState(false);

  const isAdminPage = pathname.startsWith("/admin");

  const { data: currentUserProfile, isLoading: profileLoading } = useQuery<CurrentUserProfile>({
    queryKey: ["current-user-shell-profile"],
    enabled: !isAdminPage,
    queryFn: async () => {
      const {
        data: { user },
        error: userError,
      } = await supabase.auth.getUser();

      if (userError) throw userError;

      if (!user) {
        return {
          id: "",
          full_name: "Teacher",
          email: "",
          teacher_type: "class_adviser",
          avatar_url: null,
        };
      }

      const metadataTeacherType = user.user_metadata?.teacher_type;
      const metadataFullName =
        user.user_metadata?.full_name ||
        user.user_metadata?.name ||
        user.user_metadata?.display_name;

      // The generated Supabase type is older than the deployed profile schema.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: profile, error: profileError } = await (supabase as any)
        .from("profiles")
        .select("full_name, email, teacher_type, avatar_url")
        .eq("id", user.id)
        .maybeSingle();

      const profileTeacherType = profile?.teacher_type as unknown;
      const resolvedTeacherType =
        !profileError && isTeacherType(profileTeacherType)
          ? profileTeacherType
          : isTeacherType(metadataTeacherType)
            ? metadataTeacherType
            : "class_adviser";

      const email = profile?.email || user.email || "";
      const fullName =
        profile?.full_name || metadataFullName || (email ? displayNameFromEmail(email) : "Teacher");

      return {
        id: user.id,
        full_name: fullName,
        email,
        teacher_type: resolvedTeacherType,
        avatar_url: profile?.avatar_url || null,
      };
    },
    staleTime: 5 * 60 * 1000,
  });

  const teacherType = currentUserProfile?.teacher_type ?? "class_adviser";
  const isSubjectTeacher = teacherType === "subject_teacher";

  const navigationItems = isSubjectTeacher ? subjectTeacherNav : classAdviserNav;

  const dashboardPath = isSubjectTeacher ? "/subject-teacher-dashboard" : "/dashboard";

  const signOut = async () => {
    await queryClient.cancelQueries();
    queryClient.clear();

    const { error } = await supabase.auth.signOut();

    if (error) {
      console.error("Sign-out failed:", error.message);
      return;
    }

    navigate({ to: "/auth", replace: true });
  };

  if (isAdminPage) {
    return <div className="min-h-screen bg-[#fbf7f1]">{children}</div>;
  }

  return (
    <div className="min-h-screen bg-background">
      <header className="sticky top-0 z-20 border-b bg-card/90 backdrop-blur print:hidden">
        <div className="mx-auto flex max-w-[1600px] items-stretch gap-1.5 px-2 sm:gap-3 sm:px-4">
          <Link
            to={dashboardPath}
            className="flex min-w-0 flex-1 items-center gap-2 py-3 sm:gap-3 lg:w-[300px] lg:flex-none xl:w-[400px]"
          >
            <img
              src={logo}
              alt="Agusan del Sur National Science High School"
              className="size-11 shrink-0 rounded-lg object-contain"
            />

            <div className="min-w-0 flex-1 leading-tight">
              <div className="hidden truncate text-[15px] font-semibold xl:block">
                Agusan del Sur National Science High School
              </div>

              {/* Keep ASNSHS visible on mobile/tablet. */}
              <div className="block truncate text-sm font-semibold xl:hidden">
                ASNSHS
              </div>

              <div className="block max-w-[180px] truncate text-[11px] text-muted-foreground sm:max-w-none">
                School Information, Grading &amp; Learner's Assessment
              </div>
            </div>
          </Link>

          <nav
            id="primary-navigation"
            aria-label="Main navigation"
            className="mx-auto hidden h-[72px] min-w-0 flex-1 self-center lg:grid"
            style={{
              gridTemplateColumns: `repeat(${Math.max(navigationItems.length, 1)}, minmax(0, 1fr))`,
            }}
          >
            {!profileLoading
              ? navigationItems.map((item) => {
                  const active = isNavigationItemActive(pathname, item.to);

                  const Icon = item.icon;

                  return (
                    <Link
                      key={item.to}
                      to={item.to}
                      aria-label={item.label}
                      className="group relative flex min-w-0 items-center justify-center overflow-hidden px-1 transition-colors hover:bg-muted/40 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary focus-visible:ring-inset"
                    >
                      <span
                        className={`grid size-10 shrink-0 place-items-center rounded-xl border transition-all duration-200 ease-out group-hover:-translate-y-2 group-hover:shadow-sm group-focus-visible:-translate-y-2 group-focus-visible:shadow-sm ${
                          active
                            ? "border-primary bg-primary text-primary-foreground shadow-sm"
                            : item.tone
                        }`}
                      >
                        <Icon className="size-5" />
                      </span>

                      <span
                        aria-hidden="true"
                        className={`pointer-events-none absolute inset-x-1 bottom-1.5 translate-y-2 truncate text-center text-[10px] font-semibold leading-none opacity-0 transition-all duration-200 ease-out group-hover:translate-y-0 group-hover:opacity-100 group-focus-visible:translate-y-0 group-focus-visible:opacity-100 ${
                          active ? "text-primary" : "text-foreground/75"
                        }`}
                      >
                        {item.label}
                      </span>

                      <span
                        aria-hidden="true"
                        className={`absolute inset-x-2 bottom-0 h-1 rounded-t-full transition-opacity ${
                          active ? "bg-primary opacity-100" : "opacity-0"
                        }`}
                      />
                    </Link>
                  );
                })
              : navigationItems.map((item) => (
                  <div
                    key={item.to}
                    className="grid min-w-0 place-items-center"
                    aria-label="Loading navigation"
                  >
                    <span className="size-10 animate-pulse rounded-xl bg-muted" />
                  </div>
                ))}
          </nav>

          <div className="flex shrink-0 items-center justify-end gap-1.5 py-3 sm:gap-2 lg:w-[300px] xl:w-[400px]">
            <NotificationBell />

            <div>
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <button
                    type="button"
                    aria-label="Open profile menu"
                    className="flex max-w-[250px] items-center gap-2 rounded-xl border bg-background p-1.5 text-left shadow-sm transition hover:bg-accent sm:px-3 sm:py-2"
                  >
                    {currentUserProfile?.avatar_url ? (
                      <img
                        src={currentUserProfile.avatar_url}
                        alt={currentUserProfile.full_name}
                        className="size-9 shrink-0 rounded-full object-cover"
                      />
                    ) : (
                      <div className="grid size-9 shrink-0 place-items-center rounded-full bg-primary/10 text-primary">
                        <UserRound className="size-4" />
                      </div>
                    )}

                    <div className="hidden min-w-0 leading-tight 2xl:block">
                      <div className="truncate text-xs font-semibold">
                        {profileLoading ? "Loading..." : currentUserProfile?.full_name || "Teacher"}
                      </div>
                      <div className="truncate text-[10px] text-muted-foreground">
                        {roleLabel(teacherType)}
                      </div>
                    </div>

                    <ChevronDown className="hidden size-4 shrink-0 text-muted-foreground sm:block" />
                  </button>
                </DropdownMenuTrigger>

                <DropdownMenuContent align="end" className="w-48">
                  <DropdownMenuItem asChild>
                    <Link to="/profile" className="cursor-pointer">
                      <UserRound className="mr-2 size-4" />
                      View Profile
                    </Link>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>

            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={signOut}
              className="hidden border-destructive/40 px-2 text-destructive hover:bg-destructive/10 hover:text-destructive sm:inline-flex md:px-3"
            >
              <LogOut className="size-4 2xl:mr-1" />
              <span className="hidden 2xl:inline">Logout</span>
            </Button>

            <Button
              type="button"
              variant="outline"
              size="icon"
              onClick={() => setMobileNavigationOpen((current) => !current)}
              aria-label={mobileNavigationOpen ? "Close navigation menu" : "Open navigation menu"}
              aria-expanded={mobileNavigationOpen}
              aria-controls="mobile-primary-navigation"
              className={`shrink-0 lg:hidden ${
                mobileNavigationOpen
                  ? "border-primary bg-primary text-primary-foreground hover:bg-primary/90 hover:text-primary-foreground"
                  : ""
              }`}
            >
              {mobileNavigationOpen ? <X className="size-5" /> : <Menu className="size-5" />}
            </Button>
          </div>
        </div>

        {mobileNavigationOpen && (
          <div className="border-t bg-card/95 lg:hidden">
            <nav
              id="mobile-primary-navigation"
              aria-label="Mobile navigation"
              className="mx-auto flex max-w-2xl animate-in flex-col gap-2 px-3 py-3 fade-in-0 slide-in-from-top-2 duration-200 sm:px-4"
            >
              {!profileLoading
                ? navigationItems.map((item) => {
                    const active = isNavigationItemActive(pathname, item.to);
                    const Icon = item.icon;

                    return (
                      <Link
                        key={item.to}
                        to={item.to}
                        onClick={() => setMobileNavigationOpen(false)}
                        aria-current={active ? "page" : undefined}
                        className={`flex min-h-12 items-center gap-3 rounded-xl border px-3 py-2.5 text-sm font-semibold shadow-sm transition active:scale-[0.99] ${
                          active
                            ? "border-primary bg-primary text-primary-foreground"
                            : `${item.tone} hover:brightness-[0.98]`
                        }`}
                      >
                        <span
                          className={`grid size-9 shrink-0 place-items-center rounded-lg ${
                            active ? "bg-white/15" : "bg-white/60"
                          }`}
                        >
                          <Icon className="size-5" />
                        </span>
                        <span className="min-w-0 flex-1 truncate">{item.label}</span>
                      </Link>
                    );
                  })
                : navigationItems.map((item) => (
                    <div
                      key={item.to}
                      className="flex min-h-12 items-center gap-3 rounded-xl border px-3 py-2.5"
                      aria-label="Loading navigation"
                    >
                      <span className="size-9 animate-pulse rounded-lg bg-muted" />
                      <span className="h-4 w-32 animate-pulse rounded bg-muted" />
                    </div>
                  ))}
            </nav>
          </div>
        )}
      </header>

      <div className="mx-auto max-w-7xl px-4 py-6">
        <main className="min-w-0">{children}</main>
      </div>
    </div>
  );
}