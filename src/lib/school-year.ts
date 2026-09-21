import { supabase } from "@/integrations/supabase/client";

export type SchoolYearSetting = {
  id: number;
  school_year: string | null;
  is_locked: boolean;
  updated_by: string | null;
  updated_at: string;
};

export function normalizeSchoolYearInput(value: string) {
  return value.trim().replace(/[–—]/g, "-").replace(/\s+/g, "");
}

export function isValidSchoolYear(value: string) {
  const normalized = normalizeSchoolYearInput(value);
  const match = /^(\d{4})-(\d{4})$/.exec(normalized);

  if (!match) return false;

  const startYear = Number(match[1]);
  const endYear = Number(match[2]);

  return endYear === startYear + 1;
}

export async function getSchoolYearSetting(): Promise<SchoolYearSetting> {
  // This table is added by the School Year migration and may not yet exist
  // in the generated Supabase Database type file.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data, error } = await (supabase as any)
    .from("school_year_settings")
    .select("id, school_year, is_locked, updated_by, updated_at")
    .eq("id", 1)
    .maybeSingle();

  if (error) throw error;

  return {
    id: 1,
    school_year: data?.school_year ?? null,
    is_locked: Boolean(data?.is_locked),
    updated_by: data?.updated_by ?? null,
    updated_at: data?.updated_at ?? new Date(0).toISOString(),
  };
}

export async function requireLockedSchoolYear() {
  const setting = await getSchoolYearSetting();

  if (!setting.is_locked || !setting.school_year) {
    throw new Error(
      "The administrator has not locked an active School Year yet. Please contact the administrator.",
    );
  }

  return setting.school_year;
}
