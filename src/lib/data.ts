import { supabase } from "@/integrations/supabase/client";

export const SUBJECTS = [
  "English",
  "Filipino",
  "Mathematics",
  "Science",
  "Araling Panlipunan",
  "MAPEH",
  "Values Education",
  "TLE: AFA - Agricultural Crop Production",
  "TLE: AFA - Animal Production",
  "TLE: AFA - Aquaculture",
  "TLE: AFA - Fish Capture",
  "TLE: AFA - Food and Beverage Processing",
  "TLE: FCS - Beauty Care Services",
  "TLE: FCS - Food Preparation",
  "TLE: FCS - Food Service",
  "TLE: FCS - Garments",
  "TLE: FCS - Handicraft",
  "TLE: FCS - Health and Wellness Massage",
  "TLE: FCS - Hotel Services",
  "TLE: FCS - Tourism Services",
  "TLE: IA - Automotive and Small Engine Servicing",
  "TLE: IA - Electrical and Electronics Servicing",
  "TLE: IA - Residential Carpentry",
  "TLE: IA - Residential Masonry and Tile Setting",
  "TLE: IA - Residential Plumbing",
  "TLE: IA - Shielded Metal Arc Welding (SMAW)",
  "TLE: ICT - Computer Programming",
  "TLE: ICT - Computer Systems Servicing",
  "TLE: ICT - Telecommunications",
  "TLE: ICT - Visual Arts",
];

export const TERMS = [
  { value: "1", label: "Term 1" },
  { value: "2", label: "Term 2" },
  { value: "3", label: "Term 3" },
  { value: "final", label: "Final" },
] as const;

export type ClassRow = {
  id: string;
  teacher_id: string;
  subject: string;
  grade_level: string;
  section: string;
  school_year: string;
  units: number | null;
  start_date: string | null;
  end_date: string | null;
  school_name: string | null;
  school_id: string | null;
  region: string | null;
  division: string | null;
  district: string | null;
  teacher_name: string | null;
};
export type StudentRow = {
  id: string;
  class_id: string;
  teacher_id: string;
  last_name: string;
  first_name: string;
  middle_name: string | null;
  sex: "male" | "female" | null;
  lrn: string | null;
  birthdate: string | null;
  address: string | null;
  mother_name: string | null;
  father_name: string | null;
  guardian: string | null;
  contact_number: string | null;
};
export type AttendanceRow = {
  id: string;
  student_id: string;
  class_id: string;
  teacher_id: string;
  date: string;
  status: "present" | "absent" | "late" | "excused" | "holiday";
  reason: string | null;
};
export type GradeRow = {
  id: string;
  student_id: string;
  class_id: string;
  teacher_id: string;
  subject: string;
  term: string;
  score: number | null;
};
export type GradeComponent = {
  id: string;
  class_id: string;
  teacher_id: string;
  term: string;
  component: "WW" | "PT" | "QA";
  weight: number;
};
export type GradeActivity = {
  id: string;
  class_id: string;
  teacher_id: string;
  term: string;
  component: "WW" | "PT" | "QA";
  position: number;
  title: string | null;
  hps: number;
};
export type ActivityScore = {
  id: string;
  activity_id: string;
  student_id: string;
  class_id: string;
  teacher_id: string;
  score: number | null;
};

/** DepEd K-12 transmutation (simplified linear): 60 floor, 100 ceiling. */
export function transmute(initial: number | null): number | null {
  if (initial == null) return null;
  const t = Math.round(initial * 0.4 + 60);
  return Math.min(100, Math.max(60, t));
}

export async function getUserId() {
  const { data } = await supabase.auth.getUser();
  return data.user!.id;
}

export function descriptorFor(avg: number | null) {
  if (avg == null) return { label: "—", tone: "muted" as const };
  if (avg >= 90) return { label: "Advancing", tone: "success" as const };
  if (avg >= 80) return { label: "Benchmarking", tone: "success" as const };
  if (avg >= 75) return { label: "Connecting", tone: "muted" as const };
  if (avg >= 65) return { label: "Developing", tone: "muted" as const };
  return { label: "Emerging", tone: "danger" as const };
}

export function computeAverage(scores: Array<number | null | undefined>): number | null {
  const nums = scores.filter((s): s is number => typeof s === "number" && !isNaN(s));
  if (!nums.length) return null;
  return Math.round((nums.reduce((a, b) => a + b, 0) / nums.length) * 100) / 100;
}