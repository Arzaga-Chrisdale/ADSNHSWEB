-- Admin Analytics & Insights: read-only access to analytics source tables.
-- This migration assumes public.is_admin(uuid) already exists.
-- It does NOT give Admin write access to grades or class records.

ALTER TABLE public.classes ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.students ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.grades ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.grade_activities ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.grade_components ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.activity_scores ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "admins read analytics classes" ON public.classes;
CREATE POLICY "admins read analytics classes"
ON public.classes
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins read analytics students" ON public.students;
CREATE POLICY "admins read analytics students"
ON public.students
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins read analytics profiles" ON public.profiles;
CREATE POLICY "admins read analytics profiles"
ON public.profiles
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins read analytics grades" ON public.grades;
CREATE POLICY "admins read analytics grades"
ON public.grades
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins read analytics grade activities"
ON public.grade_activities;
CREATE POLICY "admins read analytics grade activities"
ON public.grade_activities
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins read analytics grade components"
ON public.grade_components;
CREATE POLICY "admins read analytics grade components"
ON public.grade_components
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins read analytics activity scores"
ON public.activity_scores;
CREATE POLICY "admins read analytics activity scores"
ON public.activity_scores
FOR SELECT
TO authenticated
USING (public.is_admin(auth.uid()));

GRANT SELECT ON TABLE
  public.classes,
  public.students,
  public.profiles,
  public.grades,
  public.grade_activities,
  public.grade_components,
  public.activity_scores
TO authenticated;

NOTIFY pgrst, 'reload schema';
