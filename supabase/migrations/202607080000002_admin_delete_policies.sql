DROP POLICY IF EXISTS "admins delete profiles" ON public.profiles;

CREATE POLICY "admins delete profiles"
ON public.profiles
FOR DELETE
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins delete classes" ON public.classes;

CREATE POLICY "admins delete classes"
ON public.classes
FOR DELETE
TO authenticated
USING (public.is_admin(auth.uid()));

DROP POLICY IF EXISTS "admins delete students" ON public.students;

CREATE POLICY "admins delete students"
ON public.students
FOR DELETE
TO authenticated
USING (public.is_admin(auth.uid()));