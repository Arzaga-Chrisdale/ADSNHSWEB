BEGIN;

CREATE EXTENSION IF NOT EXISTS pgcrypto;


DROP TABLE IF EXISTS public.attendance_records CASCADE;

-- ============================================================
-- Profiles / School information per teacher
-- ============================================================

CREATE TABLE IF NOT EXISTS public.profiles (
  id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  full_name TEXT,
  email TEXT,
  school_name TEXT,
  school_id TEXT,
  region TEXT,
  division TEXT,
  district TEXT,
  principal TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS full_name TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS email TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS school_name TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS school_id TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS region TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS division TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS district TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS principal TEXT;
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT now();
ALTER TABLE public.profiles ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT now();

GRANT SELECT, INSERT, UPDATE, DELETE ON public.profiles TO authenticated;
GRANT ALL ON public.profiles TO service_role;

ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "own profile read" ON public.profiles;
DROP POLICY IF EXISTS "own profile write" ON public.profiles;
DROP POLICY IF EXISTS "own profile update" ON public.profiles;

CREATE POLICY "own profile read"
ON public.profiles
FOR SELECT
TO authenticated
USING (auth.uid() = id);

CREATE POLICY "own profile write"
ON public.profiles
FOR INSERT
TO authenticated
WITH CHECK (auth.uid() = id);

CREATE POLICY "own profile update"
ON public.profiles
FOR UPDATE
TO authenticated
USING (auth.uid() = id)
WITH CHECK (auth.uid() = id);

-- ============================================================
-- Classes
-- ============================================================

CREATE TABLE IF NOT EXISTS public.classes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  subject TEXT NOT NULL,
  grade_level TEXT NOT NULL,
  section TEXT NOT NULL,
  school_year TEXT NOT NULL,
  units SMALLINT,
  start_date DATE,
  end_date DATE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.classes ADD COLUMN IF NOT EXISTS units SMALLINT;

ALTER TABLE public.classes
DROP CONSTRAINT IF EXISTS classes_units_by_grade_check;

ALTER TABLE public.classes
ADD CONSTRAINT classes_units_by_grade_check CHECK (
  (grade_level = 'Grade 11' AND units IN (2, 3, 6))
  OR
  (grade_level = 'Grade 12' AND (units IS NULL OR units = 3))
  OR  
  (grade_level NOT IN ('Grade 11', 'Grade 12') AND units IS NULL)
);

COMMENT ON COLUMN public.classes.units IS
  'Grade 11: 2, 3, or 6 units; Grade 12: 3 units or none (NULL); Grade 7-10: no units (NULL).';

GRANT SELECT, INSERT, UPDATE, DELETE ON public.classes TO authenticated;
GRANT ALL ON public.classes TO service_role;

ALTER TABLE public.classes ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "own classes" ON public.classes;

CREATE POLICY "own classes"
ON public.classes
FOR ALL
TO authenticated
USING (auth.uid() = teacher_id)
WITH CHECK (auth.uid() = teacher_id);

-- ============================================================
-- Students
-- ============================================================

CREATE TABLE IF NOT EXISTS public.students (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  last_name TEXT NOT NULL,
  first_name TEXT NOT NULL,
  middle_name TEXT,
  sex TEXT CHECK (sex IN ('male', 'female')),
  lrn TEXT,
  birthdate DATE,
  address TEXT,
  mother_name TEXT,
  father_name TEXT,
  guardian TEXT,
  contact_number TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS students_class_id_idx
ON public.students(class_id);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.students TO authenticated;
GRANT ALL ON public.students TO service_role;

ALTER TABLE public.students ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "own students" ON public.students;

CREATE POLICY "own students"
ON public.students
FOR ALL
TO authenticated
USING (auth.uid() = teacher_id)
WITH CHECK (auth.uid() = teacher_id);

-- ============================================================
-- Grades
-- ============================================================

CREATE TABLE IF NOT EXISTS public.grades (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  student_id UUID NOT NULL REFERENCES public.students(id) ON DELETE CASCADE,
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  subject TEXT NOT NULL,
  term TEXT NOT NULL CHECK (term IN ('1', '2', '3', 'final')),
  score NUMERIC(5,2),
  UNIQUE(student_id, subject, term)
);

CREATE INDEX IF NOT EXISTS grades_class_id_term_idx
ON public.grades(class_id, term);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.grades TO authenticated;
GRANT ALL ON public.grades TO service_role;

ALTER TABLE public.grades ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "own grades" ON public.grades;

CREATE POLICY "own grades"
ON public.grades
FOR ALL
TO authenticated
USING (auth.uid() = teacher_id)
WITH CHECK (auth.uid() = teacher_id);

-- ============================================================
-- Grade requests
-- ============================================================

CREATE TABLE IF NOT EXISTS public.grade_requests (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  subject TEXT NOT NULL,
  term TEXT NOT NULL,
  teacher_email TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending'
    CHECK (status IN ('pending', 'submitted')),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.grade_requests TO authenticated;
GRANT ALL ON public.grade_requests TO service_role;

ALTER TABLE public.grade_requests ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "own grade requests" ON public.grade_requests;

CREATE POLICY "own grade requests"
ON public.grade_requests
FOR ALL
TO authenticated
USING (auth.uid() = teacher_id)
WITH CHECK (auth.uid() = teacher_id);

-- ============================================================
-- Individual learner status letters
-- ============================================================

CREATE TABLE IF NOT EXISTS public.learner_status_letters (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  student_id UUID NOT NULL REFERENCES public.students(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  meeting_date DATE,
  meeting_time TIME,
  meeting_venue TEXT,
  teacher_notes TEXT,
  action_plan_items JSONB NOT NULL DEFAULT '[]'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.learner_status_letters TO authenticated;
GRANT ALL ON public.learner_status_letters TO service_role;

ALTER TABLE public.learner_status_letters ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "own letters" ON public.learner_status_letters;

CREATE POLICY "own letters"
ON public.learner_status_letters
FOR ALL
TO authenticated
USING (auth.uid() = teacher_id)
WITH CHECK (auth.uid() = teacher_id);

-- ============================================================
-- Community posts
-- ============================================================

CREATE TABLE IF NOT EXISTS public.community_posts (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  author_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  author_name TEXT,
  title TEXT NOT NULL,
  content TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.community_posts TO authenticated;
GRANT ALL ON public.community_posts TO service_role;

ALTER TABLE public.community_posts ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "all can read posts" ON public.community_posts;
DROP POLICY IF EXISTS "author can insert" ON public.community_posts;
DROP POLICY IF EXISTS "author can update" ON public.community_posts;
DROP POLICY IF EXISTS "author can delete" ON public.community_posts;

CREATE POLICY "all can read posts"
ON public.community_posts
FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "author can insert"
ON public.community_posts
FOR INSERT
TO authenticated
WITH CHECK (auth.uid() = author_id);

CREATE POLICY "author can update"
ON public.community_posts
FOR UPDATE
TO authenticated
USING (auth.uid() = author_id)
WITH CHECK (auth.uid() = author_id);

CREATE POLICY "author can delete"
ON public.community_posts
FOR DELETE
TO authenticated
USING (auth.uid() = author_id);

-- ============================================================
-- Community comments
-- ============================================================

CREATE TABLE IF NOT EXISTS public.community_comments (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  post_id UUID NOT NULL REFERENCES public.community_posts(id) ON DELETE CASCADE,
  author_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  author_name TEXT,
  content TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, UPDATE, DELETE ON public.community_comments TO authenticated;
GRANT ALL ON public.community_comments TO service_role;

ALTER TABLE public.community_comments ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "all can read comments" ON public.community_comments;
DROP POLICY IF EXISTS "author can insert comment" ON public.community_comments;
DROP POLICY IF EXISTS "author can delete comment" ON public.community_comments;

CREATE POLICY "all can read comments"
ON public.community_comments
FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "author can insert comment"
ON public.community_comments
FOR INSERT
TO authenticated
WITH CHECK (auth.uid() = author_id);

CREATE POLICY "author can delete comment"
ON public.community_comments
FOR DELETE
TO authenticated
USING (auth.uid() = author_id);

-- ============================================================
-- Automatically create a profile when a user signs up
-- ============================================================

CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
  INSERT INTO public.profiles (id, email, full_name)
  VALUES (
    NEW.id,
    NEW.email,
    COALESCE(NEW.raw_user_meta_data->>'full_name', NEW.email)
  )
  ON CONFLICT (id) DO UPDATE
  SET
    email = EXCLUDED.email,
    full_name = COALESCE(public.profiles.full_name, EXCLUDED.full_name),
    updated_at = now();

  RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;

CREATE TRIGGER on_auth_user_created
AFTER INSERT ON auth.users
FOR EACH ROW
EXECUTE FUNCTION public.handle_new_user();

COMMIT;
