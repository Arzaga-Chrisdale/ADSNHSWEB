
-- Profiles / School Info per teacher
CREATE TABLE public.profiles (
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
GRANT SELECT, INSERT, UPDATE, DELETE ON public.profiles TO authenticated;
GRANT ALL ON public.profiles TO service_role;
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
CREATE POLICY "own profile read" ON public.profiles FOR SELECT TO authenticated USING (auth.uid() = id);
CREATE POLICY "own profile write" ON public.profiles FOR INSERT TO authenticated WITH CHECK (auth.uid() = id);
CREATE POLICY "own profile update" ON public.profiles FOR UPDATE TO authenticated USING (auth.uid() = id) WITH CHECK (auth.uid() = id);

-- Classes
CREATE TABLE public.classes (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  subject TEXT NOT NULL,
  grade_level TEXT NOT NULL,
  section TEXT NOT NULL,
  school_year TEXT NOT NULL,
  start_date DATE,
  end_date DATE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.classes TO authenticated;
GRANT ALL ON public.classes TO service_role;
ALTER TABLE public.classes ENABLE ROW LEVEL SECURITY;
CREATE POLICY "own classes" ON public.classes FOR ALL TO authenticated USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

-- Students
CREATE TABLE public.students (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  last_name TEXT NOT NULL,
  first_name TEXT NOT NULL,
  middle_name TEXT,
  sex TEXT CHECK (sex IN ('male','female')),
  lrn TEXT,
  birthdate DATE,
  address TEXT,
  mother_name TEXT,
  father_name TEXT,
  guardian TEXT,
  contact_number TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX ON public.students(class_id);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.students TO authenticated;
GRANT ALL ON public.students TO service_role;
ALTER TABLE public.students ENABLE ROW LEVEL SECURITY;
CREATE POLICY "own students" ON public.students FOR ALL TO authenticated USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

-- Attendance
CREATE TABLE public.attendance_records (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  student_id UUID NOT NULL REFERENCES public.students(id) ON DELETE CASCADE,
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  date DATE NOT NULL,
  status TEXT NOT NULL CHECK (status IN ('present','absent','late','excused','holiday')),
  reason TEXT,
  UNIQUE(student_id, date)
);
CREATE INDEX ON public.attendance_records(class_id, date);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.attendance_records TO authenticated;
GRANT ALL ON public.attendance_records TO service_role;
ALTER TABLE public.attendance_records ENABLE ROW LEVEL SECURITY;
CREATE POLICY "own attendance" ON public.attendance_records FOR ALL TO authenticated USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

-- Grades
CREATE TABLE public.grades (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  student_id UUID NOT NULL REFERENCES public.students(id) ON DELETE CASCADE,
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  subject TEXT NOT NULL,
  term TEXT NOT NULL CHECK (term IN ('1','2','3','final')),
  score NUMERIC(5,2),
  UNIQUE(student_id, subject, term)
);
CREATE INDEX ON public.grades(class_id, term);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.grades TO authenticated;
GRANT ALL ON public.grades TO service_role;
ALTER TABLE public.grades ENABLE ROW LEVEL SECURITY;
CREATE POLICY "own grades" ON public.grades FOR ALL TO authenticated USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

-- Grade requests
CREATE TABLE public.grade_requests (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  class_id UUID NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
  teacher_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  subject TEXT NOT NULL,
  term TEXT NOT NULL,
  teacher_email TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending','submitted')),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.grade_requests TO authenticated;
GRANT ALL ON public.grade_requests TO service_role;
ALTER TABLE public.grade_requests ENABLE ROW LEVEL SECURITY;
CREATE POLICY "own grade requests" ON public.grade_requests FOR ALL TO authenticated USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

-- Individual Learner Status letters
CREATE TABLE public.learner_status_letters (
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
CREATE POLICY "own letters" ON public.learner_status_letters FOR ALL TO authenticated USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

-- Community posts
CREATE TABLE public.community_posts (
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
CREATE POLICY "all can read posts" ON public.community_posts FOR SELECT TO authenticated USING (true);
CREATE POLICY "author can insert" ON public.community_posts FOR INSERT TO authenticated WITH CHECK (auth.uid() = author_id);
CREATE POLICY "author can update" ON public.community_posts FOR UPDATE TO authenticated USING (auth.uid() = author_id) WITH CHECK (auth.uid() = author_id);
CREATE POLICY "author can delete" ON public.community_posts FOR DELETE TO authenticated USING (auth.uid() = author_id);

-- Community comments
CREATE TABLE public.community_comments (
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
CREATE POLICY "all can read comments" ON public.community_comments FOR SELECT TO authenticated USING (true);
CREATE POLICY "author can insert comment" ON public.community_comments FOR INSERT TO authenticated WITH CHECK (auth.uid() = author_id);
CREATE POLICY "author can delete comment" ON public.community_comments FOR DELETE TO authenticated USING (auth.uid() = author_id);

-- Auto-create profile on signup
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
  INSERT INTO public.profiles (id, email, full_name)
  VALUES (NEW.id, NEW.email, COALESCE(NEW.raw_user_meta_data->>'full_name', NEW.email));
  RETURN NEW;
END;
$$;

CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();
