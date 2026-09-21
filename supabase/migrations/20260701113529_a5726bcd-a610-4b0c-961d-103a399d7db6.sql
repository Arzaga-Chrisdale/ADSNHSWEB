  ALTER TABLE public.students
    ADD COLUMN IF NOT EXISTS weight_kg numeric,
    ADD COLUMN IF NOT EXISTS height_m numeric;
    
  -- Extend classes with DepEd header fields
  ALTER TABLE public.classes
    ADD COLUMN IF NOT EXISTS school_name text,
    ADD COLUMN IF NOT EXISTS school_id text,
    ADD COLUMN IF NOT EXISTS region text,
    ADD COLUMN IF NOT EXISTS division text,
    ADD COLUMN IF NOT EXISTS district text,
    ADD COLUMN IF NOT EXISTS teacher_name text;

  -- Grade components: WW / PT / QA per class per term, with weight and activity count
  CREATE TABLE IF NOT EXISTS public.grade_components (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    class_id uuid NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
    teacher_id uuid NOT NULL,
    term text NOT NULL,
    component text NOT NULL CHECK (component IN ('WW','PT','QA')),
    weight numeric NOT NULL DEFAULT 0,
    created_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (class_id, term, component)
  );
  GRANT SELECT, INSERT, UPDATE, DELETE ON public.grade_components TO authenticated;
  GRANT ALL ON public.grade_components TO service_role;
  ALTER TABLE public.grade_components ENABLE ROW LEVEL SECURITY;
  CREATE POLICY "own grade_components" ON public.grade_components FOR ALL
    USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

  -- Activities/items under a component (e.g. WW1, WW2)
  CREATE TABLE IF NOT EXISTS public.grade_activities (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    class_id uuid NOT NULL REFERENCES public.classes(id) ON DELETE CASCADE,
    teacher_id uuid NOT NULL,
    term text NOT NULL,
    component text NOT NULL CHECK (component IN ('WW','PT','QA')),
    position int NOT NULL,
    title text,
    hps numeric NOT NULL DEFAULT 10,
    created_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (class_id, term, component, position)
  );
  GRANT SELECT, INSERT, UPDATE, DELETE ON public.grade_activities TO authenticated;
  GRANT ALL ON public.grade_activities TO service_role;
  ALTER TABLE public.grade_activities ENABLE ROW LEVEL SECURITY;
  CREATE POLICY "own grade_activities" ON public.grade_activities FOR ALL
    USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);

  -- Per-student per-activity raw scores
  CREATE TABLE IF NOT EXISTS public.activity_scores (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    activity_id uuid NOT NULL REFERENCES public.grade_activities(id) ON DELETE CASCADE,
    student_id uuid NOT NULL REFERENCES public.students(id) ON DELETE CASCADE,
    class_id uuid NOT NULL,
    teacher_id uuid NOT NULL,
    score numeric,
    created_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (activity_id, student_id)
  );
  GRANT SELECT, INSERT, UPDATE, DELETE ON public.activity_scores TO authenticated;
  GRANT ALL ON public.activity_scores TO service_role;
  ALTER TABLE public.activity_scores ENABLE ROW LEVEL SECURITY;
  CREATE POLICY "own activity_scores" ON public.activity_scores FOR ALL
    USING (auth.uid() = teacher_id) WITH CHECK (auth.uid() = teacher_id);
