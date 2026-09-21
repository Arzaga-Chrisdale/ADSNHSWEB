CREATE TABLE IF NOT EXISTS public.user_roles (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  role text NOT NULL CHECK (role IN ('admin', 'teacher')),
  created_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (user_id, role)
);

ALTER TABLE public.user_roles ENABLE ROW LEVEL SECURITY;

CREATE OR REPLACE FUNCTION public.is_admin(check_user_id uuid DEFAULT auth.uid())
RETURNS boolean
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = public
AS $$
  SELECT EXISTS (
    SELECT 1
    FROM public.user_roles
    WHERE user_id = check_user_id
      AND role = 'admin'
  );
$$;

GRANT EXECUTE ON FUNCTION public.is_admin(uuid) TO authenticated;

INSERT INTO public.user_roles (user_id, role)
VALUES ('c9da94f7-0056-4062-ad1e-9871470bfee4', 'admin')
ON CONFLICT (user_id, role) DO NOTHING;


-- Second admin account
INSERT INTO public.user_roles (user_id, role)
VALUES ('5bd017e6-9a5c-469e-98bb-20d031049b34','admin')
ON CONFLICT (user_id, role) DO NOTHING;


-- 3rd admin account
INSERT INTO public.user_roles (user_id, role)
VALUES ('9244a621-4c37-4010-9b53-80de69e22bd7','admin')
ON CONFLICT (user_id, role) DO NOTHING;