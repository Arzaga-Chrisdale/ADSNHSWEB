import { createClient } from "npm:@supabase/supabase-js@2";
import {
  corsHeaders as supabaseCorsHeaders,
} from "npm:@supabase/supabase-js@^2/cors";

const corsHeaders = {
  ...supabaseCorsHeaders,
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

type TeacherType = "class_adviser" | "subject_teacher";

type CreateUserBody = {
  full_name?: unknown;
  email?: unknown;
  password?: unknown;
  role?: unknown;
  teacher_type?: unknown;
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json",
    },
  });
}

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (request.method !== "POST") {
    return json({ error: "Method not allowed." }, 405);
  }

  const supabaseUrl = Deno.env.get("SUPABASE_URL");
  const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
  const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
  const authorization = request.headers.get("Authorization");

  if (!supabaseUrl || !anonKey || !serviceRoleKey) {
    return json(
      { error: "Supabase function environment is incomplete." },
      500,
    );
  }

  if (!authorization) {
    return json({ error: "Missing administrator authorization." }, 401);
  }

  try {
    const body = (await request.json()) as CreateUserBody;

    const fullName = String(body.full_name ?? "").trim();
    const email = String(body.email ?? "").trim().toLowerCase();
    const password = String(body.password ?? "");
    const role = String(body.role ?? "teacher").trim().toLowerCase();
    const teacherTypeValue = String(body.teacher_type ?? "")
      .trim()
      .toLowerCase();

    if (!fullName) {
      return json({ error: "Full name is required." }, 400);
    }

    if (!/^\S+@\S+\.\S+$/.test(email)) {
      return json({ error: "A valid email address is required." }, 400);
    }

    if (password.length < 8) {
      return json(
        { error: "Password must contain at least 8 characters." },
        400,
      );
    }

    if (role !== "teacher" && role !== "admin") {
      return json({ error: "Role must be teacher or admin." }, 400);
    }

    let teacherType: TeacherType | null = null;

    if (role === "teacher") {
      if (
        teacherTypeValue !== "class_adviser" &&
        teacherTypeValue !== "subject_teacher"
      ) {
        return json(
          {
            error:
              "Teacher type must be class_adviser or subject_teacher.",
          },
          400,
        );
      }

      teacherType = teacherTypeValue;
    }

    const callerClient = createClient(supabaseUrl, anonKey, {
      global: {
        headers: {
          Authorization: authorization,
        },
      },
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    });

    const {
      data: { user: caller },
      error: callerError,
    } = await callerClient.auth.getUser();

    if (callerError || !caller) {
      return json(
        { error: "Administrator session is invalid or expired." },
        401,
      );
    }

    const { data: isAdmin, error: adminCheckError } =
      await callerClient.rpc("is_admin", {
        check_user_id: caller.id,
      });

    if (adminCheckError || isAdmin !== true) {
      return json(
        { error: "Only an administrator can create user accounts." },
        403,
      );
    }

    const adminClient = createClient(supabaseUrl, serviceRoleKey, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    });

    const { data: created, error: createError } =
      await adminClient.auth.admin.createUser({
        email,
        password,
        email_confirm: true,
        user_metadata: {
          full_name: fullName,
          teacher_type: teacherType,
        },
      });

    if (createError || !created.user) {
      return json(
        {
          error:
            createError?.message || "Unable to create the Auth user.",
        },
        400,
      );
    }

    const userId = created.user.id;

    const { error: profileError } = await adminClient
      .from("profiles")
      .upsert(
        {
          id: userId,
          full_name: fullName,
          email,
          teacher_type: teacherType,
        },
        {
          onConflict: "id",
        },
      );

    if (profileError) {
      await adminClient.auth.admin.deleteUser(userId);

      return json(
        {
          error: `Unable to create the profile: ${profileError.message}`,
        },
        400,
      );
    }

    const { error: roleError } = await adminClient
      .from("user_roles")
      .insert({
        user_id: userId,
        role,
      });

    if (roleError) {
      await adminClient.from("profiles").delete().eq("id", userId);
      await adminClient.auth.admin.deleteUser(userId);

      return json(
        {
          error: `Unable to assign the account role: ${roleError.message}`,
        },
        400,
      );
    }

    return json({
      user: {
        id: userId,
        full_name: fullName,
        email,
        role,
        teacher_type: teacherType,
      },
    });
  } catch (error) {
    console.error("admin-create-user failed", error);

    return json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Unexpected server error.",
      },
      500,
    );
  }
});
