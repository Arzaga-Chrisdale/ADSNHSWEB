import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Cache-Control": "no-store",
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

function normalizeEmail(value: unknown) {
  return typeof value === "string"
    ? value.trim().toLowerCase()
    : "";
}

function randomOtp() {
  const value = crypto.getRandomValues(new Uint32Array(1))[0] % 1_000_000;
  return value.toString().padStart(6, "0");
}

function randomToken() {
  const bytes = crypto.getRandomValues(new Uint8Array(32));

  return btoa(String.fromCharCode(...bytes))
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replaceAll("=", "");
}

async function sha256(value: string) {
  const data = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", data);

  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function getServiceRoleKey() {
  return Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
}

async function secretHash(value: string) {
  const secret = getServiceRoleKey();

  if (!secret) {
    throw new Error("SUPABASE_SERVICE_ROLE_KEY is unavailable.");
  }

  return sha256(`${value}:${secret}`);
}

type AdminClient = ReturnType<typeof createClient>;

async function findAuthUserByEmail(
  admin: AdminClient,
  email: string,
) {
  const perPage = 1000;

  for (let page = 1; page <= 20; page += 1) {
    const { data, error } = await admin.auth.admin.listUsers({
      page,
      perPage,
    });

    if (error) {
      throw error;
    }

    const match = data.users.find(
      (user) => user.email?.trim().toLowerCase() === email,
    );

    if (match) {
      return match;
    }

    if (data.users.length < perPage) {
      return null;
    }
  }

  return null;
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (req.method !== "POST") {
    return json({ error: "Method not allowed." }, 405);
  }

  const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
  const serviceRoleKey = getServiceRoleKey();

  if (!supabaseUrl || !serviceRoleKey) {
    return json(
      { error: "Password-reset server configuration is incomplete." },
      500,
    );
  }

  const admin = createClient(supabaseUrl, serviceRoleKey, {
    auth: {
      autoRefreshToken: false,
      persistSession: false,
    },
  });

  let body: Record<string, unknown>;

  try {
    body = await req.json();
  } catch {
    return json({ error: "Invalid request body." }, 400);
  }

  const action =
    typeof body.action === "string"
      ? body.action.trim()
      : "";

  const email = normalizeEmail(body.email);

  if (!email) {
    return json({ error: "Email is required." }, 400);
  }

  let authUser;

  try {
    authUser = await findAuthUserByEmail(admin, email);
  } catch (error) {
    return json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Unable to read Supabase Auth users.",
      },
      500,
    );
  }

  if (!authUser) {
    return json(
      { error: "This email is not registered in Supabase Authentication." },
      404,
    );
  }

  const { data: profile, error: profileError } = await admin
    .from("profiles")
    .select("id, teacher_type, full_name")
    .eq("id", authUser.id)
    .maybeSingle();

  if (profileError) {
    return json({ error: profileError.message }, 500);
  }

  if (!profile) {
    return json(
      { error: "This account is not registered in the school portal." },
      404,
    );
  }

  const { data: adminRole, error: roleError } = await admin
    .from("user_roles")
    .select("role")
    .eq("user_id", authUser.id)
    .eq("role", "admin")
    .maybeSingle();

  if (roleError) {
    return json({ error: roleError.message }, 500);
  }

  const teacherType =
    typeof profile.teacher_type === "string"
      ? profile.teacher_type.trim().toLowerCase()
      : "";

  let roleLabel = "";

  if (adminRole?.role === "admin") {
    roleLabel = "Admin";
  } else if (teacherType === "class_adviser") {
    roleLabel = "Class Adviser";
  } else if (teacherType === "subject_teacher") {
    roleLabel = "Subject Teacher";
  }

  if (!roleLabel) {
    return json(
      {
        error:
          "Only Admin, Class Adviser, and Subject Teacher accounts can reset a password here.",
      },
      403,
    );
  }

  if (action === "request_code") {
    const oneMinuteAgo = new Date(
      Date.now() - 60_000,
    ).toISOString();

    const { data: recent, error: recentError } = await admin
      .from("password_reset_challenges")
      .select("id")
      .eq("user_id", authUser.id)
      .gte("created_at", oneMinuteAgo)
      .is("used_at", null)
      .limit(1);

    if (recentError) {
      return json({ error: recentError.message }, 500);
    }

    if (recent && recent.length > 0) {
      return json(
        {
          error:
            "A verification code was sent recently. Please wait about one minute before requesting another code.",
        },
        429,
      );
    }

    const otp = randomOtp();
    const expiresAt = new Date(Date.now() + 15 * 60_000);
    const otpHash = await secretHash(
      `${authUser.id}:${email}:${otp}`,
    );

    const { data: challenge, error: insertError } = await admin
      .from("password_reset_challenges")
      .insert({
        user_id: authUser.id,
        email,
        role_label: roleLabel,
        otp_hash: otpHash,
        expires_at: expiresAt.toISOString(),
      })
      .select("id")
      .single();

    if (insertError || !challenge) {
      return json(
        {
          error:
            insertError?.message ??
            "Unable to create a password-reset request.",
        },
        500,
      );
    }

    const serviceId = Deno.env.get("EMAILJS_SERVICE_ID") ?? "";
    const templateId = Deno.env.get("EMAILJS_TEMPLATE_ID") ?? "";
    const publicKey = Deno.env.get("EMAILJS_PUBLIC_KEY") ?? "";
    const privateKey = Deno.env.get("EMAILJS_PRIVATE_KEY") ?? "";

    if (!serviceId || !templateId || !publicKey) {
      await admin
        .from("password_reset_challenges")
        .delete()
        .eq("id", challenge.id);

      return json(
        {
          error:
            "EmailJS server secrets are not configured for reset-user-password.",
        },
        500,
      );
    }

    const time = new Intl.DateTimeFormat("en-PH", {
      timeZone: "Asia/Manila",
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
    }).format(expiresAt);

    const emailPayload: Record<string, unknown> = {
      service_id: serviceId,
      template_id: templateId,
      user_id: publicKey,
      template_params: {
        email,
        to_email: email,
        user_name:
          typeof profile.full_name === "string" &&
          profile.full_name.trim()
            ? profile.full_name.trim()
            : email,
        user_role: roleLabel,
        passcode: otp,
        time,
      },
    };

    if (privateKey) {
      emailPayload.accessToken = privateKey;
    }

    const emailResponse = await fetch(
      "https://api.emailjs.com/api/v1.0/email/send",
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(emailPayload),
      },
    );

    if (!emailResponse.ok) {
      const text = await emailResponse.text();

      await admin
        .from("password_reset_challenges")
        .delete()
        .eq("id", challenge.id);

      console.error(
        "EmailJS send failed:",
        emailResponse.status,
        text,
      );

      return json(
        { error: "The verification email could not be sent." },
        502,
      );
    }

    return json({ success: true });
  }

  if (action === "verify_code") {
    const passcode =
      typeof body.passcode === "string"
        ? body.passcode.trim()
        : "";

    if (!/^\d{6}$/.test(passcode)) {
      return json(
        { error: "Enter the complete 6-digit verification code." },
        400,
      );
    }

    const { data: challenge, error: challengeError } = await admin
      .from("password_reset_challenges")
      .select(
        "id, otp_hash, expires_at, attempts, verified_at, used_at",
      )
      .eq("user_id", authUser.id)
      .eq("email", email)
      .is("used_at", null)
      .order("created_at", { ascending: false })
      .limit(1)
      .maybeSingle();

    if (challengeError) {
      return json({ error: challengeError.message }, 500);
    }

    if (!challenge) {
      return json(
        { error: "No active verification code was found." },
        400,
      );
    }

    if (new Date(challenge.expires_at) <= new Date()) {
      return json(
        {
          error:
            "Your verification code has expired. Request a new one.",
        },
        400,
      );
    }

    if (challenge.attempts >= 5) {
      return json(
        {
          error:
            "Too many incorrect attempts. Request a new verification code.",
        },
        429,
      );
    }

    const candidateHash = await secretHash(
      `${authUser.id}:${email}:${passcode}`,
    );

    if (candidateHash !== challenge.otp_hash) {
      await admin
        .from("password_reset_challenges")
        .update({
          attempts: challenge.attempts + 1,
        })
        .eq("id", challenge.id);

      return json(
        { error: "The verification code is incorrect." },
        400,
      );
    }

    const resetToken = randomToken();
    const resetTokenHash = await secretHash(
      `${authUser.id}:${email}:${resetToken}`,
    );
    const tokenExpiresAt = new Date(Date.now() + 10 * 60_000);

    const { error: verifyUpdateError } = await admin
      .from("password_reset_challenges")
      .update({
        verified_at: new Date().toISOString(),
        reset_token_hash: resetTokenHash,
        reset_token_expires_at: tokenExpiresAt.toISOString(),
      })
      .eq("id", challenge.id);

    if (verifyUpdateError) {
      return json({ error: verifyUpdateError.message }, 500);
    }

    return json({
      success: true,
      reset_token: resetToken,
    });
  }

  if (action === "reset_password") {
    const resetToken =
      typeof body.reset_token === "string"
        ? body.reset_token.trim()
        : "";

    const newPassword =
      typeof body.new_password === "string"
        ? body.new_password
        : "";

    if (!resetToken) {
      return json(
        { error: "The verified reset session is missing." },
        400,
      );
    }

    if (newPassword.length < 6) {
      return json(
        { error: "Password must be at least 6 characters long." },
        400,
      );
    }

    const { data: challenge, error: challengeError } = await admin
      .from("password_reset_challenges")
      .select(
        "id, verified_at, reset_token_hash, reset_token_expires_at, used_at",
      )
      .eq("user_id", authUser.id)
      .eq("email", email)
      .not("verified_at", "is", null)
      .is("used_at", null)
      .order("created_at", { ascending: false })
      .limit(1)
      .maybeSingle();

    if (challengeError) {
      return json({ error: challengeError.message }, 500);
    }

    if (
      !challenge?.reset_token_hash ||
      !challenge.reset_token_expires_at
    ) {
      return json(
        { error: "No verified password-reset session was found." },
        400,
      );
    }

    if (
      new Date(challenge.reset_token_expires_at) <= new Date()
    ) {
      return json(
        {
          error:
            "Your verified reset session has expired. Request a new code.",
        },
        400,
      );
    }

    const candidateTokenHash = await secretHash(
      `${authUser.id}:${email}:${resetToken}`,
    );

    if (candidateTokenHash !== challenge.reset_token_hash) {
      return json(
        { error: "The password-reset session is invalid." },
        403,
      );
    }

    const { error: authError } =
      await admin.auth.admin.updateUserById(
        authUser.id,
        {
          password: newPassword,
        },
      );

    if (authError) {
      return json({ error: authError.message }, 500);
    }

    const { error: usedError } = await admin
      .from("password_reset_challenges")
      .update({
        used_at: new Date().toISOString(),
      })
      .eq("id", challenge.id);

    if (usedError) {
      console.error(
        "Password updated, but challenge cleanup failed:",
        usedError.message,
      );
    }

    return json({ success: true });
  }

  return json(
    { error: "Unknown password-reset action." },
    400,
  );
});
