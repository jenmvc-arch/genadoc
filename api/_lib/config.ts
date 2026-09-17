import { ApiError } from "./http.js";

const requiredPlatformVariables = [
  "DATABASE_URL",
  "AUTH_SESSION_SECRET",
  "EMAIL_CREDENTIAL_ENCRYPTION_KEY",
] as const;

const requiredGoogleVariables = [
  "GOOGLE_CLIENT_ID",
  "GOOGLE_CLIENT_SECRET",
  "GOOGLE_OAUTH_REDIRECT_URI",
] as const;

export const platformConfiguration = () => {
  const missingPlatform = requiredPlatformVariables.filter((key) => !process.env[key]?.trim());
  const missingGoogle = requiredGoogleVariables.filter((key) => !process.env[key]?.trim());
  return {
    ready: missingPlatform.length === 0,
    googleOAuthReady: missingPlatform.length === 0 && missingGoogle.length === 0,
    missing: [...missingPlatform, ...missingGoogle],
    missingPlatform,
    missingGoogle,
  };
};

export const requirePlatformConfiguration = () => {
  const configuration = platformConfiguration();
  if (!configuration.ready) {
    throw new ApiError(
      503,
      "BACKEND_NOT_CONFIGURED",
      "Cloud credential storage is not configured for this deployment.",
      { missing: configuration.missingPlatform },
    );
  }
  return configuration;
};

export const requireDatabaseConfiguration = () => {
  if (!process.env.DATABASE_URL?.trim()) {
    throw new ApiError(
      503,
      "BACKEND_NOT_CONFIGURED",
      "Supabase Postgres is not configured for this deployment.",
      { missing: ["DATABASE_URL"] },
    );
  }
};

export const requireSupabaseAuthConfiguration = () => {
  const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
  const supabasePublishableKey =
    process.env.SUPABASE_PUBLISHABLE_KEY ||
    process.env.SUPABASE_ANON_KEY ||
    process.env.VITE_SUPABASE_ANON_KEY;
  const missing = [
    !supabaseUrl?.trim() ? "VITE_SUPABASE_URL" : "",
    !supabasePublishableKey?.trim()
      ? "SUPABASE_PUBLISHABLE_KEY"
      : "",
    !process.env.AUTH_SESSION_SECRET?.trim() ? "AUTH_SESSION_SECRET" : "",
  ].filter(Boolean);
  if (missing.length) {
    throw new ApiError(
      503,
      "AUTH_NOT_CONFIGURED",
      "Supabase authentication is not configured for this deployment.",
      { missing },
    );
  }
};

export const requireSupabaseStorageConfiguration = () => {
  const supabaseUrl = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL;
  const missing = [
    !supabaseUrl?.trim() ? "VITE_SUPABASE_URL" : "",
    !(process.env.SUPABASE_SECRET_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY)?.trim()
      ? "SUPABASE_SECRET_KEY"
      : "",
  ].filter(Boolean);
  if (missing.length) {
    throw new ApiError(
      503,
      "BACKEND_NOT_CONFIGURED",
      "Supabase private file storage is not configured for this deployment.",
      { missing },
    );
  }
};

export const requireGoogleConfiguration = () => {
  const configuration = requirePlatformConfiguration();
  if (!configuration.googleOAuthReady) {
    throw new ApiError(
      503,
      "BACKEND_NOT_CONFIGURED",
      "Google OAuth is not configured for this deployment.",
      { missing: configuration.missingGoogle },
    );
  }
  return configuration;
};
