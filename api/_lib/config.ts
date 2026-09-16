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
