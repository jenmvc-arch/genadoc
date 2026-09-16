import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../../_lib/auth.js";
import { requireGoogleConfiguration } from "../../_lib/config.js";
import { encryptCredential } from "../../_lib/crypto.js";
import { upsertEmailConnection } from "../../_lib/database.js";
import { ApiError, handleApiError, requireMethod } from "../../_lib/http.js";
import { verifyOAuthState } from "../../_lib/oauth-state.js";

type TokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["GET"]);
    requireGoogleConfiguration();
    const code = Array.isArray(request.query.code) ? request.query.code[0] : request.query.code;
    const stateValue = Array.isArray(request.query.state) ? request.query.state[0] : request.query.state;
    const denied = Array.isArray(request.query.error) ? request.query.error[0] : request.query.error;
    if (denied) return response.redirect(302, "/?email=cancelled#settings");
    if (!code || !stateValue) throw new ApiError(400, "OAUTH_ERROR", "Google did not return a valid authorization response.");
    const state = verifyOAuthState(stateValue);
    const session = requireSession(request);
    if (session.userId !== state.userId || session.workspaceId !== state.workspaceId) {
      throw new ApiError(403, "FORBIDDEN", "This Google connection belongs to a different signed-in workspace session.");
    }

    const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        code,
        client_id: process.env.GOOGLE_CLIENT_ID as string,
        client_secret: process.env.GOOGLE_CLIENT_SECRET as string,
        redirect_uri: process.env.GOOGLE_OAUTH_REDIRECT_URI as string,
        grant_type: "authorization_code",
      }),
    });
    const tokens = await tokenResponse.json() as TokenResponse;
    if (!tokenResponse.ok || !tokens.access_token || !tokens.refresh_token) {
      throw new ApiError(400, "OAUTH_ERROR", tokens.error_description || "Google did not provide an offline refresh token.");
    }

    const profileResponse = await fetch("https://openidconnect.googleapis.com/v1/userinfo", {
      headers: { Authorization: `Bearer ${tokens.access_token}` },
    });
    const profile = await profileResponse.json() as { email?: string; name?: string };
    if (!profileResponse.ok || !profile.email) throw new ApiError(400, "OAUTH_ERROR", "The connected Google account has no email address.");

    const encryptedCredential = encryptCredential({ refreshToken: tokens.refresh_token }, state.userId, state.workspaceId);
    await upsertEmailConnection({
      userId: state.userId,
      workspaceId: state.workspaceId,
      authMethod: "oauth",
      senderEmail: profile.email,
      senderName: profile.name || profile.email.split("@")[0],
      encryptedCredential,
    });
    return response.redirect(302, state.returnTo);
  } catch (error) {
    if (response.headersSent) return;
    return handleApiError(response, error);
  }
}
