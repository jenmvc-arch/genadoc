import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../../_lib/auth.js";
import { requireGoogleConfiguration } from "../../_lib/config.js";
import { ApiError, handleApiError, requireMethod } from "../../_lib/http.js";
import { createOAuthState } from "../../_lib/oauth-state.js";

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["GET"]);
    requireGoogleConfiguration();
    const session = requireSession(request);
    if (session.role === "Reviewer") throw new ApiError(403, "FORBIDDEN", "Reviewers cannot manage email connections.");
    const returnTo = Array.isArray(request.query.returnTo) ? request.query.returnTo[0] : request.query.returnTo;
    const parameters = new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID as string,
      redirect_uri: process.env.GOOGLE_OAUTH_REDIRECT_URI as string,
      response_type: "code",
      access_type: "offline",
      prompt: "consent select_account",
      scope: "openid email https://www.googleapis.com/auth/gmail.send",
      state: createOAuthState(session, returnTo),
    });
    response.setHeader("Cache-Control", "no-store");
    return response.redirect(302, `https://accounts.google.com/o/oauth2/v2/auth?${parameters.toString()}`);
  } catch (error) {
    return handleApiError(response, error);
  }
}
