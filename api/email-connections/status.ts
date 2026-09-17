import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../_lib/auth.js";
import { platformConfiguration } from "../_lib/config.js";
import { publicConnection } from "../_lib/connection.js";
import { findEmailConnection } from "../_lib/database.js";
import { handleApiError, json, requireMethod } from "../_lib/http.js";
import { requireWorkspaceAccess } from "../_lib/workspace-store.js";

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["GET"]);
    const platform = platformConfiguration();
    if (!platform.ready) {
      return json(response, 503, {
        ok: false,
        code: "BACKEND_NOT_CONFIGURED",
        error: "Cloud credential storage is not configured for this deployment.",
        platform: {
          ready: false,
          googleOAuthReady: false,
          missing: platform.missing,
        },
      });
    }
    const session = requireSession(request);
    await requireWorkspaceAccess(session);
    const connection = await findEmailConnection(session.userId, session.workspaceId);
    return json(response, 200, {
      ok: true,
      workspaceId: session.workspaceId,
      platform: {
        ready: true,
        googleOAuthReady: platform.googleOAuthReady,
        missing: platform.missingGoogle,
      },
      connection: publicConnection(connection),
    });
  } catch (error) {
    return handleApiError(response, error);
  }
}
