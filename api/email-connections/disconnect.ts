import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../_lib/auth.js";
import { disconnectEmailConnection } from "../_lib/database.js";
import { ApiError, handleApiError, json, requireMethod } from "../_lib/http.js";

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["DELETE", "POST"]);
    const session = requireSession(request);
    if (session.role === "Reviewer") throw new ApiError(403, "FORBIDDEN", "Reviewers cannot disconnect email accounts.");
    const connection = await disconnectEmailConnection(session.userId, session.workspaceId);
    if (!connection) throw new ApiError(404, "CONNECTION_NOT_FOUND", "No email connection exists for this workspace.");
    return json(response, 200, { ok: true, disconnected: true });
  } catch (error) {
    return handleApiError(response, error);
  }
}
