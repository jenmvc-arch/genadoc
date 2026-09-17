import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../_lib/auth.js";
import { database } from "../_lib/database.js";
import { ApiError, handleApiError, requireMethod } from "../_lib/http.js";
import { supabaseStorage } from "../_lib/supabase.js";
import { requireWorkspaceAccess } from "../_lib/workspace-store.js";

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["GET"]);
    const session = requireSession(request);
    await requireWorkspaceAccess(session);
    const path = String(Array.isArray(request.query.path) ? request.query.path[0] : request.query.path || "");
    if (!path.startsWith(`${session.workspaceId}/`)) throw new ApiError(403, "FORBIDDEN", "This file belongs to another workspace.");
    const { sql, ready } = database();
    await ready;
    const rows = await sql<Array<{ file_name: string }>>`
      select file_name from workspace_assets
      where workspace_id = ${session.workspaceId} and storage_path = ${path} and status = 'ready'
      limit 1
    `;
    if (!rows[0]) throw new ApiError(404, "STATE_NOT_FOUND", "The private file is not available.");
    const download = String(Array.isArray(request.query.download) ? request.query.download[0] : request.query.download || "") === "1";
    const { data, error } = await supabaseStorage().createSignedUrl(path, 60, download ? { download: rows[0].file_name } : undefined);
    if (error || !data?.signedUrl) throw new ApiError(404, "STATE_NOT_FOUND", error?.message || "The private file could not be opened.");
    response.setHeader("Cache-Control", "private, no-store");
    return response.redirect(302, data.signedUrl);
  } catch (error) {
    if (response.headersSent) return;
    return handleApiError(response, error);
  }
}

