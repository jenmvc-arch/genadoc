import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../_lib/auth.js";
import { database } from "../_lib/database.js";
import { ApiError, handleApiError, json, parseBody, requireMethod } from "../_lib/http.js";
import { supabaseStorage } from "../_lib/supabase.js";
import { requireWorkspaceAccess } from "../_lib/workspace-store.js";

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["POST"]);
    const session = requireSession(request);
    const role = await requireWorkspaceAccess(session);
    if (role === "Reviewer") throw new ApiError(403, "FORBIDDEN", "Reviewers cannot upload workspace files.");
    const input = parseBody<{ assetId?: string; storagePath?: string }>(request);
    const assetId = String(input.assetId || "");
    const storagePath = String(input.storagePath || "");
    if (!assetId || !storagePath.startsWith(`${session.workspaceId}/`)) {
      throw new ApiError(400, "INVALID_REQUEST", "The uploaded file does not belong to this workspace.");
    }
    const folder = storagePath.slice(0, storagePath.lastIndexOf("/"));
    const fileName = storagePath.slice(storagePath.lastIndexOf("/") + 1);
    const { data, error } = await supabaseStorage().list(folder, { search: fileName, limit: 2 });
    if (error || !data?.some((item) => item.name === fileName)) {
      throw new ApiError(400, "INVALID_REQUEST", "The file upload did not finish. Retry the upload.");
    }
    const { sql, ready } = database();
    await ready;
    const rows = await sql<Array<{ id: string }>>`
      update workspace_assets set status = 'ready', updated_at = now()
      where id = ${assetId} and workspace_id = ${session.workspaceId} and storage_path = ${storagePath}
      returning id
    `;
    if (!rows[0]) throw new ApiError(404, "STATE_NOT_FOUND", "The pending file record was not found.");
    return json(response, 200, { ok: true, assetId, storagePath });
  } catch (error) {
    return handleApiError(response, error);
  }
}

