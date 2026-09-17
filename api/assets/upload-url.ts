import { randomUUID } from "node:crypto";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../_lib/auth.js";
import { database } from "../_lib/database.js";
import { ApiError, handleApiError, json, parseBody, requireMethod } from "../_lib/http.js";
import { supabaseStorage } from "../_lib/supabase.js";
import { requireWorkspaceAccess } from "../_lib/workspace-store.js";

type Input = {
  fileName?: string;
  contentType?: string;
  byteSize?: number;
  kind?: "letterhead" | "export" | "document-image";
};

const cleanFileName = (value: string) => {
  const cleaned = value.normalize("NFKD").replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/-+/g, "-").slice(-140);
  return cleaned || "asset.bin";
};

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["POST"]);
    const session = requireSession(request);
    const role = await requireWorkspaceAccess(session);
    if (role === "Reviewer") throw new ApiError(403, "FORBIDDEN", "Reviewers cannot upload workspace files.");
    const input = parseBody<Input>(request);
    const byteSize = Number(input.byteSize || 0);
    if (!Number.isFinite(byteSize) || byteSize <= 0 || byteSize > 50 * 1024 * 1024) {
      throw new ApiError(400, "INVALID_REQUEST", "Files must be between 1 byte and 50 MB.");
    }
    const kind = ["letterhead", "export", "document-image"].includes(String(input.kind))
      ? input.kind as NonNullable<Input["kind"]>
      : "document-image";
    const assetId = `asset_${randomUUID()}`;
    const path = `${session.workspaceId}/${kind}/${assetId}-${cleanFileName(String(input.fileName || "asset.bin"))}`;
    const { data, error } = await supabaseStorage().createSignedUploadUrl(path, { upsert: false });
    if (error || !data) throw new ApiError(502, "BACKEND_NOT_CONFIGURED", error?.message || "Could not prepare private file storage.");
    const { sql, ready } = database();
    await ready;
    await sql`
      insert into workspace_assets (
        id, workspace_id, kind, storage_path, file_name, content_type,
        byte_size, status, created_by
      ) values (
        ${assetId}, ${session.workspaceId}, ${kind}, ${path}, ${String(input.fileName || "asset.bin")},
        ${String(input.contentType || "application/octet-stream")}, ${byteSize}, 'pending', ${session.userId}
      )
    `;
    return json(response, 200, {
      ok: true,
      asset: { id: assetId, storagePath: path, signedUrl: data.signedUrl },
    });
  } catch (error) {
    return handleApiError(response, error);
  }
}

