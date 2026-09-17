import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "./_lib/auth.js";
import { handleApiError, json, parseBody, requireMethod } from "./_lib/http.js";
import { currentWorkspaceSchemaVersion, loadWorkspaceState, saveWorkspaceState } from "./_lib/workspace-store.js";

type SaveInput = {
  baseRevision?: number | null;
  payload?: {
    schemaVersion: number;
    appStore?: Record<string, unknown>;
    workspace?: Record<string, unknown>;
    preferences?: Record<string, unknown>;
  };
};

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["GET", "PUT"]);
    const session = requireSession(request);
    if (request.method === "GET") {
      const result = await loadWorkspaceState(session);
      return json(response, 200, {
        ok: true,
        schemaVersion: currentWorkspaceSchemaVersion,
        workspaceId: session.workspaceId,
        role: result.role,
        state: result.state ? {
          schemaVersion: result.state.schema_version,
          revision: Number(result.state.revision),
          payload: result.state.payload,
          updatedAt: result.state.updated_at,
          updatedBy: result.state.updated_by,
        } : null,
        preferences: result.preferences,
      });
    }

    const input = parseBody<SaveInput>(request);
    const saved = await saveWorkspaceState(session, input.payload as NonNullable<SaveInput["payload"]>, input.baseRevision ?? null);
    return json(response, 200, {
      ok: true,
      state: {
        schemaVersion: saved.schema_version,
        revision: Number(saved.revision),
        updatedAt: saved.updated_at,
        updatedBy: saved.updated_by,
      },
    });
  } catch (error) {
    return handleApiError(response, error);
  }
}
