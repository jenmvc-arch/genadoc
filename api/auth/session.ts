import { createHash } from "node:crypto";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSupabaseAuthConfiguration } from "../_lib/config.js";
import { clearSessionCookie, requireSession, setSessionCookie, type ServerSession } from "../_lib/auth.js";
import { ensureDemoWorkspace, findFirstMembership, requireWorkspaceAccess } from "../_lib/workspace-store.js";
import { ApiError, handleApiError, json, parseBody, requireMethod } from "../_lib/http.js";

type Input = { email?: string; password?: string; remember?: boolean };

const authenticateWithSupabase = async (email: string, password: string) => {
  requireSupabaseAuthConfiguration();
  const baseUrl = (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL) as string;
  const apiKey = (
    process.env.SUPABASE_PUBLISHABLE_KEY ||
    process.env.SUPABASE_ANON_KEY ||
    process.env.VITE_SUPABASE_ANON_KEY
  ) as string;
  const response = await fetch(`${baseUrl.replace(/\/$/, "")}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { apikey: apiKey, "Content-Type": "application/json" },
    body: JSON.stringify({ email, password }),
  });
  const result = await response.json() as { user?: { id: string; email?: string; user_metadata?: Record<string, unknown> }; error_description?: string; msg?: string };
  if (!response.ok || !result.user?.id) {
    throw new ApiError(401, "AUTH_REQUIRED", result.error_description || result.msg || "Email or password is incorrect.");
  }
  return result.user;
};

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["GET", "POST", "DELETE"]);
    if (request.method === "DELETE") {
      clearSessionCookie(response);
      return json(response, 200, { ok: true });
    }
    if (request.method === "GET") {
      const session = requireSession(request);
      const role = await requireWorkspaceAccess(session);
      return json(response, 200, { ok: true, session: { ...session, role } });
    }

    const input = parseBody<Input>(request);
    const email = String(input.email || "").trim().toLowerCase();
    const password = String(input.password || "");
    let session: ServerSession;
    let displayName: string;
    let workspaceName: string;

    if (!email && !password) {
      if (process.env.ALLOW_DEMO_AUTH !== "true") {
        throw new ApiError(401, "AUTH_REQUIRED", "Blank demo access is disabled on this deployment.");
      }
      const demoKey = createHash("sha256").update("zhiready-demo-user").digest("hex").slice(0, 24);
      session = { userId: `demo_${demoKey}`, workspaceId: "ws_zhiready_demo", email: "demo@zhiready.local", role: "Admin" };
      displayName = "Demo HR user";
      workspaceName = "ZhiReady Demo Workspace";
      await ensureDemoWorkspace(session, displayName);
    } else {
      if (!email || !password) throw new ApiError(400, "INVALID_REQUEST", "Enter both email and password.");
      const user = await authenticateWithSupabase(email, password);
      const membership = await findFirstMembership(user.id);
      if (!membership) throw new ApiError(403, "FORBIDDEN", "This account does not belong to an active HR workspace.");
      session = { userId: user.id, workspaceId: membership.workspace_id, email: user.email || email, role: membership.role };
      displayName = String(user.user_metadata?.full_name || email.split("@")[0]);
      workspaceName = membership.workspace_name;
    }

    setSessionCookie(response, session, input.remember !== false);
    return json(response, 200, {
      ok: true,
      session: {
        ...session,
        displayName,
        workspaceName,
        signedInAt: new Date().toISOString(),
      },
    });
  } catch (error) {
    return handleApiError(response, error);
  }
}
