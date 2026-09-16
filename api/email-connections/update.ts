import type { VercelRequest, VercelResponse } from "@vercel/node";
import { requireSession } from "../_lib/auth.js";
import { publicConnection } from "../_lib/connection.js";
import { updateEmailPreferences } from "../_lib/database.js";
import { ApiError, handleApiError, json, parseBody, requireMethod } from "../_lib/http.js";

type Input = {
  senderName?: string;
  replyToEmail?: string;
  documentInboxEmail?: string;
  notificationEmail?: string;
  sendDocuments?: boolean;
  receiveCopies?: boolean;
  notificationsEnabled?: boolean;
};

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["PATCH", "POST"]);
    const session = requireSession(request);
    if (session.role === "Reviewer") throw new ApiError(403, "FORBIDDEN", "Reviewers cannot change email delivery settings.");
    const input = parseBody<Input>(request);
    const senderName = String(input.senderName || "").trim();
    if (!senderName) throw new ApiError(400, "INVALID_REQUEST", "Enter the sender name recipients should see.");
    const connection = await updateEmailPreferences(session.userId, session.workspaceId, {
      senderName,
      replyToEmail: input.replyToEmail?.trim(),
      documentInboxEmail: input.documentInboxEmail?.trim(),
      notificationEmail: input.notificationEmail?.trim(),
      sendDocuments: input.sendDocuments,
      receiveCopies: input.receiveCopies,
      notificationsEnabled: input.notificationsEnabled,
    });
    if (!connection) throw new ApiError(404, "CONNECTION_NOT_FOUND", "Connect an email account before saving delivery preferences.");
    return json(response, 200, { ok: true, connection: publicConnection(connection) });
  } catch (error) {
    return handleApiError(response, error);
  }
}
