import type { VercelRequest, VercelResponse } from "@vercel/node";
import nodemailer from "nodemailer";
import { requireSession } from "../_lib/auth.js";
import { encryptCredential } from "../_lib/crypto.js";
import { publicConnection } from "../_lib/connection.js";
import { upsertEmailConnection } from "../_lib/database.js";
import { ApiError, handleApiError, json, parseBody, requireMethod } from "../_lib/http.js";
import { requireWorkspaceAccess } from "../_lib/workspace-store.js";

type Input = {
  gmailAddress?: string;
  gmailAppPassword?: string;
  senderName?: string;
  replyToEmail?: string;
  documentInboxEmail?: string;
  notificationEmail?: string;
  sendDocuments?: boolean;
  receiveCopies?: boolean;
  notificationsEnabled?: boolean;
};

const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    requireMethod(request, ["POST"]);
    const session = requireSession(request);
    const role = await requireWorkspaceAccess(session);
    if (role === "Reviewer") throw new ApiError(403, "FORBIDDEN", "Reviewers cannot manage email connections.");
    const input = parseBody<Input>(request);
    const gmailAddress = String(input.gmailAddress || "").trim().toLowerCase();
    const appPassword = String(input.gmailAppPassword || "").replace(/\s+/g, "");
    const senderName = String(input.senderName || "").trim();
    if (!emailPattern.test(gmailAddress)) throw new ApiError(400, "INVALID_REQUEST", "Enter a valid Gmail address.");
    if (appPassword.length !== 16) throw new ApiError(400, "INVALID_REQUEST", "Gmail App Passwords contain 16 characters.");
    if (!senderName) throw new ApiError(400, "INVALID_REQUEST", "Enter the sender name recipients should see.");

    try {
      const transport = nodemailer.createTransport({ service: "gmail", auth: { user: gmailAddress, pass: appPassword } });
      await transport.verify();
    } catch {
      throw new ApiError(400, "INVALID_REQUEST", "Google rejected this Gmail address or App Password.");
    }

    const encryptedCredential = encryptCredential({ appPassword }, session.userId, session.workspaceId);
    const connection = await upsertEmailConnection({
      userId: session.userId,
      workspaceId: session.workspaceId,
      authMethod: "app-password",
      senderEmail: gmailAddress,
      senderName,
      replyToEmail: input.replyToEmail?.trim(),
      documentInboxEmail: input.documentInboxEmail?.trim(),
      notificationEmail: input.notificationEmail?.trim(),
      sendDocuments: input.sendDocuments,
      receiveCopies: input.receiveCopies,
      notificationsEnabled: input.notificationsEnabled,
      encryptedCredential,
    });
    return json(response, 200, { ok: true, connection: publicConnection(connection) });
  } catch (error) {
    return handleApiError(response, error);
  }
}
