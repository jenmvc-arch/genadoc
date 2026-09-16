import type { VercelRequest, VercelResponse } from "@vercel/node";
import nodemailer from "nodemailer";
import { requireSession } from "../_lib/auth.js";
import { requireGoogleConfiguration } from "../_lib/config.js";
import { decryptCredential } from "../_lib/crypto.js";
import { findEmailConnection, recordConnectionTest } from "../_lib/database.js";
import { ApiError, handleApiError, json, parseBody, requireMethod } from "../_lib/http.js";

type OAuthCredential = { refreshToken: string };
type AppPasswordCredential = { appPassword: string };

const sendOAuthTest = async (connection: Awaited<ReturnType<typeof findEmailConnection>>, target: string, userId: string, workspaceId: string) => {
  if (!connection) return;
  requireGoogleConfiguration();
  const credential = decryptCredential<OAuthCredential>(connection.encrypted_credential, userId, workspaceId);
  const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.GOOGLE_CLIENT_ID as string,
      client_secret: process.env.GOOGLE_CLIENT_SECRET as string,
      refresh_token: credential.refreshToken,
      grant_type: "refresh_token",
    }),
  });
  const token = await tokenResponse.json() as { access_token?: string; error_description?: string };
  if (!tokenResponse.ok || !token.access_token) throw new Error(token.error_description || "Google could not refresh this connection.");
  const mime = [
    `From: ${connection.sender_name} <${connection.sender_email}>`,
    `To: ${target}`,
    "Subject: ZhiReady email connection test",
    "Content-Type: text/plain; charset=UTF-8",
    "",
    "Your workspace email connection is ready to send HR documents and workflow notifications.",
  ].join("\r\n");
  const sendResponse = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST",
    headers: { Authorization: `Bearer ${token.access_token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ raw: Buffer.from(mime, "utf8").toString("base64url") }),
  });
  if (!sendResponse.ok) throw new Error("Gmail rejected the test message.");
};

export default async function handler(request: VercelRequest, response: VercelResponse) {
  let connectionId = "";
  try {
    requireMethod(request, ["POST"]);
    const session = requireSession(request);
    const connection = await findEmailConnection(session.userId, session.workspaceId);
    if (!connection || connection.status === "disconnected") {
      throw new ApiError(404, "CONNECTION_NOT_FOUND", "Connect an email account before sending a test.");
    }
    if (connection.last_tested_at && Date.now() - new Date(connection.last_tested_at).getTime() < 30_000) {
      throw new ApiError(429, "RATE_LIMITED", "Wait 30 seconds before sending another test email.");
    }
    connectionId = connection.id;
    const input = parseBody<{ targetEmail?: string }>(request);
    const target = String(input.targetEmail || connection.notification_email || connection.reply_to_email || connection.sender_email).trim();
    if (connection.auth_method === "oauth") {
      await sendOAuthTest(connection, target, session.userId, session.workspaceId);
    } else {
      const credential = decryptCredential<AppPasswordCredential>(connection.encrypted_credential, session.userId, session.workspaceId);
      const transport = nodemailer.createTransport({
        service: "gmail",
        auth: { user: connection.sender_email, pass: credential.appPassword },
      });
      await transport.sendMail({
        from: { name: connection.sender_name, address: connection.sender_email },
        to: target,
        replyTo: connection.reply_to_email || undefined,
        subject: "ZhiReady email connection test",
        text: "Your workspace email connection is ready to send HR documents and workflow notifications.",
      });
    }
    await recordConnectionTest(connection.id);
    return json(response, 200, { ok: true, deliveredTo: target });
  } catch (error) {
    if (connectionId) await recordConnectionTest(connectionId, error instanceof Error ? error.message : "Test delivery failed");
    return handleApiError(response, error instanceof ApiError ? error : new ApiError(502, "DELIVERY_ERROR", error instanceof Error ? error.message : "Test delivery failed."));
  }
}
