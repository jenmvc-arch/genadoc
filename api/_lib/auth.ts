import { createHmac, timingSafeEqual } from "node:crypto";
import type { VercelRequest, VercelResponse } from "@vercel/node";
import { ApiError } from "./http.js";

export type ServerSession = {
  userId: string;
  workspaceId: string;
  email?: string;
  role: "Admin" | "Editor" | "Reviewer";
};

type SessionPayload = {
  sub: string;
  workspaceId: string;
  email?: string;
  role?: ServerSession["role"];
  exp: number;
};

const sessionCookieName = "hrdoc_session";

const encodeSession = (session: ServerSession, expiresAt: number) => {
  const secret = process.env.AUTH_SESSION_SECRET?.trim();
  if (!secret) throw new ApiError(503, "AUTH_NOT_CONFIGURED", "Trusted server authentication is not configured.");
  const payload: SessionPayload = {
    sub: session.userId,
    workspaceId: session.workspaceId,
    email: session.email,
    role: session.role,
    exp: Math.floor(expiresAt / 1000),
  };
  const encodedPayload = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const signature = createHmac("sha256", secret).update(encodedPayload).digest("base64url");
  return `${encodedPayload}.${signature}`;
};

export const setSessionCookie = (response: VercelResponse, session: ServerSession, remember = true) => {
  const maxAge = remember ? 60 * 60 * 24 * 30 : 60 * 60 * 12;
  const token = encodeSession(session, Date.now() + maxAge * 1000);
  response.setHeader(
    "Set-Cookie",
    `${sessionCookieName}=${encodeURIComponent(token)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}; ${process.env.NODE_ENV === "production" ? "Secure; " : ""}`,
  );
};

export const clearSessionCookie = (response: VercelResponse) => {
  response.setHeader(
    "Set-Cookie",
    `${sessionCookieName}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0; ${process.env.NODE_ENV === "production" ? "Secure; " : ""}`,
  );
};

const readCookie = (request: VercelRequest, key: string) => {
  const entry = (request.headers.cookie || "")
    .split(";")
    .map((item) => item.trim())
    .find((item) => item.startsWith(`${key}=`));
  return entry ? decodeURIComponent(entry.slice(key.length + 1)) : "";
};

export const requireSession = (request: VercelRequest): ServerSession => {
  const secret = process.env.AUTH_SESSION_SECRET?.trim();
  if (!secret) throw new ApiError(503, "AUTH_NOT_CONFIGURED", "Trusted server authentication is not configured.");

  const token = readCookie(request, sessionCookieName);
  if (!token) throw new ApiError(401, "AUTH_REQUIRED", "Sign in with a trusted workspace account to continue.");
  const [encodedPayload, encodedSignature] = token.split(".");
  if (!encodedPayload || !encodedSignature) throw new ApiError(401, "AUTH_REQUIRED", "The server session is invalid.");

  const expected = createHmac("sha256", secret).update(encodedPayload).digest();
  const received = Buffer.from(encodedSignature, "base64url");
  if (received.length !== expected.length || !timingSafeEqual(received, expected)) {
    throw new ApiError(401, "AUTH_REQUIRED", "The server session is invalid.");
  }

  let payload: SessionPayload;
  try {
    payload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8")) as SessionPayload;
  } catch {
    throw new ApiError(401, "AUTH_REQUIRED", "The server session is invalid.");
  }
  if (!payload.sub || !payload.workspaceId || payload.exp * 1000 <= Date.now()) {
    throw new ApiError(401, "AUTH_REQUIRED", "The server session has expired.");
  }
  return {
    userId: payload.sub,
    workspaceId: payload.workspaceId,
    email: payload.email,
    role: payload.role || "Editor",
  };
};
