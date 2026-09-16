import { createHmac, timingSafeEqual } from "node:crypto";
import { ApiError } from "./http.js";
import type { ServerSession } from "./auth.js";

type OAuthState = {
  userId: string;
  workspaceId: string;
  returnTo: string;
  exp: number;
};

const secret = () => {
  const value = process.env.AUTH_SESSION_SECRET?.trim();
  if (!value) throw new ApiError(503, "AUTH_NOT_CONFIGURED", "Trusted server authentication is not configured.");
  return value;
};

const safeReturnTo = (value?: string) => value?.startsWith("/") && !value.startsWith("//") ? value : "/?email=connected#settings";

export const createOAuthState = (session: ServerSession, returnTo?: string) => {
  const payload: OAuthState = {
    userId: session.userId,
    workspaceId: session.workspaceId,
    returnTo: safeReturnTo(returnTo),
    exp: Math.floor(Date.now() / 1000) + 10 * 60,
  };
  const encoded = Buffer.from(JSON.stringify(payload), "utf8").toString("base64url");
  const signature = createHmac("sha256", secret()).update(encoded).digest("base64url");
  return `${encoded}.${signature}`;
};

export const verifyOAuthState = (value: string): OAuthState => {
  const [encoded, signature] = value.split(".");
  if (!encoded || !signature) throw new ApiError(400, "OAUTH_ERROR", "The Google connection request is invalid.");
  const expected = createHmac("sha256", secret()).update(encoded).digest();
  const received = Buffer.from(signature, "base64url");
  if (received.length !== expected.length || !timingSafeEqual(received, expected)) {
    throw new ApiError(400, "OAUTH_ERROR", "The Google connection request could not be verified.");
  }
  const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8")) as OAuthState;
  if (!payload.userId || !payload.workspaceId || payload.exp <= Math.floor(Date.now() / 1000)) {
    throw new ApiError(400, "OAUTH_ERROR", "The Google connection request has expired.");
  }
  return { ...payload, returnTo: safeReturnTo(payload.returnTo) };
};
