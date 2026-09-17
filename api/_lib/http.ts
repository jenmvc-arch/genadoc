import type { VercelRequest, VercelResponse } from "@vercel/node";

export type ApiErrorCode =
  | "METHOD_NOT_ALLOWED"
  | "INVALID_REQUEST"
  | "AUTH_NOT_CONFIGURED"
  | "AUTH_REQUIRED"
  | "FORBIDDEN"
  | "BACKEND_NOT_CONFIGURED"
  | "STATE_NOT_FOUND"
  | "STATE_CONFLICT"
  | "STATE_TOO_LARGE"
  | "APPROVED_DOCUMENT_IMMUTABLE"
  | "SCHEMA_VERSION_UNSUPPORTED"
  | "CONNECTION_NOT_FOUND"
  | "RATE_LIMITED"
  | "OAUTH_ERROR"
  | "DELIVERY_ERROR";

export class ApiError extends Error {
  constructor(
    public status: number,
    public code: ApiErrorCode,
    message: string,
    public details?: unknown,
  ) {
    super(message);
  }
}

export const json = (response: VercelResponse, status: number, body: unknown) => {
  response.setHeader("Cache-Control", "no-store");
  return response.status(status).json(body);
};

export const requireMethod = (request: VercelRequest, methods: string[]) => {
  if (!request.method || !methods.includes(request.method)) {
    throw new ApiError(405, "METHOD_NOT_ALLOWED", `Use ${methods.join(" or ")}.`);
  }
};

export const parseBody = <T>(request: VercelRequest): T => {
  if (!request.body) return {} as T;
  if (typeof request.body === "string") {
    try {
      return JSON.parse(request.body) as T;
    } catch {
      throw new ApiError(400, "INVALID_REQUEST", "The request body is not valid JSON.");
    }
  }
  return request.body as T;
};

export const handleApiError = (response: VercelResponse, error: unknown) => {
  if (error instanceof ApiError) {
    return json(response, error.status, {
      ok: false,
      code: error.code,
      error: error.message,
      details: error.details,
    });
  }
  console.error(error);
  return json(response, 500, {
    ok: false,
    code: "BACKEND_NOT_CONFIGURED",
    error: "The cloud service could not complete this request.",
  });
};
