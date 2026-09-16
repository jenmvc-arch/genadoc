import { createCipheriv, createDecipheriv, createHash, randomBytes } from "node:crypto";
import { ApiError } from "./http.js";

const keyFromEnvironment = () => {
  const rawKey = process.env.EMAIL_CREDENTIAL_ENCRYPTION_KEY?.trim();
  if (!rawKey) throw new ApiError(503, "BACKEND_NOT_CONFIGURED", "Credential encryption is not configured.");
  return createHash("sha256").update(rawKey, "utf8").digest();
};

const associatedData = (userId: string, workspaceId: string) => Buffer.from(`${workspaceId}:${userId}`, "utf8");

export const encryptCredential = (credential: unknown, userId: string, workspaceId: string) => {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", keyFromEnvironment(), iv);
  cipher.setAAD(associatedData(userId, workspaceId));
  const ciphertext = Buffer.concat([cipher.update(JSON.stringify(credential), "utf8"), cipher.final()]);
  return ["v1", iv.toString("base64url"), cipher.getAuthTag().toString("base64url"), ciphertext.toString("base64url")].join(".");
};

export const decryptCredential = <T>(encrypted: string, userId: string, workspaceId: string): T => {
  const [version, ivValue, tagValue, ciphertextValue] = encrypted.split(".");
  if (version !== "v1" || !ivValue || !tagValue || !ciphertextValue) {
    throw new ApiError(500, "DELIVERY_ERROR", "The saved credential format is invalid.");
  }
  try {
    const decipher = createDecipheriv("aes-256-gcm", keyFromEnvironment(), Buffer.from(ivValue, "base64url"));
    decipher.setAAD(associatedData(userId, workspaceId));
    decipher.setAuthTag(Buffer.from(tagValue, "base64url"));
    const plaintext = Buffer.concat([decipher.update(Buffer.from(ciphertextValue, "base64url")), decipher.final()]);
    return JSON.parse(plaintext.toString("utf8")) as T;
  } catch {
    throw new ApiError(500, "DELIVERY_ERROR", "The saved credential could not be decrypted.");
  }
};
