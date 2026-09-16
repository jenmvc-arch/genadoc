import type { EmailConnectionRow } from "./database.js";

export const publicConnection = (connection: EmailConnectionRow | null) => {
  if (!connection) return null;
  return {
    id: connection.id,
    provider: connection.provider,
    authMethod: connection.auth_method,
    senderEmail: connection.sender_email,
    senderName: connection.sender_name,
    replyToEmail: connection.reply_to_email || "",
    documentInboxEmail: connection.document_inbox_email || "",
    notificationEmail: connection.notification_email || "",
    sendDocuments: connection.send_documents,
    receiveCopies: connection.receive_copies,
    notificationsEnabled: connection.notifications_enabled,
    status: connection.status,
    lastTestedAt: connection.last_tested_at,
    lastError: connection.last_error,
    updatedAt: connection.updated_at,
  };
};
