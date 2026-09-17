import { randomUUID } from "node:crypto";
import postgres from "postgres";
import { requireDatabaseConfiguration } from "./config.js";

export type EmailConnectionRow = {
  id: string;
  user_id: string;
  workspace_id: string;
  provider: "gmail";
  auth_method: "oauth" | "app-password";
  sender_email: string;
  sender_name: string;
  reply_to_email: string | null;
  document_inbox_email: string | null;
  notification_email: string | null;
  send_documents: boolean;
  receive_copies: boolean;
  notifications_enabled: boolean;
  status: "connected" | "expired" | "error" | "disconnected";
  encrypted_credential: string;
  last_tested_at: string | null;
  last_error: string | null;
  created_at: string;
  updated_at: string;
};

let client: ReturnType<typeof postgres> | null = null;
let schemaPromise: Promise<unknown> | null = null;

export const database = () => {
  requireDatabaseConfiguration();
  if (!client) client = postgres(process.env.DATABASE_URL as string, { max: 4, idle_timeout: 20 });
  if (!schemaPromise) {
    schemaPromise = client`
      create table if not exists email_connections (
        id text primary key,
        user_id text not null,
        workspace_id text not null,
        provider text not null check (provider = 'gmail'),
        auth_method text not null check (auth_method in ('oauth', 'app-password')),
        sender_email text not null,
        sender_name text not null,
        reply_to_email text,
        document_inbox_email text,
        notification_email text,
        send_documents boolean not null default true,
        receive_copies boolean not null default false,
        notifications_enabled boolean not null default true,
        status text not null default 'connected' check (status in ('connected', 'expired', 'error', 'disconnected')),
        encrypted_credential text not null,
        last_tested_at timestamptz,
        last_error text,
        created_at timestamptz not null default now(),
        updated_at timestamptz not null default now(),
        unique (user_id, workspace_id, provider)
      )
    `;
  }
  return { sql: client, ready: schemaPromise };
};

export const findEmailConnection = async (userId: string, workspaceId: string) => {
  const { sql, ready } = database();
  await ready;
  const rows = await sql<EmailConnectionRow[]>`
    select * from email_connections
    where user_id = ${userId} and workspace_id = ${workspaceId} and provider = 'gmail'
    limit 1
  `;
  return rows[0] || null;
};

type ConnectionInput = {
  userId: string;
  workspaceId: string;
  authMethod: "oauth" | "app-password";
  senderEmail: string;
  senderName: string;
  replyToEmail?: string;
  documentInboxEmail?: string;
  notificationEmail?: string;
  sendDocuments?: boolean;
  receiveCopies?: boolean;
  notificationsEnabled?: boolean;
  encryptedCredential: string;
};

export const upsertEmailConnection = async (input: ConnectionInput) => {
  const { sql, ready } = database();
  await ready;
  const id = `email_${randomUUID()}`;
  const rows = await sql<EmailConnectionRow[]>`
    insert into email_connections (
      id, user_id, workspace_id, provider, auth_method, sender_email, sender_name,
      reply_to_email, document_inbox_email, notification_email, send_documents,
      receive_copies, notifications_enabled, status, encrypted_credential, updated_at
    ) values (
      ${id}, ${input.userId}, ${input.workspaceId}, 'gmail', ${input.authMethod},
      ${input.senderEmail}, ${input.senderName}, ${input.replyToEmail || null},
      ${input.documentInboxEmail || null}, ${input.notificationEmail || null},
      ${input.sendDocuments ?? true}, ${input.receiveCopies ?? false},
      ${input.notificationsEnabled ?? true}, 'connected', ${input.encryptedCredential}, now()
    )
    on conflict (user_id, workspace_id, provider) do update set
      auth_method = excluded.auth_method,
      sender_email = excluded.sender_email,
      sender_name = excluded.sender_name,
      reply_to_email = excluded.reply_to_email,
      document_inbox_email = excluded.document_inbox_email,
      notification_email = excluded.notification_email,
      send_documents = excluded.send_documents,
      receive_copies = excluded.receive_copies,
      notifications_enabled = excluded.notifications_enabled,
      status = 'connected',
      encrypted_credential = excluded.encrypted_credential,
      last_error = null,
      updated_at = now()
    returning *
  `;
  return rows[0];
};

type PreferencesInput = {
  senderName: string;
  replyToEmail?: string;
  documentInboxEmail?: string;
  notificationEmail?: string;
  sendDocuments?: boolean;
  receiveCopies?: boolean;
  notificationsEnabled?: boolean;
};

export const updateEmailPreferences = async (userId: string, workspaceId: string, input: PreferencesInput) => {
  const { sql, ready } = database();
  await ready;
  const rows = await sql<EmailConnectionRow[]>`
    update email_connections set
      sender_name = ${input.senderName},
      reply_to_email = ${input.replyToEmail || null},
      document_inbox_email = ${input.documentInboxEmail || null},
      notification_email = ${input.notificationEmail || null},
      send_documents = ${input.sendDocuments ?? true},
      receive_copies = ${input.receiveCopies ?? false},
      notifications_enabled = ${input.notificationsEnabled ?? true},
      updated_at = now()
    where user_id = ${userId} and workspace_id = ${workspaceId} and provider = 'gmail'
    returning *
  `;
  return rows[0] || null;
};

export const disconnectEmailConnection = async (userId: string, workspaceId: string) => {
  const { sql, ready } = database();
  await ready;
  const rows = await sql<EmailConnectionRow[]>`
    update email_connections set status = 'disconnected', encrypted_credential = '', updated_at = now()
    where user_id = ${userId} and workspace_id = ${workspaceId} and provider = 'gmail'
    returning *
  `;
  return rows[0] || null;
};

export const recordConnectionTest = async (id: string, error?: string) => {
  const { sql, ready } = database();
  await ready;
  await sql`
    update email_connections set
      last_tested_at = now(),
      last_error = ${error || null},
      status = ${error ? "error" : "connected"},
      updated_at = now()
    where id = ${id}
  `;
};
