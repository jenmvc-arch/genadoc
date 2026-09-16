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
  status text not null default 'connected'
    check (status in ('connected', 'expired', 'error', 'disconnected')),
  encrypted_credential text not null,
  last_tested_at timestamptz,
  last_error text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (user_id, workspace_id, provider)
);

create index if not exists email_connections_workspace_idx
  on email_connections (workspace_id, status);
