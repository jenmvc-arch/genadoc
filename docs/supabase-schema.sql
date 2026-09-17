-- HR Doc Generator / ZhiReady
-- Supabase Postgres schema (idempotent)
-- Apply this file in the Supabase SQL editor before enabling cloud persistence.

create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.profiles (
  user_id text primary key,
  email text,
  display_name text not null default '',
  avatar_url text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.workspaces (
  id text primary key,
  name text not null,
  slug text not null unique,
  status text not null default 'active' check (status in ('active', 'suspended', 'archived')),
  settings jsonb not null default '{}'::jsonb,
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.workspace_members (
  id text primary key default ('member_' || replace(gen_random_uuid()::text, '-', '')),
  workspace_id text not null references public.workspaces(id) on delete cascade,
  user_id text not null references public.profiles(user_id) on delete cascade,
  role text not null check (role in ('Admin', 'Editor', 'Reviewer')),
  status text not null default 'active' check (status in ('invited', 'active', 'disabled')),
  invited_by text references public.profiles(user_id) on delete set null,
  joined_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, user_id)
);

create table if not exists public.companies (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  name text not null,
  country_code text,
  registration_number text,
  address jsonb not null default '{}'::jsonb,
  contact_details jsonb not null default '{}'::jsonb,
  status text not null default 'active' check (status in ('active', 'inactive')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, name)
);

create table if not exists public.user_preferences (
  workspace_id text not null references public.workspaces(id) on delete cascade,
  user_id text not null references public.profiles(user_id) on delete cascade,
  theme text not null default 'light' check (theme in ('light', 'dark')),
  navigation jsonb not null default '{}'::jsonb,
  editor_preferences jsonb not null default '{}'::jsonb,
  notification_preferences jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  primary key (workspace_id, user_id)
);

-- Lossless canonical snapshot. The normalized tables below support search,
-- reporting and future APIs; this row guarantees that every current editor
-- field can be restored even while the frontend domain model evolves.
create table if not exists public.workspace_states (
  workspace_id text primary key references public.workspaces(id) on delete cascade,
  schema_version integer not null default 2 check (schema_version > 0),
  revision bigint not null default 1 check (revision > 0),
  payload jsonb not null default '{}'::jsonb,
  payload_bytes integer not null default 0 check (payload_bytes >= 0),
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.templates (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  company_id text references public.companies(id) on delete set null,
  name text not null,
  document_type text not null,
  country_code text,
  language text not null default 'en',
  status text not null default 'draft' check (status in ('draft', 'published', 'inactive')),
  current_version integer not null default 1,
  stable_key text not null,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, stable_key)
);

create table if not exists public.template_versions (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  template_id text not null references public.templates(id) on delete cascade,
  version integer not null,
  definition jsonb not null,
  change_summary text,
  status text not null default 'draft' check (status in ('draft', 'published', 'superseded')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  unique (template_id, version)
);

create table if not exists public.clauses (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  title text not null,
  stable_key text not null,
  structure text not null default 'flat' check (structure in ('flat', 'nested')),
  category text,
  document_types text[] not null default '{}',
  language text not null default 'en',
  tags text[] not null default '{}',
  status text not null default 'draft' check (status in ('draft', 'published', 'inactive')),
  current_version integer not null default 1,
  definition jsonb not null default '{}'::jsonb,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, stable_key)
);

create table if not exists public.clause_versions (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  clause_id text not null references public.clauses(id) on delete cascade,
  version integer not null,
  definition jsonb not null,
  change_summary text,
  status text not null default 'draft' check (status in ('draft', 'published', 'superseded')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  unique (clause_id, version)
);

create or replace function public.protect_published_library_version()
returns trigger
language plpgsql
as $$
begin
  if old.status in ('published', 'superseded') then
    raise exception 'Published library versions are immutable';
  end if;
  if tg_op = 'DELETE' then
    return old;
  end if;
  return new;
end;
$$;

drop trigger if exists protect_template_version_update on public.template_versions;
create trigger protect_template_version_update
before update or delete on public.template_versions
for each row execute function public.protect_published_library_version();

drop trigger if exists protect_clause_version_update on public.clause_versions;
create trigger protect_clause_version_update
before update or delete on public.clause_versions
for each row execute function public.protect_published_library_version();

create table if not exists public.data_sources (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  name text not null,
  source_type text not null check (source_type in ('onboarding', 'system-table', 'csv', 'excel', 'api', 'webhook', 'manual')),
  status text not null default 'active' check (status in ('active', 'error', 'inactive')),
  connection_config jsonb not null default '{}'::jsonb,
  schema_snapshot jsonb not null default '{}'::jsonb,
  last_synced_at timestamptz,
  last_error text,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.placeholder_groups (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  data_source_id text references public.data_sources(id) on delete set null,
  name text not null,
  stable_key text not null,
  source_type text not null,
  source_table text,
  status text not null default 'active' check (status in ('active', 'inactive')),
  sort_order integer not null default 0,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, stable_key)
);

create table if not exists public.placeholder_fields (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  group_id text not null references public.placeholder_groups(id) on delete cascade,
  stable_key text not null,
  label text not null,
  field_type text not null check (field_type in ('text', 'multiline', 'number', 'currency', 'date', 'select', 'multiselect', 'repeating')),
  source_field_id text,
  description text,
  example_value text,
  default_value jsonb,
  validation jsonb not null default '{}'::jsonb,
  display_format jsonb not null default '{}'::jsonb,
  permissions jsonb not null default '{}'::jsonb,
  required boolean not null default false,
  manual_override boolean not null default true,
  mapping_status text not null default 'valid' check (mapping_status in ('valid', 'invalid', 'unmapped')),
  status text not null default 'active' check (status in ('active', 'inactive')),
  sort_order integer not null default 0,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (group_id, stable_key)
);

create table if not exists public.onboarding_people (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  external_person_id text,
  display_name text not null,
  email text,
  status text not null default 'active' check (status in ('active', 'inactive', 'archived')),
  private_metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, external_person_id)
);

create table if not exists public.onboarding_submissions (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  person_id text not null references public.onboarding_people(id) on delete cascade,
  data_source_id text references public.data_sources(id) on delete set null,
  external_submission_id text,
  submitted_at timestamptz,
  synced_at timestamptz,
  source_revision text,
  values jsonb not null default '{}'::jsonb,
  validation_errors jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, external_submission_id)
);

create table if not exists public.field_mappings (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  data_source_id text not null references public.data_sources(id) on delete cascade,
  placeholder_field_id text not null references public.placeholder_fields(id) on delete cascade,
  source_field_id text not null,
  transform jsonb not null default '{}'::jsonb,
  status text not null default 'active' check (status in ('active', 'invalid', 'inactive')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (data_source_id, placeholder_field_id)
);

create table if not exists public.dropdown_fields (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  name text not null,
  stable_key text not null,
  selection_type text not null default 'single' check (selection_type in ('single', 'multiple')),
  allow_custom_value boolean not null default false,
  prompt_text text,
  default_value jsonb not null default 'null'::jsonb,
  required boolean not null default false,
  scope text not null default 'document' check (scope in ('person', 'document', 'clause', 'section')),
  status text not null default 'draft' check (status in ('draft', 'published', 'inactive')),
  sort_order integer not null default 0,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, stable_key)
);

create table if not exists public.dropdown_options (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  dropdown_field_id text not null references public.dropdown_fields(id) on delete cascade,
  stable_key text not null,
  label text not null,
  value text not null,
  content_variant jsonb not null default '{}'::jsonb,
  status text not null default 'active' check (status in ('active', 'inactive')),
  sort_order integer not null default 0,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (dropdown_field_id, stable_key),
  unique (dropdown_field_id, value)
);

create table if not exists public.condition_rules (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  name text not null,
  stable_key text not null,
  definition jsonb not null default '{}'::jsonb,
  status text not null default 'draft' check (status in ('draft', 'published', 'inactive')),
  current_version integer not null default 1,
  priority integer not null default 0,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, stable_key)
);

create table if not exists public.condition_rule_versions (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  condition_rule_id text not null references public.condition_rules(id) on delete cascade,
  version integer not null,
  definition jsonb not null,
  change_summary text,
  status text not null default 'draft' check (status in ('draft', 'published', 'superseded')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  unique (condition_rule_id, version)
);

drop trigger if exists protect_condition_rule_version_update on public.condition_rule_versions;
create trigger protect_condition_rule_version_update
before update or delete on public.condition_rule_versions
for each row execute function public.protect_published_library_version();

create table if not exists public.layouts (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  company_id text references public.companies(id) on delete set null,
  name text not null,
  stable_key text not null,
  status text not null default 'draft' check (status in ('draft', 'published', 'inactive')),
  current_version integer not null default 1,
  definition jsonb not null default '{}'::jsonb,
  is_company_default boolean not null default false,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (workspace_id, stable_key)
);

create table if not exists public.workspace_assets (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  kind text not null check (kind in ('letterhead', 'export', 'document-image')),
  storage_path text not null unique,
  file_name text not null,
  content_type text not null,
  byte_size bigint not null check (byte_size > 0),
  checksum_sha256 text,
  status text not null default 'pending' check (status in ('pending', 'ready', 'failed', 'deleted')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.layout_versions (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  layout_id text not null references public.layouts(id) on delete cascade,
  version integer not null,
  definition jsonb not null,
  change_summary text,
  status text not null default 'draft' check (status in ('draft', 'published', 'superseded')),
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  unique (layout_id, version)
);

drop trigger if exists protect_layout_version_update on public.layout_versions;
create trigger protect_layout_version_update
before update or delete on public.layout_versions
for each row execute function public.protect_published_library_version();

create table if not exists public.documents (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  company_id text references public.companies(id) on delete set null,
  template_id text references public.templates(id) on delete set null,
  person_id text references public.onboarding_people(id) on delete set null,
  title text not null,
  document_type text not null,
  save_status text not null default 'saved' check (save_status in ('saving', 'saved', 'failed')),
  document_status text not null default 'draft' check (document_status in ('draft', 'in-review', 'approved', 'void')),
  generation_status text not null default 'not-generated' check (generation_status in ('not-generated', 'generating', 'generated', 'failed')),
  current_version integer not null default 1,
  current_snapshot jsonb not null default '{}'::jsonb,
  approved_version_id text,
  created_by text not null references public.profiles(user_id) on delete restrict,
  updated_by text not null references public.profiles(user_id) on delete restrict,
  archived_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.document_versions (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  document_id text not null references public.documents(id) on delete cascade,
  version integer not null,
  status text not null check (status in ('draft', 'in-review', 'approved', 'void')),
  snapshot jsonb not null,
  source_snapshot jsonb not null default '{}'::jsonb,
  change_summary text,
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now(),
  unique (document_id, version)
);

create or replace function public.protect_approved_document_version()
returns trigger
language plpgsql
as $$
begin
  if old.status = 'approved' then
    raise exception 'Approved document versions are immutable';
  end if;
  if tg_op = 'DELETE' then
    return old;
  end if;
  return new;
end;
$$;

drop trigger if exists protect_approved_document_version_update on public.document_versions;
create trigger protect_approved_document_version_update
before update or delete on public.document_versions
for each row execute function public.protect_approved_document_version();

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'documents_approved_version_fk'
      and conrelid = 'public.documents'::regclass
  ) then
    alter table public.documents
      add constraint documents_approved_version_fk
      foreign key (approved_version_id)
      references public.document_versions(id)
      on delete restrict
      not valid;
  end if;
end;
$$;

create or replace function public.protect_approved_document_snapshot()
returns trigger
language plpgsql
as $$
begin
  if old.document_status = 'approved'
     and new.document_status = 'approved'
     and (
       new.company_id is distinct from old.company_id
       or new.template_id is distinct from old.template_id
       or new.person_id is distinct from old.person_id
       or new.title is distinct from old.title
       or new.document_type is distinct from old.document_type
       or (new.current_snapshot - array['status', 'generationStatus', 'exports', 'updatedAt']::text[])
          is distinct from
          (old.current_snapshot - array['status', 'generationStatus', 'exports', 'updatedAt']::text[])
     ) then
    raise exception 'Approved documents are immutable; restore the document as a draft before editing';
  end if;
  return new;
end;
$$;

drop trigger if exists protect_approved_document_snapshot_update on public.documents;
create trigger protect_approved_document_snapshot_update
before update on public.documents
for each row execute function public.protect_approved_document_snapshot();

create table if not exists public.document_exports (
  id text primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  document_id text not null references public.documents(id) on delete cascade,
  document_version_id text references public.document_versions(id) on delete set null,
  format text not null check (format in ('docx', 'pdf')),
  status text not null check (status in ('generating', 'generated', 'failed')),
  file_name text not null,
  storage_path text,
  byte_size bigint,
  checksum_sha256 text,
  error text,
  created_by text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now()
);

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'document_exports_storage_path_fk'
      and conrelid = 'public.document_exports'::regclass
  ) then
    alter table public.document_exports
      add constraint document_exports_storage_path_fk
      foreign key (storage_path)
      references public.workspace_assets(storage_path)
      on delete set null
      not valid;
  end if;
end;
$$;

create table if not exists public.approval_events (
  id text primary key default ('approval_' || replace(gen_random_uuid()::text, '-', '')),
  workspace_id text not null references public.workspaces(id) on delete cascade,
  document_id text not null references public.documents(id) on delete cascade,
  document_version_id text references public.document_versions(id) on delete set null,
  action text not null check (action in ('submitted', 'approved', 'rejected', 'voided', 'restored')),
  comment text,
  actor_user_id text not null references public.profiles(user_id) on delete restrict,
  created_at timestamptz not null default now()
);

create table if not exists public.audit_logs (
  id bigint generated always as identity primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  actor_user_id text not null references public.profiles(user_id) on delete restrict,
  action text not null,
  entity_type text not null,
  entity_id text,
  before_data jsonb,
  after_data jsonb,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.email_connections (
  id text primary key,
  user_id text not null references public.profiles(user_id) on delete cascade,
  workspace_id text not null references public.workspaces(id) on delete cascade,
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
);

-- Backfill users created before this migration and preserve any existing
-- email-connection workspace IDs from the earlier schema.
insert into public.profiles (user_id, email, display_name)
select
  user_id,
  max(sender_email),
  max(sender_name)
from public.email_connections
group by user_id
on conflict (user_id) do nothing;

insert into public.profiles (user_id, email, display_name)
select
  id::text,
  email,
  coalesce(raw_user_meta_data ->> 'full_name', split_part(coalesce(email, 'HR user'), '@', 1))
from auth.users
on conflict (user_id) do update set email = excluded.email;

insert into public.workspaces (id, name, slug, created_by)
select
  connection.workspace_id,
  'Imported HR workspace',
  'imported-' || substr(md5(connection.workspace_id), 1, 20),
  min(connection.user_id)
from public.email_connections connection
where not exists (select 1 from public.workspaces workspace where workspace.id = connection.workspace_id)
group by connection.workspace_id;

insert into public.workspace_members (workspace_id, user_id, role, status, joined_at)
select distinct workspace_id, user_id, 'Admin', 'active', now()
from public.email_connections
on conflict (workspace_id, user_id) do nothing;

do $$
declare
  auth_user record;
  workspace_key text;
begin
  for auth_user in
    select profile.user_id, profile.display_name
    from public.profiles profile
    where exists (select 1 from auth.users users where users.id::text = profile.user_id)
      and not exists (
        select 1 from public.workspace_members member
        where member.user_id = profile.user_id and member.status = 'active'
      )
  loop
    workspace_key := 'ws_' || replace(gen_random_uuid()::text, '-', '');
    insert into public.workspaces (id, name, slug, created_by)
    values (
      workspace_key,
      coalesce(nullif(auth_user.display_name, ''), 'HR user') || '''s workspace',
      'workspace-' || replace(gen_random_uuid()::text, '-', ''),
      auth_user.user_id
    );
    insert into public.workspace_members (workspace_id, user_id, role, status, joined_at)
    values (workspace_key, auth_user.user_id, 'Admin', 'active', now());
  end loop;
end;
$$;

do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conname = 'email_connections_workspace_id_fkey'
      and conrelid = 'public.email_connections'::regclass
  ) then
    alter table public.email_connections
      add constraint email_connections_workspace_id_fkey
      foreign key (workspace_id) references public.workspaces(id) on delete cascade;
  end if;
  if not exists (
    select 1 from pg_constraint
    where conname = 'email_connections_user_id_fkey'
      and conrelid = 'public.email_connections'::regclass
  ) then
    alter table public.email_connections
      add constraint email_connections_user_id_fkey
      foreign key (user_id) references public.profiles(user_id) on delete cascade;
  end if;
end;
$$;

create index if not exists workspace_members_user_idx on public.workspace_members(user_id, status);
create index if not exists companies_workspace_idx on public.companies(workspace_id, status);
create index if not exists templates_workspace_idx on public.templates(workspace_id, status, document_type);
create index if not exists clauses_workspace_idx on public.clauses(workspace_id, status, category);
create index if not exists placeholder_groups_workspace_idx on public.placeholder_groups(workspace_id, status);
create index if not exists placeholder_fields_group_idx on public.placeholder_fields(group_id, status, sort_order);
create index if not exists dropdown_fields_workspace_idx on public.dropdown_fields(workspace_id, status, sort_order);
create index if not exists dropdown_options_field_idx on public.dropdown_options(dropdown_field_id, status, sort_order);
create index if not exists condition_rules_workspace_idx on public.condition_rules(workspace_id, status, priority);
create index if not exists onboarding_people_workspace_idx on public.onboarding_people(workspace_id, status);
create index if not exists onboarding_submissions_person_idx on public.onboarding_submissions(workspace_id, person_id, submitted_at desc);
create index if not exists layouts_workspace_idx on public.layouts(workspace_id, status, company_id);
create index if not exists workspace_assets_workspace_idx on public.workspace_assets(workspace_id, kind, status, created_at desc);
create index if not exists documents_workspace_idx on public.documents(workspace_id, document_status, generation_status, updated_at desc);
create index if not exists document_versions_document_idx on public.document_versions(document_id, version desc);
create index if not exists document_exports_document_idx on public.document_exports(document_id, created_at desc);
create index if not exists audit_logs_workspace_idx on public.audit_logs(workspace_id, created_at desc);
create index if not exists email_connections_workspace_idx on public.email_connections(workspace_id, status);

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'profiles', 'workspaces', 'workspace_members', 'companies', 'user_preferences',
    'workspace_states', 'templates', 'template_versions', 'clauses', 'clause_versions',
    'data_sources', 'placeholder_groups', 'placeholder_fields', 'onboarding_people',
    'onboarding_submissions', 'field_mappings', 'dropdown_fields', 'dropdown_options',
    'condition_rules', 'condition_rule_versions', 'layouts', 'layout_versions',
    'workspace_assets', 'documents', 'document_versions', 'document_exports', 'approval_events',
    'audit_logs', 'email_connections'
  ] loop
    execute format('drop trigger if exists %I_set_updated_at on public.%I', table_name, table_name);
    if table_name not in (
      'template_versions', 'clause_versions', 'condition_rule_versions', 'layout_versions', 'document_versions',
      'document_exports', 'approval_events', 'audit_logs'
    ) then
      execute format(
        'create trigger %I_set_updated_at before update on public.%I for each row execute function public.set_updated_at()',
        table_name,
        table_name
      );
    end if;
  end loop;
end;
$$;

create or replace function public.current_workspace_role(target_workspace_id text)
returns text
language sql
stable
security definer
set search_path = public
as $$
  select wm.role
  from public.workspace_members wm
  where wm.workspace_id = target_workspace_id
    and wm.user_id = auth.uid()::text
    and wm.status = 'active'
  limit 1
$$;

create or replace function public.is_workspace_member(target_workspace_id text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_workspace_role(target_workspace_id) is not null
$$;

create or replace function public.is_workspace_admin(target_workspace_id text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_workspace_role(target_workspace_id) = 'Admin'
$$;

create or replace function public.can_edit_workspace(target_workspace_id text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_workspace_role(target_workspace_id) in ('Admin', 'Editor')
$$;

grant execute on function public.current_workspace_role(text) to authenticated;
grant execute on function public.is_workspace_member(text) to authenticated;
grant execute on function public.is_workspace_admin(text) to authenticated;
grant execute on function public.can_edit_workspace(text) to authenticated;

alter table public.profiles enable row level security;
alter table public.workspaces enable row level security;
alter table public.workspace_members enable row level security;

drop policy if exists profiles_self_select on public.profiles;
create policy profiles_self_select on public.profiles for select to authenticated using (user_id = auth.uid()::text);
drop policy if exists profiles_self_update on public.profiles;
create policy profiles_self_update on public.profiles for update to authenticated using (user_id = auth.uid()::text) with check (user_id = auth.uid()::text);

drop policy if exists workspaces_member_select on public.workspaces;
create policy workspaces_member_select on public.workspaces for select to authenticated using (public.is_workspace_member(id));
drop policy if exists workspaces_admin_update on public.workspaces;
create policy workspaces_admin_update on public.workspaces for update to authenticated using (public.is_workspace_admin(id)) with check (public.is_workspace_admin(id));

drop policy if exists workspace_members_member_select on public.workspace_members;
create policy workspace_members_member_select on public.workspace_members for select to authenticated using (public.is_workspace_member(workspace_id));
drop policy if exists workspace_members_admin_write on public.workspace_members;
create policy workspace_members_admin_write on public.workspace_members for all to authenticated using (public.is_workspace_admin(workspace_id)) with check (public.is_workspace_admin(workspace_id));

alter table public.user_preferences enable row level security;
drop policy if exists user_preferences_own on public.user_preferences;
create policy user_preferences_own on public.user_preferences for all to authenticated
  using (user_id = auth.uid()::text and public.is_workspace_member(workspace_id))
  with check (user_id = auth.uid()::text and public.is_workspace_member(workspace_id));

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'companies', 'templates', 'template_versions', 'clauses', 'clause_versions',
    'data_sources', 'placeholder_groups', 'placeholder_fields', 'field_mappings',
    'dropdown_fields', 'dropdown_options', 'condition_rules', 'condition_rule_versions',
    'layouts', 'layout_versions'
  ] loop
    execute format('alter table public.%I enable row level security', table_name);
    execute format('drop policy if exists %I_member_select on public.%I', table_name, table_name);
    execute format(
      'create policy %I_member_select on public.%I for select to authenticated using (public.is_workspace_member(workspace_id))',
      table_name,
      table_name
    );
    execute format('drop policy if exists %I_admin_write on public.%I', table_name, table_name);
    execute format(
      'create policy %I_admin_write on public.%I for all to authenticated using (public.is_workspace_admin(workspace_id)) with check (public.is_workspace_admin(workspace_id))',
      table_name,
      table_name
    );
  end loop;
end;
$$;

alter table public.workspace_assets enable row level security;
drop policy if exists workspace_assets_member_select on public.workspace_assets;
create policy workspace_assets_member_select on public.workspace_assets for select to authenticated using (public.is_workspace_member(workspace_id));
drop policy if exists workspace_assets_editor_write on public.workspace_assets;
create policy workspace_assets_editor_write on public.workspace_assets for all to authenticated
  using (public.can_edit_workspace(workspace_id))
  with check (public.can_edit_workspace(workspace_id));

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'onboarding_people', 'onboarding_submissions', 'documents', 'document_versions',
    'document_exports'
  ] loop
    execute format('alter table public.%I enable row level security', table_name);
    execute format('drop policy if exists %I_member_select on public.%I', table_name, table_name);
    execute format(
      'create policy %I_member_select on public.%I for select to authenticated using (public.is_workspace_member(workspace_id))',
      table_name,
      table_name
    );
    execute format('drop policy if exists %I_editor_write on public.%I', table_name, table_name);
    execute format(
      'create policy %I_editor_write on public.%I for all to authenticated using (public.can_edit_workspace(workspace_id)) with check (public.can_edit_workspace(workspace_id))',
      table_name,
      table_name
    );
  end loop;
end;
$$;

alter table public.workspace_states enable row level security;
drop policy if exists workspace_states_member_select on public.workspace_states;
create policy workspace_states_member_select on public.workspace_states for select to authenticated using (public.is_workspace_member(workspace_id));
drop policy if exists workspace_states_editor_write on public.workspace_states;
create policy workspace_states_editor_write on public.workspace_states for all to authenticated using (public.can_edit_workspace(workspace_id)) with check (public.can_edit_workspace(workspace_id));

alter table public.approval_events enable row level security;
drop policy if exists approval_events_member_select on public.approval_events;
create policy approval_events_member_select on public.approval_events for select to authenticated using (public.is_workspace_member(workspace_id));
drop policy if exists approval_events_reviewer_insert on public.approval_events;
create policy approval_events_reviewer_insert on public.approval_events for insert to authenticated
  with check (public.current_workspace_role(workspace_id) in ('Admin', 'Reviewer'));

alter table public.audit_logs enable row level security;
drop policy if exists audit_logs_admin_select on public.audit_logs;
create policy audit_logs_admin_select on public.audit_logs for select to authenticated using (public.is_workspace_admin(workspace_id));

alter table public.email_connections enable row level security;
drop policy if exists email_connections_member_select on public.email_connections;
create policy email_connections_member_select on public.email_connections for select to authenticated using (public.is_workspace_member(workspace_id));
drop policy if exists email_connections_owner_write on public.email_connections;
create policy email_connections_owner_write on public.email_connections for all to authenticated
  using (user_id = auth.uid()::text and public.can_edit_workspace(workspace_id))
  with check (user_id = auth.uid()::text and public.can_edit_workspace(workspace_id));

-- Create a private bucket for letterheads and generated files. Object paths
-- must begin with the workspace ID: <workspace-id>/<entity>/<file>.
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
  'hr-document-assets',
  'hr-document-assets',
  false,
  52428800,
  array['image/png', 'image/jpeg', 'application/pdf', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']
)
on conflict (id) do update set
  public = excluded.public,
  file_size_limit = excluded.file_size_limit,
  allowed_mime_types = excluded.allowed_mime_types;

drop policy if exists hr_document_assets_member_read on storage.objects;
create policy hr_document_assets_member_read on storage.objects for select to authenticated
  using (bucket_id = 'hr-document-assets' and public.is_workspace_member((storage.foldername(name))[1]));
drop policy if exists hr_document_assets_editor_insert on storage.objects;
create policy hr_document_assets_editor_insert on storage.objects for insert to authenticated
  with check (bucket_id = 'hr-document-assets' and public.can_edit_workspace((storage.foldername(name))[1]));
drop policy if exists hr_document_assets_editor_update on storage.objects;
create policy hr_document_assets_editor_update on storage.objects for update to authenticated
  using (bucket_id = 'hr-document-assets' and public.can_edit_workspace((storage.foldername(name))[1]))
  with check (bucket_id = 'hr-document-assets' and public.can_edit_workspace((storage.foldername(name))[1]));
drop policy if exists hr_document_assets_editor_delete on storage.objects;
create policy hr_document_assets_editor_delete on storage.objects for delete to authenticated
  using (bucket_id = 'hr-document-assets' and public.can_edit_workspace((storage.foldername(name))[1]));

-- Supabase Auth bootstrap: each new user receives a private workspace. Teams
-- can later invite the user into other workspaces without changing stable IDs.
create or replace function public.handle_new_auth_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  workspace_key text := 'ws_' || replace(gen_random_uuid()::text, '-', '');
  display_label text := coalesce(new.raw_user_meta_data ->> 'full_name', split_part(coalesce(new.email, 'HR user'), '@', 1));
begin
  insert into public.profiles (user_id, email, display_name)
  values (new.id::text, new.email, display_label)
  on conflict (user_id) do update set email = excluded.email, display_name = excluded.display_name;

  insert into public.workspaces (id, name, slug, created_by)
  values (
    workspace_key,
    coalesce(nullif(new.raw_user_meta_data ->> 'workspace_name', ''), display_label || '''s workspace'),
    'workspace-' || replace(gen_random_uuid()::text, '-', ''),
    new.id::text
  );

  insert into public.workspace_members (workspace_id, user_id, role, status, joined_at)
  values (workspace_key, new.id::text, 'Admin', 'active', now());
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute function public.handle_new_auth_user();
