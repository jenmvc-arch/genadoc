# ZhiReady HR Doc Generator

## Supabase setup

1. Create a Supabase project and run [`docs/supabase-schema.sql`](docs/supabase-schema.sql) in the SQL editor.
2. Copy the Supabase transaction-pooler Postgres URL into `DATABASE_URL`.
3. Configure the variables documented in `.env.example` in Vercel.
4. Create users through Supabase Auth. The database trigger creates each user's first workspace and Admin membership.
5. Redeploy, sign in, and save once. Existing browser data is migrated to `workspace_states` without deleting the local backup.

Use [`docs/SUPABASE_PERSISTENCE.md`](docs/SUPABASE_PERSISTENCE.md) for the persistence coverage matrix, deployment checks, RLS/Storage tenant-isolation tests, and verification SQL.

The application uses a complete revisioned JSONB workspace snapshot for lossless restoration and synchronizes the searchable Clause, Placeholder, Layout, Template and Document records into normalized tables. Browser storage is an offline cache, not the authoritative saved state.

`ALLOW_DEMO_AUTH=true` enables blank demo login for local or controlled staging environments. Do not enable it for production.
