# Supabase persistence audit

Audit date: September 16, 2026.

## Persistence coverage

The cloud snapshot in `workspace_states.payload` stores the complete application state:

- document body, rich HTML, section and canvas ordering;
- document status, workflow stage, verification watermark and current values;
- first-page and subsequent-page Letterhead settings and uploaded data;
- Onboarding source snapshots, pending source changes and manual overrides;
- document versions, template versions and approval state;
- Clause Library, Placeholder groups and Layout Library;
- schema-ready Dropdown options and versioned conditional rules;
- Document Library records and generated export contents;
- navigation, theme and editor preferences.

The same save transaction synchronizes searchable records into normalized tables. This makes dependencies, library lists, permissions and reporting queryable without sacrificing fields that are still represented only in the current frontend state model.

## Save guarantees

- `revision` is an optimistic lock. A stale browser receives HTTP `409 STATE_CONFLICT` instead of overwriting a newer save.
- Manual Save only reports success after Postgres confirms the transaction.
- Autosave is debounced and retains the local cache if the network fails.
- The original `hr-doc-generator-state-v2` and `hr-doc-generator-app-store-v1` keys are retained for migration recovery.
- Reviewer saves cannot change templates, clauses, placeholders or layouts.
- Every state save records the actor, revision, payload size and time in `audit_logs`.
- Submit, approve and restore actions only update the UI after the Supabase transaction succeeds.
- Approved document content and approved version rows are immutable. Editing requires restoring a new draft.
- Published Template, Clause and Layout versions are retained in their version tables.
- Document exports retain their document/version relationship and private Storage asset path.

## Security model

- Supabase Auth verifies the email/password on the server.
- The server exchanges a verified user for an HttpOnly signed `hrdoc_session` cookie.
- Every API request verifies workspace membership from Postgres.
- RLS is enabled on tenant tables for future direct Supabase clients as defense in depth.
- Gmail App Passwords and OAuth refresh tokens remain encrypted server-side and are never returned by an API.
- The Supabase Storage bucket is private and paths are scoped by workspace ID.

## Deployment variables

See `.env.example`. Production requires:

- `DATABASE_URL`
- `AUTH_SESSION_SECRET`
- `SUPABASE_URL`
- `SUPABASE_PUBLISHABLE_KEY`
- `SUPABASE_SECRET_KEY`
- `EMAIL_CREDENTIAL_ENCRYPTION_KEY`

Google OAuth additionally requires the three `GOOGLE_*` variables. `ALLOW_DEMO_AUTH` must remain false in production.

## File storage

Letterhead files and generated Word/PDF files use short-lived signed upload URLs and are written directly to the private `hr-document-assets` Supabase Storage bucket. The workspace JSON stores only the private storage path. Downloads pass through a membership-checked API that creates a one-minute signed URL. This keeps large files out of the revisioned JSON payload and prevents one workspace from reading another workspace's objects.

The lossless snapshot endpoint rejects JSON payloads over 4 MB before a hosting platform can truncate the request. Binary files do not count toward this limit once uploaded to Storage.

## Apply and verify

1. Run `docs/supabase-schema.sql` in the Supabase SQL editor. It is safe to rerun after a partial deployment.
2. Configure all production variables from `.env.example`, keep `ALLOW_DEMO_AUTH=false`, and redeploy.
3. Create a Supabase Auth test user and confirm the Auth trigger creates one `profiles` row, one `workspaces` row and an Admin `workspace_members` row.
4. Sign in, change a document field, save, reload and confirm the latest `workspace_states.revision` and `payload` restore the change.
5. Save one Clause, Placeholder group and Layout. Confirm their normalized rows and version rows use the same `workspace_id`.
6. Submit and approve a test document. Confirm `approval_events`, `document_versions` and `documents.approved_version_id` are linked.
7. Export Word and PDF. Confirm `workspace_assets.status='ready'`, the bucket object is private, and `document_exports.storage_path` points to the asset.
8. Open a second browser with the same revision, save from the first browser, then save from the second. The second request must return `409 STATE_CONFLICT`.
9. Test a second workspace. Cross-workspace table reads and asset downloads must be denied.

Useful verification query:

```sql
select workspace_id, schema_version, revision, payload_bytes, updated_at, updated_by
from public.workspace_states
order by updated_at desc;

select d.id, d.document_status, d.current_version, d.approved_version_id,
       count(distinct v.id) as version_count,
       count(distinct e.id) as export_count
from public.documents d
left join public.document_versions v on v.document_id = d.id
left join public.document_exports e on e.document_id = d.id
group by d.id
order by max(d.updated_at) desc;
```

## Verification limit in this checkout

As of September 16, 2026, this checkout has no `DATABASE_URL` or Supabase project credentials in `.env.local`. The production build and local recovery workflow can be verified here, but applying the schema, exercising RLS against two real tenants, and uploading to the private bucket require the target Supabase credentials.

The current Dropdown/Rules administrator cards are still read-only demonstrations. Their selected values and resulting document content are preserved in document snapshots; the new normalized Dropdown and conditional-rule tables are ready for the future editable configuration UI, but no user-authored rule definitions can be end-to-end tested until that UI exists.
