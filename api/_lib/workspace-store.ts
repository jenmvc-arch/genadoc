import { randomUUID } from "node:crypto";
import type postgres from "postgres";
import { database } from "./database.js";
import { ApiError } from "./http.js";
import type { ServerSession } from "./auth.js";

export const currentWorkspaceSchemaVersion = 2;
export const maxWorkspacePayloadBytes = 4 * 1024 * 1024;

type WorkspaceStateRow = {
  workspace_id: string;
  schema_version: number;
  revision: number;
  payload: unknown;
  updated_at: string;
  updated_by: string;
};

type WorkspacePayload = {
  schemaVersion: number;
  appStore?: Record<string, unknown>;
  workspace?: Record<string, unknown>;
  preferences?: Record<string, unknown>;
};

type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };
const asJson = (value: unknown) => JSON.parse(JSON.stringify(value)) as JsonValue;

const qualifiedId = (workspaceId: string, value: unknown, fallback: string) =>
  `${workspaceId}:${String(value || fallback).slice(0, 180)}`;

const asRecord = (value: unknown): Record<string, unknown> =>
  value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};

const asArray = (value: unknown): Array<Record<string, unknown>> =>
  Array.isArray(value) ? value.filter((item) => item && typeof item === "object") as Array<Record<string, unknown>> : [];

export const requireWorkspaceAccess = async (session: ServerSession) => {
  const { sql, ready } = database();
  await ready;
  const rows = await sql<Array<{ role: ServerSession["role"] }>>`
    select role from workspace_members
    where workspace_id = ${session.workspaceId}
      and user_id = ${session.userId}
      and status = 'active'
    limit 1
  `;
  if (!rows[0]) throw new ApiError(403, "FORBIDDEN", "You do not have access to this workspace.");
  return rows[0].role;
};

export const ensureDemoWorkspace = async (session: ServerSession, displayName: string) => {
  const { sql, ready } = database();
  await ready;
  await sql.begin(async (transaction) => {
    await transaction`
      insert into profiles (user_id, email, display_name)
      values (${session.userId}, ${session.email || null}, ${displayName})
      on conflict (user_id) do update set
        email = excluded.email,
        display_name = excluded.display_name,
        updated_at = now()
    `;
    await transaction`
      insert into workspaces (id, name, slug, created_by)
      values (${session.workspaceId}, 'ZhiReady Demo Workspace', ${session.workspaceId}, ${session.userId})
      on conflict (id) do nothing
    `;
    await transaction`
      insert into workspace_members (workspace_id, user_id, role, status, joined_at)
      values (${session.workspaceId}, ${session.userId}, ${session.role}, 'active', now())
      on conflict (workspace_id, user_id) do update set role = excluded.role, status = 'active', updated_at = now()
    `;
  });
};

export const findFirstMembership = async (userId: string) => {
  const { sql, ready } = database();
  await ready;
  const rows = await sql<Array<{ workspace_id: string; role: ServerSession["role"]; workspace_name: string }>>`
    select wm.workspace_id, wm.role, w.name as workspace_name
    from workspace_members wm
    join workspaces w on w.id = wm.workspace_id
    where wm.user_id = ${userId} and wm.status = 'active' and w.status = 'active'
    order by case wm.role when 'Admin' then 1 when 'Editor' then 2 else 3 end, wm.created_at
    limit 1
  `;
  return rows[0] || null;
};

export const loadWorkspaceState = async (session: ServerSession) => {
  const role = await requireWorkspaceAccess(session);
  const { sql, ready } = database();
  await ready;
  const rows = await sql<WorkspaceStateRow[]>`
    select workspace_id, schema_version, revision, payload, updated_at, updated_by
    from workspace_states
    where workspace_id = ${session.workspaceId}
    limit 1
  `;
  const preferenceRows = await sql<Array<{ theme: string; navigation: unknown; editor_preferences: unknown; notification_preferences: unknown }>>`
    select theme, navigation, editor_preferences, notification_preferences
    from user_preferences
    where workspace_id = ${session.workspaceId} and user_id = ${session.userId}
    limit 1
  `;
  return { state: rows[0] || null, preferences: preferenceRows[0] || null, role };
};

const protectedConfiguration = (payload: WorkspacePayload) => {
  const appStore = asRecord(payload.appStore);
  const workspace = asRecord(payload.workspace);
  return JSON.stringify({
    clauses: appStore.clauses,
    placeholderGroups: appStore.placeholderGroups,
    dropdowns: appStore.dropdowns,
    conditionRules: appStore.conditionRules,
    layouts: appStore.layouts,
    templates: workspace.templates,
    templateVersions: workspace.templateVersions,
    templateLayouts: workspace.templateLayouts,
    customPlaceholders: workspace.customPlaceholders,
  });
};

const documentContentSignature = (payload: WorkspacePayload) => {
  const appStore = asRecord(payload.appStore);
  const workspace = asRecord(payload.workspace);
  const protectedWorkspace = { ...workspace };
  for (const key of ["docStatus", "workflowStage", "versions", "versionCounter", "lastSavedAt", "currentRole"]) {
    delete protectedWorkspace[key];
  }
  const protectedDocuments = asArray(appStore.documents).map((document) => {
    const copy = { ...document };
    for (const key of ["status", "generationStatus", "versions", "exports", "updatedAt"]) delete copy[key];
    return copy;
  });
  return JSON.stringify({ workspace: protectedWorkspace, documents: protectedDocuments });
};

const documentStatuses = (payload: WorkspacePayload) => {
  const appStore = asRecord(payload.appStore);
  return new Map(asArray(appStore.documents).map((document) => [String(document.id || ""), String(document.status || "draft")]));
};

const documentRecords = (payload: WorkspacePayload) => {
  const appStore = asRecord(payload.appStore);
  return new Map(asArray(appStore.documents).map((document) => [String(document.id || ""), document]));
};

const documentRecordContentSignature = (document: Record<string, unknown>) => {
  const copy = { ...document };
  for (const key of ["status", "generationStatus", "exports", "updatedAt"]) delete copy[key];
  return JSON.stringify(copy);
};

const documentVersionNumber = (version: Record<string, unknown>, fallback: number) => {
  const match = String(version.label || "").match(/^v(\d+)/i);
  const parsed = match ? Number(match[1]) : Number(version.version || fallback);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
};

const syncNormalizedState = async (
  transaction: postgres.TransactionSql<{}>,
  session: ServerSession,
  payload: WorkspacePayload,
) => {
  const appStore = asRecord(payload.appStore);
  const workspace = asRecord(payload.workspace);
  const companyId = qualifiedId(session.workspaceId, "northstar-labs-my", "company");
  const sourceId = qualifiedId(session.workspaceId, "onboarding-form", "data-source");
  const templatesByLocalId = new Map(
    asArray(workspace.templates).map((template) => [String(template.id || ""), template]),
  );
  const dataSourceStatus = workspace.connectionStatus === "Failed" ? "error" : workspace.connectionStatus === "Connected" ? "active" : "active";
  const dataSourceSnapshot = {
    source: "workspace_snapshot",
    connectionStatus: workspace.connectionStatus || "Connected",
    sourceUpdateMeta: asRecord(workspace.sourceUpdateMeta),
  };

  await transaction`
    insert into companies (id, workspace_id, name, country_code, status, created_by)
    values (${companyId}, ${session.workspaceId}, 'Northstar Labs', 'MY', 'active', ${session.userId})
    on conflict (id) do update set updated_at = now()
  `;
  await transaction`
    insert into data_sources (
      id, workspace_id, name, source_type, status, schema_snapshot, last_synced_at,
      last_error, created_by, updated_by
    ) values (
      ${sourceId}, ${session.workspaceId}, 'Onboarding Form', 'onboarding', ${dataSourceStatus},
      ${transaction.json(asJson(dataSourceSnapshot))}, ${workspace.lastSyncAt ? new Date().toISOString() : null},
      ${workspace.connectionError ? String(workspace.connectionError) : null}, ${session.userId}, ${session.userId}
    ) on conflict (id) do update set
      status = excluded.status,
      schema_snapshot = excluded.schema_snapshot,
      last_synced_at = excluded.last_synced_at,
      last_error = excluded.last_error,
      updated_by = excluded.updated_by,
      updated_at = now()
  `;

  // Reconcile removals without destroying version history. Items present in
  // the incoming snapshot are reactivated by the upserts below.
  await transaction`update templates set status = 'inactive' where workspace_id = ${session.workspaceId}`;
  await transaction`update clauses set status = 'inactive' where workspace_id = ${session.workspaceId}`;
  await transaction`update placeholder_groups set status = 'inactive' where workspace_id = ${session.workspaceId}`;
  await transaction`
    update placeholder_fields
    set status = 'inactive', updated_by = ${session.userId}, updated_at = now()
    where workspace_id = ${session.workspaceId}
  `;
  await transaction`
    update field_mappings
    set status = 'inactive', updated_by = ${session.userId}, updated_at = now()
    where workspace_id = ${session.workspaceId}
  `;
  await transaction`update layouts set status = 'inactive' where workspace_id = ${session.workspaceId}`;
  await transaction`
    update documents
    set archived_at = coalesce(archived_at, now())
    where workspace_id = ${session.workspaceId}
  `;

  for (const template of asArray(workspace.templates)) {
    const localId = String(template.id || randomUUID());
    const id = qualifiedId(session.workspaceId, localId, "template");
    const version = Math.max(1, Number(asRecord(workspace.templateVersions)[localId] || 1));
    const templateDefinition = {
      ...template,
      letterhead: asRecord(workspace.templateLayouts)[localId] || null,
    };
    await transaction`
      insert into templates (
        id, workspace_id, company_id, name, document_type, language, status,
        current_version, stable_key, created_by, updated_by
      ) values (
        ${id}, ${session.workspaceId}, ${companyId}, ${String(template.name || "Untitled template")},
        ${String(template.type || "HR document")}, 'en', 'published',
        ${version}, ${localId}, ${session.userId}, ${session.userId}
      ) on conflict (id) do update set
        name = excluded.name,
        document_type = excluded.document_type,
        status = excluded.status,
        current_version = excluded.current_version,
        updated_by = excluded.updated_by,
        updated_at = now()
    `;
    await transaction`
      insert into template_versions (
        id, workspace_id, template_id, version, definition, status, created_by
      ) values (
        ${qualifiedId(session.workspaceId, `${localId}:v${version}`, "template-version")},
        ${session.workspaceId}, ${id}, ${version}, ${transaction.json(asJson(templateDefinition))},
        'published', ${session.userId}
      ) on conflict (template_id, version) do nothing
    `;
  }

  for (const clause of asArray(appStore.clauses)) {
    const localId = String(clause.id || randomUUID());
    const id = qualifiedId(session.workspaceId, localId, "clause");
    const version = Math.max(1, Number(clause.version || 1));
    await transaction`
      insert into clauses (
        id, workspace_id, title, stable_key, structure, category, document_types,
        language, tags, status, current_version, definition, created_by, updated_by
      ) values (
        ${id}, ${session.workspaceId}, ${String(clause.title || "Untitled clause")}, ${localId},
        ${clause.structure === "nested" ? "nested" : "flat"}, ${String(clause.category || "")},
        ${Array.isArray(clause.documentTypes) ? clause.documentTypes.map(String) : []},
        ${String(clause.language || "en")}, ${Array.isArray(clause.tags) ? clause.tags.map(String) : []},
        ${["draft", "published", "inactive"].includes(String(clause.status)) ? String(clause.status) : "draft"},
        ${version}, ${transaction.json(asJson(clause))}, ${session.userId}, ${session.userId}
      ) on conflict (id) do update set
        title = excluded.title,
        structure = excluded.structure,
        category = excluded.category,
        document_types = excluded.document_types,
        language = excluded.language,
        tags = excluded.tags,
        status = excluded.status,
        current_version = excluded.current_version,
        definition = excluded.definition,
        updated_by = excluded.updated_by,
        updated_at = now()
    `;
    await transaction`
      insert into clause_versions (
        id, workspace_id, clause_id, version, definition, status, created_by
      ) values (
        ${qualifiedId(session.workspaceId, `${localId}:v${version}`, "clause-version")},
        ${session.workspaceId}, ${id}, ${version}, ${transaction.json(asJson(clause))},
        ${clause.status === "published" ? "published" : "draft"}, ${session.userId}
      ) on conflict (clause_id, version) do update set
        definition = excluded.definition,
        status = excluded.status
      where clause_versions.status = 'draft'
    `;
  }

  for (const group of asArray(appStore.placeholderGroups)) {
    const localId = String(group.id || randomUUID());
    const groupId = qualifiedId(session.workspaceId, localId, "placeholder-group");
    await transaction`
      insert into placeholder_groups (
        id, workspace_id, data_source_id, name, stable_key, source_type, source_table, status,
        created_by, updated_by
      ) values (
        ${groupId}, ${session.workspaceId}, ${String(group.sourceType || "Manual").toLowerCase() === "manual" ? null : sourceId},
        ${String(group.name || "Untitled group")}, ${localId},
        ${String(group.sourceType || "Manual")}, ${group.sourceTable ? String(group.sourceTable) : null},
        ${group.status === "inactive" ? "inactive" : "active"}, ${session.userId}, ${session.userId}
      ) on conflict (id) do update set
        data_source_id = excluded.data_source_id,
        name = excluded.name,
        source_type = excluded.source_type,
        source_table = excluded.source_table,
        status = excluded.status,
        updated_by = excluded.updated_by,
        updated_at = now()
    `;
    for (const [index, field] of asArray(group.fields).entries()) {
      const fieldLocalId = String(field.id || field.key || randomUUID());
      const fieldId = qualifiedId(session.workspaceId, `${localId}:${fieldLocalId}`, "placeholder-field");
      const fieldType = ["text", "multiline", "number", "currency", "date", "select", "multiselect", "repeating"].includes(String(field.type))
        ? String(field.type)
        : "text";
      await transaction`
        insert into placeholder_fields (
          id, workspace_id, group_id, stable_key, label, field_type, source_field_id,
          example_value, required, manual_override, mapping_status, status, sort_order,
          created_by, updated_by
        ) values (
          ${fieldId}, ${session.workspaceId}, ${groupId}, ${fieldLocalId},
          ${String(field.label || field.key || "Untitled field")}, ${fieldType},
          ${field.sourceField ? String(field.sourceField) : null}, ${String(field.example || "")},
          ${Boolean(field.required)}, ${field.manualOverride !== false},
          ${field.mappingStatus === "invalid" ? "invalid" : "valid"},
          ${field.status === "inactive" ? "inactive" : "active"}, ${index}, ${session.userId}, ${session.userId}
        ) on conflict (id) do update set
          label = excluded.label,
          field_type = excluded.field_type,
          source_field_id = excluded.source_field_id,
          example_value = excluded.example_value,
          required = excluded.required,
          manual_override = excluded.manual_override,
          mapping_status = excluded.mapping_status,
          status = excluded.status,
          sort_order = excluded.sort_order,
          updated_by = excluded.updated_by,
          updated_at = now()
      `;
      if (field.sourceField) {
        const mappingId = qualifiedId(session.workspaceId, `${localId}:${fieldLocalId}:mapping`, "field-mapping");
        await transaction`
          insert into field_mappings (
            id, workspace_id, data_source_id, placeholder_field_id, source_field_id,
            status, created_by, updated_by
          ) values (
            ${mappingId}, ${session.workspaceId}, ${sourceId}, ${fieldId}, ${String(field.sourceField)},
            ${field.mappingStatus === "invalid" ? "invalid" : "active"}, ${session.userId}, ${session.userId}
          ) on conflict (id) do update set
            source_field_id = excluded.source_field_id,
            status = excluded.status,
            updated_by = excluded.updated_by,
            updated_at = now()
        `;
      }
    }
  }

  const snapshots = asRecord(workspace.sourceSnapshots);
  const currentPersonId = String(workspace.personId || "unassigned");
  if (!snapshots[currentPersonId]) snapshots[currentPersonId] = asRecord(workspace.values);
  for (const document of asArray(appStore.documents)) {
    const documentPersonId = document.personId ? String(document.personId) : "";
    if (documentPersonId && !snapshots[documentPersonId]) snapshots[documentPersonId] = asRecord(document.values);
  }
  for (const [personLocalId, snapshotValue] of Object.entries(snapshots)) {
    const personId = qualifiedId(session.workspaceId, personLocalId, "person");
    const snapshot = asRecord(snapshotValue);
    const displayName = String(snapshot.full_name || snapshot.name || personLocalId);
    await transaction`
      insert into onboarding_people (
        id, workspace_id, external_person_id, display_name, email, status, private_metadata
      ) values (
        ${personId}, ${session.workspaceId}, ${personLocalId}, ${displayName},
        ${snapshot.email ? String(snapshot.email) : null}, 'active',
        ${transaction.json(asJson({ source: "workspace_snapshot" }))}
      ) on conflict (id) do update set
        display_name = excluded.display_name,
        email = excluded.email,
        updated_at = now()
    `;
    const submissionId = qualifiedId(session.workspaceId, `${personLocalId}:latest`, "submission");
    await transaction`
      insert into onboarding_submissions (
        id, workspace_id, person_id, data_source_id, external_submission_id,
        synced_at, source_revision, values
      ) values (
        ${submissionId}, ${session.workspaceId}, ${personId}, ${sourceId}, ${`${personLocalId}:latest`},
        now(), ${String(asRecord(asRecord(workspace.sourceUpdateMeta)[personLocalId]).submissionId || "snapshot")},
        ${transaction.json(asJson(snapshot))}
      ) on conflict (id) do update set
        synced_at = now(),
        source_revision = excluded.source_revision,
        values = excluded.values,
        updated_at = now()
    `;
  }

  for (const layout of asArray(appStore.layouts)) {
    const localId = String(layout.id || randomUUID());
    const id = qualifiedId(session.workspaceId, localId, "layout");
    const layoutRows = await transaction<Array<{ current_version: number }>>`
      insert into layouts (
        id, workspace_id, company_id, name, stable_key, status, current_version,
        definition, created_by, updated_by
      ) values (
        ${id}, ${session.workspaceId}, ${companyId}, ${String(layout.name || "Untitled layout")}, ${localId},
        ${["draft", "published", "inactive"].includes(String(layout.status)) ? String(layout.status) : "draft"},
        1, ${transaction.json(asJson(layout))}, ${session.userId}, ${session.userId}
      ) on conflict (id) do update set
        name = excluded.name,
        status = excluded.status,
        current_version = case
          when layouts.definition is distinct from excluded.definition then layouts.current_version + 1
          else layouts.current_version
        end,
        definition = excluded.definition,
        updated_by = excluded.updated_by,
        updated_at = now()
      returning current_version
    `;
    const version = Number(layoutRows[0]?.current_version || 1);
    await transaction`
      insert into layout_versions (
        id, workspace_id, layout_id, version, definition, status, created_by
      ) values (
        ${qualifiedId(session.workspaceId, `${localId}:v${version}`, "layout-version")},
        ${session.workspaceId}, ${id}, ${version}, ${transaction.json(asJson(layout))},
        ${layout.status === "published" ? "published" : "draft"}, ${session.userId}
      ) on conflict (layout_id, version) do nothing
    `;
  }

  for (const document of asArray(appStore.documents)) {
    const localId = String(document.id || randomUUID());
    const id = qualifiedId(session.workspaceId, localId, "document");
    const status = ["draft", "in-review", "approved", "void"].includes(String(document.status)) ? String(document.status) : "draft";
    const generation = ["not-generated", "generating", "generated", "failed"].includes(String(document.generationStatus))
      ? String(document.generationStatus)
      : "not-generated";
    const templateLocalId = document.templateId ? String(document.templateId) : "";
    const template = templatesByLocalId.get(templateLocalId);
    const templateDatabaseId = template
      ? qualifiedId(session.workspaceId, templateLocalId, "template")
      : null;
    const personDatabaseId = document.personId
      ? qualifiedId(session.workspaceId, String(document.personId), "person")
      : null;
    const documentVersions = asArray(document.versions);
    const versionNumbers = documentVersions.map((version, index) =>
      documentVersionNumber(version, Math.max(1, documentVersions.length - index)),
    );
    const currentVersion = Math.max(1, ...versionNumbers);
    await transaction`
      insert into documents (
        id, workspace_id, company_id, template_id, person_id, title, document_type, save_status,
        document_status, generation_status, current_version, current_snapshot,
        created_by, updated_by
      ) values (
        ${id}, ${session.workspaceId}, ${companyId}, ${templateDatabaseId}, ${personDatabaseId},
        ${String(document.docName || "Untitled document")},
        ${String(template?.type || document.templateId || "HR document")}, 'saved', ${status}, ${generation},
        ${currentVersion}, ${transaction.json(asJson(document))},
        ${session.userId}, ${session.userId}
      ) on conflict (id) do update set
        company_id = excluded.company_id,
        template_id = excluded.template_id,
        person_id = excluded.person_id,
        title = excluded.title,
        document_type = excluded.document_type,
        save_status = excluded.save_status,
        document_status = excluded.document_status,
        generation_status = excluded.generation_status,
        current_version = excluded.current_version,
        current_snapshot = excluded.current_snapshot,
        archived_at = null,
        updated_by = excluded.updated_by,
        updated_at = now()
    `;
    let approvedVersionId: string | null = null;
    const versionIds = new Map<string, string>();
    for (const [versionIndex, version] of documentVersions.entries()) {
      const versionLocalId = String(version.id || `${localId}:version:${versionIndex + 1}`);
      const versionId = qualifiedId(session.workspaceId, versionLocalId, "document-version");
      const versionStatus = String(version.status || "draft").toLowerCase().replace(" ", "-");
      await transaction`
        insert into document_versions (
          id, workspace_id, document_id, version, status, snapshot,
          source_snapshot, change_summary, created_by
        ) values (
          ${versionId}, ${session.workspaceId}, ${id}, ${versionNumbers[versionIndex]},
          ${["draft", "in-review", "approved", "void"].includes(versionStatus) ? versionStatus : "draft"},
          ${transaction.json(asJson(version))}, ${transaction.json(asJson(asRecord(version.sourceSnapshot)))},
          ${version.summary ? String(version.summary) : null}, ${session.userId}
        ) on conflict do nothing
      `;
      const persistedVersions = await transaction<Array<{ id: string }>>`
        select id from document_versions
        where document_id = ${id} and version = ${versionNumbers[versionIndex]}
        limit 1
      `;
      const persistedVersionId = persistedVersions[0]?.id || versionId;
      versionIds.set(versionLocalId, persistedVersionId);
      if (!approvedVersionId && versionStatus === "approved") approvedVersionId = persistedVersionId;
    }
    await transaction`
      update documents
      set approved_version_id = ${status === "approved" ? approvedVersionId : null}
      where id = ${id}
    `;
    for (const exported of asArray(document.exports)) {
      const exportLocalId = String(exported.id || randomUUID());
      const exportId = qualifiedId(session.workspaceId, exportLocalId, "document-export");
      const exportStatus = ["generating", "generated", "failed"].includes(String(exported.status)) ? String(exported.status) : "failed";
      const exportVersionLocalId = exported.versionId ? String(exported.versionId) : "";
      const exportVersionId = exportVersionLocalId ? versionIds.get(exportVersionLocalId) || null : null;
      await transaction`
        insert into document_exports (
          id, workspace_id, document_id, document_version_id, format, status, file_name, storage_path,
          byte_size, error, created_by
        ) values (
          ${exportId}, ${session.workspaceId}, ${id}, ${exportVersionId}, ${exported.format === "pdf" ? "pdf" : "docx"},
          ${exportStatus}, ${String(exported.fileName || "document")},
          ${exported.storagePath ? String(exported.storagePath) : null},
          ${exported.byteSize ? Number(exported.byteSize) : null},
          ${exported.error ? String(exported.error) : null}, ${session.userId}
        ) on conflict (id) do update set
          document_version_id = excluded.document_version_id,
          status = excluded.status,
          file_name = excluded.file_name,
          storage_path = excluded.storage_path,
          byte_size = excluded.byte_size,
          error = excluded.error
      `;
    }
  }
};

export const saveWorkspaceState = async (
  session: ServerSession,
  payload: WorkspacePayload,
  baseRevision: number | null,
) => {
  const role = await requireWorkspaceAccess(session);
  if (!payload || typeof payload !== "object") throw new ApiError(400, "INVALID_REQUEST", "Workspace payload is required.");
  if (payload.schemaVersion !== currentWorkspaceSchemaVersion) {
    throw new ApiError(400, "SCHEMA_VERSION_UNSUPPORTED", `Use workspace schema version ${currentWorkspaceSchemaVersion}.`);
  }
  if (!payload.appStore || typeof payload.appStore !== "object" || !payload.workspace || typeof payload.workspace !== "object") {
    throw new ApiError(400, "INVALID_REQUEST", "The workspace save is missing its application or document state.");
  }
  const serialized = JSON.stringify(payload);
  const payloadBytes = Buffer.byteLength(serialized, "utf8");
  if (payloadBytes > maxWorkspacePayloadBytes) {
    throw new ApiError(413, "STATE_TOO_LARGE", "Workspace data exceeds the 4 MB cloud-save limit. Move images and generated files to Supabase Storage.", {
      payloadBytes,
      maxBytes: maxWorkspacePayloadBytes,
    });
  }

  const { sql, ready } = database();
  await ready;
  return sql.begin(async (transaction) => {
    const currentRows = await transaction<WorkspaceStateRow[]>`
      select workspace_id, schema_version, revision, payload, updated_at, updated_by
      from workspace_states
      where workspace_id = ${session.workspaceId}
      for update
    `;
    const current = currentRows[0] || null;
    if (current && baseRevision !== current.revision) {
      throw new ApiError(409, "STATE_CONFLICT", "This workspace changed in another session. Reload the cloud version before saving again.", {
        revision: current.revision,
        updatedAt: current.updated_at,
        updatedBy: current.updated_by,
      });
    }
    if (!current && baseRevision !== null && baseRevision !== 0) {
      throw new ApiError(409, "STATE_CONFLICT", "The workspace cloud state was reset. Reload before saving again.", { revision: 0 });
    }
    if (role === "Reviewer" && current) {
      const previousPayload = current.payload as WorkspacePayload;
      if (protectedConfiguration(previousPayload) !== protectedConfiguration(payload)) {
        throw new ApiError(403, "FORBIDDEN", "Reviewers cannot change templates, clauses, placeholders, or layouts.");
      }
      if (documentContentSignature(previousPayload) !== documentContentSignature(payload)) {
        throw new ApiError(403, "FORBIDDEN", "Reviewers can approve or reject a submitted snapshot but cannot edit its document content or source data.");
      }
    }
    if (current) {
      const previousPayload = current.payload as WorkspacePayload;
      const previousStatuses = documentStatuses(previousPayload);
      const nextStatuses = documentStatuses(payload);
      const previousDocuments = documentRecords(previousPayload);
      const nextDocuments = documentRecords(payload);
      const documentIds = new Set([...previousStatuses.keys(), ...nextStatuses.keys()]);
      for (const documentId of documentIds) {
        const previousStatus = previousStatuses.get(documentId) || "draft";
        const nextStatus = nextStatuses.get(documentId);
        if (previousStatus === "approved" && !nextStatus) {
          throw new ApiError(403, "APPROVED_DOCUMENT_IMMUTABLE", "Approved documents cannot be deleted. Void or restore the document as a new draft first.");
        }
        if (previousStatus === "approved" && nextStatus === "approved") {
          const previousDocument = previousDocuments.get(documentId);
          const nextDocument = nextDocuments.get(documentId);
          if (previousDocument && nextDocument && documentRecordContentSignature(previousDocument) !== documentRecordContentSignature(nextDocument)) {
            throw new ApiError(403, "APPROVED_DOCUMENT_IMMUTABLE", "Approved document content is locked. Restore it as a new draft before editing.");
          }
        }
        if (!nextStatus) continue;
        if (role === "Editor" && nextStatus === "approved" && previousStatus !== "approved") {
          throw new ApiError(403, "FORBIDDEN", "Editors cannot approve document versions.");
        }
        if (role === "Reviewer" && nextStatus !== previousStatus && !(previousStatus === "in-review" && nextStatus === "approved")) {
          throw new ApiError(403, "FORBIDDEN", "Reviewers can only approve documents that are currently in review.");
        }
      }
    }

    const nextRevision = (current?.revision || 0) + 1;
    const rows = await transaction<WorkspaceStateRow[]>`
      insert into workspace_states (
        workspace_id, schema_version, revision, payload, payload_bytes, updated_by
      ) values (
        ${session.workspaceId}, ${currentWorkspaceSchemaVersion}, ${nextRevision},
        ${transaction.json(asJson(payload))}, ${payloadBytes}, ${session.userId}
      ) on conflict (workspace_id) do update set
        schema_version = excluded.schema_version,
        revision = excluded.revision,
        payload = excluded.payload,
        payload_bytes = excluded.payload_bytes,
        updated_by = excluded.updated_by,
        updated_at = now()
      returning workspace_id, schema_version, revision, payload, updated_at, updated_by
    `;

    const preferences = asRecord(payload.preferences);
    await transaction`
      insert into user_preferences (
        workspace_id, user_id, theme, navigation, editor_preferences, notification_preferences
      ) values (
        ${session.workspaceId}, ${session.userId}, ${preferences.theme === "dark" ? "dark" : "light"},
        ${transaction.json(asJson(asRecord(preferences.navigation)))},
        ${transaction.json(asJson(asRecord(preferences.editor)))},
        ${transaction.json(asJson(asRecord(preferences.notifications)))}
      ) on conflict (workspace_id, user_id) do update set
        theme = excluded.theme,
        navigation = excluded.navigation,
        editor_preferences = excluded.editor_preferences,
        notification_preferences = excluded.notification_preferences,
        updated_at = now()
    `;

    await syncNormalizedState(transaction, session, payload);
    if (current) {
      const previousStatuses = documentStatuses(current.payload as WorkspacePayload);
      const nextStatuses = documentStatuses(payload);
      for (const [documentLocalId, nextStatus] of nextStatuses) {
        const previousStatus = previousStatuses.get(documentLocalId) || "draft";
        if (previousStatus === nextStatus) continue;
        const action = nextStatus === "in-review" ? "submitted" : nextStatus === "approved" ? "approved" : nextStatus === "void" ? "voided" : "restored";
        const nextDocument = documentRecords(payload).get(documentLocalId);
        const latestVersion = nextDocument ? asArray(nextDocument.versions)[0] : undefined;
        const documentVersionId = latestVersion?.id
          ? qualifiedId(session.workspaceId, String(latestVersion.id), "document-version")
          : null;
        await transaction`
          insert into approval_events (workspace_id, document_id, document_version_id, action, actor_user_id)
          values (
            ${session.workspaceId}, ${qualifiedId(session.workspaceId, documentLocalId, "document")},
            ${documentVersionId}, ${action}, ${session.userId}
          )
        `;
      }
    }
    await transaction`
      insert into audit_logs (workspace_id, actor_user_id, action, entity_type, entity_id, metadata)
      values (
        ${session.workspaceId}, ${session.userId}, ${current ? "workspace.updated" : "workspace.migrated"},
        'workspace_state', ${session.workspaceId},
        ${transaction.json(asJson({ revision: nextRevision, schemaVersion: currentWorkspaceSchemaVersion, payloadBytes }))}
      )
    `;
    return rows[0];
  });
};
