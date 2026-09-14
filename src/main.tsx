import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  AlignmentType,
  Document,
  Footer,
  Header,
  ImageRun,
  PageNumber,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} from "docx";
import { jsPDF } from "jspdf";
import * as XLSX from "xlsx";
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Archive,
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  Bold,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleHelp,
  ClipboardPaste,
  Cloud,
  Copy,
  Download,
  FileCheck2,
  FileText,
  FolderOpen,
  GripVertical,
  Highlighter,
  History,
  ImagePlus,
  Italic,
  Link2,
  List,
  ListOrdered,
  Lock,
  Mail,
  MoreHorizontal,
  Paintbrush,
  PanelRight,
  Paperclip,
  Plus,
  RefreshCw,
  Redo2,
  RotateCcw,
  Save,
  Search,
  Settings2,
  ShieldCheck,
  Sparkles,
  Strikethrough,
  Sun,
  Table2,
  Trash2,
  Underline,
  Undo2,
  Upload,
  UserRound,
  UsersRound,
  WandSparkles,
  X,
  Moon,
  PanelLeftClose,
  PanelLeftOpen,
  Check,
  Circle,
} from "lucide-react";
import "./styles.css";

type DocumentType =
  "Offer Letter" | "Fixed-Term Employment Agreement" | "Freelance Agreement";
type Clause = {
  id: string;
  title: string;
  text: string;
  html?: string;
  included: boolean;
  tag: string;
  example?: boolean;
  sourceClauseId?: string;
  sourceClauseVersion?: number;
};
type Section = { id: string; title: string; clauses: Clause[] };
type Template = {
  id: string;
  name: string;
  type: DocumentType;
  description: string;
  sections: Section[];
  updated: string;
  color: string;
};
type Person = {
  id: string;
  submissionId: string;
  name: string;
  role: string;
  submitted: string;
  fields: Record<string, string>;
};
type Letterhead = {
  mode: "first" | "all" | "different";
  accent: string;
  opacity: number;
  top: number;
  left: number;
  width: number;
  page: "A4" | "Letter";
  margin: number;
  fileName?: string;
  fileType?: string;
  dataUrl?: string;
  firstPage?: LetterheadLayout;
  subsequentPage?: LetterheadLayout;
};
type LetterheadLayout = Pick<Letterhead, "accent" | "opacity" | "top" | "left" | "width" | "margin" | "fileName" | "fileType" | "dataUrl">;
type CustomPlaceholder = {
  key: string;
  label: string;
  source: string;
  group: string;
};
type UserRole = "Admin" | "Editor" | "Reviewer";
type SyncDiff = { key: string; oldValue: string; newValue: string };
type DocumentVersion = {
  id: string;
  label: string;
  createdAt: string;
  author: string;
  status: "Draft" | "In review" | "Approved";
  summary: string;
  sections: Section[];
  values: Record<string, string>;
  letterhead: Letterhead;
  docName: string;
  templateId?: string;
  personId?: string;
  submissionId?: string;
  sourceSnapshot?: Record<string, string>;
  sourceSyncedAt?: string;
  manualOverrides?: string[];
  sourceUpdates?: Record<string, Record<string, string>>;
  sourceUpdateMeta?: Record<string, { submissionId: string; detectedAt: string }>;
};
type PendingChange = {
  kind: "template" | "person";
  id: string;
  label: string;
};
type PromptRequest = {
  kind: "placeholder" | "replace" | "rename" | "link";
  title: string;
  description?: string;
  value: string;
  oldKey?: string;
};

type WorkflowStage = "build" | "review" | "export";
type AppModule = "workspace" | "clauses" | "placeholders" | "layouts" | "documents";
type ClauseSubsection = { id: string; title: string; contents: Clause[]; collapsed?: boolean };
type ClauseRecord = {
  id: string;
  title: string;
  structure: "flat" | "nested";
  subsections: ClauseSubsection[];
  contents: Clause[];
  category: string;
  documentTypes: string[];
  language: string;
  tags: string[];
  status: "draft" | "published" | "inactive";
  version: number;
  updatedAt: string;
};
type PlaceholderField = {
  id: string;
  key: string;
  label: string;
  type: "text" | "date" | "currency" | "number" | "select" | "multiline";
  sourceField: string;
  example: string;
  required: boolean;
  status: "active" | "inactive";
  mappingStatus: "valid" | "invalid";
  manualOverride: boolean;
};
type PlaceholderGroup = {
  id: string;
  name: string;
  sourceType: string;
  sourceTable?: string;
  fields: PlaceholderField[];
  status: "active" | "inactive";
};
type LayoutRecord = {
  id: string;
  name: string;
  companyId: string;
  letterhead: Letterhead;
  status: "draft" | "published" | "inactive";
  updatedAt: string;
};
type ExportRecord = {
  id: string;
  documentId: string;
  versionId?: string;
  format: "docx" | "pdf";
  status: "generating" | "generated" | "failed";
  fileName: string;
  createdAt: string;
  contentBase64?: string;
  error?: string;
};
type DocumentRecord = {
  id: string;
  docName: string;
  templateId?: string;
  personId?: string;
  companyId: string;
  status: "draft" | "in-review" | "approved" | "void";
  generationStatus: "not-generated" | "generating" | "generated" | "failed";
  sections: Section[];
  values: Record<string, string>;
  letterhead: Letterhead;
  versions: DocumentVersion[];
  exports: ExportRecord[];
  updatedAt: string;
};
type NavigationState = {
  module: AppModule;
  sidebarCollapsed: boolean;
  filters: Record<string, Record<string, string>>;
};
type AppStore = {
  schemaVersion: 1;
  navigation: NavigationState;
  clauses: ClauseRecord[];
  placeholderGroups: PlaceholderGroup[];
  layouts: LayoutRecord[];
  documents: DocumentRecord[];
  exports: ExportRecord[];
};

const formatDocumentStamp = (date = new Date()) => {
  const datePart = date.toLocaleDateString("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
  });
  const timePart = date.toLocaleTimeString("en-US", {
    hour: "2-digit",
    minute: "2-digit",
  });
  return `${datePart} · ${timePart}`;
};

const demoTemplates: Template[] = [
  {
    id: "offer",
    name: "Offer Letter · Malaysia",
    type: "Offer Letter",
    description:
      "Standard permanent hire offer with compensation and onboarding details.",
    updated: "Sep 08, 2026",
    color: "#1f7a62",
    sections: [
      {
        id: "intro",
        title: "Appointment",
        clauses: [
          {
            id: "offer-1",
            title: "Offer of employment",
            tag: "Core",
            included: true,
            text: "We are pleased to offer you employment with {{company_name}} as {{job_title}}, reporting to {{manager_name}}. Your proposed start date is {{start_date}}.",
          },
          {
            id: "offer-2",
            title: "Work arrangement",
            tag: "Conditional",
            included: true,
            text: "Your primary work arrangement will be {{work_arrangement}}. The Company may reasonably adjust this arrangement to meet business needs.",
          },
        ],
      },
      {
        id: "comp",
        title: "Compensation & benefits",
        clauses: [
          {
            id: "offer-3",
            title: "Salary",
            tag: "Core",
            included: true,
            text: "Your gross monthly salary will be {{monthly_salary}}, payable in accordance with the Company’s normal payroll practices.",
          },
          {
            id: "offer-4",
            title: "Probation",
            tag: "Conditional",
            included: true,
            text: "Your appointment is subject to a probationary period of {{probation_period}}.",
          },
        ],
      },
      {
        id: "close",
        title: "Acceptance",
        clauses: [
          {
            id: "offer-5",
            title: "Acceptance and signature",
            tag: "Core",
            included: true,
            text: "Please sign and return this letter by {{acceptance_deadline}}. We look forward to welcoming you to {{company_name}}.",
          },
        ],
      },
    ],
  },
  {
    id: "fixed",
    name: "Fixed-Term Employment Agreement",
    type: "Fixed-Term Employment Agreement",
    description:
      "Fixed-term agreement with term, renewal and expiry provisions.",
    updated: "Sep 04, 2026",
    color: "#3559a8",
    sections: [
      {
        id: "parties",
        title: "Parties & appointment",
        clauses: [
          {
            id: "fixed-1",
            title: "Parties",
            tag: "Core",
            included: true,
            text: "This Agreement is made between {{company_name}} (the “Company”) and {{full_name}} (the “Employee”).",
          },
          {
            id: "fixed-2",
            title: "Position and duties",
            tag: "Core",
            included: true,
            text: "The Employee is appointed as {{job_title}} and will perform the duties reasonably assigned by the Company.",
          },
        ],
      },
      {
        id: "term",
        title: "Term & probation",
        clauses: [
          {
            id: "fixed-3",
            title: "Fixed term",
            tag: "Conditional",
            included: true,
            text: "The employment commences on {{start_date}} and ends on {{end_date}}, unless terminated earlier in accordance with this Agreement.",
          },
          {
            id: "fixed-4",
            title: "Renewal",
            tag: "Optional",
            included: true,
            text: "The parties may renew this Agreement by written agreement before the expiry date.",
          },
        ],
      },
      {
        id: "pay",
        title: "Salary & working arrangements",
        clauses: [
          {
            id: "fixed-5",
            title: "Salary and payment",
            tag: "Core",
            included: true,
            text: "The Employee will receive {{monthly_salary}} per month, paid by {{payment_method}}.",
          },
          {
            id: "fixed-6",
            title: "Leave and benefits",
            tag: "Core",
            included: true,
            text: "The Employee is eligible for benefits and leave in accordance with applicable law and Company policy.",
          },
        ],
      },
      {
        id: "sign",
        title: "Signatures",
        clauses: [
          {
            id: "fixed-7",
            title: "Execution",
            tag: "Core",
            included: true,
            text: "Signed for and on behalf of {{company_name}} by {{signatory_name}}, {{signatory_title}}.",
          },
        ],
      },
    ],
  },
  {
    id: "freelance",
    name: "Freelance Agreement",
    type: "Freelance Agreement",
    description:
      "Independent contractor agreement with milestone billing and deliverables.",
    updated: "Aug 28, 2026",
    color: "#a66028",
    sections: [
      {
        id: "scope",
        title: "Scope of services",
        clauses: [
          {
            id: "free-1",
            title: "Engagement",
            tag: "Core",
            included: true,
            text: "{{company_name}} engages {{full_name}} as an independent contractor to provide {{service_scope}}.",
          },
          {
            id: "free-2",
            title: "Deliverables",
            tag: "Conditional",
            included: true,
            text: "The Contractor will deliver the agreed work products and milestones described in the attached schedule.",
          },
        ],
      },
      {
        id: "fees",
        title: "Fees & payment",
        clauses: [
          {
            id: "free-3",
            title: "Billing model",
            tag: "Conditional",
            included: true,
            text: "Fees will be calculated on a {{billing_method}} basis. The total agreed fee is {{project_fee}}.",
          },
          {
            id: "free-4",
            title: "Milestones",
            tag: "Conditional",
            included: true,
            text: "Payment milestones: {{payment_milestones}}. Invoices are payable within {{payment_terms}}.",
          },
        ],
      },
      {
        id: "conf",
        title: "Confidentiality & IP",
        clauses: [
          {
            id: "free-5",
            title: "Confidentiality",
            tag: "Core",
            included: true,
            text: "The Contractor must keep confidential all non-public information received in connection with the services.",
          },
          {
            id: "free-6",
            title: "Intellectual property",
            tag: "Core",
            included: true,
            text: "Upon payment, all agreed deliverables and related intellectual property will vest in the Company.",
          },
        ],
      },
      {
        id: "sign",
        title: "Signatures",
        clauses: [
          {
            id: "free-7",
            title: "Execution",
            tag: "Core",
            included: true,
            text: "Signed by {{full_name}} and for {{company_name}} by {{signatory_name}}.",
          },
        ],
      },
    ],
  },
];

const people: Person[] = [
  {
    id: "EMP-2048",
    submissionId: "onb-EMP-2048-20260911-01",
    name: "Aisha Rahman",
    role: "Senior Product Designer",
    submitted: "Today, 09:42",
    fields: {
      full_name: "Aisha Rahman",
      address: "17 Jalan Damai, Kuala Lumpur",
      job_title: "Senior Product Designer",
      start_date: "01 Oct 2026",
      end_date: "30 Sep 2027",
      company_name: "Northstar Labs Sdn. Bhd.",
      monthly_salary: "RM 14,500",
      manager_name: "Daniel Wong",
      work_arrangement: "Hybrid",
      probation_period: "3 months",
      acceptance_deadline: "18 Sep 2026",
      signatory_name: "Mei Lin Tan",
      signatory_title: "People Director",
      billing_method: "fixed fee",
      project_fee: "RM 28,000",
      payment_method: "monthly payroll",
      payment_terms: "30 days",
      service_scope: "brand system and product design",
      payment_milestones: "50% on kickoff; 50% on final delivery",
    },
  },
  {
    id: "EMP-2041",
    submissionId: "onb-EMP-2041-20260910-02",
    name: "Marcus Lee",
    role: "Frontend Engineer",
    submitted: "Yesterday, 16:10",
    fields: {
      full_name: "Marcus Lee",
      address: "8 Lorong 4, Petaling Jaya",
      job_title: "Frontend Engineer",
      start_date: "15 Sep 2026",
      end_date: "14 Sep 2027",
      company_name: "Northstar Labs Sdn. Bhd.",
      monthly_salary: "RM 12,800",
      manager_name: "Priya Nair",
      work_arrangement: "Remote",
      probation_period: "Not applicable",
      acceptance_deadline: "12 Sep 2026",
      signatory_name: "Mei Lin Tan",
      signatory_title: "People Director",
      billing_method: "hourly",
      project_fee: "RM 180 / hour",
      payment_method: "monthly payroll",
      payment_terms: "30 days",
      service_scope: "frontend engineering support",
      payment_milestones: "Monthly timesheet",
    },
  },
  {
    id: "CON-119",
    submissionId: "onb-CON-119-20260909-01",
    name: "Nadia Chen",
    role: "Brand Consultant",
    submitted: "Sep 09, 2026",
    fields: {
      full_name: "Nadia Chen",
      address: "22A Taman Tun, Kuala Lumpur",
      job_title: "Brand Consultant",
      start_date: "20 Sep 2026",
      end_date: "20 Dec 2026",
      company_name: "Northstar Labs Sdn. Bhd.",
      monthly_salary: "RM 9,000",
      manager_name: "Mei Lin Tan",
      work_arrangement: "Remote",
      probation_period: "Not applicable",
      acceptance_deadline: "16 Sep 2026",
      signatory_name: "Mei Lin Tan",
      signatory_title: "People Director",
      billing_method: "milestone",
      project_fee: "RM 24,000",
      payment_method: "bank transfer",
      payment_terms: "30 days",
      service_scope: "brand strategy and workshop facilitation",
      payment_milestones: "30% kickoff; 40% strategy; 30% final presentation",
    },
  },
];

const placeholderMeta = [
  ["full_name", "Full Name", "Onboarding Form", "Person"],
  ["job_title", "Job Title", "Onboarding Form", "Person"],
  ["start_date", "Start Date", "Onboarding Form", "Person"],
  ["end_date", "End Date", "Onboarding Form", "Person"],
  ["monthly_salary", "Monthly Salary", "HR input", "Compensation"],
  ["work_arrangement", "Work Arrangement", "Dropdown", "Employment"],
  ["billing_method", "Billing Method", "Dropdown", "Freelance"],
  ["payment_milestones", "Payment Milestones", "HR input", "Freelance"],
];

const clone = <T,>(value: T): T => JSON.parse(JSON.stringify(value));
const storageKey = "hr-doc-generator-state-v2";
const defaultLetterheadLayout: LetterheadLayout = {
  accent: "#3336cc",
  opacity: 0.12,
  top: 18,
  left: 22,
  width: 156,
  margin: 74,
};
const makeDefaultLetterhead = (): Letterhead => ({
  mode: "different",
  page: "A4",
  ...clone(defaultLetterheadLayout),
  firstPage: clone(defaultLetterheadLayout),
  subsequentPage: { ...clone(defaultLetterheadLayout), top: 12, opacity: 0.08 },
});
const normalizeLetterhead = (value?: Partial<Letterhead>): Letterhead => {
  const legacy = { ...defaultLetterheadLayout, ...(value || {}) };
  return {
    mode: value?.mode || "different",
    page: value?.page || "A4",
    ...legacy,
    firstPage: {
      ...legacy,
      ...(value?.firstPage || {}),
    },
    subsequentPage: {
      ...legacy,
      top: legacy.top - 6,
      opacity: Math.max(0.04, legacy.opacity - 0.04),
      ...(value?.subsequentPage || {}),
    },
  };
};
const formatFileDate = (date = new Date()) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};
const dataUrlToBytes = (dataUrl?: string) => {
  if (!dataUrl || !dataUrl.includes(",")) return null;
  try {
    const binary = atob(dataUrl.split(",", 2)[1]);
    return Uint8Array.from(binary, (character) => character.charCodeAt(0));
  } catch {
    return null;
  }
};
const blobToDataUrl = (blob: Blob) => new Promise<string>((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(String(reader.result || ""));
  reader.onerror = () => reject(reader.error || new Error("Could not read generated file"));
  reader.readAsDataURL(blob);
});
const appStoreKey = "hr-doc-generator-app-store-v1";
const navigationKey = "hr-doc-generator-navigation-v1";
const companyId = "northstar-labs-my";
const moduleFromHash = (hash: string): AppModule => {
  const value = hash.replace(/^#/, "");
  return (["workspace", "clauses", "placeholders", "layouts", "documents"] as AppModule[]).includes(value as AppModule)
    ? (value as AppModule)
    : "workspace";
};
const makeClauseRecord = (clause: Clause, template: Template, sectionTitle: string): ClauseRecord => ({
  id: clause.id,
  title: clause.title,
  structure: "flat",
  subsections: [],
  contents: [{ ...clone(clause), title: sectionTitle ? `${sectionTitle} · ${clause.title}` : clause.title }],
  category: sectionTitle,
  documentTypes: [template.type],
  language: "English",
  tags: [clause.tag],
  status: "published",
  version: 1,
  updatedAt: template.updated,
});
const makeDefaultPlaceholderGroups = (): PlaceholderGroup[] => {
  const grouped = new Map<string, PlaceholderField[]>();
  placeholderMeta.forEach(([key, label, source, group]) => {
    const fields = grouped.get(group) || [];
    fields.push({
      id: key,
      key,
      label,
      type: key.includes("date") ? "date" : key.includes("salary") || key.includes("fee") ? "currency" : "text",
      sourceField: key,
      example: people[1].fields[key] || people[0].fields[key] || "Example value",
      required: ["full_name", "job_title", "start_date", "company_name"].includes(key),
      status: "active",
      mappingStatus: "valid",
      manualOverride: false,
    });
    grouped.set(group, fields);
  });
  return Array.from(grouped.entries()).map(([name, fields]) => ({
    id: `group-${name.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
    name: name === "Person" ? "Employee details" : name === "Compensation" ? "Employment details" : name,
    sourceType: "Onboarding Form",
    sourceTable: "onboarding_submissions",
    fields,
    status: "active",
  }));
};
const readAppStore = (): AppStore => {
  const fallbackNavigation: NavigationState = {
    module: moduleFromHash(typeof window === "undefined" ? "" : window.location.hash),
    sidebarCollapsed: false,
    filters: {},
  };
  try {
    const stored = localStorage.getItem(appStoreKey);
    if (stored) {
      const parsed = JSON.parse(stored) as Partial<AppStore>;
      return {
        schemaVersion: 1,
        navigation: { ...fallbackNavigation, ...(parsed.navigation || {}) },
        clauses: parsed.clauses || [],
        placeholderGroups: parsed.placeholderGroups || makeDefaultPlaceholderGroups(),
        layouts: parsed.layouts || [],
        documents: parsed.documents || [],
        exports: parsed.exports || [],
      };
    }
    const legacyRaw = localStorage.getItem(storageKey);
    const legacy = legacyRaw ? JSON.parse(legacyRaw) : {};
    const legacyTemplates: Template[] = legacy.templates || demoTemplates;
    const legacySections: Section[] = legacy.sections || clone(demoTemplates[1].sections);
    const legacyValues = legacy.values || clone(people[1].fields);
    const legacyLetterhead = normalizeLetterhead(legacy.letterhead);
    const clauses = legacyTemplates.flatMap((template) => template.sections.flatMap((section) => section.clauses.map((clause) => makeClauseRecord(clause, template, section.title))));
    const layouts: LayoutRecord[] = [{
      id: "layout-default",
      name: "Northstar default letterhead",
      companyId,
      letterhead: legacyLetterhead,
      status: "published",
      updatedAt: formatDocumentStamp(),
    }];
    const document: DocumentRecord = {
      id: "doc-current",
      docName: legacy.docName || "Marcus Lee · Fixed-Term Agreement",
      templateId: legacy.templateId || "fixed",
      personId: legacy.personId || "EMP-2041",
      companyId,
      status: legacy.docStatus === "In review" ? "in-review" : legacy.docStatus === "Approved" ? "approved" : "draft",
      generationStatus: "not-generated",
      sections: legacySections,
      values: legacyValues,
      letterhead: legacyLetterhead,
      versions: legacy.versions || [],
      exports: [],
      updatedAt: legacy.lastSavedAt || formatDocumentStamp(),
    };
    if (legacyRaw) {
      localStorage.setItem(`${storageKey}-backup-${Date.now()}`, legacyRaw);
    }
    const migratedGroups = makeDefaultPlaceholderGroups();
    const customGroupMap = new Map<string, PlaceholderField[]>();
    (legacy.customPlaceholders || []).forEach((item: CustomPlaceholder) => {
      const fields = customGroupMap.get(item.group) || [];
      fields.push({ id: item.key, key: item.key, label: item.label, type: "text", sourceField: item.key, example: "Custom value", required: false, status: "active", mappingStatus: "valid", manualOverride: true });
      customGroupMap.set(item.group, fields);
    });
    customGroupMap.forEach((fields, name) => migratedGroups.push({ id: `group-${name.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`, name, sourceType: "Manual", sourceTable: "custom_fields", fields, status: "active" }));
    const next: AppStore = {
      schemaVersion: 1,
      navigation: fallbackNavigation,
      clauses,
      placeholderGroups: migratedGroups,
      layouts,
      documents: [document],
      exports: [],
    };
    localStorage.setItem(appStoreKey, JSON.stringify(next));
    return next;
  } catch {
    return { schemaVersion: 1, navigation: fallbackNavigation, clauses: [], placeholderGroups: makeDefaultPlaceholderGroups(), layouts: [], documents: [], exports: [] };
  }
};

function App() {
  const [appStore, setAppStore] = useState<AppStore>(() => readAppStore());
  const [activeModule, setActiveModule] = useState<AppModule>(() => moduleFromHash(window.location.hash));
  const [sidebarCollapsed, setSidebarCollapsed] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem(navigationKey);
      return saved ? Boolean(JSON.parse(saved).sidebarCollapsed) : false;
    } catch {
      return false;
    }
  });
  const [moduleSearch, setModuleSearch] = useState("");
  const [selectedClauseRecordId, setSelectedClauseRecordId] = useState<string | null>(null);
  const [selectedPlaceholderGroupId, setSelectedPlaceholderGroupId] = useState<string | null>(null);
  const [moduleDrawer, setModuleDrawer] = useState<"clause" | "placeholder" | "layout" | null>(null);
  const [moduleNotice, setModuleNotice] = useState("");
  const [moduleFilter, setModuleFilter] = useState("all");
  const [clauseDraft, setClauseDraft] = useState<ClauseRecord | null>(null);
  const [placeholderDraft, setPlaceholderDraft] = useState<PlaceholderGroup | null>(null);
  const [layoutDraft, setLayoutDraft] = useState<LayoutRecord | null>(null);
  const [importPreview, setImportPreview] = useState<{ fileName: string; headers: string[]; sample: Record<string, string>[] } | null>(null);
  const [templates, setTemplates] = useState<Template[]>(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved).templates : demoTemplates;
  });
  const [templateId, setTemplateId] = useState("fixed");
  const [personId, setPersonId] = useState("EMP-2041");
  const [sections, setSections] = useState<Section[]>(() =>
    clone(demoTemplates[1].sections),
  );
  const [values, setValues] = useState<Record<string, string>>(() =>
    clone(people[1].fields),
  );
  const [letterhead, setLetterhead] = useState<Letterhead>({
    ...makeDefaultLetterhead(),
  });
  const [letterheadEditorPage, setLetterheadEditorPage] = useState<"first" | "subsequent">("first");
  const [workflowStage, setWorkflowStage] = useState<WorkflowStage>("build");
  const [leftRailCollapsed, setLeftRailCollapsed] = useState(false);
  const [activeSection, setActiveSection] = useState("parties");
  const [activeClause, setActiveClause] = useState("fixed-1");
  const [activeRightTab, setActiveRightTab] = useState<
    "person" | "placeholders" | "clauses" | "letterhead"
  >("person");
  const [search, setSearch] = useState("");
  const [saving, setSaving] = useState<"saved" | "saving" | "error">("saved");
  const [lastSavedAt, setLastSavedAt] = useState("");
  const [toast, setToast] = useState("");
  const [showPreview, setShowPreview] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);
  const [adminView, setAdminView] = useState<string | null>(null);
  const [companyMenuOpen, setCompanyMenuOpen] = useState(false);
  const [documentMenuOpen, setDocumentMenuOpen] = useState(false);
  const [pendingChange, setPendingChange] = useState<PendingChange | null>(
    null,
  );
  const [promptRequest, setPromptRequest] = useState<PromptRequest | null>(null);
  const [docStatus, setDocStatus] = useState<
    "Draft" | "In review" | "Approved"
  >("Draft");
  const [customPlaceholders, setCustomPlaceholders] = useState<
    CustomPlaceholder[]
  >([]);
  const [docName, setDocName] = useState("Marcus Lee · Fixed-Term Agreement");
  const [currentRole, setCurrentRole] = useState<UserRole>("Admin");
  const [roleMenuOpen, setRoleMenuOpen] = useState(false);
  const [theme, setTheme] = useState<"light" | "dark">(() => {
    const savedTheme = localStorage.getItem("hr-doc-generator-theme");
    return savedTheme === "dark" ? "dark" : "light";
  });
  const [connectionStatus, setConnectionStatus] = useState<"Connected" | "Syncing" | "Failed">("Connected");
  const [connectionError, setConnectionError] = useState("");
  const [lastSyncAt, setLastSyncAt] = useState("");
  // Last accepted source snapshot per onboarding submission. Draft edits and
  // pending source updates stay separate so approved files can be restored
  // without silently rewriting onboarding data.
  const [sourceSnapshots, setSourceSnapshots] = useState<Record<string, Record<string, string>>>({});
  const [sourceUpdates, setSourceUpdates] = useState<Record<string, Record<string, string>>>({});
  const [sourceUpdateMeta, setSourceUpdateMeta] = useState<Record<string, { submissionId: string; detectedAt: string }>>({});
  const [manualOverrides, setManualOverrides] = useState<string[]>([]);
  const [showSyncDiff, setShowSyncDiff] = useState(false);
  const [selectedSyncKeys, setSelectedSyncKeys] = useState<string[]>([]);
  const [versions, setVersions] = useState<DocumentVersion[]>([]);
  const [compareVersionId, setCompareVersionId] = useState<string | null>(null);
  const [restoreVersionId, setRestoreVersionId] = useState<string | null>(null);
  const [templateVersions, setTemplateVersions] = useState<Record<string, number>>({});
  const [versionCounter, setVersionCounter] = useState(0);
  const [dependencyTarget, setDependencyTarget] = useState<string | null>(null);
  const [exportState, setExportState] = useState<"idle" | "exporting" | "error">("idle");
  const [exportError, setExportError] = useState("");
  const [zoom, setZoom] = useState(92);
  const [fontSize, setFontSize] = useState(11);
  const [format, setFormat] = useState({
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    align: "left",
  });
  const editorRef = useRef<HTMLDivElement>(null);
  const letterheadInputRef = useRef<HTMLInputElement>(null);
  const placeholderImportInputRef = useRef<HTMLInputElement>(null);
  const moduleDrawerRef = useRef<HTMLElement | null>(null);
  const activeEditorRef = useRef<HTMLDivElement | null>(null);
  const savedRangeRef = useRef<Range | null>(null);
  const versionCounterRef = useRef(0);
  const selectionBookmarkRef = useRef<{
    clauseId: string;
    start: number;
    end: number;
    text: string;
  } | null>(null);

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("hr-doc-generator-theme", theme);
  }, [theme]);

  useEffect(() => {
    const onHashChange = () => setActiveModule(moduleFromHash(window.location.hash));
    window.addEventListener("hashchange", onHashChange);
    if (!window.location.hash) window.history.replaceState(null, "", "#workspace");
    return () => window.removeEventListener("hashchange", onHashChange);
  }, []);

  useEffect(() => {
    if (!moduleDrawer) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setModuleDrawer(null);
    };
    window.addEventListener("keydown", onKeyDown);
    window.setTimeout(() => moduleDrawerRef.current?.focus(), 0);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [moduleDrawer]);

  useEffect(() => {
    try {
      localStorage.setItem(navigationKey, JSON.stringify({
        module: activeModule,
        sidebarCollapsed,
        filters: appStore.navigation.filters,
      } satisfies NavigationState));
      setAppStore((current) => ({
        ...current,
        navigation: { ...current.navigation, module: activeModule, sidebarCollapsed },
      }));
    } catch {
      setModuleNotice("Navigation preference could not be saved");
    }
  }, [activeModule, sidebarCollapsed]);

  useEffect(() => {
    try {
      localStorage.setItem(appStoreKey, JSON.stringify(appStore));
    } catch {
      setModuleNotice("Some module changes could not be saved in browser storage");
    }
  }, [appStore]);

  const template =
    templates.find((item) => item.id === templateId) || templates[1];
  const person = people.find((item) => item.id === personId) || people[0];
  const onboardingKeys = useMemo(
    () => new Set(Object.keys(person.fields)),
    [person.fields],
  );
  const hasTemplateEdits = useMemo(() => {
    const normalize = (items: Section[]) =>
      items.map((section) => ({
        id: section.id,
        title: section.title,
        clauses: section.clauses.map((clause) => ({
          id: clause.id,
          title: clause.title,
          text: clause.text,
          html: clause.html || "",
          included: clause.included,
          tag: clause.tag,
          example: clause.example,
        })),
      }));
    return JSON.stringify(normalize(sections)) !== JSON.stringify(normalize(template.sections));
  }, [sections, template.sections]);
  const hasValueOverrides = useMemo(
    () =>
      Object.keys(values).some(
        (key) => onboardingKeys.has(key) && values[key] !== person.fields[key],
      ),
    [onboardingKeys, person.fields, values],
  );
  const activeLetterheadLayout: LetterheadLayout =
    letterheadEditorPage === "first"
      ? letterhead.firstPage || letterhead
      : letterhead.subsequentPage || letterhead;
  const canEditDocument = currentRole !== "Reviewer" && docStatus !== "Approved";
  const guardEdit = () => {
    if (canEditDocument) return true;
    setToast(docStatus === "Approved" ? "Approved documents are locked. Restore as a new draft to edit." : "Reviewer access is read-only.");
    window.setTimeout(() => setToast(""), 2400);
    return false;
  };
  const updateActiveLetterhead = (patch: Partial<LetterheadLayout>) =>
    (guardEdit() && setLetterhead((current) => {
      const key = letterheadEditorPage === "first" ? "firstPage" : "subsequentPage";
      const base = current[key] || current;
      const nextLayout = { ...base, ...patch };
      // “Every page” is an explicit shared mode: editing either page must
      // update both layouts so the next page cannot silently diverge.
      if (current.mode === "all") {
        return {
          ...current,
          ...patch,
          firstPage: { ...(current.firstPage || current), ...patch },
          subsequentPage: { ...(current.subsequentPage || current), ...patch },
        };
      }
      return { ...current, ...patch, [key]: nextLayout };
    }));
  const syncDiffs = useMemo<SyncDiff[]>(() => {
    const incoming = sourceUpdates[personId] || {};
    return Object.keys(incoming)
      .filter((key) => incoming[key] !== values[key])
      .map((key) => ({ key, oldValue: values[key] || "", newValue: incoming[key] }));
  }, [personId, sourceUpdates, values]);
  const hasUnreviewedSync = syncDiffs.length > 0;
  const compareVersion = versions.find((version) => version.id === compareVersionId) || null;
  const restoreVersion = versions.find((version) => version.id === restoreVersionId) || null;
  const compareChanges = useMemo(() => {
    if (!compareVersion) return [] as Array<{ label: string; before: string; after: string }>;
    const changes: Array<{ label: string; before: string; after: string }> = [];
    const allKeys = new Set([
      ...Object.keys(compareVersion.values),
      ...Object.keys(values),
    ]);
    allKeys.forEach((key) => {
      const before = compareVersion.values[key] || "Empty";
      const after = values[key] || "Empty";
      if (before !== after) changes.push({ label: key.replaceAll("_", " "), before, after });
    });
    const beforeClauses = new Map(compareVersion.sections.flatMap((s) => s.clauses).map((c) => [c.id, c]));
    const afterClauses = new Map(sections.flatMap((s) => s.clauses).map((c) => [c.id, c]));
    new Set([...beforeClauses.keys(), ...afterClauses.keys()]).forEach((id) => {
      const before = beforeClauses.get(id);
      const after = afterClauses.get(id);
      const beforeState = before ? `${before.title} · ${before.included ? "included" : "excluded"}` : "Removed";
      const afterState = after ? `${after.title} · ${after.included ? "included" : "excluded"}` : "Removed";
      if (beforeState !== afterState) changes.push({ label: "Clause", before: beforeState, after: afterState });
    });
    if (compareVersion.letterhead.mode !== letterhead.mode) {
      changes.push({ label: "Letterhead mode", before: compareVersion.letterhead.mode, after: letterhead.mode });
    }
    if (compareVersion.docName !== docName) {
      changes.push({ label: "Document name", before: compareVersion.docName, after: docName });
    }
    return changes;
  }, [compareVersion, docName, letterhead.mode, sections, values]);
  const includedClauses = useMemo(
    () =>
      sections
        .flatMap((section) => section.clauses)
        .filter((clause) => clause.included),
    [sections],
  );
  const shouldShowClause = (clause: Clause) => {
    if (!clause.included) return false;
    if (templateId === "freelance" && clause.id === "free-4")
      return values.billing_method === "milestone";
    if (templateId === "fixed" && clause.id === "fixed-4")
      return values.probation_period !== "Not applicable";
    return true;
  };

  useEffect(() => {
    const saved = localStorage.getItem(storageKey);
    if (saved) {
      try {
        const state = JSON.parse(saved);
        setTemplateId(state.templateId || "fixed");
        setPersonId(state.personId || "EMP-2041");
        setValues(state.values || people[1].fields);
        setLetterhead(normalizeLetterhead(state.letterhead));
        setSections(state.sections || clone(demoTemplates[1].sections));
        setDocName(state.docName || "Marcus Lee · Fixed-Term Agreement");
        setDocStatus(state.docStatus || "Draft");
        setWorkflowStage(state.workflowStage || "build");
        setCustomPlaceholders(state.customPlaceholders || []);
        setCurrentRole(state.currentRole || "Admin");
        setConnectionStatus(state.connectionStatus || "Connected");
        setConnectionError(state.connectionError || "");
        setLastSyncAt(state.lastSyncAt || "");
        setLastSavedAt(state.lastSavedAt || "");
        setSourceSnapshots(state.sourceSnapshots || {});
        setSourceUpdates(state.sourceUpdates || {});
        setSourceUpdateMeta(state.sourceUpdateMeta || {});
        setManualOverrides(state.manualOverrides || []);
        setVersions(state.versions || []);
        setTemplateVersions(state.templateVersions || {});
        const restoredCounter = Number(state.versionCounter || state.versions?.length || 0);
        versionCounterRef.current = restoredCounter;
        setVersionCounter(restoredCounter);
      } catch {
        setSaving("error");
        setToast("Saved workspace could not be read. Start a fresh draft or clear local data.");
      }
    }
  }, []);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      setSaving("saving");
      window.setTimeout(() => {
        try {
          const savedStamp = formatDocumentStamp();
          localStorage.setItem(
            storageKey,
            JSON.stringify({
              templates,
              templateId,
              personId,
              values,
              letterhead,
              sections,
              docName,
              docStatus,
              workflowStage,
              customPlaceholders,
              currentRole,
              connectionStatus,
              connectionError,
              lastSyncAt,
              lastSavedAt: savedStamp,
              sourceSnapshots,
              sourceUpdates,
              sourceUpdateMeta,
              manualOverrides,
              versions,
              templateVersions,
              versionCounter,
            }),
          );
          setLastSavedAt(savedStamp);
          setSaving("saved");
        } catch {
          setSaving("error");
          setToast("Auto-save failed. Use Save to retry.");
          window.setTimeout(() => setToast(""), 2400);
        }
      }, 350);
    }, 800);
    return () => window.clearTimeout(timer);
  }, [
    templates,
    templateId,
    personId,
    values,
    letterhead,
    sections,
    docName,
    docStatus,
    workflowStage,
    customPlaceholders,
    currentRole,
    connectionStatus,
    connectionError,
    lastSyncAt,
    sourceSnapshots,
    sourceUpdates,
    sourceUpdateMeta,
    manualOverrides,
    versions,
    templateVersions,
    versionCounter,
  ]);

  const resolve = (text: string) =>
    text.replace(
      /{{(.*?)}}/g,
      (_, key) => values[key.trim()] || `{{${key.trim()}}}`,
    );
  const escapeHtml = (text: string) =>
    text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  const clauseHtml = (clause: Clause) => {
    const base =
      clause.html ||
      escapeHtml(clause.text).replace(/{{(.*?)}}/g, (_, rawKey) => {
        const key = rawKey.trim();
        return `<span class="placeholder-token" data-placeholder="${escapeHtml(key)}" contenteditable="false">${escapeHtml(values[key] || `{{${key}}}`)}</span>`;
      });
    const holder = document.createElement("div");
    holder.innerHTML = base;
    holder
      .querySelectorAll<HTMLElement>("[data-placeholder]")
      .forEach((node) => {
        const key = node.dataset.placeholder || "";
        const value = values[key] || `{{${key}}}`;
        const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
        const textNodes: Text[] = [];
        let current = walker.nextNode();
        while (current) {
          textNodes.push(current as Text);
          current = walker.nextNode();
        }
        if (textNodes.length) {
          textNodes[0].data = value;
          textNodes.slice(1).forEach((textNode) => (textNode.data = ""));
        } else node.textContent = value;
      });
    return holder.innerHTML;
  };
  const clausePlainText = (clause: Clause) => {
    const holder = document.createElement("div");
    holder.innerHTML = clauseHtml(clause);
    return holder.textContent || "";
  };
  const clauseTextRuns = (clause: Clause) => {
    const holder = document.createElement("div");
    holder.innerHTML = clauseHtml(clause);
    const runs: TextRun[] = [];
    const walk = (
      node: Node,
      style: {
        bold?: boolean;
        italics?: boolean;
        underline?: boolean;
        strike?: boolean;
      } = {},
    ) => {
      if (node.nodeType === Node.TEXT_NODE && node.textContent) {
        runs.push(
          new TextRun({
            text: node.textContent,
            bold: style.bold,
            italics: style.italics,
            underline: style.underline ? {} : undefined,
            strike: style.strike,
          }),
        );
        return;
      }
      if (!(node instanceof HTMLElement)) return;
      if (node.tagName === "BR") {
        runs.push(new TextRun({ break: 1 }));
        return;
      }
      const next = {
        bold:
          style.bold ||
          ["B", "STRONG"].includes(node.tagName) ||
          node.style.fontWeight === "bold" ||
          Number(node.style.fontWeight) >= 600,
        italics:
          style.italics ||
          ["I", "EM"].includes(node.tagName) ||
          node.style.fontStyle === "italic",
        underline:
          style.underline ||
          node.tagName === "U" ||
          node.style.textDecoration.includes("underline"),
        strike:
          style.strike ||
          ["S", "STRIKE", "DEL"].includes(node.tagName) ||
          node.style.textDecoration.includes("line-through"),
      };
      node.childNodes.forEach((child) => walk(child, next));
    };
    holder.childNodes.forEach((child) => walk(child));
    return runs.length ? runs : [new TextRun(clausePlainText(clause))];
  };
  const rebindSections = (items: Section[]) =>
    items.map((section) => ({
      ...section,
      clauses: section.clauses.map((clause) => {
        const keys = [...clause.text.matchAll(/{{(.*?)}}/g)].map((match) =>
          match[1].trim(),
        );
        const hasAllTokens =
          !!clause.html &&
          keys.every((key) => clause.html?.includes(`data-placeholder="${key}"`));
        return keys.length && !hasAllTokens ? { ...clause, html: undefined } : clause;
      }),
    }));
  const updateClauseHtml = (id: string, html: string) =>
    guardEdit() && setSections((prev) =>
      prev.map((section) => ({
        ...section,
        clauses: section.clauses.map((clause) =>
          clause.id === id ? { ...clause, html } : clause,
        ),
      })),
    );
  const rememberSelection = () => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    const anchor =
      range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
        ? (range.commonAncestorContainer as Element)
        : range.commonAncestorContainer.parentElement;
    const editor = anchor?.closest(
      ".editable-paragraph",
    ) as HTMLDivElement | null;
    if (!editor) return;
    activeEditorRef.current = editor;
    savedRangeRef.current = range.cloneRange();
    const clauseId = editor.dataset.clauseId;
    if (clauseId) {
      const before = range.cloneRange();
      before.selectNodeContents(editor);
      before.setEnd(range.startContainer, range.startOffset);
      const start = before.toString().length;
      selectionBookmarkRef.current = {
        clauseId,
        start,
        end: start + range.toString().length,
        text: range.toString(),
      };
      setActiveClause(clauseId);
    }
    const nextFormat = {
      bold: document.queryCommandState("bold"),
      italic: document.queryCommandState("italic"),
      underline: document.queryCommandState("underline"),
      strike: document.queryCommandState("strikeThrough"),
      align: document.queryCommandState("justifyCenter")
        ? "center"
        : document.queryCommandState("justifyRight")
          ? "right"
          : document.queryCommandState("justifyFull")
            ? "justify"
            : "left",
    };
    setFormat((current) =>
      current.bold === nextFormat.bold &&
      current.italic === nextFormat.italic &&
      current.underline === nextFormat.underline &&
      current.strike === nextFormat.strike &&
      current.align === nextFormat.align
        ? current
        : nextFormat,
    );
  };
  const applyEditorCommand = (command: string, value?: string): boolean => {
    if (!guardEdit()) return false;
    const historyCommand = command === "undo" || command === "redo";
    const blockCommand = [
      "justifyLeft",
      "justifyCenter",
      "justifyRight",
      "justifyFull",
      "insertUnorderedList",
      "insertOrderedList",
    ].includes(command);
    let bookmark = selectionBookmarkRef.current;
    const rememberedBookmark = bookmark;
    let activeRange: Range | null = null;
    let editor = rememberedBookmark
      ? (Array.from(
          document.querySelectorAll<HTMLDivElement>(".editable-paragraph"),
        ).find(
          (node) => node.dataset.clauseId === rememberedBookmark.clauseId,
        ) ?? null)
      : activeEditorRef.current;

    // Toolbar clicks can happen without the editor's mouseup handler firing
    // (for example, when a browser automation or keyboard preserves a live
    // selection). Resolve the current DOM selection before falling back to the
    // last remembered bookmark so formatting always targets the selected text.
    const liveSelection = window.getSelection();
    if (liveSelection && liveSelection.rangeCount > 0) {
      const liveRange = liveSelection.getRangeAt(0);
      const anchor =
        liveRange.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
          ? (liveRange.commonAncestorContainer as Element)
          : liveRange.commonAncestorContainer.parentElement;
      const liveEditor = anchor?.closest(
        ".editable-paragraph",
      ) as HTMLDivElement | null;
      if (liveEditor) {
        const before = liveRange.cloneRange();
        before.selectNodeContents(liveEditor);
        before.setEnd(liveRange.startContainer, liveRange.startOffset);
        const start = before.toString().length;
        editor = liveEditor;
        activeRange = liveRange.cloneRange();
        bookmark = {
          clauseId: liveEditor.dataset.clauseId || "",
          start,
          end: start + liveRange.toString().length,
          text: liveRange.toString(),
        };
        selectionBookmarkRef.current = bookmark;
        activeEditorRef.current = liveEditor;
      }
    }
    if (!activeRange && savedRangeRef.current && editor) {
      const rememberedRange = savedRangeRef.current;
      const startNode = rememberedRange.startContainer;
      const endNode = rememberedRange.endContainer;
      if (editor.contains(startNode) && editor.contains(endNode)) {
        activeRange = rememberedRange.cloneRange();
      }
    }
    if (historyCommand) {
      if (activeRange && editor) {
        editor.focus({ preventScroll: true });
        const selection = window.getSelection();
        selection?.removeAllRanges();
        selection?.addRange(activeRange);
      }
      document.execCommand(command, false);
      if (editor?.dataset.clauseId) {
        updateClauseHtml(editor.dataset.clauseId, editor.innerHTML);
      }
      return true;
    }
    if (
      !editor ||
      (!activeRange && !bookmark) ||
      (!blockCommand &&
        ((activeRange && activeRange.collapsed) ||
          (!activeRange && bookmark?.start === bookmark?.end)))
    ) {
      setToast("Select text in the document first");
      window.setTimeout(() => setToast(""), 1800);
      return false;
    }
    const locateOffset = (offset: number) => {
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
      let consumed = 0;
      let node = walker.nextNode() as Text | null;
      let last: Text | null = null;
      while (node) {
        last = node;
        const next = consumed + node.data.length;
        if (offset <= next)
          return { node, offset: Math.max(0, offset - consumed) };
        consumed = next;
        node = walker.nextNode() as Text | null;
      }
      return last ? { node: last, offset: last.data.length } : null;
    };
    const range = activeRange || document.createRange();
    if (!activeRange) {
      if (!bookmark) return false;
      const startPoint = locateOffset(bookmark.start);
      const endPoint = locateOffset(bookmark.end);
      if (!startPoint || !endPoint) return false;
      range.setStart(startPoint.node, startPoint.offset);
      range.setEnd(endPoint.node, endPoint.offset);
    }
    // Placeholder tokens are intentionally locked. Mouse selection can still
    // begin or end inside their non-editable text node, which would otherwise
    // make Range.extractContents() split the token and lose its binding. Snap
    // those boundaries to the complete token before applying inline styles.
    const closestPlaceholder = (node: Node) => {
      const element =
        node.nodeType === Node.ELEMENT_NODE
          ? (node as Element)
          : node.parentElement;
      return element?.closest("[data-placeholder]") as HTMLElement | null;
    };
    const startPlaceholder = closestPlaceholder(range.startContainer);
    const endPlaceholder = closestPlaceholder(range.endContainer);
    if (startPlaceholder) {
      if (endPlaceholder === startPlaceholder) range.selectNode(startPlaceholder);
      else range.setStartBefore(startPlaceholder);
    }
    if (endPlaceholder && endPlaceholder !== startPlaceholder) {
      range.setEndAfter(endPlaceholder);
    }
    editor.focus({ preventScroll: true });
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    const inlineFormats: Record<
      string,
      { tag: string; style?: Partial<CSSStyleDeclaration>; href?: string }
    > = {
      bold: { tag: "strong" },
      italic: { tag: "em" },
      underline: { tag: "u" },
      strikeThrough: { tag: "s" },
      hiliteColor: {
        tag: "mark",
        style: { backgroundColor: value || "#fff0a6" },
      },
      fontName: { tag: "span", style: { fontFamily: value || "Aptos" } },
      fontSize: {
        tag: "span",
        style: {
          fontSize: value === "2" ? "10px" : value === "4" ? "16px" : "13px",
        },
      },
      createLink: { tag: "a", href: value },
    };
    const inline = inlineFormats[command];
    const removableSelectors: Record<string, string> = {
      bold: "strong,b",
      italic: "em,i",
      underline: "u",
      strikeThrough: "s,strike,del",
      hiliteColor: "mark",
    };
    const removeSelector =
      command === "removeFormat"
        ? "strong,b,em,i,u,s,strike,del,mark,a,font,span:not(.placeholder-token)"
        : removableSelectors[command] && document.queryCommandState(command)
          ? removableSelectors[command]
          : "";
    if (removeSelector && !range.collapsed) {
      const startElement =
        range.startContainer.nodeType === Node.ELEMENT_NODE
          ? (range.startContainer as Element)
          : range.startContainer.parentElement;
      const endElement =
        range.endContainer.nodeType === Node.ELEMENT_NODE
          ? (range.endContainer as Element)
          : range.endContainer.parentElement;
      const startFormat = startElement?.closest(removeSelector) || null;
      const endFormat = endElement?.closest(removeSelector) || null;
      const inserted: Node[] = [];
      if (startFormat && startFormat === endFormat) {
        const beforeRange = document.createRange();
        beforeRange.selectNodeContents(startFormat);
        beforeRange.setEnd(range.startContainer, range.startOffset);
        const afterRange = document.createRange();
        afterRange.selectNodeContents(startFormat);
        afterRange.setStart(range.endContainer, range.endOffset);
        const before = beforeRange.cloneContents();
        const selected = range.cloneContents();
        const after = afterRange.cloneContents();
        const replacement = document.createDocumentFragment();
        const appendFormatted = (fragment: DocumentFragment) => {
          if (
            !fragment.textContent &&
            !fragment.querySelector("br,img,table,hr")
          )
            return;
          const shell = startFormat.cloneNode(false) as Element;
          shell.appendChild(fragment);
          replacement.appendChild(shell);
        };
        appendFormatted(before);
        const cleanContainer = document.createElement("div");
        cleanContainer.appendChild(selected);
        cleanContainer.querySelectorAll(removeSelector).forEach((node) => {
          node.replaceWith(...Array.from(node.childNodes));
        });
        while (cleanContainer.firstChild) {
          inserted.push(cleanContainer.firstChild);
          replacement.appendChild(cleanContainer.firstChild);
        }
        appendFormatted(after);
        startFormat.replaceWith(replacement);
      } else {
        const container = document.createElement("div");
        container.appendChild(range.extractContents());
        container.querySelectorAll(removeSelector).forEach((node) => {
          node.replaceWith(...Array.from(node.childNodes));
        });
        const cleaned = document.createDocumentFragment();
        while (container.firstChild) {
          inserted.push(container.firstChild);
          cleaned.appendChild(container.firstChild);
        }
        range.insertNode(cleaned);
      }
      if (inserted.length) {
        const nextRange = document.createRange();
        nextRange.setStartBefore(inserted[0]);
        nextRange.setEndAfter(inserted[inserted.length - 1]);
        selection?.removeAllRanges();
        selection?.addRange(nextRange);
        savedRangeRef.current = nextRange.cloneRange();
      }
    } else if (inline && !range.collapsed) {
      const wrapper = document.createElement(inline.tag);
      if (inline.style) Object.assign(wrapper.style, inline.style);
      if (inline.href && wrapper instanceof HTMLAnchorElement) {
        wrapper.href = inline.href;
        wrapper.target = "_blank";
        wrapper.rel = "noreferrer";
      }
      wrapper.appendChild(range.extractContents());
      range.insertNode(wrapper);
      const nextRange = document.createRange();
      nextRange.selectNodeContents(wrapper);
      selection?.removeAllRanges();
      selection?.addRange(nextRange);
      savedRangeRef.current = nextRange.cloneRange();
    } else {
      document.execCommand("styleWithCSS", false, "true");
      document.execCommand(command, false, value);
    }
    editor
      .querySelectorAll(
        "strong:empty,b:empty,em:empty,i:empty,u:empty,s:empty,strike:empty,del:empty,mark:empty,span:empty,a:empty,font:empty",
      )
      .forEach((node) => node.remove());
    const clauseId = editor.dataset.clauseId;
    if (clauseId) updateClauseHtml(clauseId, editor.innerHTML);
    rememberSelection();
    return true;
  };
  const toolbarMouseDown = (event: React.MouseEvent) => event.preventDefault();

  const commitPersonChange = (id: string) => {
    if (!guardEdit()) return;
    const next = people.find((item) => item.id === id);
    if (!next) return;
    setPersonId(id);
    setValues(clone(next.fields));
    setManualOverrides([]);
    setSourceSnapshots((prev) => ({
      ...prev,
      [id]: prev[id] || clone(next.fields),
    }));
    // Keep pending changes for every person. A context switch must not discard
    // a source diff that still needs an explicit HR decision.
    setSelectedSyncKeys(Object.keys(sourceUpdates[id] || {}));
    setConnectionError("");
    setConnectionStatus("Connected");
    setLastSyncAt(next.submitted);
    setSections((prev) => rebindSections(prev));
    setDocName(`${next.name} · ${template.type}`);
    setToast(`Loaded ${next.name}'s onboarding submission`);
    window.setTimeout(() => setToast(""), 2200);
  };
  const changePerson = (id: string) => {
    if (id === personId) return;
    const next = people.find((item) => item.id === id);
    if (!next) return;
    if (hasValueOverrides) {
      setPendingChange({ kind: "person", id, label: next.name });
      return;
    }
    commitPersonChange(id);
  };
  const commitTemplateChange = (id: string) => {
    if (!guardEdit()) return;
    const next = templates.find((item) => item.id === id);
    if (!next) return;
    setTemplateId(id);
    setSections(clone(next.sections));
    setActiveSection(next.sections[0].id);
    setActiveClause(next.sections[0].clauses[0]?.id || "");
    setDocName(`${person.name} · ${next.type}`);
    setToast(`Template switched to ${next.name}`);
    window.setTimeout(() => setToast(""), 2200);
  };
  const changeTemplate = (id: string) => {
    if (id === templateId) return;
    const next = templates.find((item) => item.id === id);
    if (!next) return;
    if (hasTemplateEdits) {
      setPendingChange({ kind: "template", id, label: next.type });
      return;
    }
    commitTemplateChange(id);
  };
  const confirmPendingChange = () => {
    if (!pendingChange) return;
    if (pendingChange.kind === "template") commitTemplateChange(pendingChange.id);
    else commitPersonChange(pendingChange.id);
    setPendingChange(null);
  };
  const updateValue = (key: string, value: string) =>
    guardEdit() && setValues((prev) => ({ ...prev, [key]: value }));
  const updateOnboardingValue = (key: string, value: string) => {
    if (!guardEdit()) return;
    updateValue(key, value);
    if (onboardingKeys.has(key)) {
      setManualOverrides((prev) =>
        prev.includes(key) ? prev : [...prev, key],
      );
    }
  };
  const restoreOnboardingValue = (key: string) => {
    if (!guardEdit()) return;
    if (!onboardingKeys.has(key)) return;
    const acceptedValue = (sourceSnapshots[personId] || person.fields)[key] || "";
    setValues((prev) => ({ ...prev, [key]: acceptedValue }));
    setManualOverrides((prev) => prev.filter((item) => item !== key));
    setToast(`${key.replaceAll("_", " ")} restored from the accepted onboarding value`);
    window.setTimeout(() => setToast(""), 2200);
  };
  const checkOnboardingUpdates = () => {
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      setConnectionStatus("Failed");
      setConnectionError("Your browser is offline. Reconnect before checking the form.");
      setToast("Onboarding sync unavailable while offline");
      window.setTimeout(() => setToast(""), 2400);
      return;
    }
    setConnectionStatus("Syncing");
    setConnectionError("");
    window.setTimeout(() => {
      try {
        const accepted = sourceSnapshots[personId] || person.fields;
        const nextRole = person.fields.job_title?.includes("updated")
          ? person.fields.job_title
          : `${person.fields.job_title} (updated)`;
        const latestSource = {
          ...person.fields,
          job_title: nextRole,
          start_date: person.fields.start_date,
        };
        const incoming = Object.fromEntries(
          Object.entries(latestSource).filter(([key, value]) => value !== accepted[key]),
        );
        if (!Object.keys(incoming).length) {
          setConnectionStatus("Connected");
          setLastSyncAt(formatDocumentStamp());
          setSelectedSyncKeys([]);
          setShowSyncDiff(true);
          setToast("Onboarding is up to date");
          window.setTimeout(() => setToast(""), 2200);
          return;
        }
        const detectedAt = formatDocumentStamp();
        setSourceUpdates((prev) => ({
          ...prev,
          [personId]: { ...(prev[personId] || {}), ...incoming },
        }));
        setSourceUpdateMeta((prev) => ({
          ...prev,
          [personId]: { submissionId: person.submissionId, detectedAt },
        }));
        setConnectionStatus("Connected");
        setLastSyncAt(detectedAt);
        setSelectedSyncKeys(Object.keys(incoming).filter((key) => !manualOverrides.includes(key)));
        setShowSyncDiff(true);
        setToast("New onboarding changes found. Review before applying.");
        window.setTimeout(() => setToast(""), 2600);
      } catch {
        setConnectionStatus("Failed");
        setConnectionError("The onboarding response could not be read.");
        setToast("Onboarding sync failed. Retry when the source is available.");
        window.setTimeout(() => setToast(""), 2600);
      }
    }, 500);
  };
  const retryConnection = () => {
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      setConnectionStatus("Failed");
      setConnectionError("Your browser is still offline.");
      return;
    }
    setConnectionStatus("Syncing");
    setConnectionError("");
    window.setTimeout(() => {
      setConnectionStatus("Connected");
      setLastSyncAt(formatDocumentStamp());
      setToast("Onboarding connection restored");
      window.setTimeout(() => setToast(""), 2200);
    }, 500);
  };
  const applySelectedSync = () => {
    if (!guardEdit()) return;
    const incoming = sourceUpdates[personId] || {};
    // Protected fields are unchecked by default. If HR explicitly checks one,
    // that is an intentional decision to replace the local override.
    const keysToApply = selectedSyncKeys;
    setValues((prev) => {
      const next = { ...prev };
      keysToApply.forEach((key) => {
        if (incoming[key] !== undefined) next[key] = incoming[key];
      });
      return next;
    });
    setSourceSnapshots((prev) => ({
      ...prev,
      [personId]: {
        ...(prev[personId] || person.fields),
        ...Object.fromEntries(keysToApply.map((key) => [key, incoming[key]])),
      },
    }));
    setManualOverrides((prev) => prev.filter((key) => !keysToApply.includes(key)));
    setSourceUpdates((prev) => {
      const next = { ...prev };
      const pending = { ...(next[personId] || {}) };
      keysToApply.forEach((key) => delete pending[key]);
      next[personId] = pending;
      return next;
    });
    setShowSyncDiff(false);
    setSelectedSyncKeys([]);
    const includedOverrides = keysToApply.filter((key) => manualOverrides.includes(key)).length;
    setToast(
      keysToApply.length
        ? includedOverrides
          ? `Applied ${keysToApply.length} updates, including ${includedOverrides} manual override`
          : "Selected onboarding changes applied to this draft"
        : "No onboarding changes selected",
    );
    window.setTimeout(() => setToast(""), 2400);
  };
  const dismissSyncDiff = () => {
    // Keep the incoming snapshot pending. Dismissing the modal is not the
    // same as rejecting source changes; HR can reopen the diff later.
    setShowSyncDiff(false);
    setToast("Onboarding changes remain pending for this draft");
    window.setTimeout(() => setToast(""), 2400);
  };
  const setLetterheadMode = (mode: Letterhead["mode"]) => {
    if (!guardEdit()) return;
    setLetterhead((current) => {
      if (mode !== "all") return { ...current, mode };
      const source = current.firstPage || current;
      return {
        ...current,
        mode,
        firstPage: clone(source),
        subsequentPage: clone(source),
      };
    });
    setToast(mode === "all" ? "Every page now uses the same Letterhead layout" : mode === "first" ? "Letterhead applied to the first page only" : "First and subsequent pages use independent layouts");
    window.setTimeout(() => setToast(""), 2400);
  };
  const updateBillingMethod = (billingMethod: string) => {
    if (!guardEdit()) return;
    const defaults = person.fields;
    setValues((prev) => ({
      ...prev,
      billing_method: billingMethod,
      project_fee:
        billingMethod === defaults.billing_method ? defaults.project_fee : "",
      payment_milestones:
        billingMethod === defaults.billing_method
          ? defaults.payment_milestones
          : "",
    }));
  };
  const toggleClause = (clauseId: string) =>
    guardEdit() && setSections((prev) =>
      prev.map((section) => ({
        ...section,
        clauses: section.clauses.map((clause) =>
          clause.id === clauseId
            ? { ...clause, included: !clause.included }
            : clause,
        ),
      })),
    );
  const addClause = (sectionId: string, custom = false) => {
    if (!guardEdit()) return;
    const id = `${sectionId}-${Date.now()}`;
    const clause: Clause = {
      id,
      title: custom ? "New custom paragraph" : "Confidentiality addendum",
      tag: custom ? "Custom" : "Library",
      included: true,
      text: custom
        ? "Add your own paragraph here. Use {{full_name}} and other placeholders from the insert menu."
        : "The parties will treat all non-public information as confidential and use it only for the purposes of this Agreement.",
    };
    setSections((prev) =>
      prev.map((section) =>
        section.id === sectionId
          ? { ...section, clauses: [...section.clauses, clause] }
          : section,
      ),
    );
    setActiveClause(id);
  };
  const addSection = () => {
    if (!guardEdit()) return;
    const id = `section-${Date.now()}`;
    setSections((prev) => [
      ...prev,
      {
        id,
        title: "New section",
        clauses: [
          {
            id: `${id}-clause`,
            title: "New paragraph",
            tag: "Custom",
            included: true,
            text: "Start writing your paragraph here.",
          },
        ],
      },
    ]);
    setActiveSection(id);
  };
  const extractPlaceholderKeys = (text: string) =>
    [...text.matchAll(/{{(.*?)}}/g)].map((match) => match[1].trim());
  const documentIssues = useMemo(() => {
    const issues: string[] = [];
    const missing = new Set<string>();
    const active = sections.flatMap((section) =>
      section.clauses.filter(shouldShowClause),
    );
    active.forEach((clause) => {
      extractPlaceholderKeys(clause.text).forEach((key) => {
        if (!values[key]?.trim()) missing.add(key);
      });
      if (clausePlainText(clause).includes("{{")) {
        issues.push(`Unresolved placeholder in “${clause.title}”`);
      }
    });
    if (missing.size) {
      issues.unshift(`Missing required fields: ${[...missing].join(", ")}`);
    }
    const otherNames = people
      .filter((candidate) => candidate.id !== personId)
      .flatMap((candidate) => [candidate.name, candidate.role]);
    const staleIdentity = active.some((clause) => {
      const text = clausePlainText(clause);
      return otherNames.some(
        (value) => value && text.includes(value) && value !== person.name && value !== person.role,
      );
    });
    if (staleIdentity) {
      issues.push("Document contains another person's name or role");
    }
    const billingClause = active.find((clause) => clause.id === "free-3");
    if (
      templateId === "freelance" &&
      values.billing_method === "milestone" &&
      billingClause &&
      /hour/i.test(clausePlainText(billingClause))
    ) {
      issues.push("Milestone billing still contains an hourly amount");
    }
    const firstLayout = letterhead.firstPage || letterhead;
    if (firstLayout.margin <= firstLayout.top + 8) {
      issues.push("Letterhead overlaps the document body");
    }
    return issues;
  }, [letterhead, person.id, person.name, person.role, personId, sections, templateId, values]);
  const blockingIssues = documentIssues;
  const workflowSummary = useMemo(() => {
    if (docStatus === "Approved") {
      return {
        stage: "export" as WorkflowStage,
        targetStage: "export" as WorkflowStage,
        title: "Approved and ready to export",
        detail: "This version is locked. Export the approved copy or restore it as a new draft.",
        tone: "success" as const,
        action: "Open export",
      };
    }
    if (docStatus === "In review") {
      return {
        stage: "review" as WorkflowStage,
        targetStage: "review" as WorkflowStage,
        title: "Waiting for review",
        detail: currentRole === "Reviewer" ? "Review the checks below, then approve this version." : "An editor has submitted this version for approval.",
        tone: "review" as const,
        action: currentRole === "Reviewer" ? "Open review" : "View submission",
      };
    }
    if (blockingIssues.length) {
      return {
        stage: "build" as WorkflowStage,
        targetStage: "review" as WorkflowStage,
        title: `${blockingIssues.length} item${blockingIssues.length === 1 ? "" : "s"} to resolve`,
        detail: "Finish the highlighted fields and layout checks before sending this document for review.",
        tone: "warning" as const,
        action: "Review checks",
      };
    }
    if (workflowStage === "review") {
      return {
        stage: "review" as WorkflowStage,
        targetStage: "build" as WorkflowStage,
        title: "Review this draft",
        detail: "Check required fields, clause inclusion and the final page layout before submission.",
        tone: "review" as const,
        action: "Back to build",
      };
    }
    if (workflowStage === "export") {
      return {
        stage: "export" as WorkflowStage,
        targetStage: "export" as WorkflowStage,
        title: "Export this draft",
        detail: "The preview is the final source of truth for Word and PDF output.",
        tone: "ready" as const,
        action: "Open export",
      };
    }
    return {
      stage: "build" as WorkflowStage,
      targetStage: "review" as WorkflowStage,
      title: "Ready for review",
      detail: "Source data, clauses and layout checks are complete.",
      tone: "ready" as const,
      action: "Open preview",
    };
  }, [blockingIssues.length, currentRole, docStatus, workflowStage]);
  const effectiveWorkflowStage: WorkflowStage = docStatus === "Approved"
    ? "export"
    : docStatus === "In review"
      ? "review"
      : workflowStage;
  const workflowStageIndex = effectiveWorkflowStage === "build" ? 0 : effectiveWorkflowStage === "review" ? 1 : 2;
  const goToWorkflowStage = (stage: WorkflowStage) => {
    setWorkflowStage(stage);
    if (stage === "build") {
      setActiveRightTab("person");
      setShowPreview(false);
      return;
    }
    setShowPreview(true);
  };
  const openNewBlankDocument = () => {
    if (!guardEdit()) return;
    const id = `blank-${Date.now()}`;
    setSections([
      {
        id,
        title: "New section",
        clauses: [
          {
            id: `${id}-clause`,
            title: "New paragraph",
            tag: "Custom",
            included: true,
            text: "Start writing your paragraph here.",
          },
        ],
      },
    ]);
    setActiveSection(id);
    setActiveClause(`${id}-clause`);
    setDocName("Untitled HR document");
    setDocumentMenuOpen(false);
  };
  const copyCurrentDocument = () => {
    if (!guardEdit()) return;
    setSections((prev) =>
      prev.map((section) => ({
        ...section,
        id: `${section.id}-copy-${Date.now()}`,
        clauses: section.clauses.map((clause) => ({
          ...clause,
          id: `${clause.id}-copy-${Date.now()}`,
        })),
      })),
    );
    setDocName(`${docName} · Copy`);
    setDocumentMenuOpen(false);
  };
  const updateClauseText = (id: string, text: string) =>
    (guardEdit() && setSections((prev) =>
      prev.map((section) => ({
        ...section,
        clauses: section.clauses.map((clause) =>
          clause.id === id ? { ...clause, text } : clause,
        ),
      })),
    ));
  const duplicateClause = (sectionId: string, clause: Clause) =>
    (guardEdit() && setSections((prev) =>
      prev.map((section) =>
        section.id === sectionId
          ? {
              ...section,
              clauses: [
                ...section.clauses,
                {
                  ...clone(clause),
                  id: `${clause.id}-copy-${Date.now()}`,
                  title: `${clause.title} copy`,
                },
              ],
            }
          : section,
      ),
    ));
  const moveSection = (index: number, direction: -1 | 1) =>
    guardEdit() && setSections((prev) => {
      const next = [...prev];
      const target = index + direction;
      if (target < 0 || target >= next.length) return prev;
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  const buildVersion = (
    summary: string,
    status: DocumentVersion["status"],
    payload: Partial<DocumentVersion> = {},
  ): DocumentVersion => {
    const revision = Math.max(versionCounterRef.current + 1, versionCounter + 1);
    versionCounterRef.current = revision;
    setVersionCounter(revision);
    return {
      id: `v-${Date.now()}-${revision}`,
      label: `v${revision}.0 · ${status}`,
      createdAt: formatDocumentStamp(),
      author: currentRole === "Reviewer" ? "Reviewer" : "JL",
      status,
      summary,
      sections: clone(payload.sections ?? sections),
      values: clone(payload.values ?? values),
      letterhead: clone(payload.letterhead ?? letterhead),
      docName: payload.docName ?? docName,
      templateId: payload.templateId ?? templateId,
      personId: payload.personId ?? personId,
      submissionId: payload.submissionId ?? person.submissionId,
      sourceSnapshot: clone(payload.sourceSnapshot ?? person.fields),
      sourceSyncedAt: payload.sourceSyncedAt ?? lastSyncAt,
      manualOverrides: clone(payload.manualOverrides ?? manualOverrides),
      sourceUpdates: clone(payload.sourceUpdates ?? sourceUpdates),
      sourceUpdateMeta: clone(payload.sourceUpdateMeta ?? sourceUpdateMeta),
    };
  };
  const createVersion = (
    summary: string,
    status: DocumentVersion["status"] = docStatus,
    payload: Partial<DocumentVersion> = {},
  ) => {
    const version = buildVersion(summary, status, payload);
    setVersions((prev) => [version, ...prev].slice(0, 12));
    return version;
  };
  const restoreVersionNow = (version: DocumentVersion) => {
    const restoredSections = clone(version.sections);
    const restoredValues = clone(version.values);
    const restoredLetterhead = normalizeLetterhead(version.letterhead);
    const restoredDocName = version.docName;
    const restoredPersonId = version.personId || personId;
    const restoredTemplateId = version.templateId || templateId;
    const restoredOverrides = clone(version.manualOverrides || []);
    const restoredUpdates = clone(version.sourceUpdates || {});
    const restoredMeta = clone(version.sourceUpdateMeta || sourceUpdateMeta);
    setSections(restoredSections);
    setValues(restoredValues);
    setLetterhead(restoredLetterhead);
    setDocName(restoredDocName);
    setPersonId(restoredPersonId);
    setTemplateId(restoredTemplateId);
    setActiveSection(restoredSections[0]?.id || "");
    setActiveClause(restoredSections[0]?.clauses[0]?.id || "");
    setSourceSnapshots((prev) => ({
      ...prev,
      [restoredPersonId]: clone(
        version.sourceSnapshot ||
          prev[restoredPersonId] ||
          people.find((item) => item.id === restoredPersonId)?.fields ||
          {},
      ),
    }));
    setManualOverrides(restoredOverrides);
    setSourceUpdates(restoredUpdates);
    setSourceUpdateMeta(restoredMeta);
    setDocStatus("Draft");
    const restoredSnapshot = buildVersion(`Restored from ${version.label}`, "Draft", {
      sections: restoredSections,
      values: restoredValues,
      letterhead: restoredLetterhead,
      docName: restoredDocName,
      personId: restoredPersonId,
      templateId: restoredTemplateId,
      submissionId: version.submissionId,
      sourceSnapshot: version.sourceSnapshot,
      sourceSyncedAt: version.sourceSyncedAt,
      manualOverrides: restoredOverrides,
      sourceUpdates: restoredUpdates,
      sourceUpdateMeta: restoredMeta,
    });
    setVersions((prev) => [restoredSnapshot, ...prev].slice(0, 12));
    setRestoreVersionId(null);
    setShowHistory(false);
    setToast(`${version.label} restored as a new draft`);
    window.setTimeout(() => setToast(""), 2600);
  };
  const saveNow = () => {
    try {
      const nextStatus = docStatus === "Approved" ? "Draft" : docStatus;
      const version = buildVersion(
        docStatus === "Approved" ? "Approved file edited; new draft created" : "Manual save",
        nextStatus,
      );
      const nextVersions = [version, ...versions].slice(0, 12);
      localStorage.setItem(
        storageKey,
        JSON.stringify({
          templates,
          templateId,
          personId,
          values,
          letterhead,
          sections,
          docName,
          docStatus: nextStatus,
          customPlaceholders,
          currentRole,
          connectionStatus,
          connectionError,
          lastSyncAt,
          lastSavedAt: formatDocumentStamp(),
          sourceSnapshots,
          sourceUpdates,
          sourceUpdateMeta,
          manualOverrides,
          versions: nextVersions,
          templateVersions,
          versionCounter: versionCounterRef.current,
        }),
      );
      setVersions(nextVersions);
      if (docStatus === "Approved") setDocStatus("Draft");
      setLastSavedAt(formatDocumentStamp());
      setSaving("saved");
      setToast("Saved and versioned in local workspace");
    } catch {
      setSaving("error");
      setToast("Save failed. Check browser storage and retry.");
    }
    window.setTimeout(() => setToast(""), 2400);
  };
  const createPlaceholder = (labelInput?: string) => {
    if (!guardEdit()) return;
    if (labelInput === undefined) {
      setPromptRequest({
        kind: "placeholder",
        title: "Create placeholder",
        description: "Give this field a human-readable name. Its stable key will be generated automatically.",
        value: "New HR field",
      });
      return;
    }
    const label = labelInput.trim();
    if (!label) return;
    const key =
      label
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "_")
        .replace(/^_|_$/g, "") || `field_${Date.now()}`;
    setCustomPlaceholders((prev) => [
      ...prev,
      { key, label, source: "HR input", group: "Custom" },
    ]);
    setValues((prev) => ({ ...prev, [key]: "" }));
    setToast(`Created {{${key}}}`);
    window.setTimeout(() => setToast(""), 2200);
  };
  const replacePlaceholderReferences = (oldKey: string, replacement?: string) => {
    if (currentRole !== "Admin") {
      setToast("Admin role required to replace field references");
      window.setTimeout(() => setToast(""), 2200);
      return;
    }
    if (replacement === undefined) {
      setPromptRequest({
        kind: "replace",
        title: `Replace {{${oldKey}}}`,
        description: "Enter the existing placeholder key that should receive these references.",
        value: placeholderMeta.find(([key]) => key !== oldKey)?.[0] || customPlaceholders.find((item) => item.key !== oldKey)?.key || "new_field",
        oldKey,
      });
      return;
    }
    const nextKey = replacement.trim();
    if (!nextKey || nextKey === oldKey) return;
    const tokenPattern = new RegExp(`{{\\s*${oldKey.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&")}\\s*}}`, "g");
    setSections((prev) => prev.map((section) => ({
      ...section,
      clauses: section.clauses.map((clause) => ({
        ...clause,
        text: clause.text.replace(tokenPattern, `{{${nextKey}}}`),
        html: clause.html?.replace(tokenPattern, `{{${nextKey}}}`),
      })),
    })));
    setValues((prev) => {
      if (prev[nextKey] || !prev[oldKey]) return prev;
      return { ...prev, [nextKey]: prev[oldKey] };
    });
    setDependencyTarget(null);
    setToast(`Replaced {{${oldKey}}} references with {{${nextKey}}}`);
    window.setTimeout(() => setToast(""), 2600);
  };
  const renameDocument = (nextInput?: string) => {
    if (nextInput === undefined) {
      setPromptRequest({
        kind: "rename",
        title: "Rename document",
        description: "Use a name your team can find later in document history.",
        value: docName,
      });
      return;
    }
    const next = nextInput.trim();
    if (next) setDocName(next);
  };
  const insertLink = (urlInput?: string) => {
    if (urlInput === undefined) {
      setPromptRequest({
        kind: "link",
        title: "Insert link",
        description: "Paste a secure URL for the selected text.",
        value: "https://",
      });
      return;
    }
    const url = urlInput.trim();
    if (url) applyEditorCommand("createLink", url);
  };
  const submitPromptRequest = () => {
    if (!promptRequest) return;
    const request = promptRequest;
    setPromptRequest(null);
    if (request.kind === "placeholder") createPlaceholder(request.value);
    if (request.kind === "replace" && request.oldKey) replacePlaceholderReferences(request.oldKey, request.value);
    if (request.kind === "rename") renameDocument(request.value);
    if (request.kind === "link") insertLink(request.value);
  };
  const uploadLetterhead = (file?: File) => {
    if (!guardEdit()) return;
    if (!file) return;
    const supported = [
      "image/png",
      "image/jpeg",
      "application/pdf",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ];
    if (!supported.includes(file.type)) {
      setToast("Use PNG, JPG, PDF, or DOCX for a letterhead");
      window.setTimeout(() => setToast(""), 2600);
      return;
    }
    if (file.size > 2_000_000) {
      setToast(
        "Letterhead files must be 2 MB or smaller in this local workspace",
      );
      window.setTimeout(() => setToast(""), 2800);
      return;
    }
    if (file.type.startsWith("image/")) {
      const reader = new FileReader();
      reader.onload = () =>
        updateActiveLetterhead({
          fileName: file.name,
          fileType: file.type,
          dataUrl: String(reader.result),
        });
      reader.readAsDataURL(file);
    } else
      updateActiveLetterhead({
        fileName: file.name,
        fileType: file.type,
        dataUrl: undefined,
      });
    setToast(
      file.type === "application/pdf"
        ? "PDF selected. Choose its source page when publishing."
        : file.type.includes("wordprocessingml")
          ? "DOCX selected. Header and footer content may need manual adjustment."
          : "Letterhead image uploaded and applied",
    );
    window.setTimeout(() => setToast(""), 3200);
  };
  const layoutForPage = (pageIndex: number) =>
    pageIndex === 0
      ? letterhead.firstPage || letterhead
      : letterhead.subsequentPage || letterhead.firstPage || letterhead;
  const letterheadImageType = (layout: LetterheadLayout) => {
    if (!layout.fileType?.startsWith("image/")) return null;
    if (layout.fileType.includes("jpeg") || layout.fileType.includes("jpg")) return "jpg" as const;
    if (layout.fileType.includes("gif")) return "gif" as const;
    if (layout.fileType.includes("bmp")) return "bmp" as const;
    return "png" as const;
  };
  const renderDocxLetterhead = (layout: LetterheadLayout) => {
    const bytes = dataUrlToBytes(layout.dataUrl);
    const type = letterheadImageType(layout);
    if (bytes && type) {
      return new Paragraph({
        children: [
          new ImageRun({
            type,
            data: bytes,
            transformation: {
              width: Math.max(80, Math.round(layout.width)),
              height: Math.max(24, Math.round(layout.width * 0.28)),
            },
          }),
        ],
      });
    }
    return new Paragraph({
      children: [
        new TextRun({
          text: values.company_name || "Company Letterhead",
          bold: true,
          color: layout.accent.replace("#", ""),
        }),
      ],
    });
  };
  const estimatePageCount = () => {
    const characters = sections
      .flatMap((section) => [section.title, ...section.clauses.filter(shouldShowClause).map(clausePlainText)])
      .join(" ").length;
    return Math.max(1, Math.ceil((characters + docName.length) / 1550));
  };
  const makeDocx = async () => {
    if (blockingIssues.length) {
      setToast("Fix the highlighted checks before exporting");
      window.setTimeout(() => setToast(""), 2400);
      return;
    }
    setExportState("exporting");
    setExportError("");
    try {
      const firstLayout = layoutForPage(0);
      const subsequentLayout = layoutForPage(1);
      const firstHeader = new Header({ children: [renderDocxLetterhead(firstLayout)] });
      const subsequentHeader = new Header({ children: [renderDocxLetterhead(letterhead.mode === "all" ? firstLayout : subsequentLayout)] });
      const footer = new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `${values.company_name || "Company"} · Confidential · Page ` }), new TextRun({ children: [PageNumber.CURRENT] })],
          }),
        ],
      });
      const children = [
        new Paragraph({ text: docName, heading: HeadingLevel.TITLE }),
        ...sections.flatMap((section) => [
          new Paragraph({ text: section.title, heading: HeadingLevel.HEADING_1 }),
          ...section.clauses
            .filter(shouldShowClause)
            .map((clause) => new Paragraph({ children: clauseTextRuns(clause) })),
        ]),
        new Paragraph({
          text: `Signed for ${values.company_name} by ${values.signatory_name}, ${values.signatory_title}`,
        }),
        new Paragraph({ text: "Employee / Contractor signature: ______________________________" }),
        new Paragraph({ text: "Date: ____________________" }),
      ];
      const blob = await Packer.toBlob(new Document({
        sections: [{
          headers: letterhead.mode === "first"
            ? { first: firstHeader }
            : letterhead.mode === "all"
              ? { default: subsequentHeader, first: firstHeader }
              : { default: subsequentHeader, first: firstHeader },
          footers: { default: footer, first: footer },
          properties: {
            titlePage: letterhead.mode !== "all",
            page: {
              margin: {
                top: Math.round((firstLayout.margin || 74) * 15),
                right: 1080,
                bottom: 900,
                left: 1080,
              },
            },
          },
          children,
        }],
      }));
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const fileName = `${template.type}_${person.name.replaceAll(" ", "_")}_${formatFileDate()}.docx`;
      a.download = fileName;
      a.click();
      URL.revokeObjectURL(url);
      recordExport("docx", "generated", fileName, undefined, await blobToDataUrl(blob));
      setExportState("idle");
      setToast("Word document exported");
    } catch (error) {
      setExportState("error");
      setExportError(`Word export failed: ${error instanceof Error ? error.message : "Unknown error"}`);
      recordExport("docx", "failed", `${template.type}_${person.name.replaceAll(" ", "_")}_${formatFileDate()}.docx`, error instanceof Error ? error.message : "Unknown error");
      setToast("Word export failed. Retry after resolving the error.");
    }
    window.setTimeout(() => setToast(""), 2600);
  };
  const makePdf = async () => {
    if (blockingIssues.length) {
      setToast("Fix the highlighted checks before exporting");
      window.setTimeout(() => setToast(""), 2400);
      return;
    }
    setExportState("exporting");
    setExportError("");
    try {
      const pdf = new jsPDF({ format: letterhead.page === "A4" ? "a4" : "letter", unit: "pt" });
      const pageHeight = pdf.internal.pageSize.getHeight();
      const pageWidth = pdf.internal.pageSize.getWidth();
      const addPageFurniture = (pageIndex: number) => {
        const layout = layoutForPage(pageIndex);
        const imageType = letterheadImageType(layout);
        if (layout.dataUrl && imageType) {
          const imageHeight = Math.max(24, layout.width * 0.28);
          pdf.addImage(layout.dataUrl, imageType === "jpg" ? "JPEG" : imageType.toUpperCase(), layout.left, layout.top, layout.width, imageHeight, undefined, "FAST");
        } else {
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(10);
          pdf.setTextColor(layout.accent);
          pdf.text(values.company_name || "Company Letterhead", layout.left, layout.top + 12);
          pdf.setTextColor("#222222");
        }
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(8);
        pdf.setTextColor("#6b7280");
        pdf.text(`${values.company_name || "Company"} · Confidential · Page ${pageIndex + 1}`, pageWidth / 2, pageHeight - 28, { align: "center" });
        pdf.setTextColor("#222222");
      };
      let pageIndex = 0;
      addPageFurniture(pageIndex);
      const firstLayout = layoutForPage(0);
      let y = Math.max(70, firstLayout.margin || 74);
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(16);
      pdf.text(docName, 54, y);
      y += 28;
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(10);
      sections.forEach((section) => {
        const visible = section.clauses.filter(shouldShowClause);
        if (!visible.length) return;
        if (y > pageHeight - 90) { pdf.addPage(); pageIndex += 1; addPageFurniture(pageIndex); y = Math.max(58, layoutForPage(pageIndex).margin || 58); }
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(12);
        pdf.text(section.title, 54, y);
        y += 18;
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);
        visible.forEach((clause) => {
          const lines = pdf.splitTextToSize(clausePlainText(clause), 490);
          lines.forEach((line: string) => {
            if (y > pageHeight - 70) { pdf.addPage(); pageIndex += 1; addPageFurniture(pageIndex); y = Math.max(58, layoutForPage(pageIndex).margin || 58); }
            pdf.text(line, 54, y);
            y += 14;
          });
          y += 6;
        });
      });
      if (y > pageHeight - 130) { pdf.addPage(); pageIndex += 1; addPageFurniture(pageIndex); y = Math.max(58, layoutForPage(pageIndex).margin || 58); }
      pdf.setFont("helvetica", "bold");
      pdf.text(`Signed for ${values.company_name} by ${values.signatory_name}, ${values.signatory_title}`, 54, y + 20);
      pdf.setFont("helvetica", "normal");
      pdf.text("Employee / Contractor signature: ______________________________", 54, y + 48);
      pdf.text("Date: ____________________", 54, y + 68);
      const fileName = `${template.type}_${person.name.replaceAll(" ", "_")}_${formatFileDate()}.pdf`;
      const pdfContentBase64 = await blobToDataUrl(pdf.output("blob"));
      pdf.save(fileName);
      recordExport("pdf", "generated", fileName, undefined, pdfContentBase64);
      setExportState("idle");
      setToast("PDF exported");
    } catch (error) {
      setExportState("error");
      setExportError(`PDF export failed: ${error instanceof Error ? error.message : "Unknown error"}`);
      recordExport("pdf", "failed", `${template.type}_${person.name.replaceAll(" ", "_")}_${formatFileDate()}.pdf`, error instanceof Error ? error.message : "Unknown error");
      setToast("PDF export failed. Retry after resolving the error.");
    }
    window.setTimeout(() => setToast(""), 2600);
  };

  const selectedClause = sections
    .flatMap((s) => s.clauses)
    .find((c) => c.id === activeClause);
  const reviewActionDisabled =
    (currentRole === "Reviewer" && docStatus !== "In review") ||
    (currentRole !== "Reviewer" && docStatus === "Approved");
  const adminCards: Array<
    [string, string, React.ComponentType<{ size?: number }>]
  > = [
    ["Templates", `${templates.length} active · v${templateVersions[templateId] || 1} current`, FileText],
    ["Placeholders", `${placeholderMeta.length + customPlaceholders.length} configured`, Sparkles],
    ["Dropdowns", "6 configured", ListOrdered],
    ["Connections", `Onboarding · ${connectionStatus}`, Cloud],
    ["Letterheads", "2 saved", ImagePlus],
    ["Rules", "9 active · 1 draft", WandSparkles],
  ];
  const navItems: Array<[AppModule, string, React.ComponentType<{ size?: number }>]> = [
    ["workspace", "Workspace", FileText],
    ["clauses", "Clauses Library", Archive],
    ["placeholders", "Placeholder Management", Sparkles],
    ["layouts", "Layout Management", ImagePlus],
    ["documents", "Document Library", FolderOpen],
  ];
  const currentDocumentRecord = (): DocumentRecord => ({
    id: "doc-current",
    docName,
    templateId,
    personId,
    companyId,
    status: docStatus === "In review" ? "in-review" : docStatus === "Approved" ? "approved" : "draft",
    generationStatus: appStore.documents.find((item) => item.id === "doc-current")?.generationStatus || "not-generated",
    sections: clone(sections),
    values: clone(values),
    letterhead: clone(letterhead),
    versions: clone(versions),
    exports: clone(appStore.documents.find((item) => item.id === "doc-current")?.exports || []),
    updatedAt: formatDocumentStamp(),
  });
  const persistCurrentDocumentRecord = () => {
    try {
      const record = currentDocumentRecord();
      const next: AppStore = {
        ...appStore,
        navigation: { ...appStore.navigation, module: activeModule, sidebarCollapsed },
        documents: [record, ...appStore.documents.filter((item) => item.id !== record.id)],
      };
      localStorage.setItem(appStoreKey, JSON.stringify(next));
      setAppStore(next);
      return true;
    } catch {
      setSaving("error");
      setModuleNotice("Save failed. Retry or stay on this page.");
      return false;
    }
  };
  const switchModule = (module: AppModule) => {
    if (module === activeModule) return;
    if (activeModule === "workspace" && !persistCurrentDocumentRecord()) return;
    window.location.hash = module;
    setActiveModule(module);
  };
  const moduleTitle = navItems.find(([id]) => id === activeModule)?.[1] || "Workspace";
  const filteredClauses = appStore.clauses.filter((clause) => {
    const matchesSearch = !moduleSearch || `${clause.title} ${clause.category} ${clause.tags.join(" ")}`.toLowerCase().includes(moduleSearch.toLowerCase());
    const matchesFilter = moduleFilter === "all" || clause.status === moduleFilter;
    return matchesSearch && matchesFilter;
  });
  const selectedGroup = appStore.placeholderGroups.find((group) => group.id === (selectedPlaceholderGroupId || appStore.placeholderGroups[0]?.id));
  const updateClauseDraft = (patch: Partial<ClauseRecord>) => setClauseDraft((current) => current ? { ...current, ...patch, updatedAt: formatDocumentStamp() } : current);
  const saveClauseDraft = (publish = false) => {
    if (!clauseDraft || currentRole !== "Admin" && publish) return;
    const nextClause = { ...clauseDraft, status: publish ? "published" : clauseDraft.status, version: clauseDraft.version + (publish ? 1 : 0), updatedAt: formatDocumentStamp() } as ClauseRecord;
    setAppStore((current) => ({ ...current, clauses: [nextClause, ...current.clauses.filter((item) => item.id !== nextClause.id)] }));
    setSelectedClauseRecordId(nextClause.id);
    setClauseDraft(null);
    setModuleDrawer(null);
    setModuleNotice(publish ? "Clause published for new documents" : "Clause draft saved");
  };
  const openClauseEditor = (record?: ClauseRecord) => {
    const next = record ? clone(record) : {
      id: `clause-${Date.now()}`,
      title: "New clause",
      structure: "flat" as const,
      subsections: [],
      contents: [{ id: `content-${Date.now()}`, title: "Content block", text: "Add clause text here.", included: true, tag: "Core" }],
      category: "General",
      documentTypes: [template.type],
      language: "English",
      tags: ["Draft"],
      status: "draft" as const,
      version: 1,
      updatedAt: formatDocumentStamp(),
    };
    setClauseDraft(next);
    setModuleDrawer("clause");
  };
  const addClauseContent = (subsectionId?: string) => {
    if (!clauseDraft) return;
    const item: Clause = { id: `content-${Date.now()}`, title: "New content block", text: "Write reusable content here.", included: true, tag: "Core" };
    if (clauseDraft.structure === "nested" && subsectionId) {
      updateClauseDraft({ subsections: clauseDraft.subsections.map((sub) => sub.id === subsectionId ? { ...sub, contents: [...sub.contents, item] } : sub) });
    } else updateClauseDraft({ contents: [...clauseDraft.contents, item] });
  };
  const copyClauseRecord = (record: ClauseRecord) => {
    const copy = { ...clone(record), id: `clause-${Date.now()}`, title: `${record.title} copy`, status: "draft" as const, version: 1, updatedAt: formatDocumentStamp() };
    setAppStore((current) => ({ ...current, clauses: [copy, ...current.clauses] }));
    setModuleNotice("Clause copied as a draft");
  };
  const toggleClauseStatus = (record: ClauseRecord) => {
    if (currentRole !== "Admin") return setModuleNotice("Admin role required to change clause status");
    setAppStore((current) => ({ ...current, clauses: current.clauses.map((item) => item.id === record.id ? { ...item, status: item.status === "inactive" ? "published" : "inactive", updatedAt: formatDocumentStamp() } : item) }));
  };
  const openPlaceholderEditor = (group?: PlaceholderGroup) => {
    setPlaceholderDraft(group ? clone(group) : { id: `group-${Date.now()}`, name: "New placeholder group", sourceType: "Manual", sourceTable: "custom_fields", fields: [], status: "active" });
    setModuleDrawer("placeholder");
  };
  const addPlaceholderField = () => setPlaceholderDraft((current) => current ? { ...current, fields: [...current.fields, { id: `field-${Date.now()}`, key: `custom_${current.fields.length + 1}`, label: "New field", type: "text", sourceField: "", example: "Example value", required: false, status: "active", mappingStatus: "valid", manualOverride: true }] } : current);
  const previewPlaceholderImport = async (file?: File) => {
    if (!file) return;
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
      const headers = rows.length ? Object.keys(rows[0]) : [];
      const sample = rows.slice(0, 3).map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, String(value)])));
      setImportPreview({ fileName: file.name, headers, sample });
      setPlaceholderDraft((current) => current ? { ...current, sourceType: file.name.toLowerCase().endsWith(".csv") ? "CSV import" : "Excel workbook", sourceTable: sheetName } : current);
      setModuleNotice(`${file.name} loaded for preview. Confirm the fields before saving.`);
    } catch {
      setModuleNotice("Import preview failed. Use a CSV or Excel workbook with a header row.");
    }
  };
  const confirmPlaceholderImport = () => {
    if (!importPreview) return;
    setPlaceholderDraft((current) => current ? { ...current, fields: importPreview.headers.map((header, index) => ({ id: `${current.id}-${header}-${index}`, key: header.toLowerCase().replace(/[^a-z0-9]+/g, "_"), label: header, type: "text", sourceField: header, example: importPreview.sample[0]?.[header] || "", required: false, status: "active", mappingStatus: "valid", manualOverride: false })) } : current);
    setImportPreview(null);
    setModuleNotice("Imported fields added to this placeholder group. Review and save the group.");
  };
  const savePlaceholderDraft = () => {
    if (!placeholderDraft) return;
    setAppStore((current) => ({ ...current, placeholderGroups: [placeholderDraft, ...current.placeholderGroups.filter((item) => item.id !== placeholderDraft.id)] }));
    setSelectedPlaceholderGroupId(placeholderDraft.id);
    setPlaceholderDraft(null);
    setModuleDrawer(null);
    setModuleNotice("Placeholder group saved");
  };
  const openLayoutEditor = (record?: LayoutRecord) => {
    setLayoutDraft(record ? clone(record) : { id: `layout-${Date.now()}`, name: "New letterhead", companyId, letterhead: makeDefaultLetterhead(), status: "draft", updatedAt: formatDocumentStamp() });
    setModuleDrawer("layout");
  };
  const saveLayoutDraft = (publish = false) => {
    if (!layoutDraft) return;
    const next = { ...layoutDraft, status: publish ? "published" : layoutDraft.status, updatedAt: formatDocumentStamp() } as LayoutRecord;
    setAppStore((current) => ({ ...current, layouts: [next, ...current.layouts.filter((item) => item.id !== next.id)] }));
    setLayoutDraft(null);
    setModuleDrawer(null);
    setModuleNotice(publish ? "Layout published" : "Layout saved");
  };
  const applyLayoutRecord = (record: LayoutRecord) => {
    if (!guardEdit()) return;
    setLetterhead(normalizeLetterhead(record.letterhead));
    setActiveRightTab("letterhead");
    setModuleNotice(`${record.name} applied to this document snapshot`);
  };
  const insertClauseRecord = (record: ClauseRecord) => {
    if (!guardEdit()) return;
    const sourceClauses = record.structure === "nested"
      ? record.subsections.flatMap((subsection) => subsection.contents)
      : record.contents;
    const insertedClauses = sourceClauses.map((clause) => ({
      ...clone(clause),
      id: `${record.id}-${clause.id}-${Date.now()}`,
      tag: "Library",
      included: true,
      sourceClauseId: record.id,
      sourceClauseVersion: record.version,
    }));
    const section: Section = {
      id: `library-section-${Date.now()}`,
      title: record.title,
      clauses: insertedClauses,
    };
    setSections((current) => [...current, section]);
    setActiveSection(section.id);
    setActiveClause(insertedClauses[0]?.id || "");
    setModuleNotice(`${record.title} inserted as a document copy`);
    switchModule("workspace");
  };
  const recordExport = (format: ExportRecord["format"], status: ExportRecord["status"], fileName: string, error?: string, contentBase64?: string) => {
    const exportRecord: ExportRecord = {
      id: `export-${Date.now()}-${format}`,
      documentId: "doc-current",
      versionId: versions[0]?.id,
      format,
      status,
      fileName,
      createdAt: formatDocumentStamp(),
      contentBase64,
      error,
    };
    setAppStore((current) => {
      const existing = current.documents.find((item) => item.id === "doc-current") || currentDocumentRecord();
      const nextDocument: DocumentRecord = {
        ...existing,
        sections: clone(sections),
        values: clone(values),
        letterhead: clone(letterhead),
        docName,
        templateId,
        personId,
        status: docStatus === "In review" ? "in-review" : docStatus === "Approved" ? "approved" : "draft",
        generationStatus: status === "generated" ? "generated" : "failed",
        exports: status === "failed" ? existing.exports : [exportRecord, ...existing.exports],
        updatedAt: formatDocumentStamp(),
      };
      return { ...current, exports: [exportRecord, ...current.exports], documents: [nextDocument, ...current.documents.filter((item) => item.id !== "doc-current")] };
    });
  };
  const formatDocStatus = (status: DocumentRecord["status"]) => status === "in-review" ? "In review" : status.charAt(0).toUpperCase() + status.slice(1);
  const downloadStoredExport = (record: DocumentRecord, format: ExportRecord["format"]) => {
    const exportRecord = record.exports.find((item) => item.format === format && item.status === "generated" && item.contentBase64);
    if (!exportRecord?.contentBase64) {
      setModuleNotice(`No stored ${format.toUpperCase()} file is available. Regenerate from Workspace.`);
      return;
    }
    const bytes = dataUrlToBytes(exportRecord.contentBase64);
    if (!bytes) return setModuleNotice("Stored file could not be read. Regenerate from Workspace.");
    const blob = new Blob([bytes], { type: format === "pdf" ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exportRecord.fileName;
    link.click();
    URL.revokeObjectURL(url);
  };
  const renderModulePage = () => {
    if (activeModule === "clauses") {
      return <section className="module-page">
        <div className="module-header"><div><span className="eyebrow">CONTENT SYSTEM</span><h1>Clauses Library</h1><p>Maintain reusable, versioned content without changing existing document copies.</p></div><button className="primary-btn" onClick={() => openClauseEditor()}><Plus size={15} /> New clause</button></div>
        <div className="module-toolbar"><label className="module-search"><Search size={15} /><input value={moduleSearch} onChange={(e) => setModuleSearch(e.target.value)} placeholder="Search clauses" /></label><select value={moduleFilter} onChange={(e) => setModuleFilter(e.target.value)}><option value="all">All statuses</option><option value="published">Published</option><option value="draft">Draft</option><option value="inactive">Inactive</option></select><span className="module-count">{filteredClauses.length} clauses</span></div>
        <div className="library-table"><div className="library-table-head"><span>Clause</span><span>Structure</span><span>Use</span><span>Status</span><span>Actions</span></div>{filteredClauses.map((record) => <div className="library-row" key={record.id}><div><strong>{record.title}</strong><small>{record.category} · {record.language} · v{record.version}</small></div><span>{record.structure === "nested" ? `${record.subsections.length} subsections` : "Flat"}</span><span>{record.contents.length + record.subsections.reduce((count, sub) => count + sub.contents.length, 0)} content blocks</span><span className={`status-pill ${record.status}`}>{record.status}</span><div className="row-actions"><button className="icon-btn" title="Edit" onClick={() => openClauseEditor(record)}><Settings2 size={14} /></button><button className="icon-btn" title="Copy" onClick={() => copyClauseRecord(record)}><Copy size={14} /></button><button className="icon-btn" title="Preview" onClick={() => { setSelectedClauseRecordId(record.id); setModuleNotice("Preview loaded below"); }}><FileCheck2 size={14} /></button><button className="icon-btn" title={record.status === "inactive" ? "Activate" : "Deactivate"} onClick={() => toggleClauseStatus(record)}><Archive size={14} /></button></div></div>)}</div>
        {selectedClauseRecordId && <div className="module-preview"><div><div><span className="eyebrow">PREVIEW</span><h3>{appStore.clauses.find((item) => item.id === selectedClauseRecordId)?.title}</h3></div><div className="row-actions"><button className="primary-btn" onClick={() => { const record = appStore.clauses.find((item) => item.id === selectedClauseRecordId); if (record) insertClauseRecord(record); }}><Plus size={14} /> Insert in Workspace</button><button className="icon-btn" onClick={() => setSelectedClauseRecordId(null)}><X size={15} /></button></div></div><div className="preview-outline">{(() => { const record = appStore.clauses.find((item) => item.id === selectedClauseRecordId); if (!record) return null; return <>{record.structure === "nested" ? record.subsections.map((sub, index) => <div key={sub.id}><strong>{index + 1}. {sub.title}</strong>{sub.contents.map((item, itemIndex) => <p key={item.id}>{index + 1}.{itemIndex + 1} {item.text}</p>)}</div>) : record.contents.map((item, index) => <p key={item.id}>{index + 1}. {item.text}</p>)}</>; })()}</div></div>}
      </section>;
    }
    if (activeModule === "placeholders") {
      return <section className="module-page"><div className="module-header"><div><span className="eyebrow">DATA MODEL</span><h1>Placeholder Management</h1><p>Map reusable fields to onboarding and company records with stable IDs.</p></div><button className="primary-btn" onClick={() => openPlaceholderEditor()}><Plus size={15} /> New group</button></div><div className="split-module"><aside className="module-list">{appStore.placeholderGroups.map((group) => <button className={selectedGroup?.id === group.id ? "selected" : ""} key={group.id} onClick={() => setSelectedPlaceholderGroupId(group.id)}><span>{group.name}</span><small>{group.fields.length} fields · {group.sourceType}</small></button>)}</aside><div className="module-detail">{selectedGroup ? <><div className="detail-heading"><div><h2>{selectedGroup.name}</h2><span>{selectedGroup.sourceType} · {selectedGroup.sourceTable}</span></div><div className="row-actions"><button className="outline-btn" onClick={() => openPlaceholderEditor(selectedGroup)}><Settings2 size={14} /> Edit group</button><button className="primary-btn" onClick={() => { setPlaceholderDraft(clone(selectedGroup)); addPlaceholderField(); setModuleDrawer("placeholder"); }}><Plus size={14} /> Add field</button></div></div><div className="field-table"><div className="field-table-head"><span>Name</span><span>Placeholder</span><span>Type</span><span>Source</span><span>Example</span><span>Status</span></div>{selectedGroup.fields.map((field) => <div className="field-table-row" key={field.id}><strong>{field.label}</strong><code>{`{{${selectedGroup.name}.${field.label}}}`}</code><span>{field.type}</span><span>{field.sourceField || "Manual"}</span><span>{field.example}</span><span className={`status-pill ${field.mappingStatus === "valid" ? "published" : "inactive"}`}>{field.mappingStatus === "valid" ? "Connected" : "Mapping invalid"}</span></div>)}</div></> : <div className="empty-state"><Sparkles size={18} /> Create a placeholder group to get started.</div>}</div></div></section>;
    }
    if (activeModule === "layouts") {
      return <section className="module-page"><div className="module-header"><div><span className="eyebrow">PAGE SYSTEM</span><h1>Layout Management</h1><p>Keep company letterheads and page safety settings reusable, previewable, and versioned.</p></div><div className="row-actions"><button className="outline-btn" onClick={() => openLayoutEditor()}><Plus size={15} /> Create letterhead</button><button className="primary-btn" onClick={() => { letterheadInputRef.current?.click(); }}><Upload size={15} /> Upload existing</button></div></div><div className="layout-grid">{appStore.layouts.map((record) => <article className="layout-card" key={record.id}><div className="layout-thumb" style={{ borderTopColor: record.letterhead.accent }}><div className="layout-thumb-brand"><span style={{ background: record.letterhead.accent }}>N</span><strong>{values.company_name || "Northstar Labs"}</strong></div><div className="layout-thumb-lines" /></div><div className="layout-card-body"><div><strong>{record.name}</strong><small>{record.letterhead.page} · {record.status} · updated {record.updatedAt}</small></div><div className="row-actions"><button className="icon-btn" title="Apply to workspace" onClick={() => { applyLayoutRecord(record); switchModule("workspace"); }}><Check size={14} /></button><button className="icon-btn" title="Edit" onClick={() => openLayoutEditor(record)}><Settings2 size={14} /></button><button className="icon-btn" title="Copy" onClick={() => setAppStore((current) => ({ ...current, layouts: [{ ...clone(record), id: `layout-${Date.now()}`, name: `${record.name} copy`, status: "draft" }, ...current.layouts] }))}><Copy size={14} /></button></div></div></article>)}</div></section>;
    }
    if (activeModule === "documents") {
      const docs = appStore.documents.filter((doc) => {
        const matchesSearch = !moduleSearch || `${doc.docName} ${doc.personId}`.toLowerCase().includes(moduleSearch.toLowerCase());
        const matchesFilter = moduleFilter === "all" || (moduleFilter === "generated" ? doc.generationStatus === "generated" : doc.status === moduleFilter);
        return matchesSearch && matchesFilter;
      });
      return <section className="module-page"><div className="module-header"><div><span className="eyebrow">FILES & VERSIONS</span><h1>Document Library</h1><p>Saved records, approval states, and generated files stay connected.</p></div><button className="primary-btn" onClick={() => switchModule("workspace")}><Plus size={15} /> New document</button></div><div className="module-toolbar"><label className="module-search"><Search size={15} /><input value={moduleSearch} onChange={(e) => setModuleSearch(e.target.value)} placeholder="Search documents" /></label><select value={moduleFilter} onChange={(e) => setModuleFilter(e.target.value)}><option value="all">All documents</option><option value="draft">Drafts</option><option value="in-review">In review</option><option value="approved">Approved</option><option value="generated">Generated</option></select></div><div className="library-table"><div className="library-table-head"><span>Document</span><span>Person</span><span>Saved status</span><span>Generation</span><span>Actions</span></div>{docs.length ? docs.map((doc) => <div className="library-row" key={doc.id}><div><strong>{doc.docName}</strong><small>{templates.find((item) => item.id === doc.templateId)?.type || "HR document"} · v{doc.versions.length || 1} · {doc.exports.length} export{doc.exports.length === 1 ? "" : "s"}</small></div><span>{people.find((item) => item.id === doc.personId)?.name || doc.personId || "Unassigned"}</span><span className={`status-pill ${doc.status}`}>{formatDocStatus(doc.status)}</span><span className={`status-pill ${doc.generationStatus === "generated" ? "published" : doc.generationStatus === "failed" ? "inactive" : "draft"}`}>{doc.generationStatus}</span><div className="row-actions"><button className="outline-btn" onClick={() => { switchModule("workspace"); }}>Open</button>{doc.exports.some((item) => item.format === "docx" && item.status === "generated") && <button className="icon-btn" title="Download Word" onClick={() => downloadStoredExport(doc, "docx")}><Download size={14} /></button>}{doc.exports.some((item) => item.format === "pdf" && item.status === "generated") && <button className="icon-btn" title="Download PDF" onClick={() => downloadStoredExport(doc, "pdf")}><FileCheck2 size={14} /></button>}<button className="icon-btn" title="Copy as new" onClick={() => { setDocName(`${doc.docName} copy`); switchModule("workspace"); }}><Copy size={14} /></button></div></div>) : <div className="empty-state"><FolderOpen size={18} /> No saved documents yet.</div>}</div></section>;
    }
    return null;
  };
  const renderModuleDrawer = () => {
    if (!moduleDrawer) return null;
    if (moduleDrawer === "clause" && clauseDraft) {
      return <div className="module-drawer-backdrop" onClick={() => setModuleDrawer(null)}><aside className="module-drawer" onClick={(event) => event.stopPropagation()}><div className="drawer-heading"><div><span className="eyebrow">CLAUSE EDITOR · 1–4</span><h2>{clauseDraft.title || "New clause"}</h2><p>Build a reusable clause without changing document copies already in use.</p></div><button className="icon-btn" onClick={() => setModuleDrawer(null)}><X size={17} /></button></div><div className="drawer-steps"><span className="active">1 Details</span><span>2 Structure</span><span>3 Content</span><span>4 Preview</span></div><label className="drawer-field"><span>Primary title</span><input value={clauseDraft.title} onChange={(e) => updateClauseDraft({ title: e.target.value })} /></label><div className="drawer-grid"><label className="drawer-field"><span>Category</span><input value={clauseDraft.category} onChange={(e) => updateClauseDraft({ category: e.target.value })} /></label><label className="drawer-field"><span>Language</span><select value={clauseDraft.language} onChange={(e) => updateClauseDraft({ language: e.target.value })}><option>English</option><option>中文</option><option>Bahasa Melayu</option></select></label></div><label className="drawer-field"><span>Structure</span><div className="segmented-control"><button className={clauseDraft.structure === "flat" ? "active" : ""} onClick={() => updateClauseDraft({ structure: "flat" })}>Flat content</button><button className={clauseDraft.structure === "nested" ? "active" : ""} onClick={() => updateClauseDraft({ structure: "nested" })}>Nested subsections</button></div></label><div className="drawer-section-heading"><div><strong>Content blocks</strong><small>{clauseDraft.structure === "nested" ? "Organise content under subsections." : "Add content directly under the primary title."}</small></div><button className="add-field" onClick={() => addClauseContent()}><Plus size={14} /> Add content</button></div>{clauseDraft.structure === "nested" ? <div className="drawer-subsections">{clauseDraft.subsections.map((subsection, index) => <div className="drawer-subsection" key={subsection.id}><div className="drawer-subsection-head"><input value={subsection.title} onChange={(e) => updateClauseDraft({ subsections: clauseDraft.subsections.map((item) => item.id === subsection.id ? { ...item, title: e.target.value } : item) })} /><div className="row-actions"><button className="icon-btn" title="Move up" onClick={() => updateClauseDraft({ subsections: clauseDraft.subsections.map((item, itemIndex, items) => itemIndex === index && index > 0 ? items[index - 1] : itemIndex === index - 1 ? items[index] : item) })}><ArrowUp size={13} /></button><button className="icon-btn" title="Move down" onClick={() => updateClauseDraft({ subsections: clauseDraft.subsections.map((item, itemIndex, items) => itemIndex === index && index < items.length - 1 ? items[index + 1] : itemIndex === index + 1 ? items[index] : item) })}><ArrowDown size={13} /></button></div></div>{subsection.contents.map((item) => <textarea key={item.id} value={item.text} onChange={(e) => updateClauseDraft({ subsections: clauseDraft.subsections.map((group) => group.id === subsection.id ? { ...group, contents: group.contents.map((content) => content.id === item.id ? { ...content, text: e.target.value } : content) } : group) })} />)}<button className="text-btn" onClick={() => addClauseContent(subsection.id)}><Plus size={13} /> Add content to subsection</button></div>)}<button className="add-field" onClick={() => updateClauseDraft({ subsections: [...clauseDraft.subsections, { id: `sub-${Date.now()}`, title: "New subsection", contents: [] }] })}><Plus size={14} /> Add subsection</button></div> : <div className="drawer-contents">{clauseDraft.contents.map((item, index) => <div className="drawer-content-row" key={item.id}><span>{index + 1}</span><textarea value={item.text} onChange={(e) => updateClauseDraft({ contents: clauseDraft.contents.map((content) => content.id === item.id ? { ...content, text: e.target.value } : content) })} /><div className="row-actions"><button className="icon-btn" title="Move up" onClick={() => updateClauseDraft({ contents: clauseDraft.contents.map((content, itemIndex, items) => itemIndex === index && index > 0 ? items[index - 1] : itemIndex === index - 1 ? items[index] : content) })}><ArrowUp size={13} /></button><button className="icon-btn" title="Move down" onClick={() => updateClauseDraft({ contents: clauseDraft.contents.map((content, itemIndex, items) => itemIndex === index && index < items.length - 1 ? items[index + 1] : itemIndex === index + 1 ? items[index] : content) })}><ArrowDown size={13} /></button><button className="icon-btn" title="Delete" onClick={() => updateClauseDraft({ contents: clauseDraft.contents.filter((content) => content.id !== item.id) })}><Trash2 size={13} /></button></div></div>)}</div>}<div className="drawer-footer"><button className="outline-btn" onClick={() => { setClauseDraft(null); setModuleDrawer(null); }}>Cancel</button><button className="outline-btn" onClick={() => saveClauseDraft(false)}><Save size={14} /> Save draft</button><button className="primary-btn" disabled={currentRole !== "Admin"} onClick={() => saveClauseDraft(true)}><ShieldCheck size={14} /> Publish</button></div></aside></div>;
    }
    if (moduleDrawer === "placeholder" && placeholderDraft) {
      return <div className="module-drawer-backdrop" onClick={() => setModuleDrawer(null)}><aside className="module-drawer" onClick={(event) => event.stopPropagation()}><div className="drawer-heading"><div><span className="eyebrow">PLACEHOLDER GROUP</span><h2>{placeholderDraft.name}</h2><p>Stable field IDs keep existing references working when display labels change.</p></div><button className="icon-btn" onClick={() => setModuleDrawer(null)}><X size={17} /></button></div><label className="drawer-field"><span>Group name</span><input value={placeholderDraft.name} onChange={(e) => setPlaceholderDraft({ ...placeholderDraft, name: e.target.value })} /></label><div className="drawer-grid"><label className="drawer-field"><span>Source type</span><select value={placeholderDraft.sourceType} onChange={(e) => setPlaceholderDraft({ ...placeholderDraft, sourceType: e.target.value })}><option>Onboarding Form</option><option>CSV import</option><option>Excel workbook</option><option>Manual</option></select></label><label className="drawer-field"><span>Source table / sheet</span><input value={placeholderDraft.sourceTable || ""} onChange={(e) => setPlaceholderDraft({ ...placeholderDraft, sourceTable: e.target.value })} /></label></div><div className="drawer-section-heading"><div><strong>Fields</strong><small>{placeholderDraft.fields.length} reusable fields</small></div><button className="add-field" onClick={addPlaceholderField}><Plus size={14} /> Add field</button></div><div className="drawer-field-list">{placeholderDraft.fields.map((field, index) => <div className="drawer-field-card" key={field.id}><input value={field.label} aria-label={`Field ${index + 1} name`} onChange={(e) => setPlaceholderDraft({ ...placeholderDraft, fields: placeholderDraft.fields.map((item) => item.id === field.id ? { ...item, label: e.target.value } : item) })} /><code>{`{{${placeholderDraft.name}.${field.label}}}`}</code><select value={field.type} onChange={(e) => setPlaceholderDraft({ ...placeholderDraft, fields: placeholderDraft.fields.map((item) => item.id === field.id ? { ...item, type: e.target.value as PlaceholderField["type"] } : item) })}><option value="text">Text</option><option value="date">Date</option><option value="currency">Currency</option><option value="number">Number</option><option value="multiline">Multiline</option></select><input value={field.sourceField} placeholder="Source field" onChange={(e) => setPlaceholderDraft({ ...placeholderDraft, fields: placeholderDraft.fields.map((item) => item.id === field.id ? { ...item, sourceField: e.target.value } : item) })} /><button className="icon-btn" onClick={() => setPlaceholderDraft({ ...placeholderDraft, fields: placeholderDraft.fields.filter((item) => item.id !== field.id) })}><Trash2 size={13} /></button></div>)}</div><div className="drawer-footer"><button className="outline-btn" onClick={() => { setPlaceholderDraft(null); setModuleDrawer(null); }}>Cancel</button><button className="primary-btn" onClick={savePlaceholderDraft}><Save size={14} /> Save group</button></div></aside></div>;
    }
    if (moduleDrawer === "layout" && layoutDraft) {
      const draftLayout = layoutDraft.letterhead.firstPage || layoutDraft.letterhead;
      return <div className="module-drawer-backdrop" onClick={() => setModuleDrawer(null)}><aside className="module-drawer" onClick={(event) => event.stopPropagation()}><div className="drawer-heading"><div><span className="eyebrow">LAYOUT EDITOR</span><h2>{layoutDraft.name}</h2><p>Save a reusable layout or apply a document-level snapshot from Workspace.</p></div><button className="icon-btn" onClick={() => setModuleDrawer(null)}><X size={17} /></button></div><label className="drawer-field"><span>Layout name</span><input value={layoutDraft.name} onChange={(e) => setLayoutDraft({ ...layoutDraft, name: e.target.value })} /></label><div className="drawer-grid"><label className="drawer-field"><span>Accent color</span><input type="color" value={draftLayout.accent} onChange={(e) => setLayoutDraft({ ...layoutDraft, letterhead: normalizeLetterhead({ ...layoutDraft.letterhead, accent: e.target.value }) })} /></label><label className="drawer-field"><span>Page</span><select value={layoutDraft.letterhead.page} onChange={(e) => setLayoutDraft({ ...layoutDraft, letterhead: { ...layoutDraft.letterhead, page: e.target.value as Letterhead["page"] } })}><option value="A4">A4</option><option value="Letter">Letter</option></select></label></div><div className="layout-editor-preview"><div className="layout-thumb" style={{ borderTopColor: draftLayout.accent }}><div className="layout-thumb-brand"><span style={{ background: draftLayout.accent }}>N</span><strong>{values.company_name || "Northstar Labs"}</strong></div><div className="layout-thumb-lines" /></div></div><div className="drawer-grid"><label className="drawer-field"><span>Top offset</span><input type="number" value={draftLayout.top} onChange={(e) => setLayoutDraft({ ...layoutDraft, letterhead: normalizeLetterhead({ ...layoutDraft.letterhead, top: Number(e.target.value) }) })} /></label><label className="drawer-field"><span>Body margin</span><input type="number" value={draftLayout.margin} onChange={(e) => setLayoutDraft({ ...layoutDraft, letterhead: normalizeLetterhead({ ...layoutDraft.letterhead, margin: Number(e.target.value) }) })} /></label></div><label className="drawer-field"><span>Page usage</span><div className="segmented-control"><button className={layoutDraft.letterhead.mode === "first" ? "active" : ""} onClick={() => setLayoutDraft({ ...layoutDraft, letterhead: { ...layoutDraft.letterhead, mode: "first" } })}>First page</button><button className={layoutDraft.letterhead.mode === "all" ? "active" : ""} onClick={() => setLayoutDraft({ ...layoutDraft, letterhead: { ...layoutDraft.letterhead, mode: "all" } })}>Every page</button><button className={layoutDraft.letterhead.mode === "different" ? "active" : ""} onClick={() => setLayoutDraft({ ...layoutDraft, letterhead: { ...layoutDraft.letterhead, mode: "different" } })}>Different pages</button></div></label><div className="drawer-footer"><button className="outline-btn" onClick={() => { setLayoutDraft(null); setModuleDrawer(null); }}>Cancel</button><button className="outline-btn" onClick={() => saveLayoutDraft(false)}><Save size={14} /> Save draft</button><button className="primary-btn" disabled={currentRole !== "Admin"} onClick={() => saveLayoutDraft(true)}><ShieldCheck size={14} /> Publish</button></div></aside></div>;
    }
    return null;
  };
  const primaryActionLabel =
    docStatus === "Approved"
      ? "Approved"
      : blockingIssues.length
      ? "Review issues"
      : currentRole === "Reviewer" && docStatus === "In review"
        ? "Approve version"
        : docStatus === "In review"
          ? "View submission"
          : "Send for review";

  return (
    <div className="app-shell">
      <input ref={letterheadInputRef} type="file" accept=".png,.jpg,.jpeg,.pdf,.docx" hidden onChange={(e) => uploadLetterhead(e.target.files?.[0])} />
      <input ref={placeholderImportInputRef} type="file" accept=".csv,.xlsx,.xls" hidden onChange={(e) => previewPlaceholderImport(e.target.files?.[0])} />
      <aside className={`global-sidebar ${sidebarCollapsed ? "collapsed" : ""}`} aria-label="Primary navigation">
        <div className="global-sidebar-brand"><div className="brand-mark"><FileText size={16} /></div>{!sidebarCollapsed && <div><strong>HR Doc Generator</strong><small>Northstar workspace</small></div>}</div>
        <div className="global-workspace-switch"><div className="company-dot">N</div>{!sidebarCollapsed && <div><strong>Northstar Labs</strong><span>Malaysia · MY</span></div>}<ChevronDown size={13} /></div>
        <nav className="global-nav">
          {navItems.map(([id, label, Icon]) => <button key={id} className={activeModule === id ? "active" : ""} onClick={() => switchModule(id as AppModule)} title={sidebarCollapsed ? label : undefined} aria-label={label} aria-current={activeModule === id ? "page" : undefined}><Icon size={17} /><span>{label}</span></button>)}
        </nav>
        <div className="global-sidebar-spacer" />
        {!sidebarCollapsed && <div className="global-help"><span className="eyebrow">WORKSPACE</span><strong>Keep your document flow moving</strong><small>Build, review and export from one place.</small></div>}
        <div className="global-nav global-nav-secondary"><button title="Account" aria-label="Account"><UserRound size={16} /><span>Account</span></button><button title="Members and roles" aria-label="Members and roles"><UsersRound size={16} /><span>Members & roles</span></button><button title="Subscription" aria-label="Subscription"><ShieldCheck size={16} /><span>Subscription</span></button><button title="Settings" aria-label="Settings"><Settings2 size={16} /><span>Settings</span></button><button title={theme === "light" ? "Switch to dark theme" : "Switch to light theme"} aria-label="Toggle theme" onClick={() => setTheme((current) => current === "light" ? "dark" : "light")}>{theme === "light" ? <Moon size={16} /> : <Sun size={16} />}<span>Theme</span></button></div>
        <button className="sidebar-collapse" onClick={() => setSidebarCollapsed((current) => !current)} aria-label={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}>{sidebarCollapsed ? <PanelLeftOpen size={16} /> : <><PanelLeftClose size={16} /><span>Collapse</span></>}</button>
      </aside>
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark">
            <FileText size={17} />
          </div>
          <span>HR Doc Generator</span>
          <span className="beta">Workspace</span>
        </div>
        <div className="top-context">
          <div className="company-switch-wrap">
          <button
            className="company-switch"
            aria-haspopup="menu"
            aria-expanded={companyMenuOpen}
            onClick={() => setCompanyMenuOpen((open) => !open)}
          >
            <div className="company-dot">N</div>
            <div>
              <strong>Northstar Labs</strong>
              <span>Malaysia · MY</span>
            </div>
            <ChevronDown size={14} />
          </button>
          {companyMenuOpen && (
            <div className="company-menu" role="menu">
              <button role="menuitem" className="company-menu-item" onClick={() => setCompanyMenuOpen(false)}>
                <strong>Northstar Labs</strong>
                <span>Malaysia · MY · Current</span>
              </button>
              <div className="company-menu-note">Additional companies can be added from Manage.</div>
            </div>
          )}
          </div>
          <span className="crumb">/</span>
          {activeModule === "workspace" ? <button
            className="doc-title"
            onClick={() => renameDocument()}
          >
            {docName}
            <ChevronDown size={14} />
          </button> : <span className="module-top-title">{moduleTitle}</span>}
        </div>
        <div className="top-actions">
          {activeModule === "workspace" && <><div className={`save-state ${saving}`}>
            {saving === "error" ? <AlertTriangle size={14} /> : <Cloud size={14} />}
            {saving === "saving"
              ? "Saving…"
              : saving === "error"
                ? "Save failed"
                : lastSavedAt
                  ? `Saved ${lastSavedAt}`
                  : "Saved locally"}
          </div>
          <button
            className="icon-btn"
            title="Version history"
            onClick={() => setShowHistory(true)}
          >
            <History size={17} />
          </button></>}
          <button
            className="icon-btn theme-toggle"
            title={theme === "light" ? "Switch to dark theme" : "Switch to light theme"}
            aria-label={theme === "light" ? "Switch to dark theme" : "Switch to light theme"}
            onClick={() => setTheme((current) => current === "light" ? "dark" : "light")}
          >
            {theme === "light" ? <Moon size={16} /> : <Sun size={16} />}
          </button>
          {activeModule === "workspace" && <button className="outline-btn" onClick={() => setShowPreview(true)}>
            <FileCheck2 size={15} />
            Preview
          </button>}
          {activeModule === "workspace" && <button
            className="primary-btn"
            disabled={reviewActionDisabled}
            title={
              currentRole === "Reviewer" && docStatus !== "In review"
                ? "Reviewers can approve documents after an editor submits them"
                : docStatus === "Approved"
                  ? "Approved documents are locked; restore as a new draft to submit changes"
                  : undefined
            }
            onClick={() => {
              if (blockingIssues.length) {
                setToast("Fix the highlighted checks before review");
                window.setTimeout(() => setToast(""), 2400);
                setShowPreview(true);
                return;
              }
              if (docStatus === "Approved") return;
              if (docStatus === "In review" && currentRole !== "Reviewer") {
                setWorkflowStage("review");
                setShowPreview(true);
                return;
              }
              if (currentRole === "Reviewer" && docStatus === "In review") {
                setDocStatus("Approved");
                setWorkflowStage("export");
                createVersion("Approved by reviewer", "Approved");
                setToast("Document approved and version locked");
              } else {
                setDocStatus("In review");
                setWorkflowStage("review");
                createVersion("Submitted for review", "In review");
                setToast("Sent for review");
              }
              window.setTimeout(() => setToast(""), 2200);
            }}
          >
            <ShieldCheck size={15} />
            {primaryActionLabel}
          </button>}
          <div className="role-switch-wrap">
            <button
              className="avatar role-avatar"
              title="Switch workspace role"
              onClick={() => setRoleMenuOpen((open) => !open)}
            >
              {currentRole.slice(0, 1)}
            </button>
            {roleMenuOpen && (
              <div className="role-menu" role="menu">
                {(["Admin", "Editor", "Reviewer"] as UserRole[]).map((role) => (
                  <button
                    key={role}
                    role="menuitem"
                    className={currentRole === role ? "selected" : ""}
                    onClick={() => {
                      setCurrentRole(role);
                      setRoleMenuOpen(false);
                      setToast(`Role switched to ${role}`);
                      window.setTimeout(() => setToast(""), 1800);
                    }}
                  >
                    <strong>{role}</strong>
                    <span>{role === "Admin" ? "Manage configuration" : role === "Editor" ? "Edit documents" : "Review and approve"}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </header>
      {activeModule === "workspace" && <div className="subbar">
        <div className="flow">
          {(["build", "review", "export"] as WorkflowStage[]).map((stage, index) => {
            const isComplete = index < workflowStageIndex || (stage === "export" && docStatus === "Approved");
            const isCurrent = stage === effectiveWorkflowStage;
            const label = stage === "build" ? "Build" : stage === "review" ? "Review" : "Export";
            return (
              <React.Fragment key={stage}>
                <button
                  className={`flow-step ${isCurrent ? "active" : ""} ${isComplete ? "complete" : ""}`}
                  onClick={() => goToWorkflowStage(stage)}
                  aria-current={isCurrent ? "step" : undefined}
                >
                  {isComplete ? <Check size={12} /> : <span className="flow-number">{index + 1}</span>}
                  <span>{label}</span>
                </button>
                {index < 2 && <span className="flow-line" />}
              </React.Fragment>
            );
          })}
        </div>
        <div className={`workflow-context ${workflowSummary.tone}`}>
          <span className="workflow-context-icon">
            {workflowSummary.tone === "warning" ? <AlertTriangle size={14} /> : workflowSummary.tone === "success" ? <Check size={14} /> : workflowSummary.tone === "review" ? <ShieldCheck size={14} /> : <Circle size={10} />}
          </span>
          <span className="workflow-context-copy">
            <strong>{workflowSummary.title}</strong>
            <span>{workflowSummary.detail}</span>
          </span>
          <button className="workflow-context-action" onClick={() => goToWorkflowStage(workflowSummary.targetStage)}>
            {workflowSummary.action}
            <ChevronRight size={13} />
          </button>
        </div>
        <div className="sub-actions">
          <span
            className={`status-pill ${docStatus.toLowerCase().replace(" ", "-")}`}
          >
            {docStatus}
          </span>
          <button className="sub-btn" onClick={saveNow}>
            <Save size={14} />
            Save
          </button>
          <button className="sub-btn manage-btn" disabled={currentRole !== "Admin"} onClick={() => { setAdminView(null); setShowAdmin(true); }} title={currentRole !== "Admin" ? "Admin role required" : "Open template and workspace settings"}>
            <Settings2 size={14} />
            Admin settings
          </button>
        </div>
      </div>}
      {activeModule === "workspace" ? <div className={`workspace ${leftRailCollapsed ? "left-rail-collapsed" : ""}`}>
        <aside className="left-rail">
          <div className="rail-section">
            <div className="rail-heading">
              <span>DOCUMENT SETUP</span>
              <div className="rail-heading-actions">
                <button
                  className="mini-icon"
                  title="Open document"
                  aria-haspopup="menu"
                  aria-expanded={documentMenuOpen}
                  onClick={() => setDocumentMenuOpen((open) => !open)}
                >
                  <FolderOpen size={14} />
                </button>
                <button
                  className="mini-icon rail-collapse-btn"
                  title={leftRailCollapsed ? "Expand section navigator" : "Collapse section navigator"}
                  aria-label={leftRailCollapsed ? "Expand section navigator" : "Collapse section navigator"}
                  onClick={() => setLeftRailCollapsed((collapsed) => !collapsed)}
                >
                  {leftRailCollapsed ? <PanelLeftOpen size={14} /> : <PanelLeftClose size={14} />}
                </button>
              </div>
            </div>
            {documentMenuOpen && (
              <div className="document-menu" role="menu">
                <button role="menuitem" onClick={openNewBlankDocument}>
                  <Plus size={14} /> New blank document
                </button>
                <button role="menuitem" onClick={copyCurrentDocument}>
                  <Copy size={14} /> Copy current document
                </button>
                <div className="document-menu-note">Current draft stays in this workspace.</div>
              </div>
            )}
            <div className="select-box">
              <FileText size={15} />
              <select
                value={templateId}
                onChange={(e) => changeTemplate(e.target.value)}
              >
                {templates.map((item) => (
                  <option key={item.id} value={item.id}>
                    {item.type}
                  </option>
                ))}
              </select>
              <ChevronDown size={14} />
            </div>
            <div className="template-context">
              <span className="template-context-dot" style={{ background: template.color }} />
              <span>
                <strong>{template.name}</strong>
                <small>{template.description}</small>
              </span>
            </div>
          </div>
          <div className="rail-section">
            <div className="rail-heading">
              <span>
                CONTENT <em>{sections.length} sections</em>
              </span>
              <button
                className="mini-icon"
                title="Add section"
                onClick={addSection}
              >
                <Plus size={15} />
              </button>
            </div>
            <div className="section-list">
              {sections.map((section, index) => (
                <div
                  className={`section-row ${activeSection === section.id ? "selected" : ""}`}
                  key={section.id}
                  onClick={() => {
                    setActiveSection(section.id);
                    setActiveClause(section.clauses[0]?.id || "");
                  }}
                >
                  <GripVertical size={14} className="drag" />
                  <div className="section-info">
                    <strong>
                      {String(index + 1).padStart(2, "0")}{" "}
                      <span
                        contentEditable={canEditDocument}
                        suppressContentEditableWarning
                        onBlur={(e) => {
                          const title =
                            e.currentTarget.textContent || section.title;
                          setSections((prev) =>
                            prev.map((s) =>
                              s.id === section.id ? { ...s, title } : s,
                            ),
                          );
                        }}
                      >
                        {section.title}
                      </span>
                    </strong>
                    <small>
                      {section.clauses.filter((c) => c.included).length} of{" "}
                      {section.clauses.length} paragraphs
                    </small>
                  </div>
                  <button
                    className="row-more"
                    onClick={(e) => {
                      e.stopPropagation();
                      moveSection(index, index === 0 ? 1 : -1);
                    }}
                    title="Move section"
                  >
                    <MoreHorizontal size={15} />
                  </button>
                </div>
              ))}
            </div>
            <button className="add-section" onClick={addSection}>
              <Plus size={14} />
              Add section
            </button>
          </div>
          <div className="rail-bottom">
            <div
              className="library-link"
              onClick={() => setActiveRightTab("clauses")}
            >
              <Archive size={15} />
              <span>Clause library</span>
              <ChevronRight size={14} />
            </div>
            <div className="library-link" onClick={() => { setAdminView(null); setShowAdmin(true); }}>
              <WandSparkles size={15} />
              <span>Template settings</span>
              <ChevronRight size={14} />
            </div>
          </div>
        </aside>
        <main className="editor-area">
          <div className="editor-toolbar">
            <div className="toolbar-group">
              <button
                className="toolbar-btn"
                title="Undo"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("undo")}
              >
                <Undo2 size={16} />
              </button>
              <button
                className="toolbar-btn"
                title="Redo"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("redo")}
              >
                <Redo2 size={16} />
              </button>
              <span className="toolbar-divider" />
              <button
                className="toolbar-select"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("fontName", "Aptos")}
              >
                Aptos <ChevronDown size={13} />
              </button>
              <button
                className="toolbar-select size-select"
                onMouseDown={toolbarMouseDown}
                onClick={() => {
                  const next = fontSize === 11 ? 13 : fontSize === 13 ? 16 : 11;
                  const applied = applyEditorCommand(
                    "fontSize",
                    next === 11 ? "2" : next === 13 ? "3" : "4",
                  );
                  if (applied) setFontSize(next);
                }}
              >
                {fontSize}
                <ChevronDown size={13} />
              </button>
              <button
                className="toolbar-btn"
                title="Increase selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => {
                  const applied = applyEditorCommand("fontSize", "4");
                  if (applied) setFontSize((v) => Math.min(24, v + 2));
                }}
              >
                <ArrowUp size={14} />
              </button>
              <button
                className="toolbar-btn"
                title="Decrease selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => {
                  const applied = applyEditorCommand("fontSize", "2");
                  if (applied) setFontSize((v) => Math.max(9, v - 2));
                }}
              >
                <ArrowDown size={14} />
              </button>
            </div>
            <div className="toolbar-group">
              <button
                className={`toolbar-btn ${format.bold ? "active" : ""}`}
                title="Bold selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("bold")}
              >
                <Bold size={16} />
              </button>
              <button
                className={`toolbar-btn ${format.italic ? "active" : ""}`}
                title="Italicize selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("italic")}
              >
                <Italic size={16} />
              </button>
              <button
                className={`toolbar-btn ${format.underline ? "active" : ""}`}
                title="Underline selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("underline")}
              >
                <Underline size={16} />
              </button>
              <button
                className={`toolbar-btn ${format.strike ? "active" : ""}`}
                title="Strikethrough selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("strikeThrough")}
              >
                <Strikethrough size={16} />
              </button>
              <button
                className="toolbar-btn"
                title="Highlight selected text"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("hiliteColor", "#fff0a6")}
              >
                <Highlighter size={16} />
              </button>
              <button
                className="toolbar-btn"
                title="Clear selected formatting"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("removeFormat")}
              >
                <Paintbrush size={16} />
              </button>
            </div>
            <div className="toolbar-group">
              <button
                className={`toolbar-btn ${format.align === "left" ? "active" : ""}`}
                title="Align paragraph left"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("justifyLeft")}
              >
                <AlignLeft size={16} />
              </button>
              <button
                className={`toolbar-btn ${format.align === "center" ? "active" : ""}`}
                title="Center paragraph"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("justifyCenter")}
              >
                <AlignCenter size={16} />
              </button>
              <button
                className={`toolbar-btn ${format.align === "right" ? "active" : ""}`}
                title="Align paragraph right"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("justifyRight")}
              >
                <AlignRight size={16} />
              </button>
              <button
                className={`toolbar-btn ${format.align === "justify" ? "active" : ""}`}
                title="Justify paragraph"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("justifyFull")}
              >
                <AlignJustify size={16} />
              </button>
              <button
                className="toolbar-btn"
                title="Bulleted list"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("insertUnorderedList")}
              >
                <List size={16} />
              </button>
              <button
                className="toolbar-btn"
                title="Numbered list"
                onMouseDown={toolbarMouseDown}
                onClick={() => applyEditorCommand("insertOrderedList")}
              >
                <ListOrdered size={16} />
              </button>
            </div>
            <div className="toolbar-group toolbar-right">
              <button className="toolbar-btn" title="Insert table">
                <Table2 size={16} />
              </button>
              <button className="toolbar-btn" title="Insert image">
                <ImagePlus size={16} />
              </button>
              <button
                className="toolbar-btn"
                title="Insert link"
                onMouseDown={toolbarMouseDown}
                onClick={() => insertLink()}
              >
                <Link2 size={16} />
              </button>
              <span className="toolbar-divider" />
              <button
                className="zoom-control"
                onClick={() => setZoom(zoom === 92 ? 105 : 92)}
              >
                {zoom}%
              </button>
            </div>
          </div>
          <div className="editor-canvas">
            <div
              className="page-wrap"
              style={{
                transform: `scale(${zoom / 100})`,
                transformOrigin: "top center",
              }}
            >
              <div
                className="paper"
                ref={editorRef}
                style={
                  {
                    "--letter-top": `${(letterhead.firstPage || letterhead).top}px`,
                    "--letter-left": `${(letterhead.firstPage || letterhead).left}px`,
                    "--letter-width": `${(letterhead.firstPage || letterhead).width}px`,
                    "--letter-opacity": (letterhead.firstPage || letterhead).opacity,
                    paddingTop: `${(letterhead.firstPage || letterhead).margin}px`,
                  } as React.CSSProperties
                }
              >
                <div
                  className="letterhead-preview"
                  style={{
                    left: `${(letterhead.firstPage || letterhead).left}px`,
                    top: `${(letterhead.firstPage || letterhead).top}px`,
                    width: `${(letterhead.firstPage || letterhead).width}px`,
                    opacity: (letterhead.firstPage || letterhead).opacity,
                    backgroundImage: (letterhead.firstPage || letterhead).dataUrl
                      ? `url(${(letterhead.firstPage || letterhead).dataUrl})`
                      : undefined,
                    backgroundSize: "contain",
                    backgroundRepeat: "no-repeat",
                  }}
                >
                  {!(letterhead.firstPage || letterhead).dataUrl && <>
                    <div className="lh-brand">
                      <div
                        className="lh-logo"
                        style={{ background: (letterhead.firstPage || letterhead).accent }}
                      >
                        N
                      </div>
                      <div>
                        <strong>Northstar Labs</strong>
                        <span>People & Culture</span>
                      </div>
                    </div>
                    <div
                      className="lh-lines"
                      style={{ background: (letterhead.firstPage || letterhead).accent }}
                    />
                  </>}
                </div>
                <div className="page-meta">
                  <span>{template.type.toUpperCase()}</span>
                  <span>MY · 2026</span>
                </div>
                <h1>{template.type}</h1>
                <p className="lede">
                  Between <mark>{values.company_name}</mark> and{" "}
                  <mark>{values.full_name}</mark>
                </p>
                <div className="doc-body">
                  {sections.map((section, sIndex) => (
                    <section
                      key={section.id}
                      className={`doc-section ${activeSection === section.id ? "active" : ""}`}
                      onClick={() => {
                        setActiveSection(section.id);
                        setActiveClause(section.clauses[0]?.id || "");
                      }}
                    >
                      <div className="section-label">
                        <span>{String(sIndex + 1).padStart(2, "0")}</span>
                        <h2>{section.title}</h2>
                        <button
                          className="section-options"
                          title="Section options"
                        >
                          <MoreHorizontal size={14} />
                        </button>
                      </div>
                      {section.clauses.map((clause, cIndex) =>
                        shouldShowClause(clause) ? (
                          <div
                            className={`clause ${activeClause === clause.id ? "active" : ""}`}
                            key={clause.id}
                            onClick={(e) => {
                              e.stopPropagation();
                              setActiveClause(clause.id);
                              setActiveSection(section.id);
                            }}
                          >
                            <div className="clause-number">
                              {sIndex + 1}.{cIndex + 1}
                            </div>
                            <div className="clause-content">
                              <div className="clause-title">
                                {clause.title}
                                <span
                                  className={`tag ${clause.tag.toLowerCase()}`}
                                >
                                  {clause.tag}
                                </span>
                              </div>
                              <div
                                className="editable-paragraph"
                                data-clause-id={clause.id}
                                contentEditable={canEditDocument}
                                suppressContentEditableWarning
                                onFocus={(e) => {
                                  activeEditorRef.current = e.currentTarget;
                                  rememberSelection();
                                }}
                                onMouseUp={rememberSelection}
                                onKeyUp={rememberSelection}
                                onBlur={(e) =>
                                  updateClauseHtml(
                                    clause.id,
                                    e.currentTarget.innerHTML,
                                  )
                                }
                                dangerouslySetInnerHTML={{
                                  __html: clauseHtml(clause),
                                }}
                              />
                              <div className="clause-actions">
                                <button onClick={() => toggleClause(clause.id)}>
                                  <X size={12} />
                                  Exclude
                                </button>
                                <button
                                  onClick={() =>
                                    duplicateClause(section.id, clause)
                                  }
                                >
                                  <Copy size={12} />
                                  Duplicate
                                </button>
                                <button
                                  onClick={() => addClause(section.id, true)}
                                >
                                  <Plus size={12} />
                                  Insert below
                                </button>
                              </div>
                            </div>
                          </div>
                        ) : clause.included ? null : (
                          <div
                            className="excluded-clause"
                            key={clause.id}
                            onClick={() => toggleClause(clause.id)}
                          >
                            <Archive size={13} />
                            {clause.title} excluded · click to restore
                          </div>
                        ),
                      )}
                    </section>
                  ))}
                  <button className="inline-add" onClick={addSection}>
                    <Plus size={14} />
                    Add a section
                  </button>
                </div>
                <div className="signatures">
                  <div>
                    <div className="sign-line" />
                    <strong>{values.signatory_name}</strong>
                    <span>{values.signatory_title}</span>
                    <small>Date: __________________</small>
                  </div>
                  <div>
                    <div className="sign-line" />
                    <strong>{values.full_name}</strong>
                    <span>Employee</span>
                    <small>Date: __________________</small>
                  </div>
                </div>
                <div className="page-footer">
                  <span>Northstar Labs · Confidential</span>
                  <span>Page 1 of {estimatePageCount()}</span>
                </div>
              </div>
            </div>
          </div>
        </main>
        <aside className="right-panel">
          <div className="right-tabs">
            <button
              className={activeRightTab === "person" ? "active" : ""}
              onClick={() => setActiveRightTab("person")}
            >
              <UserRound size={15} />
              Source
            </button>
            <button
              className={activeRightTab === "placeholders" ? "active" : ""}
              onClick={() => setActiveRightTab("placeholders")}
            >
              <Sparkles size={15} />
              Fields
            </button>
            <button
              className={activeRightTab === "clauses" ? "active" : ""}
              onClick={() => setActiveRightTab("clauses")}
            >
              <Archive size={15} />
              Clauses
            </button>
            <button
              className={activeRightTab === "letterhead" ? "active" : ""}
              onClick={() => setActiveRightTab("letterhead")}
            >
              <PanelRight size={15} />
              Layout
            </button>
          </div>
          {activeRightTab === "person" && (
            <div className="panel-content">
              <div className="panel-title-row">
                <div>
                  <span className="eyebrow">ONBOARDING</span>
                  <h3>Source & submission</h3>
                </div>
                <button className="mini-icon">
                  <MoreHorizontal size={15} />
                </button>
              </div>
              <div className="person-card">
                <div className="person-avatar">
                  {person.name
                    .split(" ")
                    .map((n) => n[0])
                    .join("")}
                </div>
                <div>
                  <strong>{person.name}</strong>
                  <span>{person.role}</span>
                  <small>
                    <span className="sync-dot" />
                    Source: Onboarding Form · {person.id} · submission {person.submissionId}
                  </small>
                </div>
                <ChevronDown size={15} />
              </div>
              <label className="field-label">
                Submission record
                <select
                  value={personId}
                  onChange={(e) => changePerson(e.target.value)}
                >
                  {people.map((item) => (
                    <option key={item.id} value={item.id}>
                      {item.id} · {item.name}
                    </option>
                  ))}
                </select>
              </label>
              <div className={`source-note connection-${connectionStatus.toLowerCase()}`}>
                {connectionStatus === "Failed" ? <AlertTriangle size={14} /> : connectionStatus === "Syncing" ? <RefreshCw size={14} className="spin" /> : <Cloud size={14} />}
                <div>
                  <strong>{connectionStatus === "Failed" ? "Onboarding connection failed" : connectionStatus === "Syncing" ? "Checking onboarding updates…" : "Connected via Onboarding Form"}</strong>
                  <span>
                    {connectionStatus === "Failed"
                      ? connectionError || "Retry to check source data"
                      : hasUnreviewedSync
                        ? `${syncDiffs.length} change${syncDiffs.length === 1 ? "" : "s"} waiting for review${sourceUpdateMeta[personId]?.detectedAt ? ` · detected ${sourceUpdateMeta[personId].detectedAt}` : ""}`
                        : lastSyncAt
                          ? `Last sync: ${lastSyncAt}`
                          : `Source record ${person.submissionId} · not synced in this session`}
                  </span>
                </div>
                <button
                  className="mini-icon"
                  title={connectionStatus === "Failed" ? "Retry connection" : "Check for updates"}
                  aria-label={connectionStatus === "Failed" ? "Retry onboarding connection" : "Check onboarding for updates"}
                  disabled={connectionStatus === "Syncing"}
                  onClick={connectionStatus === "Failed" ? retryConnection : checkOnboardingUpdates}
                >
                  {connectionStatus === "Failed" ? <RefreshCw size={14} /> : <ChevronRight size={14} />}
                </button>
              </div>
              {hasUnreviewedSync && (
                <div className="sync-callout">
                  <div>
                    <strong>Source changes need review</strong>
                    <span>Manual overrides stay protected until you choose what to apply.</span>
                  </div>
                  <button className="text-btn" onClick={() => setShowSyncDiff(true)}>Review changes <ChevronRight size={13} /></button>
                </div>
              )}
              <div className="panel-divider" />
              <div className="panel-title-row compact">
                <h4>Mapped fields</h4>
                <button
                  className="text-btn"
                  onClick={() => setActiveRightTab("placeholders")}
                >
                  Manage fields <ChevronRight size={13} />
                </button>
              </div>
              <div className="field-list">
                {[
                  ["full_name", "Full Name"],
                  ["job_title", "Job Title"],
                  ["start_date", "Start Date"],
                  ["end_date", "End Date"],
                  ["monthly_salary", "Monthly Salary"],
                  ["work_arrangement", "Work Arrangement"],
                ].map(([key, label]) => (
                  <label className="field-row" key={key}>
                    <span>
                      {label}
                      <small>
                        <Cloud size={11} />
                        {manualOverrides.includes(key)
                          ? "Manual override"
                          : sourceUpdates[personId]?.[key] !== undefined && sourceUpdates[personId]?.[key] !== values[key]
                            ? "Onboarding · update pending"
                            : "Onboarding"}
                      </small>
                    </span>
                    <input
                      value={values[key] || ""}
                      onChange={(e) => updateOnboardingValue(key, e.target.value)}
                    />
                    <button
                      type="button"
                      className={`lock-btn ${manualOverrides.includes(key) ? "overridden" : ""}`}
                      title={manualOverrides.includes(key) ? "Restore accepted onboarding value" : "Onboarding value; edit to create a protected override"}
                      aria-label={manualOverrides.includes(key) ? `Restore ${label} from onboarding` : `${label} is synced from onboarding`}
                      onClick={() => manualOverrides.includes(key)
                        ? restoreOnboardingValue(key)
                        : setToast(`${label} is currently linked to onboarding`)}
                    >
                      {manualOverrides.includes(key) ? <ShieldCheck size={12} /> : <Lock size={12} />}
                    </button>
                  </label>
                ))}
                {templateId === "freelance" && (
                  <label className="field-row">
                    <span>
                      Billing method
                      <small>
                        <List size={11} />
                        Rule input
                      </small>
                    </span>
                    <select
                      value={values.billing_method || ""}
                      onChange={(e) => updateBillingMethod(e.target.value)}
                    >
                      <option value="fixed fee">Fixed fee</option>
                      <option value="hourly">Hourly</option>
                      <option value="milestone">Milestone</option>
                    </select>
                    <span className="lock-btn" />
                  </label>
                )}
                {templateId === "fixed" && (
                  <label className="field-row">
                    <span>
                      Probation
                      <small>
                        <List size={11} />
                        Rule input
                      </small>
                    </span>
                    <select
                      value={values.probation_period || ""}
                      onChange={(e) =>
                        updateValue("probation_period", e.target.value)
                      }
                    >
                      <option value="3 months">3 months</option>
                      <option value="6 months">6 months</option>
                      <option value="Not applicable">Not applicable</option>
                    </select>
                    <span className="lock-btn" />
                  </label>
                )}
              </div>
              <button
                className="add-field"
                onClick={() => setActiveRightTab("placeholders")}
              >
                <Plus size={14} />
                Add HR-only field
              </button>
            </div>
          )}
          {activeRightTab === "placeholders" && (
            <div className="panel-content">
              <div className="panel-title-row">
                <div>
                  <span className="eyebrow">VARIABLES</span>
                  <h3>Placeholder manager</h3>
                </div>
                <button className="mini-icon" onClick={() => createPlaceholder()}>
                  <Plus size={15} />
                </button>
              </div>
              <div className="search-box">
                <Search size={14} />
                <input
                  placeholder="Search fields"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <div className="helper-callout">
                <Sparkles size={15} />
                <span>
                  Shared variables update every linked clause instantly.
                </span>
              </div>
              <div className="placeholder-list">
                {[
                  ...placeholderMeta,
                  ...customPlaceholders.map(
                    (row) =>
                      [row.key, row.label, row.source, row.group] as string[],
                  ),
                ]
                  .filter(
                    (row) =>
                      row[1].toLowerCase().includes(search.toLowerCase()) ||
                      row[0].includes(search.toLowerCase()),
                  )
                  .map(([key, label, source, group]) => (
                    <div className="placeholder-row" key={key}>
                      <div className="ph-icon">
                        {source === "Dropdown" ? (
                          <ChevronDown size={14} />
                        ) : (
                          <FileText size={14} />
                        )}
                      </div>
                      <div>
                        <strong>{label}</strong>
                        <span>{`{{${key}}}`}</span>
                      </div>
                      <div className="ph-meta">
                        <small>{group}</small>
                        <em>{source}</em>
                      </div>
                      <button className="mini-icon" title="Placeholder options">
                        <MoreHorizontal size={14} />
                      </button>
                    </div>
                  ))}
              </div>
              <button className="add-field" onClick={() => createPlaceholder()}>
                <Plus size={14} />
                Create placeholder
              </button>
              <button className="text-btn" onClick={() => switchModule("placeholders")}>Open Placeholder Management <ChevronRight size={13} /></button>
            </div>
          )}
          {activeRightTab === "clauses" && (
            <div className="panel-content">
              <div className="panel-title-row">
                <div>
                  <span className="eyebrow">LIBRARY</span>
                  <h3>Clause selection</h3>
                </div>
                <div className="row-actions"><button className="mini-icon" title="Add custom clause" onClick={() => addClause(activeSection)}><Plus size={15} /></button><button className="mini-icon" title="Open Clauses Library" onClick={() => switchModule("clauses")}><FolderOpen size={14} /></button></div>
              </div>
              <button className="add-field" onClick={() => switchModule("clauses")}><Archive size={14} /> Add from Clauses Library</button>
              <div className="helper-callout amber">
                <CircleHelp size={15} />
                <span>
                  Excluded clauses remain in the draft and can be restored.
                </span>
              </div>
              {sections.map((section) => (
                <div className="clause-group" key={section.id}>
                  <div className="group-title">
                    <span>{section.title}</span>
                    <small>
                      {section.clauses.filter((c) => c.included).length}/
                      {section.clauses.length}
                    </small>
                  </div>
                  {section.clauses.map((clause) => (
                    <label className="clause-toggle" key={clause.id}>
                      <input
                        type="checkbox"
                        checked={clause.included}
                        onChange={() => toggleClause(clause.id)}
                      />
                      <span className="fake-check">✓</span>
                      <span>
                        <strong>{clause.title}</strong>
                        <small>
                          {clause.tag} ·{" "}
                          {clause.example ? "Example" : "Approved"}
                        </small>
                      </span>
                    </label>
                  ))}
                </div>
              ))}
              <button
                className="add-field"
                onClick={() => addClause(activeSection, true)}
              >
                <Plus size={14} />
                Write custom paragraph
              </button>
            </div>
          )}
          {activeRightTab === "letterhead" && (
            <div className="panel-content">
              <div className="panel-title-row">
                <div>
                  <span className="eyebrow">PAGE DESIGN</span>
                  <h3>Letterhead settings</h3>
                </div>
                <button
                  className="mini-icon"
                  disabled={!canEditDocument}
                      onClick={() => {
                        if (!guardEdit()) return;
                        setLetterhead(makeDefaultLetterhead());
                        setToast("Letterhead settings restored for this draft");
                        window.setTimeout(() => setToast(""), 2200);
                      }}
                  title="Reset defaults"
                >
                  <RotateCcw size={14} />
                </button>
              </div>
              <input
                ref={letterheadInputRef}
                type="file"
                accept=".png,.jpg,.jpeg,.pdf,.docx"
                hidden
                onChange={(e) => uploadLetterhead(e.target.files?.[0])}
              />
              <button
                className="upload-box upload-trigger"
                onClick={() => letterheadInputRef.current?.click()}
              >
                <div className="upload-icon">
                  <Upload size={17} />
                </div>
                <div>
                  <strong>
                    {activeLetterheadLayout.fileName || "Northstar letterhead"}
                  </strong>
                  <span>
                    {activeLetterheadLayout.fileName
                      ? `${activeLetterheadLayout.fileType?.split("/").pop()?.toUpperCase()} · selected locally`
                      : "PNG, JPG, PDF, or DOCX"}
                  </span>
                </div>
                <Paperclip size={14} />
              </button>
              <div className="segmented-control" aria-label="Letterhead page style">
                <button className={letterheadEditorPage === "first" ? "active" : ""} onClick={() => setLetterheadEditorPage("first")}>First page</button>
                <button className={letterheadEditorPage === "subsequent" ? "active" : ""} onClick={() => setLetterheadEditorPage("subsequent")}>Subsequent pages</button>
              </div>
              <div className="helper-callout">
                <FileCheck2 size={15} />
                <span>Editing {letterheadEditorPage === "first" ? "first page" : "subsequent pages"} only. Other pages keep their own layout.</span>
              </div>
              <label className="field-label">
                Apply to
                <select
                  value={letterhead.mode}
                  onChange={(e) => setLetterheadMode(e.target.value as Letterhead["mode"])}
                >
                  <option value="first">First page only</option>
                  <option value="all">Every page</option>
                  <option value="different">First + subsequent pages</option>
                </select>
              </label>
              <div className="control-row">
                <label>
                  Opacity{" "}
                  <strong>{Math.round(activeLetterheadLayout.opacity * 100)}%</strong>
                  <input
                    type="range"
                    min="0"
                    max="0.5"
                    step="0.01"
                    value={activeLetterheadLayout.opacity}
                    onChange={(e) => updateActiveLetterhead({ opacity: Number(e.target.value) })}
                  />
                </label>
                <label>
                  Width <strong>{activeLetterheadLayout.width} px</strong>
                  <input
                    type="range"
                    min="100"
                    max="220"
                    value={activeLetterheadLayout.width}
                    onChange={(e) => updateActiveLetterhead({ width: Number(e.target.value) })}
                  />
                </label>
              </div>
              <div className="control-row">
                <label>
                  Top offset <strong>{activeLetterheadLayout.top} px</strong>
                  <input
                    type="range"
                    min="0"
                    max="80"
                    value={activeLetterheadLayout.top}
                    onChange={(e) => updateActiveLetterhead({ top: Number(e.target.value) })}
                  />
                </label>
                <label>
                  Body margin <strong>{activeLetterheadLayout.margin} px</strong>
                  <input
                    type="range"
                    min="50"
                    max="120"
                    value={activeLetterheadLayout.margin}
                    onChange={(e) => updateActiveLetterhead({ margin: Number(e.target.value) })}
                  />
                </label>
              </div>
              <div className="layout-preview">
                <div className="mini-page">
                  <div
                    className="mini-lh"
                    style={{
                      width: `${activeLetterheadLayout.width / 3}px`,
                      top: `${activeLetterheadLayout.top / 2}px`,
                      left: `${activeLetterheadLayout.left / 2}px`,
                      opacity: activeLetterheadLayout.opacity,
                      background: activeLetterheadLayout.accent,
                    }}
                  />
                  <div
                    className="mini-lines"
                    style={{ marginTop: `${activeLetterheadLayout.margin / 3}px` }}
                  />
                </div>
                <div>
                  <strong>{letterhead.page} · Portrait</strong>
                  <span>Live pagination preview</span>
                  <small>
                    <span className="warning-dot" />
                    Body starts {activeLetterheadLayout.margin}px below top
                  </small>
                </div>
              </div>
              <div className="button-stack">
                <button className="add-field" disabled={currentRole !== "Admin" || !canEditDocument} onClick={() => {
                  if (currentRole !== "Admin") {
                    setToast("Admin role required to publish template layouts");
                    window.setTimeout(() => setToast(""), 2200);
                    return;
                  }
                  if (!guardEdit()) return;
                  setTemplateVersions((prev) => ({ ...prev, [templateId]: (prev[templateId] || 0) + 1 }));
                  setToast(`Published ${template.type} layout v${(templateVersions[templateId] || 0) + 1}`);
                  window.setTimeout(() => setToast(""), 2400);
                }}>
                  <Save size={14} />
                  Publish template layout v{(templateVersions[templateId] || 0) + 1}
                </button>
              <button className="text-btn" onClick={() => { setModuleNotice("Current document keeps this layout independently"); }}><Check size={13} /> Save only to current document</button>
              <button className="text-btn" onClick={() => switchModule("layouts")}>Open Layout Management <ChevronRight size={13} /></button>
              </div>
            </div>
          )}
        </aside>
      </div> : renderModulePage()}
      {activeModule === "placeholders" && <button className="module-import-float" onClick={() => { openPlaceholderEditor(); window.setTimeout(() => placeholderImportInputRef.current?.click(), 0); }}><Upload size={14} /> Import CSV / Excel</button>}
      {importPreview && activeModule === "placeholders" && <div className="import-preview-panel"><div><span className="eyebrow">SOURCE PREVIEW</span><h3>{importPreview.fileName}</h3><p>{importPreview.headers.length} fields detected · first three records shown</p></div><div className="import-preview-fields">{importPreview.headers.slice(0, 8).map((header) => <code key={header}>{header}</code>)}</div><div className="row-actions"><button className="outline-btn" onClick={() => setImportPreview(null)}>Cancel</button><button className="primary-btn" onClick={confirmPlaceholderImport}><Check size={14} /> Use detected fields</button></div></div>}
      {renderModuleDrawer()}
      {moduleNotice && <div className="module-notice" role="status"><CheckCircle2 size={15} />{moduleNotice}<button className="icon-btn" onClick={() => setModuleNotice("")}><X size={13} /></button></div>}
      {showPreview && (
        <div className="modal-backdrop" onClick={() => setShowPreview(false)}>
          <div className="preview-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <span className="eyebrow">FINAL CHECK</span>
                <h2>Preview & export</h2>
                <p>
                  Excluded clauses and unresolved variables are removed from the
                  final file.
                </p>
              </div>
              <button
                className="icon-btn"
                aria-label="Close preview"
                onClick={() => setShowPreview(false)}
              >
                <X size={18} />
              </button>
            </div>
            <div className="check-grid">
              <div className={`check-item ${blockingIssues.some((issue) => issue.startsWith("Missing") || issue.startsWith("Unresolved")) ? "warn" : "ok"}`}>
                <FileCheck2 size={16} />
                <span>
                  <strong>Required fields</strong>
                  <small>
                    {blockingIssues.some((issue) => issue.startsWith("Missing") || issue.startsWith("Unresolved"))
                      ? "Missing or unresolved values"
                      : "All active required fields complete"}
                  </small>
                </span>
              </div>
              <div className={`check-item ${blockingIssues.some((issue) => issue.includes("another person's")) ? "warn" : "ok"}`}>
                <ShieldCheck size={16} />
                <span>
                  <strong>Document consistency</strong>
                  <small>
                    {blockingIssues.some((issue) => issue.includes("another person's"))
                      ? "Another person's data found"
                      : "No identity conflicts found"}
                  </small>
                </span>
              </div>
              <div className={`check-item ${blockingIssues.some((issue) => issue.includes("overlaps")) ? "warn" : "ok"}`}>
                <CircleHelp size={16} />
                <span>
                  <strong>Layout</strong>
                  <small>
                    {blockingIssues.some((issue) => issue.includes("overlaps"))
                      ? "Letterhead overlaps the document body"
                      : "No letterhead overlap detected"}
                  </small>
                </span>
                {blockingIssues.some((issue) => issue.includes("overlaps")) && (
                  <button
                    className="text-btn"
                    onClick={() => {
                      updateActiveLetterhead({
                        margin: Math.min(
                          120,
                          Math.max(activeLetterheadLayout.margin, activeLetterheadLayout.top + 16),
                        ),
                      });
                      setShowPreview(false);
                      setActiveRightTab("letterhead");
                      setToast("Body margin adjusted to avoid letterhead overlap");
                      window.setTimeout(() => setToast(""), 2400);
                    }}
                  >
                    Fix overlap
                  </button>
                )}
              </div>
            </div>
            {blockingIssues.length > 0 && (
              <div className="issue-list" role="alert">
                <strong>Resolve before export</strong>
                {blockingIssues.map((issue) => (
                  <span key={issue}>• {issue}</span>
                ))}
              </div>
            )}
            {exportState === "error" && (
              <div className="issue-list export-error" role="alert">
                <strong>Export failed</strong>
                <span>{exportError}</span>
                <button className="text-btn" onClick={() => setExportState("idle")}>Dismiss error</button>
              </div>
            )}
            <div className="preview-sheet">
              <div className="preview-sheet-head">
                <span>{template.type}</span>
                  <span>Page 1 of {estimatePageCount()}</span>
              </div>
              <h1>{template.type}</h1>
              <p>
                Between {values.company_name} and {values.full_name}
              </p>
              {sections.map((section) => (
                <div className="preview-section" key={section.id}>
                  <h3>{section.title}</h3>
                  {section.clauses.filter(shouldShowClause).map((c) => (
                    <p
                      key={c.id}
                      dangerouslySetInnerHTML={{ __html: clauseHtml(c) }}
                    />
                  ))}
                </div>
              ))}
            </div>
            <div className="modal-footer">
              <button className="outline-btn" onClick={makeDocx} disabled={blockingIssues.length > 0 || exportState === "exporting"}>
                <Download size={15} />
                {exportState === "exporting" ? "Exporting…" : "Export Word"}
              </button>
              <button className="primary-btn" onClick={makePdf} disabled={blockingIssues.length > 0 || exportState === "exporting"}>
                <Download size={15} />
                {exportState === "exporting" ? "Exporting…" : "Export PDF"}
              </button>
            </div>
          </div>
        </div>
      )}
      {showSyncDiff && (
        <div className="modal-backdrop" onClick={() => setShowSyncDiff(false)}>
          <div className="side-modal sync-modal" onClick={(e) => e.stopPropagation()}>
          <div className="modal-header">
              <div>
                <span className="eyebrow">ONBOARDING UPDATE</span>
                <h2>Review source changes</h2>
                <p>Choose which updates to apply. Manual overrides are marked and never replaced silently.</p>
                <small className="sync-source-meta">
                  Submission {person.submissionId} · detected {sourceUpdateMeta[personId]?.detectedAt || "just now"}
                </small>
              </div>
              <button className="icon-btn" onClick={() => setShowSyncDiff(false)}><X size={18} /></button>
            </div>
            <div className="sync-list">
              {syncDiffs.length ? syncDiffs.map((diff) => {
                const protectedValue = manualOverrides.includes(diff.key);
                return <label className="sync-row" key={diff.key}>
                  <input type="checkbox" checked={selectedSyncKeys.includes(diff.key) && !protectedValue} disabled={protectedValue || !canEditDocument} onChange={() => setSelectedSyncKeys((prev) => prev.includes(diff.key) ? prev.filter((key) => key !== diff.key) : [...prev, diff.key])} />
                  <span className="fake-check">✓</span>
                  <span><strong>{diff.key.replaceAll("_", " ")}</strong><small>Draft: {diff.oldValue || "Empty"} → <b>Onboarding: {diff.newValue || "Empty"}</b></small></span>
                  {protectedValue && <em><Lock size={11} /> Manual override · restore field first</em>}
                </label>;
              }) : <div className="empty-state"><CheckCircle2 size={18} /> No new source changes</div>}
            </div>
            <div className="modal-footer">
              <button className="outline-btn" onClick={dismissSyncDiff}>Keep current draft values</button>
              <button className="primary-btn" disabled={!selectedSyncKeys.some((key) => !manualOverrides.includes(key)) || !canEditDocument} onClick={applySelectedSync}><RefreshCw size={14} /> Apply selected updates</button>
            </div>
          </div>
        </div>
      )}
      {showHistory && (
        <div className="modal-backdrop" onClick={() => setShowHistory(false)}>
          <div className="side-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <span className="eyebrow">DOCUMENT HISTORY</span>
                <h2>Versions</h2>
              </div>
              <button
                className="icon-btn"
                aria-label="Close version history"
                onClick={() => setShowHistory(false)}
              >
                <X size={18} />
              </button>
            </div>
            <div className="history-banner">
              <History size={15} />
              <span>Approved versions stay unchanged. Restore always creates a new draft.</span>
            </div>
            {(versions.length ? versions : [{
              id: "current",
              label: `v0.1 · ${docStatus}`,
              createdAt: "Sep 11, 2026 · 10:15",
              author: "JL",
              status: docStatus,
              summary: "Current document state",
              sections,
              values,
              letterhead,
              docName,
            } as DocumentVersion]).map((version, i) => (
              <div className="version-row" key={version.id}>
                <div className="version-dot" />
                <div>
                  <strong>{version.label}</strong>
                  <span>{version.createdAt} · {version.author}</span>
                  <small>{version.summary}</small>
                </div>
                <div className="version-actions">
                  <button className="text-btn" onClick={() => setCompareVersionId(version.id)}>Compare</button>
                  {i === 0 ? <span className="status-pill draft">Current</span> : <button className="text-btn" disabled={currentRole !== "Editor" && currentRole !== "Admin"} onClick={() => setRestoreVersionId(version.id)}>Restore</button>}
                </div>
              </div>
            ))}
            {compareVersion && (
              <div className="version-compare">
                <div className="modal-header">
                  <div><span className="eyebrow">VERSION DIFF</span><h3>{compareVersion.label}</h3></div>
                  <button className="icon-btn" onClick={() => setCompareVersionId(null)}><X size={15} /></button>
                </div>
                <div className="diff-grid">
                  <div><strong>Person</strong><span>{compareVersion.values.full_name || "Not set"} → {values.full_name || "Not set"}</span></div>
                  <div><strong>Clauses</strong><span>{compareVersion.sections.flatMap((s) => s.clauses).filter((c) => c.included).length} saved → {sections.flatMap((s) => s.clauses).filter((c) => c.included).length} current</span></div>
                  <div><strong>Letterhead</strong><span>{compareVersion.letterhead.mode} → {letterhead.mode}</span></div>
                </div>
                <div className="version-change-list">
                  <strong>Changed values and content</strong>
                  {compareChanges.length ? compareChanges.slice(0, 12).map((change, index) => (
                    <div className="version-change" key={`${change.label}-${index}`}>
                      <span>{change.label}</span>
                      <small>{change.before} <b>→</b> {change.after}</small>
                    </div>
                  )) : <span className="empty-state">No differences from the current draft.</span>}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
      {restoreVersion && (
        <div className="modal-backdrop" onClick={() => setRestoreVersionId(null)}>
          <div className="confirm-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div><span className="eyebrow">RESTORE VERSION</span><h2>Restore {restoreVersion.label}?</h2><p>Your current edits remain in history. Restoring creates a new draft and does not overwrite the saved version.</p></div>
              <button className="icon-btn" onClick={() => setRestoreVersionId(null)}><X size={17} /></button>
            </div>
            <div className="modal-footer"><button className="outline-btn" onClick={() => setRestoreVersionId(null)}>Cancel</button><button className="primary-btn" onClick={() => restoreVersionNow(restoreVersion)}><RotateCcw size={14} /> Restore as draft</button></div>
          </div>
        </div>
      )}
      {showAdmin && (
        <div className="modal-backdrop" onClick={() => setShowAdmin(false)}>
          <div className="admin-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <span className="eyebrow">ADMIN CONSOLE</span>
                <h2>Workspace management</h2>
                <p>Configuration changes apply to new documents only.</p>
              </div>
              <button
                className="icon-btn"
                aria-label="Close workspace management"
                onClick={() => setShowAdmin(false)}
              >
                <X size={18} />
              </button>
            </div>
            {!adminView ? (
              <>
                <div className="admin-grid">
                  {adminCards.map(([title, meta, Icon]) => (
                    <button
                      className="admin-card"
                      key={title}
                      onClick={() => setAdminView(title)}
                    >
                      <div className="admin-icon">
                        <Icon size={17} />
                      </div>
                      <div>
                        <strong>{title}</strong>
                        <span>{meta}</span>
                      </div>
                      <ChevronRight size={15} />
                    </button>
                  ))}
                </div>
                <div className="rule-preview">
                  <div className="rule-head">
                    <span>Active rule</span>
                    <span className="rule-status">Live</span>
                  </div>
                  <p>
                    <strong>When</strong> billing_method <strong>equals</strong>{" "}
                    milestone <strong>show</strong> payment_milestones and milestone
                    clause
                  </p>
                  <div className="rule-test">
                    <span>Test mode</span>
                    <span className="test-chip">
                      Aisha Rahman · fixed fee <b>→ hidden</b>
                    </span>
                    <span className="test-chip active">
                      Nadia Chen · milestone <b>→ shown</b>
                    </span>
                  </div>
                </div>
              </>
            ) : (
              <div className="admin-detail">
                <button className="back-link" onClick={() => setAdminView(null)}>
                  <ArrowUp size={14} /> Back to management
                </button>
                <h3>{adminView}</h3>
                {adminView === "Templates" && (
                  <div className="admin-list">
                    {templates.map((item) => (
                      <div className="admin-list-row" key={item.id}>
                        <div><strong>{item.type}</strong><span>{item.description}</span></div>
                        <button className="text-btn" onClick={() => { setShowAdmin(false); changeTemplate(item.id); }}>Use template</button>
                      </div>
                    ))}
                  </div>
                )}
                {adminView === "Placeholders" && (
                  <div className="admin-list">
                    {placeholderMeta.map(([key, label, source, group]) => (
                      <div className="admin-list-row" key={key}>
                        <div><strong>{label}</strong><span>{`{{${key}}}`} · {group} · {source}</span></div>
                        <button className="text-btn" onClick={() => setDependencyTarget(key)}>View usage</button>
                        <span className="status-pill draft">Active</span>
                      </div>
                    ))}
                    {customPlaceholders.map((item) => (
                      <div className="admin-list-row" key={item.key}>
                        <div><strong>{item.label}</strong><span>{`{{${item.key}}}`} · {item.group} · {item.source}</span></div>
                        <button className="text-btn" onClick={() => setDependencyTarget(item.key)}>View usage</button>
                        <span className="status-pill draft">Custom</span>
                      </div>
                    ))}
                    <button className="add-field" onClick={() => createPlaceholder()}><Plus size={14} /> Create placeholder</button>
                  </div>
                )}
                {adminView === "Dropdowns" && (
                  <div className="admin-list">
                    {["Work arrangement: On-site, Remote, Hybrid", "Billing method: Fixed fee, Hourly, Milestone", "Probation: 3 months, Not applicable"].map((item) => <div className="admin-list-row" key={item}><div><strong>{item.split(":")[0]}</strong><span>{item.split(":")[1]}</span></div><span className="status-pill draft">Configured</span></div>)}
                  </div>
                )}
                {adminView === "Connections" && <div className={`admin-status-panel ${connectionStatus === "Failed" ? "admin-status-error" : ""}`}>
                  {connectionStatus === "Failed" ? <AlertTriangle size={18} /> : <Cloud size={18} />}
                  <strong>Onboarding Form · {connectionStatus}</strong>
                  <span>{connectionStatus === "Failed" ? "The last sync failed. No source values were changed." : `Last sync: ${lastSyncAt || sourceUpdateMeta[personId]?.detectedAt || person.submitted}`}</span>
                  <div className="button-stack">
                    {connectionStatus === "Failed" ? <button className="add-field" onClick={retryConnection}><RefreshCw size={14} /> Retry connection</button> : <button className="add-field" onClick={() => { setConnectionStatus("Failed"); setToast("Simulated connection failure. Draft values are safe."); window.setTimeout(() => setToast(""), 2400); }}><AlertTriangle size={14} /> Test failure state</button>}
                    <button className="text-btn" onClick={checkOnboardingUpdates}>Check for updates</button>
                  </div>
                </div>}
                {adminView === "Letterheads" && <div className="admin-status-panel"><ImagePlus size={18} /><strong>{activeLetterheadLayout.fileName || "Northstar letterhead"}</strong><span>Applied to {letterhead.mode === "first" ? "first page" : letterhead.mode === "all" ? "every page" : "first and subsequent pages"}</span><button className="add-field" onClick={() => { setShowAdmin(false); setActiveRightTab("letterhead"); }}>Open layout editor</button></div>}
                {adminView === "Rules" && <div className="admin-status-panel"><WandSparkles size={18} /><strong>Billing method rule · Live</strong><span>Milestone shows payment milestones; other methods hide it.</span><button className="add-field" onClick={() => setToast("Rule test uses the selected document values")}>Run test</button></div>}
                {dependencyTarget && (
                  <div className="dependency-panel">
                    <div className="modal-header"><div><span className="eyebrow">DEPENDENCIES</span><h3>{`{{${dependencyTarget}}}`}</h3></div><button className="icon-btn" onClick={() => setDependencyTarget(null)}><X size={15} /></button></div>
                    <p>This field is referenced by the current template, active clauses and rule checks. Choose a replacement or deactivate it before removal.</p>
                    <div className="dependency-item"><FileText size={14} /><span><strong>{template.type}</strong><small>Current document template</small></span></div>
                    {sections.flatMap((section) => section.clauses).filter((clause) => clause.text.includes(`{{${dependencyTarget}}}`)).slice(0, 4).map((clause) => <div className="dependency-item" key={clause.id}><Archive size={14} /><span><strong>{clause.title}</strong><small>Clause content reference</small></span></div>)}
                    <div className="modal-footer"><button className="outline-btn" onClick={() => setDependencyTarget(null)}>Cancel</button><button className="outline-btn" disabled={currentRole !== "Admin"} onClick={() => { if (currentRole !== "Admin") return; setToast("Field deactivated for new documents; existing drafts remain unchanged"); window.setTimeout(() => setToast(""), 2400); }}>Deactivate</button><button className="primary-btn" disabled={currentRole !== "Admin"} onClick={() => replacePlaceholderReferences(dependencyTarget)}>Replace reference</button></div>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}
      {pendingChange && (
        <div className="modal-backdrop" onClick={() => setPendingChange(null)}>
          <div className="confirm-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <span className="eyebrow">UNSAVED CONTEXT</span>
                <h2>{pendingChange.kind === "template" ? "Start with another template?" : "Load another person?"}</h2>
                <p>
                  {pendingChange.kind === "template"
                    ? "Switching templates replaces the current section structure and custom clauses. Your current draft will not be merged automatically."
                    : "This reloads mapped onboarding values. Current document overrides will be replaced so the document cannot mix people."
                  }
                </p>
              </div>
              <button className="icon-btn" aria-label="Close confirmation" onClick={() => setPendingChange(null)}><X size={18} /></button>
            </div>
            <div className="modal-footer">
              <button className="outline-btn" onClick={() => setPendingChange(null)}>Cancel</button>
              <button className="primary-btn" onClick={confirmPendingChange}>{pendingChange.kind === "template" ? "Switch and reset" : `Load ${pendingChange.label}`}</button>
            </div>
          </div>
        </div>
      )}
      {promptRequest && (
        <div className="modal-backdrop" onClick={() => setPromptRequest(null)}>
          <div className="confirm-modal prompt-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <span className="eyebrow">WORKSPACE INPUT</span>
                <h2>{promptRequest.title}</h2>
                <p>{promptRequest.description}</p>
              </div>
              <button className="icon-btn" aria-label="Close input dialog" onClick={() => setPromptRequest(null)}><X size={18} /></button>
            </div>
            <label className="prompt-field">
              <span>{promptRequest.kind === "placeholder" ? "Placeholder name" : promptRequest.kind === "replace" ? "Existing placeholder key" : promptRequest.kind === "link" ? "URL" : "Document name"}</span>
              <input
                autoFocus
                value={promptRequest.value}
                onChange={(e) => setPromptRequest((current) => current ? { ...current, value: e.target.value } : current)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    submitPromptRequest();
                  }
                  if (e.key === "Escape") setPromptRequest(null);
                }}
              />
            </label>
            <div className="modal-footer">
              <button className="outline-btn" onClick={() => setPromptRequest(null)}>Cancel</button>
              <button className="primary-btn" disabled={!promptRequest.value.trim()} onClick={submitPromptRequest}>Continue</button>
            </div>
          </div>
        </div>
      )}
      {toast && (
        <div className="toast">
          <FileCheck2 size={15} />
          {toast}
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
