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
  Bell,
  Bold,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleHelp,
  ClipboardPaste,
  Cloud,
  Copy,
  Download,
  Eye,
  EyeOff,
  FileCheck2,
  FileText,
  FolderOpen,
  GripVertical,
  Highlighter,
  History,
  ImagePlus,
  Inbox,
  Italic,
  KeyRound,
  Link2,
  List,
  ListOrdered,
  Lock,
  LogIn,
  LogOut,
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
  Send,
  Settings2,
  ShieldCheck,
  Sparkles,
  Strikethrough,
  Sun,
  Table2,
  Type,
  PenLine,
  PanelBottom,
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
type CanvasBlockKind =
  | "letterhead"
  | "meta"
  | "title"
  | "lede"
  | "section"
  | "signatures"
  | "footer";
type CanvasBlock = {
  id: string;
  kind: CanvasBlockKind;
  sectionId?: string;
};
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
type VerificationWatermark = {
  placement: "header" | "footer";
  alignment: "left" | "center" | "right";
  text: string;
  timestampFormat: string;
  verificationId: string;
};
type CustomPlaceholder = {
  key: string;
  label: string;
  source: string;
  group: string;
};
type UserRole = "Admin" | "Editor" | "Reviewer";
type AuthSession = {
  displayName: string;
  email: string;
  signedInAt: string;
};
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
  watermark?: VerificationWatermark;
  docName: string;
  templateId?: string;
  personId?: string;
  submissionId?: string;
  sourceSnapshot?: Record<string, string>;
  sourceSyncedAt?: string;
  manualOverrides?: string[];
  sourceUpdates?: Record<string, Record<string, string>>;
  sourceUpdateMeta?: Record<string, { submissionId: string; detectedAt: string }>;
  canvasBlocks?: CanvasBlock[];
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
type AppModule = "workspace" | "clauses" | "placeholders" | "layouts" | "documents" | "settings";
type EmailSettings = {
  provider: "gmail";
  authMethod: "oauth" | "app-password";
  gmailAddress: string;
  gmailAppPassword: string;
  senderName: string;
  replyToEmail: string;
  documentInboxEmail: string;
  notificationEmail: string;
  sendDocuments: boolean;
  receiveCopies: boolean;
  notificationsEnabled: boolean;
};
type EmailSettingsErrors = Partial<Record<keyof EmailSettings, string>>;
type EmailConnectionSummary = {
  id: string;
  provider: "gmail";
  authMethod: "oauth" | "app-password";
  senderEmail: string;
  senderName: string;
  replyToEmail: string;
  documentInboxEmail: string;
  notificationEmail: string;
  sendDocuments: boolean;
  receiveCopies: boolean;
  notificationsEnabled: boolean;
  status: "connected" | "expired" | "error" | "disconnected";
  lastTestedAt?: string | null;
  lastError?: string | null;
  updatedAt?: string;
};
type EmailPlatformStatus = {
  ready: boolean;
  googleOAuthReady: boolean;
  missing: string[];
  code?: string;
};
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
  watermark: VerificationWatermark;
  versions: DocumentVersion[];
  exports: ExportRecord[];
  updatedAt: string;
  canvasBlocks?: CanvasBlock[];
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
const fixedCanvasBlock = (kind: Exclude<CanvasBlockKind, "section">): CanvasBlock => ({
  id: `canvas-${kind}`,
  kind,
});
const makeDefaultCanvasBlocks = (items: Section[]): CanvasBlock[] => [
  fixedCanvasBlock("letterhead"),
  fixedCanvasBlock("meta"),
  fixedCanvasBlock("title"),
  fixedCanvasBlock("lede"),
  ...items.map((section) => ({
    id: `canvas-section-${section.id}`,
    kind: "section" as const,
    sectionId: section.id,
  })),
  fixedCanvasBlock("signatures"),
  fixedCanvasBlock("footer"),
];
const storageKey = "hr-doc-generator-state-v2";
const authStorageKey = "hr-doc-generator-auth-v1";
const defaultEmailSettings: EmailSettings = {
  provider: "gmail",
  authMethod: "oauth",
  gmailAddress: "",
  gmailAppPassword: "",
  senderName: "",
  replyToEmail: "",
  documentInboxEmail: "",
  notificationEmail: "",
  sendDocuments: true,
  receiveCopies: false,
  notificationsEnabled: true,
};
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
const makeVerificationId = () => {
  const randomPart = typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID().replaceAll("-", "").slice(0, 12)
    : `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`.slice(0, 12);
  return `DOC-${randomPart.toUpperCase()}`;
};
const makeDefaultWatermark = (): VerificationWatermark => ({
  placement: "footer",
  alignment: "center",
  text: "VERIFIED DOCUMENT",
  timestampFormat: "YYYY-MM-DD HH:mm:ss Z",
  verificationId: makeVerificationId(),
});
const normalizeWatermark = (value?: Partial<VerificationWatermark>): VerificationWatermark => ({
  placement: value?.placement === "header" ? "header" : "footer",
  alignment: value?.alignment === "left" || value?.alignment === "right" ? value.alignment : "center",
  text: value?.text ?? "VERIFIED DOCUMENT",
  timestampFormat: value?.timestampFormat ?? "YYYY-MM-DD HH:mm:ss Z",
  verificationId: value?.verificationId || makeVerificationId(),
});
const formatWatermarkTimestamp = (date: Date, pattern: string) => {
  const pad = (value: number) => String(value).padStart(2, "0");
  const hours = date.getHours();
  const offsetMinutes = -date.getTimezoneOffset();
  const offsetSign = offsetMinutes >= 0 ? "+" : "-";
  const absoluteOffset = Math.abs(offsetMinutes);
  const tokens: Record<string, string> = {
    YYYY: String(date.getFullYear()),
    YY: String(date.getFullYear()).slice(-2),
    MMM: date.toLocaleDateString("en-US", { month: "short" }),
    MM: pad(date.getMonth() + 1),
    DD: pad(date.getDate()),
    HH: pad(hours),
    hh: pad(hours % 12 || 12),
    mm: pad(date.getMinutes()),
    ss: pad(date.getSeconds()),
    A: hours >= 12 ? "PM" : "AM",
    Z: `${offsetSign}${pad(Math.floor(absoluteOffset / 60))}:${pad(absoluteOffset % 60)}`,
  };
  return (pattern.trim() || "YYYY-MM-DD HH:mm:ss Z").replace(/YYYY|MMM|YY|MM|DD|HH|hh|mm|ss|A|Z/g, (token) => tokens[token]);
};
const watermarkDisplayText = (settings: VerificationWatermark, date = new Date()) =>
  `${settings.text.trim() || "VERIFIED DOCUMENT"} · ${settings.verificationId} · ${formatWatermarkTimestamp(date, settings.timestampFormat)}`;
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
const readAuthSession = (): AuthSession | null => {
  try {
    const raw = sessionStorage.getItem(authStorageKey) || localStorage.getItem(authStorageKey);
    return raw ? JSON.parse(raw) as AuthSession : null;
  } catch {
    return null;
  }
};
const moduleFromHash = (hash: string): AppModule => {
  const value = hash.replace(/^#/, "");
  return (["workspace", "clauses", "placeholders", "layouts", "documents", "settings"] as AppModule[]).includes(value as AppModule)
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
        documents: (parsed.documents || []).map((document) => ({
          ...document,
          watermark: normalizeWatermark(document.watermark),
          versions: (document.versions || []).map((version) => ({
            ...version,
            watermark: normalizeWatermark(version.watermark || document.watermark),
          })),
        })),
        exports: parsed.exports || [],
      };
    }
    const legacyRaw = localStorage.getItem(storageKey);
    const legacy = legacyRaw ? JSON.parse(legacyRaw) : {};
    const legacyTemplates: Template[] = legacy.templates || demoTemplates;
    const legacySections: Section[] = legacy.sections || clone(demoTemplates[1].sections);
    const legacyValues = legacy.values || clone(people[1].fields);
    const legacyLetterhead = normalizeLetterhead(legacy.letterhead);
    const legacyWatermark = normalizeWatermark(legacy.watermark);
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
      watermark: legacyWatermark,
      versions: (legacy.versions || []).map((version: DocumentVersion) => ({
        ...version,
        watermark: normalizeWatermark(version.watermark || legacyWatermark),
      })),
      exports: [],
      updatedAt: legacy.lastSavedAt || formatDocumentStamp(),
      canvasBlocks: legacy.canvasBlocks || makeDefaultCanvasBlocks(legacySections),
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

function LoginScreen({
  theme,
  onThemeChange,
  onLogin,
}: {
  theme: "light" | "dark";
  onThemeChange: () => void;
  onLogin: (session: AuthSession, remember: boolean) => void;
}) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [remember, setRemember] = useState(true);
  const [submitting, setSubmitting] = useState(false);

  const submitLogin = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (submitting) return;
    setSubmitting(true);
    const normalizedEmail = email.trim();
    const label = normalizedEmail
      ? normalizedEmail.split("@")[0].replace(/[._-]+/g, " ").trim()
      : "Demo HR user";
    window.setTimeout(() => {
      setSubmitting(false);
      onLogin({
        displayName: label || "Demo HR user",
        email: normalizedEmail,
        signedInAt: new Date().toISOString(),
      }, remember);
    }, 320);
  };

  return (
    <main className="login-screen">
      <section className="login-context" aria-label="ZhiReady workspace information">
        <div className="login-brand">
          <img
            src={theme === "light" ? "/brand/zhiready-light.png" : "/brand/zhiready-dark.png"}
            alt="ZhiReady"
          />
          <small>Document operations workspace</small>
        </div>

        <div className="login-context-copy">
          <span className="login-kicker">Controlled document workflow</span>
          <h1>HR documents, ready for review.</h1>
          <p>Create, verify and export employment documents in one focused workspace.</p>

          <div className="login-workspace-summary">
            <div className="login-workspace-heading">
              <span className="login-workspace-mark">N</span>
              <span><strong>Northstar Labs</strong><small>Malaysia workspace</small></span>
              <CheckCircle2 size={17} aria-label="Workspace available" />
            </div>
            <div className="login-capability-list">
              <span><FileText size={15} /><span><strong>Document workspace</strong><small>Build, review and export</small></span></span>
              <span><Cloud size={15} /><span><strong>Onboarding sources</strong><small>Mapped data and overrides</small></span></span>
              <span><ShieldCheck size={15} /><span><strong>Approval controls</strong><small>Versions and document status</small></span></span>
            </div>
          </div>
        </div>

        <p className="login-context-note">Demo environment. No employee notification is sent from this screen.</p>
      </section>

      <section className="login-form-region">
        <button
          className="login-theme-toggle"
          type="button"
          onClick={onThemeChange}
          aria-label={theme === "light" ? "Switch to dark theme" : "Switch to light theme"}
          title={theme === "light" ? "Switch to dark theme" : "Switch to light theme"}
        >
          {theme === "light" ? <Moon size={17} /> : <Sun size={17} />}
        </button>

        <form className="login-form" onSubmit={submitLogin}>
          <div className="login-form-heading">
            <span>Welcome back</span>
            <h2>Sign in to your workspace</h2>
            <p>Use your work account, or continue with the blank demo access.</p>
          </div>

          <label className="login-field">
            <span>Work email <small>Optional</small></span>
            <input
              type="email"
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              autoComplete="username"
              placeholder="name@company.com"
              autoFocus
            />
          </label>

          <label className="login-field">
            <span>Password <small>Optional</small></span>
            <span className="login-password-field">
              <input
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                autoComplete="current-password"
                placeholder="Enter your password"
              />
              <button
                type="button"
                onClick={() => setShowPassword((current) => !current)}
                aria-label={showPassword ? "Hide password" : "Show password"}
                title={showPassword ? "Hide password" : "Show password"}
              >
                {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </span>
          </label>

          <label className="login-remember">
            <input type="checkbox" checked={remember} onChange={(event) => setRemember(event.target.checked)} />
            <span>Keep me signed in on this device</span>
          </label>

          <button className="login-submit" type="submit" disabled={submitting}>
            <LogIn size={17} />
            {submitting ? "Signing in..." : "Sign in"}
          </button>

          <div className="login-demo-note" aria-live="polite">
            <ShieldCheck size={17} />
            <span><strong>Blank access is enabled</strong><small>Leave both fields empty and select Sign in. Passwords are not stored in this browser demo.</small></span>
          </div>
        </form>
      </section>
    </main>
  );
}

function App() {
  const [theme, setTheme] = useState<"light" | "dark">(() => {
    const savedTheme = localStorage.getItem("hr-doc-generator-theme");
    return savedTheme === "dark" ? "dark" : "light";
  });
  const [authSession, setAuthSession] = useState<AuthSession | null>(() => readAuthSession());

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem("hr-doc-generator-theme", theme);
  }, [theme]);

  const handleLogin = (session: AuthSession, remember: boolean) => {
    try {
      localStorage.removeItem(authStorageKey);
      sessionStorage.removeItem(authStorageKey);
      const storage = remember ? localStorage : sessionStorage;
      storage.setItem(authStorageKey, JSON.stringify(session));
    } finally {
      setAuthSession(session);
    }
  };

  const handleSignOut = () => {
    localStorage.removeItem(authStorageKey);
    sessionStorage.removeItem(authStorageKey);
    setAuthSession(null);
  };

  if (!authSession) {
    return (
      <LoginScreen
        theme={theme}
        onThemeChange={() => setTheme((current) => current === "light" ? "dark" : "light")}
        onLogin={handleLogin}
      />
    );
  }

  return (
    <WorkspaceApp
      authSession={authSession}
      theme={theme}
      setTheme={setTheme}
      onSignOut={handleSignOut}
    />
  );
}

function WorkspaceApp({
  authSession,
  theme,
  setTheme,
  onSignOut,
}: {
  authSession: AuthSession;
  theme: "light" | "dark";
  setTheme: React.Dispatch<React.SetStateAction<"light" | "dark">>;
  onSignOut: () => void;
}) {
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
  const [emailSettings, setEmailSettings] = useState<EmailSettings>(() => clone(defaultEmailSettings));
  const [emailSettingsStatus, setEmailSettingsStatus] = useState<"idle" | "loading" | "saving" | "saved" | "error">("idle");
  const [emailSettingsError, setEmailSettingsError] = useState("");
  const [emailSettingsErrors, setEmailSettingsErrors] = useState<EmailSettingsErrors>({});
  const [emailConnection, setEmailConnection] = useState<EmailConnectionSummary | null>(null);
  const [emailPlatform, setEmailPlatform] = useState<EmailPlatformStatus>({ ready: false, googleOAuthReady: false, missing: [] });
  const [emailAdvancedOpen, setEmailAdvancedOpen] = useState(false);
  const [showGmailAppPassword, setShowGmailAppPassword] = useState(false);
  const [templates, setTemplates] = useState<Template[]>(() => {
    const saved = localStorage.getItem(storageKey);
    return saved ? JSON.parse(saved).templates : demoTemplates;
  });
  const [templateId, setTemplateId] = useState("fixed");
  const [personId, setPersonId] = useState("EMP-2041");
  const [sections, setSections] = useState<Section[]>(() =>
    clone(demoTemplates[1].sections),
  );
  const [canvasBlocks, setCanvasBlocks] = useState<CanvasBlock[]>(() =>
    makeDefaultCanvasBlocks(demoTemplates[1].sections),
  );
  const [canvasDrag, setCanvasDrag] = useState<
    | { type: "block"; id: string }
    | { type: "clause"; id: string; sectionId: string }
    | null
  >(null);
  const [activeCanvasBlock, setActiveCanvasBlock] = useState("canvas-title");
  const [addBlockMenuOpen, setAddBlockMenuOpen] = useState(false);
  const [deletedCanvasSnapshot, setDeletedCanvasSnapshot] = useState<{
    sections: Section[];
    blocks: CanvasBlock[];
    label: string;
  } | null>(null);
  const [values, setValues] = useState<Record<string, string>>(() =>
    clone(people[1].fields),
  );
  const [letterhead, setLetterhead] = useState<Letterhead>({
    ...makeDefaultLetterhead(),
  });
  const [watermark, setWatermark] = useState<VerificationWatermark>(() => {
    try {
      const saved = localStorage.getItem(storageKey);
      return normalizeWatermark(saved ? JSON.parse(saved).watermark : undefined);
    } catch {
      return makeDefaultWatermark();
    }
  });
  const [letterheadEditorPage, setLetterheadEditorPage] = useState<"first" | "subsequent">("first");
  const [workflowStage, setWorkflowStage] = useState<WorkflowStage>("build");
  const [workspacePanelOpen, setWorkspacePanelOpen] = useState(false);
  const [activeSection, setActiveSection] = useState("parties");
  const [activeClause, setActiveClause] = useState("fixed-1");
  const [activeRightTab, setActiveRightTab] = useState<
    "outline" | "person" | "placeholders" | "clauses" | "letterhead" | "review"
  >("outline");
  const [workspaceLibraryOpen, setWorkspaceLibraryOpen] = useState<
    "clauses" | "placeholders" | "layouts" | null
  >(null);
  const [workspaceLibrarySearch, setWorkspaceLibrarySearch] = useState("");
  const [workspaceClausePreviewId, setWorkspaceClausePreviewId] = useState<string | null>(null);
  const [workspaceToolMenuOpen, setWorkspaceToolMenuOpen] = useState(false);
  const [clauseScope, setClauseScope] = useState<"all" | "current">("all");
  const [search, setSearch] = useState("");
  const [saving, setSaving] = useState<"saved" | "saving" | "error">("saved");
  const [lastSavedAt, setLastSavedAt] = useState("");
  const [toast, setToast] = useState("");
  const [showPreview, setShowPreview] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);
  const [adminView, setAdminView] = useState<string | null>(null);
  const [companyMenuOpen, setCompanyMenuOpen] = useState(false);
  const [accountMenuOpen, setAccountMenuOpen] = useState(false);
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
    const onHashChange = () => setActiveModule(moduleFromHash(window.location.hash));
    window.addEventListener("hashchange", onHashChange);
    if (!window.location.hash) window.history.replaceState(null, "", "#workspace");
    return () => window.removeEventListener("hashchange", onHashChange);
  }, []);

  useEffect(() => {
    if (activeModule !== "settings") return;
    let cancelled = false;
    setEmailSettingsStatus("loading");
    setEmailSettingsError("");
    fetch("/api/email-connections/status", { headers: { Accept: "application/json" } })
      .then(async (response) => {
        const data = await response.json().catch(() => ({
          ok: false,
          code: "BACKEND_NOT_CONFIGURED",
          error: "The cloud email service is not available from this development server.",
          platform: { ready: false, googleOAuthReady: false, missing: ["DATABASE_URL", "AUTH_SESSION_SECRET", "EMAIL_CREDENTIAL_ENCRYPTION_KEY"] },
        })) as {
          ok?: boolean;
          code?: string;
          error?: string;
          platform?: EmailPlatformStatus;
          connection?: EmailConnectionSummary | null;
        };
        return { response, data };
      })
      .then(({ response, data }) => {
        if (cancelled) return;
        setEmailPlatform(data.platform || { ready: false, googleOAuthReady: false, missing: [], code: data.code });
        setEmailConnection(data.connection || null);
        if (data.connection) {
          setEmailSettings({
            ...clone(defaultEmailSettings),
            authMethod: data.connection.authMethod,
            gmailAddress: data.connection.senderEmail,
            senderName: data.connection.senderName,
            replyToEmail: data.connection.replyToEmail,
            documentInboxEmail: data.connection.documentInboxEmail,
            notificationEmail: data.connection.notificationEmail,
            sendDocuments: data.connection.sendDocuments,
            receiveCopies: data.connection.receiveCopies,
            notificationsEnabled: data.connection.notificationsEnabled,
          });
        }
        setEmailSettingsStatus("idle");
        if (!response.ok) setEmailSettingsError(data.error || "The cloud email service is not configured.");
      })
      .catch(() => {
        if (cancelled) return;
        setEmailPlatform({ ready: false, googleOAuthReady: false, missing: ["DATABASE_URL", "AUTH_SESSION_SECRET", "EMAIL_CREDENTIAL_ENCRYPTION_KEY"] });
        setEmailSettingsStatus("idle");
        setEmailSettingsError("The cloud email service is not available from this development server.");
      });
    return () => {
      cancelled = true;
    };
  }, [activeModule]);

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
  const updateWatermark = (patch: Partial<VerificationWatermark>) => {
    if (!guardEdit()) return;
    setWatermark((current) => ({ ...current, ...patch }));
  };
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
  const hasCanvasBlock = (kind: CanvasBlockKind) =>
    canvasBlocks.some((block) => block.kind === kind);
  const orderedSections = canvasBlocks
    .filter((block) => block.kind === "section" && block.sectionId)
    .map((block) => sections.find((section) => section.id === block.sectionId))
    .filter((section): section is Section => Boolean(section));
  const clausePanelSections = clauseScope === "current"
    ? orderedSections.filter((section) => section.id === activeSection)
    : orderedSections;
  const focusSection = (sectionId: string, openClauses = true) => {
    const section = sections.find((item) => item.id === sectionId);
    if (!section) return;
    setActiveSection(section.id);
    setActiveClause(section.clauses[0]?.id || "");
    const block = canvasBlocks.find((item) => item.kind === "section" && item.sectionId === section.id);
    if (block) setActiveCanvasBlock(block.id);
    if (openClauses) setActiveRightTab("clauses");
    window.requestAnimationFrame(() => {
      document.querySelector(`[data-section-id="${section.id}"]`)?.scrollIntoView({ behavior: "smooth", block: "center" });
    });
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
        setWatermark(normalizeWatermark(state.watermark));
        const restoredSections = state.sections || clone(demoTemplates[1].sections);
        setSections(restoredSections);
        setCanvasBlocks(state.canvasBlocks || makeDefaultCanvasBlocks(restoredSections));
        setActiveSection(restoredSections[0]?.id || "");
        setActiveClause(restoredSections[0]?.clauses[0]?.id || "");
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
              watermark,
              sections,
              canvasBlocks,
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
    watermark,
    sections,
    canvasBlocks,
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

  useEffect(() => {
    setCanvasBlocks((current) => {
      const sectionIds = new Set(sections.map((section) => section.id));
      const retained = current.filter(
        (block) =>
          block.kind !== "section" ||
          Boolean(block.sectionId && sectionIds.has(block.sectionId)),
      );
      const placed = new Set(
        retained
          .filter((block) => block.kind === "section")
          .map((block) => block.sectionId),
      );
      const missing = sections
        .filter((section) => !placed.has(section.id))
        .map((section) => ({
          id: `canvas-section-${section.id}`,
          kind: "section" as const,
          sectionId: section.id,
        }));
      if (!missing.length && retained.length === current.length) return current;
      const footerIndex = retained.findIndex(
        (block) => block.kind === "signatures" || block.kind === "footer",
      );
      if (footerIndex < 0) return [...retained, ...missing];
      return [
        ...retained.slice(0, footerIndex),
        ...missing,
        ...retained.slice(footerIndex),
      ];
    });
  }, [sections]);

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
    const nextSections = clone(next.sections);
    setSections(nextSections);
    setCanvasBlocks(makeDefaultCanvasBlocks(nextSections));
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
    setAddBlockMenuOpen(false);
  };
  const addSection = () => {
    if (!guardEdit()) return;
    const id = `section-${Date.now()}`;
    const section: Section = {
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
    };
    setSections((prev) => [...prev, section]);
    const block: CanvasBlock = {
      id: `canvas-section-${id}`,
      kind: "section",
      sectionId: id,
    };
    setCanvasBlocks((current) => {
      const footerIndex = current.findIndex(
        (item) => item.kind === "signatures" || item.kind === "footer",
      );
      if (footerIndex < 0) return [...current, block];
      return [
        ...current.slice(0, footerIndex),
        block,
        ...current.slice(footerIndex),
      ];
    });
    setActiveSection(id);
    setActiveClause(`${id}-clause`);
    setActiveCanvasBlock(block.id);
    setAddBlockMenuOpen(false);
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
    if (hasCanvasBlock("letterhead") && firstLayout.margin <= firstLayout.top + 8) {
      issues.push("Letterhead overlaps the document body");
    }
    if (!watermark.text.trim()) {
      issues.push("Verification watermark text is required");
    }
    if (!watermark.timestampFormat.trim()) {
      issues.push("Verification timestamp format is required");
    }
    return issues;
  }, [canvasBlocks, letterhead, person.id, person.name, person.role, personId, sections, templateId, values, watermark]);
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
      action: "Run review",
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
      setActiveRightTab("outline");
      setWorkspacePanelOpen(false);
      setShowPreview(false);
      return;
    }
    if (stage === "review") {
      setActiveRightTab("review");
      setWorkspacePanelOpen(true);
      setShowPreview(false);
      return;
    }
    setShowPreview(true);
  };
  const rememberCanvasBeforeDelete = (label: string) => {
    setDeletedCanvasSnapshot({
      sections: clone(sections),
      blocks: clone(canvasBlocks),
      label,
    });
  };
  const restoreDeletedCanvasItem = () => {
    if (!deletedCanvasSnapshot) return;
    setSections(clone(deletedCanvasSnapshot.sections));
    setCanvasBlocks(clone(deletedCanvasSnapshot.blocks));
    setDeletedCanvasSnapshot(null);
    setToast(`${deletedCanvasSnapshot.label} restored`);
    window.setTimeout(() => setToast(""), 2200);
  };
  const removeCanvasBlock = (block: CanvasBlock) => {
    if (!guardEdit()) return;
    const section = block.sectionId
      ? sections.find((item) => item.id === block.sectionId)
      : null;
    const label = section?.title || block.kind.replaceAll("-", " ");
    rememberCanvasBeforeDelete(label);
    if (block.kind === "section" && block.sectionId) {
      setSections((current) =>
        current.filter((item) => item.id !== block.sectionId),
      );
      setActiveSection("");
      setActiveClause("");
    }
    setCanvasBlocks((current) => current.filter((item) => item.id !== block.id));
    setActiveCanvasBlock("");
    setToast(`Removed ${label} from this document`);
    window.setTimeout(() => setToast(""), 4200);
  };
  const clearCanvas = () => {
    if (!guardEdit() || (!canvasBlocks.length && !sections.length)) return;
    rememberCanvasBeforeDelete("Canvas content");
    setSections([]);
    setCanvasBlocks([]);
    setActiveSection("");
    setActiveClause("");
    setActiveCanvasBlock("");
    setAddBlockMenuOpen(false);
    setToast("Canvas cleared");
    window.setTimeout(() => setToast(""), 4200);
  };
  const addFixedCanvasBlock = (kind: Exclude<CanvasBlockKind, "section">) => {
    if (!guardEdit()) return;
    const existing = canvasBlocks.find((block) => block.kind === kind);
    if (existing) {
      setActiveCanvasBlock(existing.id);
      setToast(`${kind.replaceAll("-", " ")} is already on the page`);
      window.setTimeout(() => setToast(""), 1800);
      return;
    }
    const next = fixedCanvasBlock(kind);
    setCanvasBlocks((current) => [...current, next]);
    setActiveCanvasBlock(next.id);
    setAddBlockMenuOpen(false);
  };
  const addParagraphElement = () => {
    if (!guardEdit()) return;
    const targetSection = activeSection && sections.some((section) => section.id === activeSection)
      ? activeSection
      : sections[0]?.id;
    if (targetSection) {
      addClause(targetSection, true);
      const block = canvasBlocks.find((item) => item.kind === "section" && item.sectionId === targetSection);
      if (block) setActiveCanvasBlock(block.id);
      setAddBlockMenuOpen(false);
      return;
    }
    addSection();
    setAddBlockMenuOpen(false);
  };
  const moveCanvasBlock = (sourceId: string, targetId: string) => {
    if (!guardEdit() || sourceId === targetId) return;
    setCanvasBlocks((current) => {
      const sourceIndex = current.findIndex((block) => block.id === sourceId);
      const targetIndex = current.findIndex((block) => block.id === targetId);
      if (sourceIndex < 0 || targetIndex < 0) return current;
      const next = [...current];
      const [moved] = next.splice(sourceIndex, 1);
      next.splice(targetIndex, 0, moved);
      return next;
    });
  };
  const nudgeCanvasBlock = (blockId: string, direction: -1 | 1) => {
    if (!guardEdit()) return;
    const index = canvasBlocks.findIndex((block) => block.id === blockId);
    const target = canvasBlocks[index + direction];
    if (index < 0 || !target) return;
    moveCanvasBlock(blockId, target.id);
  };
  const removeClause = (sectionId: string, clauseId: string) => {
    if (!guardEdit()) return;
    const clause = sections
      .find((section) => section.id === sectionId)
      ?.clauses.find((item) => item.id === clauseId);
    rememberCanvasBeforeDelete(clause?.title || "Paragraph");
    setSections((current) =>
      current.map((section) =>
        section.id === sectionId
          ? {
              ...section,
              clauses: section.clauses.filter((item) => item.id !== clauseId),
            }
          : section,
      ),
    );
    setActiveClause("");
    setToast(`Removed ${clause?.title || "paragraph"} from this document`);
    window.setTimeout(() => setToast(""), 4200);
  };
  const moveClauseTo = (
    sourceSectionId: string,
    clauseId: string,
    targetSectionId: string,
    targetClauseId?: string,
  ) => {
    if (!guardEdit()) return;
    setSections((current) => {
      const sourceSection = current.find((section) => section.id === sourceSectionId);
      const moved = sourceSection?.clauses.find((clause) => clause.id === clauseId);
      if (!moved) return current;
      const without = current.map((section) =>
        section.id === sourceSectionId
          ? { ...section, clauses: section.clauses.filter((clause) => clause.id !== clauseId) }
          : section,
      );
      return without.map((section) => {
        if (section.id !== targetSectionId) return section;
        const nextClauses = [...section.clauses];
        const targetIndex = targetClauseId
          ? nextClauses.findIndex((clause) => clause.id === targetClauseId)
          : nextClauses.length;
        nextClauses.splice(targetIndex < 0 ? nextClauses.length : targetIndex, 0, moved);
        return { ...section, clauses: nextClauses };
      });
    });
    setActiveSection(targetSectionId);
    setActiveClause(clauseId);
  };
  const openNewBlankDocument = () => {
    if (!guardEdit()) return;
    setSections([]);
    setCanvasBlocks([]);
    setActiveSection("");
    setActiveClause("");
    setActiveCanvasBlock("");
    setDocName("Untitled HR document");
    setWatermark(makeDefaultWatermark());
    setDocumentMenuOpen(false);
    setToast("Blank document ready. Add only the elements you need.");
    window.setTimeout(() => setToast(""), 2400);
  };
  const copyCurrentDocument = () => {
    if (!guardEdit()) return;
    const stamp = Date.now();
    const sectionIdMap = new Map<string, string>();
    const copiedSections = sections.map((section) => {
      const nextSectionId = `${section.id}-copy-${stamp}`;
      sectionIdMap.set(section.id, nextSectionId);
      return {
        ...clone(section),
        id: nextSectionId,
        clauses: section.clauses.map((clause) => ({
          ...clone(clause),
          id: `${clause.id}-copy-${stamp}`,
        })),
      };
    });
    setSections(copiedSections);
    setCanvasBlocks((current) =>
      current.map((block) => {
        if (block.kind !== "section" || !block.sectionId) return clone(block);
        const nextSectionId = sectionIdMap.get(block.sectionId);
        return nextSectionId
          ? {
              ...block,
              id: `canvas-section-${nextSectionId}`,
              sectionId: nextSectionId,
            }
          : block;
      }),
    );
    setDocName(`${docName} · Copy`);
    setWatermark((current) => ({
      ...current,
      verificationId: makeVerificationId(),
    }));
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
      const sourceId = next[index].id;
      const targetId = next[target].id;
      [next[index], next[target]] = [next[target], next[index]];
      setCanvasBlocks((blocks) => {
        const sourceBlock = blocks.find((block) => block.sectionId === sourceId);
        const targetBlock = blocks.find((block) => block.sectionId === targetId);
        if (!sourceBlock || !targetBlock) return blocks;
        const copy = [...blocks];
        const sourceBlockIndex = copy.findIndex((block) => block.id === sourceBlock.id);
        const targetBlockIndex = copy.findIndex((block) => block.id === targetBlock.id);
        [copy[sourceBlockIndex], copy[targetBlockIndex]] = [copy[targetBlockIndex], copy[sourceBlockIndex]];
        return copy;
      });
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
      watermark: clone(payload.watermark ?? watermark),
      docName: payload.docName ?? docName,
      templateId: payload.templateId ?? templateId,
      personId: payload.personId ?? personId,
      submissionId: payload.submissionId ?? person.submissionId,
      sourceSnapshot: clone(payload.sourceSnapshot ?? person.fields),
      sourceSyncedAt: payload.sourceSyncedAt ?? lastSyncAt,
      manualOverrides: clone(payload.manualOverrides ?? manualOverrides),
      sourceUpdates: clone(payload.sourceUpdates ?? sourceUpdates),
      sourceUpdateMeta: clone(payload.sourceUpdateMeta ?? sourceUpdateMeta),
      canvasBlocks: clone(payload.canvasBlocks ?? canvasBlocks),
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
    const restoredWatermark = normalizeWatermark(version.watermark || watermark);
    const restoredDocName = version.docName;
    const restoredPersonId = version.personId || personId;
    const restoredTemplateId = version.templateId || templateId;
    const restoredOverrides = clone(version.manualOverrides || []);
    const restoredUpdates = clone(version.sourceUpdates || {});
    const restoredMeta = clone(version.sourceUpdateMeta || sourceUpdateMeta);
    const restoredCanvasBlocks = clone(
      version.canvasBlocks || makeDefaultCanvasBlocks(restoredSections),
    );
    setSections(restoredSections);
    setCanvasBlocks(restoredCanvasBlocks);
    setValues(restoredValues);
    setLetterhead(restoredLetterhead);
    setWatermark(restoredWatermark);
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
      watermark: restoredWatermark,
      docName: restoredDocName,
      personId: restoredPersonId,
      templateId: restoredTemplateId,
      submissionId: version.submissionId,
      sourceSnapshot: version.sourceSnapshot,
      sourceSyncedAt: version.sourceSyncedAt,
      manualOverrides: restoredOverrides,
      sourceUpdates: restoredUpdates,
      sourceUpdateMeta: restoredMeta,
      canvasBlocks: restoredCanvasBlocks,
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
          watermark,
          sections,
          canvasBlocks,
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
  const watermarkAlignment = () =>
    watermark.alignment === "left"
      ? AlignmentType.LEFT
      : watermark.alignment === "right"
        ? AlignmentType.RIGHT
        : AlignmentType.CENTER;
  const renderDocxWatermark = (text: string) => new Paragraph({
    alignment: watermarkAlignment(),
    children: [new TextRun({
      text,
      size: 14,
      color: "7C8699",
      italics: true,
    })],
  });
  const estimatePageCount = () => {
    const characters = orderedSections
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
      const exportedAt = new Date();
      const verificationText = watermarkDisplayText(watermark, exportedAt);
      const firstLayout = layoutForPage(0);
      const subsequentLayout = layoutForPage(1);
      const firstHeaderChildren = [
        ...(hasCanvasBlock("letterhead") ? [renderDocxLetterhead(firstLayout)] : []),
        ...(watermark.placement === "header" ? [renderDocxWatermark(verificationText)] : []),
      ];
      const subsequentHeaderChildren = [
        ...(hasCanvasBlock("letterhead") && letterhead.mode !== "first"
          ? [renderDocxLetterhead(letterhead.mode === "all" ? firstLayout : subsequentLayout)]
          : []),
        ...(watermark.placement === "header" ? [renderDocxWatermark(verificationText)] : []),
      ];
      const firstHeader = new Header({ children: firstHeaderChildren });
      const subsequentHeader = new Header({ children: subsequentHeaderChildren });
      const footer = new Footer({
        children: [
          ...(hasCanvasBlock("footer") ? [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `${values.company_name || "Company"} · Confidential · Page ` }), new TextRun({ children: [PageNumber.CURRENT] })],
          })] : []),
          ...(watermark.placement === "footer" ? [renderDocxWatermark(verificationText)] : []),
        ],
      });
      const children = canvasBlocks.flatMap((block): Paragraph[] => {
        if (block.kind === "meta") {
          return [new Paragraph({ text: `${template.type.toUpperCase()} · MY · 2026` })];
        }
        if (block.kind === "title") {
          return [new Paragraph({ text: docName, heading: HeadingLevel.TITLE })];
        }
        if (block.kind === "lede") {
          return [new Paragraph({ text: `Between ${values.company_name || "Company"} and ${values.full_name || "Recipient"}` })];
        }
        if (block.kind === "section" && block.sectionId) {
          const section = sections.find((item) => item.id === block.sectionId);
          if (!section) return [];
          return [
            new Paragraph({ text: section.title, heading: HeadingLevel.HEADING_1 }),
            ...section.clauses
              .filter(shouldShowClause)
              .map((clause) => new Paragraph({ children: clauseTextRuns(clause) })),
          ];
        }
        if (block.kind === "signatures") {
          return [
            new Paragraph({
              text: `Signed for ${values.company_name || "Company"} by ${values.signatory_name || "Company representative"}, ${values.signatory_title || ""}`,
            }),
            new Paragraph({ text: "Employee / Contractor signature: ______________________________" }),
            new Paragraph({ text: "Date: ____________________" }),
          ];
        }
        return [];
      });
      const blob = await Packer.toBlob(new Document({
        sections: [{
          headers: firstHeaderChildren.length || subsequentHeaderChildren.length
            ? { default: subsequentHeader, first: firstHeader }
            : {},
          footers: hasCanvasBlock("footer") || watermark.placement === "footer"
            ? { default: footer, first: footer }
            : {},
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
      const exportedAt = new Date();
      const verificationText = watermarkDisplayText(watermark, exportedAt);
      const pdf = new jsPDF({ format: letterhead.page === "A4" ? "a4" : "letter", unit: "pt" });
      const pageHeight = pdf.internal.pageSize.getHeight();
      const pageWidth = pdf.internal.pageSize.getWidth();
      const addPageFurniture = (pageIndex: number) => {
        const layout = layoutForPage(pageIndex);
        const imageType = letterheadImageType(layout);
        if (hasCanvasBlock("letterhead")) {
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
        }
        if (hasCanvasBlock("footer")) {
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(8);
          pdf.setTextColor("#6b7280");
          pdf.text(`${values.company_name || "Company"} · Confidential · Page ${pageIndex + 1}`, pageWidth / 2, pageHeight - 28, { align: "center" });
          pdf.setTextColor("#222222");
        }
        pdf.setFont("helvetica", "italic");
        pdf.setFontSize(7);
        pdf.setTextColor("#7c8699");
        const verificationLines = pdf.splitTextToSize(verificationText, pageWidth - 108) as string[];
        const verificationX = watermark.alignment === "left" ? 54 : watermark.alignment === "right" ? pageWidth - 54 : pageWidth / 2;
        const verificationY = watermark.placement === "header" ? 14 : pageHeight - 14 - Math.max(0, verificationLines.length - 1) * 8;
        const verificationAlign = watermark.alignment === "left" ? "left" : watermark.alignment === "right" ? "right" : "center";
        verificationLines.forEach((line, lineIndex) => {
          pdf.text(line, verificationX, verificationY + lineIndex * 8, { align: verificationAlign });
        });
        pdf.setFont("helvetica", "normal");
        pdf.setTextColor("#222222");
      };
      let pageIndex = 0;
      addPageFurniture(pageIndex);
      const firstLayout = layoutForPage(0);
      let y = hasCanvasBlock("letterhead") ? Math.max(70, firstLayout.margin || 74) : 54;
      const ensureSpace = (height: number) => {
        if (y <= pageHeight - height) return;
        pdf.addPage();
        pageIndex += 1;
        addPageFurniture(pageIndex);
        y = hasCanvasBlock("letterhead") ? Math.max(58, layoutForPage(pageIndex).margin || 58) : 54;
      };
      canvasBlocks.forEach((block) => {
        if (block.kind === "meta") {
          ensureSpace(28);
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(8);
          pdf.setTextColor("#64748b");
          pdf.text(`${template.type.toUpperCase()} · MY · 2026`, 54, y);
          pdf.setTextColor("#222222");
          y += 24;
        } else if (block.kind === "title") {
          ensureSpace(42);
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(16);
          pdf.text(docName, 54, y);
          y += 28;
        } else if (block.kind === "lede") {
          ensureSpace(34);
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(10);
          pdf.text(`Between ${values.company_name || "Company"} and ${values.full_name || "Recipient"}`, 54, y);
          y += 26;
        } else if (block.kind === "section" && block.sectionId) {
          const section = sections.find((item) => item.id === block.sectionId);
          if (!section) return;
          const visible = section.clauses.filter(shouldShowClause);
          ensureSpace(44);
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(12);
          pdf.text(section.title, 54, y);
          y += 18;
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(10);
          visible.forEach((clause) => {
            const lines = pdf.splitTextToSize(clausePlainText(clause), 490);
            lines.forEach((line: string) => {
              ensureSpace(70);
              pdf.text(line, 54, y);
              y += 14;
            });
            y += 6;
          });
        } else if (block.kind === "signatures") {
          ensureSpace(130);
          pdf.setFont("helvetica", "bold");
          pdf.text(`Signed for ${values.company_name || "Company"} by ${values.signatory_name || "Company representative"}, ${values.signatory_title || ""}`, 54, y + 20);
          pdf.setFont("helvetica", "normal");
          pdf.text("Employee / Contractor signature: ______________________________", 54, y + 48);
          pdf.text("Date: ____________________", 54, y + 68);
          y += 88;
        }
      });
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
    watermark: clone(watermark),
    versions: clone(versions),
    exports: clone(appStore.documents.find((item) => item.id === "doc-current")?.exports || []),
    updatedAt: formatDocumentStamp(),
    canvasBlocks: clone(canvasBlocks),
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
  const signOutOfWorkspace = () => {
    if (activeModule === "workspace" && !persistCurrentDocumentRecord()) {
      setAccountMenuOpen(false);
      return;
    }
    setAccountMenuOpen(false);
    onSignOut();
  };
  const switchModule = (module: AppModule) => {
    if (module === activeModule) return;
    if (activeModule === "workspace" && !persistCurrentDocumentRecord()) return;
    window.location.hash = module;
    setActiveModule(module);
  };
  const openWorkspaceTool = (
    tab: "outline" | "person" | "placeholders" | "clauses" | "letterhead" | "review",
    library?: "clauses" | "placeholders" | "layouts",
  ) => {
    setActiveRightTab(tab);
    setWorkspacePanelOpen(true);
    setWorkspaceLibraryOpen(library || null);
    setShowPreview(false);
  };
  const handlePrimaryNavigation = (module: AppModule) => {
    if (activeModule === "workspace") {
      const target = module === "workspace"
        ? { tab: "outline" as const, library: undefined }
        : module === "clauses"
          ? { tab: "clauses" as const, library: "clauses" as const }
          : module === "placeholders"
            ? { tab: "placeholders" as const, library: "placeholders" as const }
            : module === "layouts"
              ? { tab: "letterhead" as const, library: "layouts" as const }
              : null;
      if (target) {
        if (workspacePanelOpen && activeRightTab === target.tab) {
          setWorkspacePanelOpen(false);
          setWorkspaceToolMenuOpen(false);
          return;
        }
        return openWorkspaceTool(target.tab, target.library);
      }
    }
    if (module === "workspace") {
      switchModule(module);
      setActiveRightTab("outline");
      setWorkspacePanelOpen(true);
      return;
    }
    switchModule(module);
  };
  const moduleTitle = activeModule === "settings"
    ? "Email Settings"
    : navItems.find(([id]) => id === activeModule)?.[1] || "Workspace";
  const filteredClauses = appStore.clauses.filter((clause) => {
    const matchesSearch = !moduleSearch || `${clause.title} ${clause.category} ${clause.tags.join(" ")}`.toLowerCase().includes(moduleSearch.toLowerCase());
    const matchesFilter = moduleFilter === "all" || clause.status === moduleFilter;
    return matchesSearch && matchesFilter;
  });
  const selectedGroup = appStore.placeholderGroups.find((group) => group.id === (selectedPlaceholderGroupId || appStore.placeholderGroups[0]?.id));
  const workspaceClauseRecords = appStore.clauses.filter((record) =>
    record.status !== "inactive" &&
    (!workspaceLibrarySearch || `${record.title} ${record.category} ${record.tags.join(" ")}`.toLowerCase().includes(workspaceLibrarySearch.toLowerCase())),
  );
  const workspacePlaceholderGroups = appStore.placeholderGroups.filter((group) =>
    group.status !== "inactive" &&
    (!workspaceLibrarySearch || `${group.name} ${group.sourceType} ${group.fields.map((field) => field.label).join(" ")}`.toLowerCase().includes(workspaceLibrarySearch.toLowerCase())),
  );
  const workspaceLayoutRecords = appStore.layouts.filter((record) =>
    record.status !== "inactive" &&
    (!workspaceLibrarySearch || record.name.toLowerCase().includes(workspaceLibrarySearch.toLowerCase())),
  );
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
    setCanvasBlocks((current) => [
      ...current,
      { id: `canvas-section-${section.id}`, kind: "section", sectionId: section.id },
    ]);
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
        watermark: clone(watermark),
        docName,
        templateId,
        personId,
        status: docStatus === "In review" ? "in-review" : docStatus === "Approved" ? "approved" : "draft",
        generationStatus: status === "generated" ? "generated" : "failed",
        exports: status === "failed" ? existing.exports : [exportRecord, ...existing.exports],
        updatedAt: formatDocumentStamp(),
        canvasBlocks: clone(canvasBlocks),
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
  const renderCanvasBlock = (block: CanvasBlock, blockIndex: number) => {
    const selected = activeCanvasBlock === block.id;
    const section = block.sectionId
      ? sections.find((item) => item.id === block.sectionId)
      : null;
    const dropCanvasItem = (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      if (!canvasDrag) return;
      if (canvasDrag.type === "block") {
        moveCanvasBlock(canvasDrag.id, block.id);
      } else if (section) {
        moveClauseTo(canvasDrag.sectionId, canvasDrag.id, section.id);
      }
      setCanvasDrag(null);
    };
    const controls = (
      <div className="canvas-block-controls" aria-label="Element controls">
        <button
          className="canvas-drag-handle"
          title="Drag to reposition"
          aria-label="Drag element"
          draggable={canEditDocument}
          onDragStart={(event) => {
            event.dataTransfer.effectAllowed = "move";
            setCanvasDrag({ type: "block", id: block.id });
          }}
          onDragEnd={() => setCanvasDrag(null)}
        >
          <GripVertical size={14} />
        </button>
        <span>{block.kind === "section" ? "Section" : block.kind}</span>
        <button
          title="Move up"
          aria-label="Move element up"
          disabled={blockIndex === 0}
          onClick={(event) => {
            event.stopPropagation();
            nudgeCanvasBlock(block.id, -1);
          }}
        >
          <ArrowUp size={13} />
        </button>
        <button
          title="Move down"
          aria-label="Move element down"
          disabled={blockIndex === canvasBlocks.length - 1}
          onClick={(event) => {
            event.stopPropagation();
            nudgeCanvasBlock(block.id, 1);
          }}
        >
          <ArrowDown size={13} />
        </button>
        <button
          className="danger"
          title="Delete from document"
          aria-label="Delete element"
          onClick={(event) => {
            event.stopPropagation();
            removeCanvasBlock(block);
          }}
        >
          <Trash2 size={13} />
        </button>
      </div>
    );
    const shell = (content: React.ReactNode, className = "") => (
      <div
        className={`canvas-block ${className} ${selected ? "selected" : ""} ${canvasDrag?.type === "block" && canvasDrag.id === block.id ? "dragging" : ""}`}
        key={block.id}
        data-canvas-block={block.kind}
        data-section-id={section?.id}
        onClick={(event) => {
          event.stopPropagation();
          setActiveCanvasBlock(block.id);
          if (section) {
            setActiveSection(section.id);
            setActiveClause(section.clauses[0]?.id || "");
          }
        }}
        onDragOver={(event) => {
          event.preventDefault();
          event.dataTransfer.dropEffect = "move";
        }}
        onDrop={dropCanvasItem}
      >
        {controls}
        {content}
      </div>
    );

    if (block.kind === "letterhead") {
      const layout = letterhead.firstPage || letterhead;
      return shell(
        <div
          className="canvas-letterhead-space"
          style={{ minHeight: `${Math.max(56, layout.top + 46)}px` }}
        >
          <div
            className="letterhead-preview"
            style={{
              left: `${layout.left}px`,
              top: `${layout.top}px`,
              width: `${layout.width}px`,
              opacity: layout.opacity,
              backgroundImage: layout.dataUrl ? `url(${layout.dataUrl})` : undefined,
              backgroundSize: "contain",
              backgroundRepeat: "no-repeat",
            }}
          >
            {!layout.dataUrl && <>
              <div className="lh-brand">
                <div className="lh-logo" style={{ background: layout.accent }}>N</div>
                <div>
                  <strong>Northstar Labs</strong>
                  <span>People & Culture</span>
                </div>
              </div>
              <div className="lh-lines" style={{ background: layout.accent }} />
            </>}
          </div>
        </div>,
        "canvas-letterhead-block",
      );
    }
    if (block.kind === "meta") {
      return shell(
        <div className="page-meta">
          <span>{template.type.toUpperCase()}</span>
          <span>MY · 2026</span>
        </div>,
      );
    }
    if (block.kind === "title") {
      return shell(
        <h1
          contentEditable={canEditDocument}
          suppressContentEditableWarning
          onBlur={(event) => renameDocument(event.currentTarget.textContent || docName)}
        >
          {docName}
        </h1>,
        "canvas-title-block",
      );
    }
    if (block.kind === "lede") {
      return shell(
        <p className="lede">
          Between <mark>{values.company_name || "Company"}</mark> and{" "}
          <mark>{values.full_name || "Recipient"}</mark>
        </p>,
      );
    }
    if (block.kind === "section" && section) {
      const sectionNumber = orderedSections.findIndex((item) => item.id === section.id) + 1;
      return shell(
        <section className={`doc-section ${activeSection === section.id ? "active" : ""}`}>
          <div className="section-label">
            <span>{String(Math.max(1, sectionNumber)).padStart(2, "0")}</span>
            <h2
              contentEditable={canEditDocument}
              suppressContentEditableWarning
              onBlur={(event) => {
                const title = event.currentTarget.textContent?.trim() || section.title;
                setSections((current) =>
                  current.map((item) => item.id === section.id ? { ...item, title } : item),
                );
              }}
            >
              {section.title}
            </h2>
            <button
              className="section-options"
              title="Add paragraph"
              onClick={(event) => {
                event.stopPropagation();
                addClause(section.id, true);
              }}
            >
              <Plus size={14} />
            </button>
          </div>
          {section.clauses.map((clause, cIndex) =>
            shouldShowClause(clause) ? (
              <div
                className={`clause ${activeClause === clause.id ? "active" : ""} ${canvasDrag?.type === "clause" && canvasDrag.id === clause.id ? "dragging" : ""}`}
                key={clause.id}
                onClick={(event) => {
                  event.stopPropagation();
                  setActiveClause(clause.id);
                  setActiveSection(section.id);
                  setActiveCanvasBlock(block.id);
                }}
                onDragOver={(event) => {
                  if (canvasDrag?.type !== "clause") return;
                  event.preventDefault();
                  event.stopPropagation();
                  event.dataTransfer.dropEffect = "move";
                }}
                onDrop={(event) => {
                  if (canvasDrag?.type !== "clause") return;
                  event.preventDefault();
                  event.stopPropagation();
                  moveClauseTo(canvasDrag.sectionId, canvasDrag.id, section.id, clause.id);
                  setCanvasDrag(null);
                }}
              >
                <button
                  className="clause-drag-handle"
                  title="Drag paragraph"
                  aria-label="Drag paragraph"
                  draggable={canEditDocument}
                  onDragStart={(event) => {
                    event.stopPropagation();
                    event.dataTransfer.effectAllowed = "move";
                    setCanvasDrag({ type: "clause", id: clause.id, sectionId: section.id });
                  }}
                  onDragEnd={() => setCanvasDrag(null)}
                >
                  <GripVertical size={13} />
                </button>
                <div className="clause-number">{sectionNumber}.{cIndex + 1}</div>
                <div className="clause-content">
                  <div className="clause-title">
                    <span
                      contentEditable={canEditDocument}
                      suppressContentEditableWarning
                      onBlur={(event) => {
                        const title = event.currentTarget.textContent?.trim() || clause.title;
                        setSections((current) => current.map((item) => ({
                          ...item,
                          clauses: item.clauses.map((candidate) => candidate.id === clause.id ? { ...candidate, title } : candidate),
                        })));
                      }}
                    >
                      {clause.title}
                    </span>
                    <span className={`tag ${clause.tag.toLowerCase()}`}>{clause.tag}</span>
                  </div>
                  <div
                    className="editable-paragraph"
                    data-clause-id={clause.id}
                    contentEditable={canEditDocument}
                    suppressContentEditableWarning
                    onFocus={(event) => {
                      activeEditorRef.current = event.currentTarget;
                      rememberSelection();
                    }}
                    onMouseUp={rememberSelection}
                    onKeyUp={rememberSelection}
                    onBlur={(event) => updateClauseHtml(clause.id, event.currentTarget.innerHTML)}
                    dangerouslySetInnerHTML={{ __html: clauseHtml(clause) }}
                  />
                  <div className="clause-actions">
                    <button onClick={() => toggleClause(clause.id)}><Archive size={12} /> Exclude</button>
                    <button onClick={() => duplicateClause(section.id, clause)}><Copy size={12} /> Duplicate</button>
                    <button onClick={() => addClause(section.id, true)}><Plus size={12} /> Insert below</button>
                    <button className="delete-action" onClick={() => removeClause(section.id, clause.id)}><Trash2 size={12} /> Delete</button>
                  </div>
                </div>
              </div>
            ) : clause.included ? null : (
              <div className="excluded-clause" key={clause.id}>
                <Archive size={13} />
                <span>{clause.title} excluded</span>
                <button onClick={() => toggleClause(clause.id)}>Restore</button>
                <button className="delete-action" onClick={() => removeClause(section.id, clause.id)}>Delete</button>
              </div>
            ),
          )}
          {!section.clauses.length && (
            <button className="empty-section-add" onClick={() => addClause(section.id, true)}>
              <Plus size={14} /> Add paragraph
            </button>
          )}
        </section>,
        "canvas-section-block",
      );
    }
    if (block.kind === "signatures") {
      return shell(
        <div className="signatures">
          <div>
            <div className="sign-line" />
            <strong>{values.signatory_name || "Company representative"}</strong>
            <span>{values.signatory_title || "Title"}</span>
            <small>Date: __________________</small>
          </div>
          <div>
            <div className="sign-line" />
            <strong>{values.full_name || "Employee / Contractor"}</strong>
            <span>Employee / Contractor</span>
            <small>Date: __________________</small>
          </div>
        </div>,
      );
    }
    if (block.kind === "footer") {
      return shell(
        <div className="page-footer">
          <span>{values.company_name || "Company"} · Confidential</span>
          <span>Page 1 of {estimatePageCount()}</span>
        </div>,
      );
    }
    return null;
  };
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const updateEmailSetting = <Key extends keyof EmailSettings>(key: Key, value: EmailSettings[Key]) => {
    setEmailSettings((current) => ({ ...current, [key]: value }));
    setEmailSettingsErrors((current) => ({ ...current, [key]: undefined }));
    setEmailSettingsStatus("idle");
  };
  const validateEmailSetup = (includeCredential = false) => {
    const errors: EmailSettingsErrors = {};
    if (includeCredential) {
      if (!emailSettings.gmailAddress.trim()) errors.gmailAddress = "Enter the Gmail address used to send documents.";
      else if (!emailPattern.test(emailSettings.gmailAddress.trim())) errors.gmailAddress = "Enter a valid email address.";
      if (emailSettings.gmailAppPassword.replace(/\s+/g, "").length !== 16) errors.gmailAppPassword = "Paste the 16-character Gmail App Password.";
    }
    if (!emailSettings.senderName.trim()) errors.senderName = "Enter the sender name recipients should see.";
    if (emailSettings.replyToEmail.trim() && !emailPattern.test(emailSettings.replyToEmail.trim())) {
      errors.replyToEmail = "Enter a valid reply-to email address.";
    }
    if (emailSettings.receiveCopies && !emailSettings.documentInboxEmail.trim()) errors.documentInboxEmail = "Choose where document copies should be delivered.";
    else if (emailSettings.documentInboxEmail.trim() && !emailPattern.test(emailSettings.documentInboxEmail.trim())) errors.documentInboxEmail = "Enter a valid document inbox address.";
    if (emailSettings.notificationsEnabled && !emailSettings.notificationEmail.trim()) errors.notificationEmail = "Choose where system notifications should be delivered.";
    else if (emailSettings.notificationEmail.trim() && !emailPattern.test(emailSettings.notificationEmail.trim())) errors.notificationEmail = "Enter a valid notification email address.";
    return errors;
  };
  const applyEmailConnection = (connection: EmailConnectionSummary) => {
    setEmailConnection(connection);
    setEmailSettings({
      ...clone(defaultEmailSettings),
      authMethod: connection.authMethod,
      gmailAddress: connection.senderEmail,
      senderName: connection.senderName,
      replyToEmail: connection.replyToEmail,
      documentInboxEmail: connection.documentInboxEmail,
      notificationEmail: connection.notificationEmail,
      sendDocuments: connection.sendDocuments,
      receiveCopies: connection.receiveCopies,
      notificationsEnabled: connection.notificationsEnabled,
    });
  };
  const readEmailApiResponse = async (response: Response) => {
    const result = await response.json().catch(() => ({ ok: false, error: "The email service returned an invalid response." })) as {
      ok?: boolean;
      error?: string;
      connection?: EmailConnectionSummary;
      deliveredTo?: string;
    };
    if (!response.ok || !result.ok) throw new Error(result.error || "The email service could not complete this action.");
    return result;
  };
  const connectGoogleEmail = () => {
    if (!emailPlatform.googleOAuthReady) {
      setEmailSettingsError("Google OAuth must be configured by a platform administrator before users can connect accounts.");
      return;
    }
    window.location.assign("/api/email-connections/oauth/start?returnTo=%2F%3Femail%3Dconnected%23settings");
  };
  const saveAppPasswordConnection = async () => {
    const errors = validateEmailSetup(true);
    setEmailSettingsErrors(errors);
    if (Object.keys(errors).length) return;
    setEmailSettingsStatus("saving");
    setEmailSettingsError("");
    try {
      const response = await fetch("/api/email-connections/app-password", {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify(emailSettings),
      });
      const result = await readEmailApiResponse(response);
      if (result.connection) applyEmailConnection(result.connection);
      setEmailSettings((current) => ({ ...current, gmailAppPassword: "" }));
      setShowGmailAppPassword(false);
      setEmailAdvancedOpen(false);
      setEmailSettingsStatus("saved");
      setModuleNotice("Gmail connected to this workspace");
    } catch (error) {
      setEmailSettingsStatus("error");
      setEmailSettingsError(error instanceof Error ? error.message : "Could not save email settings");
    }
  };
  const saveEmailPreferences = async () => {
    const errors = validateEmailSetup(false);
    setEmailSettingsErrors(errors);
    if (Object.keys(errors).length) return;
    setEmailSettingsStatus("saving");
    setEmailSettingsError("");
    try {
      const response = await fetch("/api/email-connections/update", {
        method: "PATCH",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify(emailSettings),
      });
      const result = await readEmailApiResponse(response);
      if (result.connection) applyEmailConnection(result.connection);
      setEmailSettingsStatus("saved");
      setModuleNotice("Email delivery settings saved");
    } catch (error) {
      setEmailSettingsStatus("error");
      setEmailSettingsError(error instanceof Error ? error.message : "Could not save email settings");
    }
  };
  const sendEmailConnectionTest = async () => {
    setEmailSettingsStatus("saving");
    setEmailSettingsError("");
    try {
      const response = await fetch("/api/email-connections/test", {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ targetEmail: emailSettings.notificationEmail || emailSettings.replyToEmail || emailSettings.gmailAddress }),
      });
      const result = await readEmailApiResponse(response);
      setEmailSettingsStatus("saved");
      setModuleNotice(`Test email sent to ${result.deliveredTo || "the configured inbox"}`);
    } catch (error) {
      setEmailSettingsStatus("error");
      setEmailSettingsError(error instanceof Error ? error.message : "The test email could not be delivered.");
    }
  };
  const disconnectEmailAccount = async () => {
    if (!window.confirm("Disconnect this Gmail account from the current workspace? Existing documents will not be changed.")) return;
    setEmailSettingsStatus("saving");
    setEmailSettingsError("");
    try {
      const response = await fetch("/api/email-connections/disconnect", { method: "DELETE", headers: { Accept: "application/json" } });
      await readEmailApiResponse(response);
      setEmailConnection(null);
      setEmailSettings(clone(defaultEmailSettings));
      setEmailSettingsStatus("idle");
      setModuleNotice("Email account disconnected");
    } catch (error) {
      setEmailSettingsStatus("error");
      setEmailSettingsError(error instanceof Error ? error.message : "The account could not be disconnected.");
    }
  };
  const renderEmailSettingsPage = () => {
    const connected = Boolean(emailConnection && emailConnection.status !== "disconnected");
    const connectionHealthy = emailConnection?.status === "connected";
    const settingsLocked = currentRole === "Reviewer" || emailSettingsStatus === "saving";
    return (
      <section className="module-page email-settings-page">
        <div className="module-header email-settings-header">
          <div>
            <span className="eyebrow">WORKSPACE SETTINGS</span>
            <h1>Email connections</h1>
            <p>Connect a sender once, then use it for document delivery and workflow notifications in this workspace.</p>
          </div>
          <div className={`email-environment-status ${connected && connectionHealthy ? "ready" : emailPlatform.ready ? "pending" : "remote"}`}>
            {emailSettingsStatus === "loading" ? <RefreshCw size={15} /> : connected && connectionHealthy ? <CheckCircle2 size={15} /> : emailPlatform.ready ? <Mail size={15} /> : <Cloud size={15} />}
            <span><strong>{emailSettingsStatus === "loading" ? "Checking connection" : connected && connectionHealthy ? "Gmail connected" : emailPlatform.ready ? "Not connected" : "Cloud setup required"}</strong><small>{connected ? `${emailConnection?.senderEmail} · ${emailConnection?.authMethod === "oauth" ? "Google OAuth" : "App Password"}` : "Northstar Labs Malaysia workspace"}</small></span>
          </div>
        </div>
        <div className="email-saas-context" aria-label="Email connection ownership">
          <div><Cloud size={15} /><span><small>Workspace</small><strong>Northstar Labs Malaysia</strong></span></div>
          <div><UserRound size={15} /><span><small>Connected by</small><strong>{authSession.displayName}</strong></span></div>
          <div><ShieldCheck size={15} /><span><small>Credential storage</small><strong>{emailPlatform.ready ? "Encrypted cloud record" : "Not configured"}</strong></span></div>
        </div>

        {!emailPlatform.ready && (
          <div className="email-platform-blocker" role="alert">
            <span className="email-platform-icon"><Cloud size={19} /></span>
            <div><strong>Cloud email connections need platform setup</strong><p>{currentRole === "Admin" ? "Add the platform variables below in Vercel, configure trusted server authentication, then redeploy. User credentials will be stored per workspace in the database, not in .env files." : "A workspace administrator must finish cloud credential storage and trusted authentication before you can connect Gmail."}</p>
              {currentRole === "Admin" && emailPlatform.missing.length > 0 && <div className="email-config-chips">{emailPlatform.missing.map((item) => <code key={item}>{item}</code>)}</div>}
            </div>
          </div>
        )}

        {emailSettingsError && <div className="email-settings-error" role="alert"><AlertTriangle size={16} /><span>{emailSettingsError}</span></div>}
        {emailSettingsStatus === "saved" && <div className="email-settings-success" role="status"><CheckCircle2 size={16} /><span>The workspace email connection has been updated.</span></div>}

        <div className="email-saas-layout">
          <section className="email-connection-panel">
            <div className="email-section-heading"><div><span className="eyebrow">SENDING ACCOUNT</span><h2>{connected ? "Connected Gmail account" : "Connect Gmail"}</h2><p>{connected ? "This account is available to permitted members in the current workspace." : "Google OAuth is recommended because users can revoke access without changing passwords."}</p></div>{connected && <span className={`status-pill ${connectionHealthy ? "published" : "inactive"}`}>{emailConnection?.status}</span>}</div>

            {connected ? (
              <>
                <div className="email-connected-account"><span className="email-provider-mark"><Mail size={19} /></span><div><strong>{emailConnection?.senderEmail}</strong><small>{emailConnection?.authMethod === "oauth" ? "Google OAuth" : "Gmail App Password"} · scoped to this workspace</small></div><button className="outline-btn" type="button" disabled={!emailPlatform.googleOAuthReady || settingsLocked} onClick={connectGoogleEmail}><RefreshCw size={14} /> Reconnect</button></div>
                {emailConnection?.lastError && <div className="email-info-box warning"><AlertTriangle size={16} /><span><strong>Last delivery check failed</strong><small>{emailConnection.lastError}</small></span></div>}
                <div className="email-form-grid">
                  <label className="email-field"><span>Sender name</span><input value={emailSettings.senderName} disabled={settingsLocked} onChange={(event) => updateEmailSetting("senderName", event.target.value)} placeholder="Northstar Labs HR" />{emailSettingsErrors.senderName && <small role="alert">{emailSettingsErrors.senderName}</small>}</label>
                  <label className="email-field"><span>Reply-to email <small>Optional</small></span><input type="email" value={emailSettings.replyToEmail} disabled={settingsLocked} onChange={(event) => updateEmailSetting("replyToEmail", event.target.value)} placeholder={emailSettings.gmailAddress || "hr@company.com"} />{emailSettingsErrors.replyToEmail && <small role="alert">{emailSettingsErrors.replyToEmail}</small>}</label>
                </div>
                <div className="email-sender-preview"><span className="email-sender-avatar">{(emailSettings.senderName || "H").slice(0, 1).toUpperCase()}</span><span><small>Recipients will see</small><strong>{emailSettings.senderName || "Your sender name"}</strong><small>{emailSettings.gmailAddress}</small></span></div>
              </>
            ) : (
              <>
                <div className="email-provider-row"><span className="email-provider-mark"><Mail size={18} /></span><span><strong>Google Workspace or Gmail</strong><small>Secure OAuth connection with Gmail send permission</small></span><button className="primary-btn" type="button" disabled={!emailPlatform.googleOAuthReady || settingsLocked} onClick={connectGoogleEmail}><Cloud size={14} /> Connect Gmail</button></div>
                <div className="email-info-box"><Lock size={16} /><span><strong>No Gmail password is shared with ZhiReady</strong><small>Google issues a revocable token. The refresh token is encrypted on the server and never returned to the browser.</small></span></div>
                <button className="email-advanced-toggle" type="button" aria-expanded={emailAdvancedOpen} onClick={() => setEmailAdvancedOpen((current) => !current)}><KeyRound size={15} /><span><strong>Use a Gmail App Password instead</strong><small>Advanced compatibility option for accounts where OAuth is unavailable</small></span><ChevronDown size={15} /></button>
                {emailAdvancedOpen && <div className="email-advanced-panel">
                  <div className="gmail-tutorial compact"><div><span>1</span><p>Turn on 2-Step Verification in the Google account.</p></div><div><span>2</span><p>Create an App Password named <strong>ZhiReady</strong>.</p></div><div><span>3</span><p>Paste the 16-character password below, then connect.</p></div></div>
                  <a className="email-help-link" href="https://support.google.com/accounts/answer/185833" target="_blank" rel="noreferrer">Google App Password instructions <ChevronRight size={14} /></a>
                  <div className="email-form-grid">
                    <label className="email-field"><span>Gmail address</span><input type="email" value={emailSettings.gmailAddress} disabled={settingsLocked} onChange={(event) => updateEmailSetting("gmailAddress", event.target.value)} placeholder="you@gmail.com" autoComplete="email" />{emailSettingsErrors.gmailAddress && <small role="alert">{emailSettingsErrors.gmailAddress}</small>}</label>
                    <label className="email-field"><span>Sender name</span><input value={emailSettings.senderName} disabled={settingsLocked} onChange={(event) => updateEmailSetting("senderName", event.target.value)} placeholder="Northstar Labs HR" />{emailSettingsErrors.senderName && <small role="alert">{emailSettingsErrors.senderName}</small>}</label>
                  </div>
                  <label className="email-field"><span>Gmail App Password</span><span className="email-secret-field"><input type={showGmailAppPassword ? "text" : "password"} value={emailSettings.gmailAppPassword} disabled={settingsLocked} onChange={(event) => updateEmailSetting("gmailAppPassword", event.target.value)} placeholder="16-character App Password" autoComplete="new-password" /><button type="button" onClick={() => setShowGmailAppPassword((current) => !current)} aria-label={showGmailAppPassword ? "Hide App Password" : "Show App Password"}>{showGmailAppPassword ? <EyeOff size={15} /> : <Eye size={15} />}</button></span>{emailSettingsErrors.gmailAppPassword && <small role="alert">{emailSettingsErrors.gmailAppPassword}</small>}</label>
                  <div className="email-panel-actions"><span><Lock size={13} /> Cleared from the form after encrypted storage</span><button className="primary-btn" type="button" disabled={!emailPlatform.ready || settingsLocked} onClick={saveAppPasswordConnection}><KeyRound size={14} /> {emailSettingsStatus === "saving" ? "Connecting..." : "Connect with App Password"}</button></div>
                </div>}
              </>
            )}
          </section>

          <aside className="email-security-panel">
            <ShieldCheck size={19} />
            <h3>SaaS credential boundary</h3>
            <p>Each connection belongs to one user and one company workspace. Switching companies never reuses another workspace's sender.</p>
            <dl><div><dt>Storage</dt><dd>AES-256-GCM encrypted</dd></div><div><dt>Browser</dt><dd>No password persistence</dd></div><div><dt>Access</dt><dd>Server session required</dd></div><div><dt>Revocation</dt><dd>Reconnect or disconnect anytime</dd></div></dl>
          </aside>
        </div>

        {connected && <section className="email-delivery-panel">
          <div className="email-section-heading"><div><span className="eyebrow">DELIVERY & NOTIFICATIONS</span><h2>Workspace delivery rules</h2><p>Control which emails this connection may send and where operational copies arrive.</p></div></div>
          <div className="email-preferences-grid">
            <div className="email-preference-stack">
              <label className="email-preference-row"><input type="checkbox" disabled={settingsLocked} checked={emailSettings.sendDocuments} onChange={(event) => updateEmailSetting("sendDocuments", event.target.checked)} /><span className="email-preference-icon"><Send size={16} /></span><span><strong>Send generated documents</strong><small>Deliver approved Word and PDF files from this sender.</small></span></label>
              <label className="email-preference-row"><input type="checkbox" disabled={settingsLocked} checked={emailSettings.receiveCopies} onChange={(event) => updateEmailSetting("receiveCopies", event.target.checked)} /><span className="email-preference-icon"><Inbox size={16} /></span><span><strong>Receive document copies</strong><small>Keep a copy of completed document delivery emails.</small></span></label>
              <label className="email-preference-row"><input type="checkbox" disabled={settingsLocked} checked={emailSettings.notificationsEnabled} onChange={(event) => updateEmailSetting("notificationsEnabled", event.target.checked)} /><span className="email-preference-icon"><Bell size={16} /></span><span><strong>Workflow notifications</strong><small>Receive approval, export failure and status notifications.</small></span></label>
            </div>
            <div className="email-routing-fields">
              <label className="email-field"><span>Document copy inbox</span><input type="email" disabled={settingsLocked || !emailSettings.receiveCopies} value={emailSettings.documentInboxEmail} onChange={(event) => updateEmailSetting("documentInboxEmail", event.target.value)} placeholder={emailSettings.gmailAddress} />{emailSettingsErrors.documentInboxEmail && <small role="alert">{emailSettingsErrors.documentInboxEmail}</small>}</label>
              <label className="email-field"><span>Notification email</span><input type="email" disabled={settingsLocked || !emailSettings.notificationsEnabled} value={emailSettings.notificationEmail} onChange={(event) => updateEmailSetting("notificationEmail", event.target.value)} placeholder={emailSettings.gmailAddress} />{emailSettingsErrors.notificationEmail && <small role="alert">{emailSettingsErrors.notificationEmail}</small>}</label>
              <button className="email-use-gmail" type="button" disabled={settingsLocked} onClick={() => { if (emailSettings.receiveCopies) updateEmailSetting("documentInboxEmail", emailSettings.gmailAddress); if (emailSettings.notificationsEnabled) updateEmailSetting("notificationEmail", emailSettings.gmailAddress); }}>Use connected Gmail for enabled inboxes</button>
            </div>
          </div>
          <div className="email-delivery-actions"><button className="delete-action" type="button" disabled={settingsLocked} onClick={disconnectEmailAccount}>Disconnect account</button><div><button className="outline-btn" type="button" disabled={settingsLocked} onClick={sendEmailConnectionTest}><Send size={14} /> Send test email</button><button className="primary-btn" type="button" disabled={settingsLocked} onClick={saveEmailPreferences}><Save size={14} /> {emailSettingsStatus === "saving" ? "Saving..." : "Save settings"}</button></div></div>
        </section>}
      </section>
    );
  };
  const renderModulePage = () => {
    if (activeModule === "settings") return renderEmailSettingsPage();
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
  const workspaceTools = [
    { id: "outline" as const, label: "Document outline", description: "Template and section order", Icon: PanelLeftOpen },
    { id: "person" as const, label: "Source data", description: "Onboarding record and sync", Icon: UserRound },
    { id: "placeholders" as const, label: "Fields", description: "Values and Placeholder groups", Icon: Sparkles },
    { id: "clauses" as const, label: "Clauses", description: "Included content and library", Icon: Archive },
    { id: "letterhead" as const, label: "Page layout", description: "Letterhead, margins and watermark", Icon: PanelRight },
    { id: "review" as const, label: "Review checks", description: "Validate the current document", Icon: FileCheck2 },
  ];
  const currentWorkspaceTool = workspaceTools.find((tool) => tool.id === activeRightTab) || workspaceTools[0];
  const CurrentWorkspaceToolIcon = currentWorkspaceTool.Icon;
  const chooseWorkspaceTool = (toolId: typeof workspaceTools[number]["id"]) => {
    if (toolId === "review") goToWorkflowStage("review");
    else openWorkspaceTool(toolId);
    setWorkspaceToolMenuOpen(false);
  };
  const renderWorkspaceToolSwitcher = () => (
    <div
      className="workspace-tool-switcher"
      onBlur={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setWorkspaceToolMenuOpen(false);
      }}
      onKeyDown={(event) => {
        if (event.key === "Escape") setWorkspaceToolMenuOpen(false);
      }}
    >
      <div className="workspace-tool-current-row">
        <button
          className="workspace-tool-current"
          type="button"
          aria-haspopup="menu"
          aria-expanded={workspaceToolMenuOpen}
          onClick={() => setWorkspaceToolMenuOpen((open) => !open)}
        >
          <span className="workspace-tool-current-icon"><CurrentWorkspaceToolIcon size={15} /></span>
          <span className="workspace-tool-current-copy">
            <small>Workspace tool</small>
            <strong>{currentWorkspaceTool.label}</strong>
          </span>
          {activeRightTab === "review" && (
            <span className={`workspace-tool-status ${blockingIssues.length ? "warning" : "ready"}`}>
              {blockingIssues.length ? blockingIssues.length : <Check size={11} />}
            </span>
          )}
          <ChevronDown size={14} />
        </button>
        <button
          className="workspace-tool-close"
          type="button"
          title="Close workspace tool"
          aria-label="Close workspace tool"
          onClick={() => {
            setWorkspacePanelOpen(false);
            setWorkspaceToolMenuOpen(false);
          }}
        >
          <X size={15} />
        </button>
      </div>
      {workspaceToolMenuOpen && (
        <div className="workspace-tool-menu" role="menu">
          <div className="workspace-tool-menu-heading">Switch tool</div>
          {workspaceTools.map(({ id, label, description, Icon }) => (
            <button
              key={id}
              className={activeRightTab === id ? "active" : ""}
              type="button"
              role="menuitem"
              onClick={() => chooseWorkspaceTool(id)}
            >
              <span className="workspace-tool-menu-icon"><Icon size={14} /></span>
              <span><strong>{label}</strong><small>{description}</small></span>
              {id === "review" && blockingIssues.length > 0
                ? <span className="workspace-tool-menu-count">{blockingIssues.length}</span>
                : activeRightTab === id
                  ? <Check size={13} />
                  : <ChevronRight size={12} />}
            </button>
          ))}
        </div>
      )}
    </div>
  );

  return (
    <div className={`app-shell ${sidebarCollapsed ? "sidebar-is-collapsed" : ""}`}>
      <input ref={letterheadInputRef} type="file" accept=".png,.jpg,.jpeg,.pdf,.docx" hidden onChange={(e) => uploadLetterhead(e.target.files?.[0])} />
      <input ref={placeholderImportInputRef} type="file" accept=".csv,.xlsx,.xls" hidden onChange={(e) => previewPlaceholderImport(e.target.files?.[0])} />
      <aside className={`global-sidebar ${sidebarCollapsed ? "collapsed" : ""}`} aria-label="Primary navigation">
        <nav className="global-nav">
          {navItems.map(([id, label, Icon]) => {
            const linkedTab = id === "clauses" ? "clauses" : id === "placeholders" ? "placeholders" : id === "layouts" ? "letterhead" : null;
            const linkedActive = activeModule === "workspace" && linkedTab === activeRightTab;
            return <button key={id} className={`${activeModule === id ? "active" : ""} ${linkedActive ? "linked-active" : ""}`} onClick={() => handlePrimaryNavigation(id as AppModule)} title={sidebarCollapsed ? label : undefined} aria-label={label} aria-current={activeModule === id ? "page" : undefined}><Icon size={17} /><span>{label}</span></button>;
          })}
        </nav>
        <div className="global-sidebar-spacer" />
        {!sidebarCollapsed && <div className="global-help"><span className="eyebrow">WORKSPACE</span><strong>Keep your document flow moving</strong><small>Build, review and export from one place.</small></div>}
        <div className="global-nav global-nav-secondary">
          <div
            className="global-account-item"
            onBlur={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setAccountMenuOpen(false);
            }}
            onKeyDown={(event) => {
              if (event.key === "Escape") setAccountMenuOpen(false);
            }}
          >
            <button
              type="button"
              title="Account"
              aria-label="Account"
              aria-haspopup="menu"
              aria-expanded={accountMenuOpen}
              onClick={() => setAccountMenuOpen((current) => !current)}
            >
              <UserRound size={16} /><span>Account</span>
            </button>
            {accountMenuOpen && (
              <div className="global-account-menu" role="menu">
                <div className="global-account-profile">
                  <span>{authSession.displayName.slice(0, 1).toUpperCase()}</span>
                  <div><strong>{authSession.displayName}</strong><small>{authSession.email || "Blank demo access"}</small></div>
                </div>
                <button type="button" role="menuitem" onClick={signOutOfWorkspace}>
                  <LogOut size={15} /> Sign out
                </button>
              </div>
            )}
          </div>
          <button title="Members and roles" aria-label="Members and roles"><UsersRound size={16} /><span>Members & roles</span></button>
          <button title="Subscription" aria-label="Subscription"><ShieldCheck size={16} /><span>Subscription</span></button>
          <button className={activeModule === "settings" ? "active" : ""} title="Settings" aria-label="Settings" aria-current={activeModule === "settings" ? "page" : undefined} onClick={() => switchModule("settings")}><Settings2 size={16} /><span>Settings</span></button>
          <button title={theme === "light" ? "Switch to dark theme" : "Switch to light theme"} aria-label="Toggle theme" onClick={() => setTheme((current) => current === "light" ? "dark" : "light")}>{theme === "light" ? <Moon size={16} /> : <Sun size={16} />}<span>Theme</span></button>
        </div>
        <button className="sidebar-collapse" onClick={() => setSidebarCollapsed((current) => !current)} aria-label={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}>{sidebarCollapsed ? <PanelLeftOpen size={16} /> : <><PanelLeftClose size={16} /><span>Collapse</span></>}</button>
      </aside>
      <header className="topbar">
        <div className="brand">
          <img
            className="brand-logo"
            src={theme === "light" ? "/brand/zhiready-light.png" : "/brand/zhiready-dark.png"}
            alt="ZhiReady"
          />
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
                goToWorkflowStage("review");
                return;
              }
              if (docStatus === "Approved") return;
              if (docStatus === "In review" && currentRole !== "Reviewer") {
                goToWorkflowStage("review");
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
      {activeModule === "workspace" ? <div className={`workspace workspace-panel-${activeRightTab} ${workspacePanelOpen ? "workspace-panel-open" : "workspace-panel-closed"}`}>
        <aside className="left-rail">
          {renderWorkspaceToolSwitcher()}
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
              {orderedSections.map((section, index) => (
                <div
                  className={`section-row ${activeSection === section.id ? "selected" : ""}`}
                  key={section.id}
                  onClick={() => focusSection(section.id)}
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
              <div className="canvas-add-wrap">
                <button
                  className="toolbar-btn canvas-add-btn"
                  title="Add document element"
                  aria-haspopup="menu"
                  aria-expanded={addBlockMenuOpen}
                  onClick={() => setAddBlockMenuOpen((open) => !open)}
                >
                  <Plus size={16} />
                  <span>Add element</span>
                  <ChevronDown size={12} />
                </button>
                {addBlockMenuOpen && (
                  <div className="canvas-add-menu" role="menu">
                    <button role="menuitem" onClick={() => addFixedCanvasBlock("title")}><Type size={14} /> Add title</button>
                    <button role="menuitem" onClick={() => addFixedCanvasBlock("meta")}><FileCheck2 size={14} /> Add document meta</button>
                    <button role="menuitem" onClick={() => addFixedCanvasBlock("lede")}><AlignLeft size={14} /> Add summary</button>
                    <button role="menuitem" onClick={() => addSection()}><Plus size={14} /> Add section</button>
                    <button role="menuitem" onClick={addParagraphElement}><AlignLeft size={14} /> Add paragraph</button>
                    <button role="menuitem" onClick={() => addFixedCanvasBlock("signatures")}><PenLine size={14} /> Add signatures</button>
                    <button role="menuitem" onClick={() => addFixedCanvasBlock("letterhead")}><ImagePlus size={14} /> Add letterhead</button>
                    <button role="menuitem" onClick={() => addFixedCanvasBlock("footer")}><PanelBottom size={14} /> Add footer</button>
                  </div>
                )}
              </div>
              <button
                className="toolbar-btn canvas-clear-btn"
                title="Clear canvas"
                aria-label="Clear canvas"
                onClick={clearCanvas}
              >
                <Trash2 size={16} />
              </button>
              {deletedCanvasSnapshot && (
                <button
                  className="toolbar-btn"
                  title={`Restore ${deletedCanvasSnapshot.label}`}
                  aria-label="Restore last removed element"
                  onClick={restoreDeletedCanvasItem}
                >
                  <RotateCcw size={16} />
                </button>
              )}
              <span className="toolbar-divider" />
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
                  className={`verification-watermark verification-watermark-${watermark.placement} verification-watermark-${watermark.alignment}`}
                  aria-label="Document verification watermark"
                >
                  {watermarkDisplayText(watermark)}
                </div>
                <div className="doc-body canvas-document-body">
                  {canvasBlocks.length ? (
                    canvasBlocks.map((block, index) => renderCanvasBlock(block, index))
                  ) : (
                    <div className="canvas-empty-state">
                      <FileText size={24} />
                      <h2>Blank canvas</h2>
                      <p>Add only the elements this document needs.</p>
                      <div className="canvas-empty-actions">
                        <button className="outline-btn" onClick={() => addSection()}><Plus size={14} /> Add section</button>
                        <button className="outline-btn" onClick={() => addFixedCanvasBlock("title")}><Type size={14} /> Add title</button>
                        <button className="outline-btn" onClick={addParagraphElement}><AlignLeft size={14} /> Add paragraph</button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </main>
        <aside className="right-panel">
          {renderWorkspaceToolSwitcher()}
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
              <button
                className="text-btn workspace-library-toggle"
                onClick={() => {
                  setWorkspaceLibrarySearch("");
                  setWorkspaceLibraryOpen((current) => current === "placeholders" ? null : "placeholders");
                }}
              >
                <Sparkles size={13} />
                {workspaceLibraryOpen === "placeholders" ? "Hide data groups" : "Browse data groups"}
                <ChevronRight size={13} />
              </button>
              {workspaceLibraryOpen === "placeholders" && (
                <div className="workspace-library-browser">
                  <div className="workspace-library-heading">
                    <div><span>CONNECTED DATA</span><strong>Placeholder groups</strong></div>
                    <button className="icon-btn" title="Create group" onClick={() => openPlaceholderEditor()}><Plus size={13} /></button>
                  </div>
                  <label className="workspace-library-search">
                    <Search size={13} />
                    <input value={workspaceLibrarySearch} onChange={(event) => setWorkspaceLibrarySearch(event.target.value)} placeholder="Search groups or fields" />
                  </label>
                  <div className="workspace-library-list">
                    {workspacePlaceholderGroups.map((group) => (
                      <div className="workspace-library-card" key={group.id}>
                        <div className="workspace-library-card-head">
                          <span><strong>{group.name}</strong><small>{group.sourceType} · {group.fields.length} fields</small></span>
                          <button className="icon-btn" title={`Edit ${group.name}`} onClick={() => openPlaceholderEditor(group)}><Settings2 size={13} /></button>
                        </div>
                        <div className="workspace-field-chips">
                          {group.fields.slice(0, 5).map((field) => (
                            <span className={field.mappingStatus === "valid" ? "" : "invalid"} key={field.id}>{field.label}</span>
                          ))}
                          {group.fields.length > 5 && <span>+{group.fields.length - 5}</span>}
                        </div>
                      </div>
                    ))}
                  </div>
                  <button className="workspace-full-management" onClick={() => switchModule("placeholders")}>Open full management <ChevronRight size={12} /></button>
                </div>
              )}
            </div>
          )}
          {activeRightTab === "clauses" && (
            <div className="panel-content">
              <div className="panel-title-row clause-panel-title-row">
                <div>
                  <span className="eyebrow">CONTENT MAP</span>
                  <h3>Clauses & sections</h3>
                  <p className="panel-subtitle">Select a section on the left, then include or exclude its paragraphs here.</p>
                </div>
                <div className="row-actions"><button className="mini-icon" title="Add custom clause" onClick={activeSection ? () => addClause(activeSection) : addParagraphElement}><Plus size={15} /></button><button className="mini-icon" title="Browse Clauses Library here" onClick={() => { setWorkspaceLibrarySearch(""); setWorkspaceLibraryOpen((current) => current === "clauses" ? null : "clauses"); }}><FolderOpen size={14} /></button></div>
              </div>
              <div className="clause-panel-toolbar">
                <div className="clause-focus-copy">
                  <span>Editing</span>
                  <strong>{orderedSections.find((section) => section.id === activeSection)?.title || "No section selected"}</strong>
                </div>
                <div className="segmented-control clause-scope-toggle" aria-label="Clause scope">
                  <button className={clauseScope === "current" ? "active" : ""} aria-pressed={clauseScope === "current"} disabled={!activeSection} onClick={() => setClauseScope("current")}>Current</button>
                  <button className={clauseScope === "all" ? "active" : ""} aria-pressed={clauseScope === "all"} onClick={() => setClauseScope("all")}>All</button>
                </div>
              </div>
              <div className="clause-panel-actions">
                <button className="add-field" onClick={() => { setWorkspaceLibrarySearch(""); setWorkspaceLibraryOpen((current) => current === "clauses" ? null : "clauses"); }}><Archive size={14} /> {workspaceLibraryOpen === "clauses" ? "Hide Clauses Library" : "Add from Clauses Library"}</button>
                <button className="text-btn" onClick={activeSection ? () => addClause(activeSection, true) : addParagraphElement}><Plus size={13} /> Write custom paragraph</button>
              </div>
              {workspaceLibraryOpen === "clauses" && (
                <div className="workspace-library-browser clause-browser">
                  <div className="workspace-library-heading">
                    <div><span>REUSABLE CONTENT</span><strong>Clauses Library</strong></div>
                    <button className="icon-btn" title="Create reusable clause" onClick={() => openClauseEditor()}><Plus size={13} /></button>
                  </div>
                  <label className="workspace-library-search">
                    <Search size={13} />
                    <input value={workspaceLibrarySearch} onChange={(event) => setWorkspaceLibrarySearch(event.target.value)} placeholder="Search clauses" />
                  </label>
                  <div className="workspace-library-list">
                    {workspaceClauseRecords.length ? workspaceClauseRecords.map((record) => {
                      const selected = workspaceClausePreviewId === record.id;
                      const sourceClauses = record.structure === "nested" ? record.subsections.flatMap((subsection) => subsection.contents) : record.contents;
                      return (
                        <div className={`workspace-library-card ${selected ? "selected" : ""}`} key={record.id}>
                          <button className="workspace-library-card-button" onClick={() => setWorkspaceClausePreviewId(selected ? null : record.id)}>
                            <span><strong>{record.title}</strong><small>{record.category} · v{record.version} · {sourceClauses.length} blocks</small></span>
                            <ChevronRight size={13} />
                          </button>
                          {selected && (
                            <div className="workspace-clause-preview">
                              <p>{sourceClauses[0]?.text || "No content blocks"}</p>
                              <button className="primary-btn" disabled={!sourceClauses.length} onClick={() => insertClauseRecord(record)}><Plus size={13} /> Insert copy</button>
                            </div>
                          )}
                        </div>
                      );
                    }) : <div className="workspace-library-empty">No matching clauses.</div>}
                  </div>
                  <button className="workspace-full-management" onClick={() => switchModule("clauses")}>Open full management <ChevronRight size={12} /></button>
                </div>
              )}
              <div className="helper-callout amber compact-callout">
                <CircleHelp size={15} />
                <span>Excluded paragraphs stay in the draft and can be restored.</span>
              </div>
              {clausePanelSections.length ? clausePanelSections.map((section) => {
                const sectionNumber = orderedSections.findIndex((item) => item.id === section.id) + 1;
                return (
                <div className={`clause-group ${activeSection === section.id ? "active" : ""}`} key={section.id}>
                  <button className="group-title group-title-button" onClick={() => focusSection(section.id)}>
                    <span><b>{String(sectionNumber).padStart(2, "0")}</b>{section.title}</span>
                    <small>
                      {section.clauses.filter((c) => c.included).length}/
                      {section.clauses.length}
                      <ChevronRight size={13} />
                    </small>
                  </button>
                  {section.clauses.map((clause) => (
                    <label className={`clause-toggle ${activeClause === clause.id ? "selected" : ""}`} key={clause.id} onClick={() => {
                      setActiveSection(section.id);
                      setActiveClause(clause.id);
                      const block = canvasBlocks.find((item) => item.kind === "section" && item.sectionId === section.id);
                      if (block) setActiveCanvasBlock(block.id);
                    }}>
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
                );
              }) : <div className="clauses-empty-state"><Archive size={18} /><strong>No sections yet</strong><span>Add a section to start building the document content.</span><button className="add-field" onClick={addSection}><Plus size={14} /> Add section</button></div>}
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
              <div className="watermark-settings">
                <div className="watermark-settings-heading">
                  <span className="watermark-settings-icon"><ShieldCheck size={15} /></span>
                  <div>
                    <strong>Verification watermark</strong>
                    <small>Required on every page and included in Word and PDF.</small>
                  </div>
                </div>
                <div className="drawer-grid watermark-grid">
                  <label className="drawer-field">
                    <span>Placement</span>
                    <select
                      value={watermark.placement}
                      disabled={!canEditDocument}
                      onChange={(event) => updateWatermark({ placement: event.target.value as VerificationWatermark["placement"] })}
                    >
                      <option value="header">Header</option>
                      <option value="footer">Footer</option>
                    </select>
                  </label>
                  <label className="drawer-field">
                    <span>Alignment</span>
                    <div className="watermark-alignment" role="group" aria-label="Watermark alignment">
                      {([
                        ["left", AlignLeft],
                        ["center", AlignCenter],
                        ["right", AlignRight],
                      ] as const).map(([alignment, Icon]) => (
                        <button
                          key={alignment}
                          type="button"
                          className={watermark.alignment === alignment ? "active" : ""}
                          disabled={!canEditDocument}
                          title={`${alignment[0].toUpperCase()}${alignment.slice(1)} align`}
                          aria-label={`${alignment} align watermark`}
                          onClick={() => updateWatermark({ alignment })}
                        >
                          <Icon size={14} />
                        </button>
                      ))}
                    </div>
                  </label>
                </div>
                <label className="drawer-field">
                  <span>Anti-counterfeit text</span>
                  <input
                    value={watermark.text}
                    disabled={!canEditDocument}
                    required
                    placeholder="VERIFIED DOCUMENT"
                    onChange={(event) => updateWatermark({ text: event.target.value })}
                  />
                </label>
                <label className="drawer-field">
                  <span>Timestamp format</span>
                  <input
                    value={watermark.timestampFormat}
                    disabled={!canEditDocument}
                    required
                    list="watermark-timestamp-formats"
                    placeholder="YYYY-MM-DD HH:mm:ss Z"
                    onChange={(event) => updateWatermark({ timestampFormat: event.target.value })}
                  />
                  <small>Tokens: YYYY, YY, MMM, MM, DD, HH, hh, mm, ss, A, Z</small>
                </label>
                <datalist id="watermark-timestamp-formats">
                  <option value="YYYY-MM-DD HH:mm:ss Z" />
                  <option value="DD/MM/YYYY HH:mm" />
                  <option value="MMM DD, YYYY hh:mm A" />
                </datalist>
                <div className={`watermark-live-preview watermark-live-preview-${watermark.placement}`}>
                  <span style={{ textAlign: watermark.alignment }}>{watermarkDisplayText(watermark)}</span>
                </div>
                <div className="watermark-code-row">
                  <span>Verification code</span>
                  <code>{watermark.verificationId}</code>
                  <button
                    className="icon-btn"
                    type="button"
                    disabled={!canEditDocument}
                    title="Generate a new verification code"
                    aria-label="Generate a new verification code"
                    onClick={() => updateWatermark({ verificationId: makeVerificationId() })}
                  >
                    <RefreshCw size={13} />
                  </button>
                </div>
                <p className="watermark-note">The displayed timestamp is captured when each file is exported.</p>
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
              <button
                className="text-btn workspace-library-toggle"
                onClick={() => {
                  setWorkspaceLibrarySearch("");
                  setWorkspaceLibraryOpen((current) => current === "layouts" ? null : "layouts");
                }}
              >
                <PanelRight size={13} />
                {workspaceLibraryOpen === "layouts" ? "Hide saved layouts" : "Browse saved layouts"}
                <ChevronRight size={13} />
              </button>
              </div>
              {workspaceLibraryOpen === "layouts" && (
                <div className="workspace-library-browser layout-browser">
                  <div className="workspace-library-heading">
                    <div><span>LAYOUT LIBRARY</span><strong>Saved page designs</strong></div>
                    <button className="icon-btn" title="Create layout" onClick={() => openLayoutEditor()}><Plus size={13} /></button>
                  </div>
                  <label className="workspace-library-search">
                    <Search size={13} />
                    <input value={workspaceLibrarySearch} onChange={(event) => setWorkspaceLibrarySearch(event.target.value)} placeholder="Search layouts" />
                  </label>
                  <div className="workspace-library-list">
                    {workspaceLayoutRecords.map((record) => (
                      <div className="workspace-library-card workspace-layout-option" key={record.id}>
                        <div className="workspace-layout-swatch" style={{ borderTopColor: (record.letterhead.firstPage || record.letterhead).accent }} />
                        <span><strong>{record.name}</strong><small>{record.letterhead.page} · {record.status}</small></span>
                        <button className="outline-btn" onClick={() => applyLayoutRecord(record)}>Apply</button>
                        <button className="icon-btn" title={`Edit ${record.name}`} onClick={() => openLayoutEditor(record)}><Settings2 size={13} /></button>
                      </div>
                    ))}
                  </div>
                  <button className="workspace-full-management" onClick={() => switchModule("layouts")}>Open full management <ChevronRight size={12} /></button>
                </div>
              )}
            </div>
          )}
          {activeRightTab === "review" && (
            <div className="panel-content workspace-review-panel">
              <div className="panel-title-row">
                <div>
                  <span className="eyebrow">SELF REVIEW</span>
                  <h3>Document checks</h3>
                  <p className="panel-subtitle">Review source data, content and page output while keeping the document visible.</p>
                </div>
                <span className={`review-count ${blockingIssues.length ? "has-issues" : "ready"}`}>{blockingIssues.length}</span>
              </div>
              <div className={`review-summary ${blockingIssues.length ? "warning" : "ready"}`}>
                {blockingIssues.length ? <AlertTriangle size={16} /> : <ShieldCheck size={16} />}
                <span>
                  <strong>{blockingIssues.length ? `${blockingIssues.length} issue${blockingIssues.length === 1 ? "" : "s"} need attention` : "Ready for final preview"}</strong>
                  <small>{blockingIssues.length ? "Open a check below to fix it without leaving Workspace." : "All automated checks currently pass."}</small>
                </span>
              </div>
              <div className="workspace-review-list">
                {([
                  {
                    id: "source",
                    label: "Onboarding source",
                    detail: connectionStatus === "Failed" ? "Connection needs attention" : hasUnreviewedSync ? `${syncDiffs.length} source update${syncDiffs.length === 1 ? "" : "s"} waiting` : `${person.name} · ${person.submissionId}`,
                    issue: connectionStatus === "Failed" || hasUnreviewedSync,
                    tab: "person",
                    Icon: UserRound,
                  },
                  {
                    id: "fields",
                    label: "Required fields",
                    detail: blockingIssues.find((issue) => issue.startsWith("Missing") || issue.startsWith("Unresolved")) || "All active values resolved",
                    issue: blockingIssues.some((issue) => issue.startsWith("Missing") || issue.startsWith("Unresolved")),
                    tab: "placeholders",
                    Icon: Sparkles,
                  },
                  {
                    id: "clauses",
                    label: "Clauses and identity",
                    detail: blockingIssues.find((issue) => issue.includes("another person's") || issue.includes("billing")) || `${sections.flatMap((section) => section.clauses).filter((clause) => clause.included).length} included paragraphs`,
                    issue: blockingIssues.some((issue) => issue.includes("another person's") || issue.includes("billing")),
                    tab: "clauses",
                    Icon: Archive,
                  },
                  {
                    id: "layout",
                    label: "Layout and Letterhead",
                    detail: blockingIssues.find((issue) => issue.includes("overlaps")) || `${letterhead.page} · ${letterhead.mode === "different" ? "different page styles" : letterhead.mode === "all" ? "every page" : "first page"}`,
                    issue: blockingIssues.some((issue) => issue.includes("overlaps")),
                    tab: "letterhead",
                    Icon: PanelRight,
                  },
                  {
                    id: "verification",
                    label: "Verification watermark",
                    detail: blockingIssues.find((issue) => issue.includes("Verification")) || `${watermark.placement} · ${watermark.alignment} · ${watermark.verificationId}`,
                    issue: blockingIssues.some((issue) => issue.includes("Verification")),
                    tab: "letterhead",
                    Icon: ShieldCheck,
                  },
                ] as const).map(({ id, label, detail, issue, tab, Icon }) => (
                  <button className={`workspace-review-row ${issue ? "issue" : "ok"}`} key={id} onClick={() => openWorkspaceTool(tab)}>
                    <span className="workspace-review-icon"><Icon size={14} /></span>
                    <span><strong>{label}</strong><small>{detail}</small></span>
                    {issue ? <AlertTriangle size={14} /> : <Check size={14} />}
                    <ChevronRight size={13} />
                  </button>
                ))}
              </div>
              {blockingIssues.length > 0 && (
                <div className="workspace-review-issues" role="alert">
                  <strong>Issues found</strong>
                  {blockingIssues.map((issue) => <span key={issue}>{issue}</span>)}
                </div>
              )}
              <div className="workspace-review-actions">
                <button className="primary-btn" onClick={() => setShowPreview(true)}><FileCheck2 size={14} /> Preview final document</button>
                <button className="outline-btn" onClick={() => openWorkspaceTool("outline")}><PanelLeftOpen size={14} /> Back to outline</button>
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
              <div
                className={`verification-watermark verification-watermark-${watermark.placement} verification-watermark-${watermark.alignment}`}
                aria-label="Document verification watermark"
              >
                {watermarkDisplayText(watermark)}
              </div>
              {canvasBlocks.length ? canvasBlocks.map((block) => {
                if (block.kind === "letterhead") {
                  const layout = letterhead.firstPage || letterhead;
                  return <div className="preview-letterhead" key={block.id} style={{ borderColor: layout.accent }}>
                    {layout.dataUrl ? <img src={layout.dataUrl} alt="Letterhead" /> : <><strong>{values.company_name || "Company"}</strong><span>People & Culture</span></>}
                  </div>;
                }
                if (block.kind === "meta") return <div className="preview-sheet-head" key={block.id}><span>{template.type}</span><span>Page 1 of {estimatePageCount()}</span></div>;
                if (block.kind === "title") return <h1 key={block.id}>{docName}</h1>;
                if (block.kind === "lede") return <p key={block.id}>Between {values.company_name || "Company"} and {values.full_name || "Recipient"}</p>;
                if (block.kind === "section" && block.sectionId) {
                  const section = sections.find((item) => item.id === block.sectionId);
                  if (!section) return null;
                  return <div className="preview-section" key={block.id}>
                    <h3>{section.title}</h3>
                    {section.clauses.filter(shouldShowClause).map((c) => <p key={c.id} dangerouslySetInnerHTML={{ __html: clauseHtml(c) }} />)}
                  </div>;
                }
                if (block.kind === "signatures") return <div className="preview-signatures" key={block.id}><span>Company representative: {values.signatory_name || "________________"}</span><span>Employee / Contractor: {values.full_name || "________________"}</span></div>;
                if (block.kind === "footer") return <div className="preview-sheet-footer" key={block.id}><span>{values.company_name || "Company"} · Confidential</span><span>Page 1 of {estimatePageCount()}</span></div>;
                return null;
              }) : <div className="preview-empty-state"><FileText size={20} /><span>Blank canvas</span></div>}
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
