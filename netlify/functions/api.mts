import type { Config, Context } from "@netlify/functions";
import { getDeployStore, getStore } from "@netlify/blobs";
import * as XLSX from "xlsx";

declare const Netlify: {
  context?: { deploy?: { context?: string } };
  env?: { get?: (name: string) => string | undefined };
};

type Role = "admin" | "user";

type StoredUser = {
  id: string;
  username: string;
  usernameKey: string;
  role: Role;
  salt: string;
  passwordHash: string;
  createdAt: string;
};

type PublicUser = Pick<StoredUser, "id" | "username" | "role" | "createdAt">;

type SessionRecord = {
  userId: string;
  createdAt: string;
  expiresAt: string;
};

type ImportStats = {
  clientsCreated: number;
  clientsExisting: number;
  sitesCreated: number;
  sitesExisting: number;
  assetsCreated: number;
  assetsExisting: number;
  cellsFilled: number;
  cellsKept: number;
  placeholderAssetsRemoved: number;
};

const STORE_NAME = "greenacre-client-db";
const DATA_KEY = "data";
const USERS_KEY = "users";
const ASSET_FIELDS_KEY = "asset-fields";
const SESSION_COOKIE = "greenacre_session";
const SESSION_DAYS = 7;
const MAX_ATTACHMENT_BYTES = 12 * 1024 * 1024;
const ALLOWED_ATTACHMENT_TYPES = new Set(["application/pdf", "image/jpeg", "image/png", "image/gif", "image/webp"]);
const STANDARD_ASSET_CUSTOM_FIELDS = [
  { id: "custom_recent_project_numbers", label: "Recent Project Numbers", type: "text", standard: true },
  { id: "custom_w3w_asset", label: "W3W (Asset)", type: "text", standard: true },
  { id: "custom_w3w_site_gate", label: "W3W (Site Gate)", type: "text", standard: true },
  { id: "custom_batch", label: "Batch", type: "text", standard: true },
  { id: "custom_frequency", label: "Frequency", type: "text", standard: true },
  { id: "custom_system_type", label: "System Type", type: "text", standard: true },
  { id: "custom_operating_state", label: "Operating State", type: "text", standard: true },
  { id: "custom_dco", label: "DCO", type: "text", standard: true },
  { id: "custom_last_test", label: "Last Test", type: "date", standard: true },
  { id: "custom_due_date", label: "Due Date", type: "date", standard: true },
  { id: "custom_status", label: "Status", type: "text", standard: true },
  { id: "custom_job_numbers", label: "Job Numbers / Project Numbers", type: "text", standard: true },
  { id: "custom_outstanding_reports", label: "Outstanding Reports", type: "text", standard: true },
  { id: "custom_odour_control", label: "Odour Control", type: "text", standard: true },
  { id: "custom_details", label: "Details", type: "text", standard: true },
  { id: "custom_recent_changes", label: "Recent Changes", type: "text", standard: true },
  { id: "custom_site_contacts", label: "Site Contacts", type: "text", standard: true }
];

const STANDARD_ASSET_FIELD_ALIASES: Record<string, string[]> = {
  custom_recent_project_numbers: ["Recent Project Numbers", "Recent Project Numbers.", "Project Numbers", "Previous Project Numbers"],
  custom_w3w_asset: ["W3W (Asset)", "What3Words Asset", "Asset W3W"],
  custom_w3w_site_gate: ["W3W (Site Gate)", "What3Words Site Gate", "Site Gate W3W"],
  custom_batch: ["Batch", "Batch Number"],
  custom_frequency: ["Frequency", "Service Frequency", "Test Frequency", "Inspection Frequency"],
  custom_system_type: ["System Type", "System"],
  custom_operating_state: ["Operating State", "Operational State", "Running State"],
  custom_dco: ["DCO", "D C O"],
  custom_last_test: ["Last Test", "Last Test Date"],
  custom_due_date: ["Due Date", "Next Test Due", "Next Due Date"],
  custom_status: ["Status", "Asset Status"],
  custom_job_numbers: ["Job Numbers / Project Numbers", "Job Numbers", "Project Numbers", "Recent Project Numbers", "Recent Project Numbers."],
  custom_outstanding_reports: ["Outstanding Reports", "Outstanding Report"],
  custom_odour_control: ["Odour Control"],
  custom_details: ["Details", "Asset Details", "Notes"],
  custom_recent_changes: ["Recent Changes", "Recent Change", "Changes"],
  custom_site_contacts: ["Site Contacts", "Site Contact"]
};

function store() {
  const isProduction = Netlify?.context?.deploy?.context === "production";
  return isProduction ? getStore(STORE_NAME, { consistency: "strong" }) : getDeployStore({ name: STORE_NAME });
}

function env(name: string): string {
  return Netlify?.env?.get?.(name) || "";
}

function json(data: unknown, init: ResponseInit = {}): Response {
  return new Response(JSON.stringify(data), {
    ...init,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      ...(init.headers || {})
    }
  });
}

function text(message: string, status = 400): Response {
  return new Response(message, {
    status,
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "Cache-Control": "no-store"
    }
  });
}

function escapeIcs(value = ""): string {
  return String(value).replace(/\\/g, "\\\\").replace(/,/g, "\\,").replace(/;/g, "\\;").replace(/\r?\n/g, "\\n");
}

function htmlEscape(value = ""): string {
  return String(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[char] || char));
}

function compactDate(value = ""): string {
  return String(value || "").replace(/-/g, "") || new Date().toISOString().slice(0, 10).replace(/-/g, "");
}

function sixWeeksBeforeIso(value = ""): string {
  const date = new Date(`${value}T09:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  date.setDate(date.getDate() - 42);
  return date.toISOString().slice(0, 10);
}

function randomId(prefix: string): string {
  return `${prefix}-${crypto.randomUUID()}`;
}

function safeFileName(value = "attachment"): string {
  const name = String(value || "attachment").replace(/[^\w.\- ()]+/g, "_").replace(/^_+|_+$/g, "");
  return name || "attachment";
}

function attachmentPrefix(assetId: string): string {
  return `attachments/${encodeURIComponent(assetId)}/`;
}

function attachmentKey(assetId: string, fileId: string): string {
  return `${attachmentPrefix(assetId)}${encodeURIComponent(fileId)}`;
}

function toBase64(bytes: Uint8Array): string {
  return btoa(String.fromCharCode(...bytes));
}

function fromBase64(value: string): Uint8Array {
  return Uint8Array.from(atob(value), (char) => char.charCodeAt(0));
}

async function sha256(value: string): Promise<string> {
  const data = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return toBase64(new Uint8Array(hash));
}

async function hashPassword(password: string, saltBase64?: string): Promise<{ salt: string; passwordHash: string }> {
  const salt = saltBase64 ? fromBase64(saltBase64) : crypto.getRandomValues(new Uint8Array(16));
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveBits"]);
  const saltBuffer = salt.buffer.slice(salt.byteOffset, salt.byteOffset + salt.byteLength) as ArrayBuffer;
  const bits = await crypto.subtle.deriveBits({ name: "PBKDF2", salt: saltBuffer, iterations: 210000, hash: "SHA-256" }, key, 256);
  return { salt: toBase64(salt), passwordHash: toBase64(new Uint8Array(bits)) };
}

async function verifyPassword(password: string, user: StoredUser): Promise<boolean> {
  const result = await hashPassword(password, user.salt);
  return result.passwordHash === user.passwordHash;
}

async function readUsers(): Promise<StoredUser[]> {
  return (await store().get(USERS_KEY, { type: "json" })) || [];
}

async function writeUsers(users: StoredUser[]): Promise<void> {
  await store().setJSON(USERS_KEY, users);
}

function publicUser(user: StoredUser): PublicUser {
  return {
    id: user.id,
    username: user.username,
    role: user.role,
    createdAt: user.createdAt
  };
}

function cookieToken(req: Request): string {
  const cookie = req.headers.get("cookie") || "";
  const match = cookie.match(new RegExp(`${SESSION_COOKIE}=([^;]+)`));
  return match ? decodeURIComponent(match[1]) : "";
}

function cookieHeader(token: string, req: Request): string {
  const secure = new URL(req.url).protocol === "https:" ? "; Secure" : "";
  return `${SESSION_COOKIE}=${encodeURIComponent(token)}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${SESSION_DAYS * 24 * 60 * 60}${secure}`;
}

function clearCookieHeader(req: Request): string {
  const secure = new URL(req.url).protocol === "https:" ? "; Secure" : "";
  return `${SESSION_COOKIE}=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0${secure}`;
}

async function currentUser(req: Request): Promise<StoredUser | null> {
  const token = cookieToken(req);
  if (!token) return null;
  const session = await store().get(`sessions/${await sha256(token)}`, { type: "json" }) as SessionRecord | null;
  if (!session || new Date(session.expiresAt).getTime() < Date.now()) return null;
  const users = await readUsers();
  return users.find((user) => user.id === session.userId) || null;
}

async function requireUser(req: Request): Promise<StoredUser | Response> {
  const user = await currentUser(req);
  return user || text("Not signed in", 401);
}

async function requireAdmin(req: Request): Promise<StoredUser | Response> {
  const user = await currentUser(req);
  if (!user) return text("Not signed in", 401);
  if (user.role !== "admin") return text("Admin access required", 403);
  return user;
}

async function createUser(username: string, password: string, role: Role): Promise<StoredUser> {
  const now = new Date().toISOString();
  const { salt, passwordHash } = await hashPassword(password);
  return {
    id: randomId("user"),
    username,
    usernameKey: username.trim().toLowerCase(),
    role,
    salt,
    passwordHash,
    createdAt: now
  };
}

async function login(req: Request): Promise<Response> {
  const body = await req.json().catch(() => ({}));
  const username = String(body.username || "").trim();
  const password = String(body.password || "");
  if (!username || !password) return text("Enter username and password", 400);

  let users = await readUsers();
  const usernameKey = username.toLowerCase();
  let user = users.find((item) => item.usernameKey === usernameKey) || null;

  if (!users.length) {
    const adminUsername = env("GREENACRE_ADMIN_USERNAME") || "admin";
    const adminPassword = env("GREENACRE_ADMIN_PASSWORD") || "GreenacreAdmin123!";
    if (username !== adminUsername || password !== adminPassword) {
      return text("First login must use the configured admin username and password", 401);
    }
    user = await createUser(username, password, "admin");
    users = [user];
    await writeUsers(users);
  }

  if (!user || !(await verifyPassword(password, user))) return text("Invalid username or password", 401);

  const token = crypto.randomUUID() + crypto.randomUUID();
  const expiresAt = new Date(Date.now() + SESSION_DAYS * 24 * 60 * 60 * 1000).toISOString();
  await store().setJSON(`sessions/${await sha256(token)}`, { userId: user.id, createdAt: new Date().toISOString(), expiresAt });
  return json({ user: publicUser(user) }, { headers: { "Set-Cookie": cookieHeader(token, req) } });
}

async function logout(req: Request): Promise<Response> {
  const token = cookieToken(req);
  if (token) await store().delete(`sessions/${await sha256(token)}`);
  return json({ ok: true }, { headers: { "Set-Cookie": clearCookieHeader(req) } });
}

async function getData(req: Request): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const stored = await store().get(DATA_KEY, { type: "json" }) || { version: 0, state: {}, assetCustomFields: [] };
  return json(stored);
}

async function putData(req: Request): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const body = await req.json().catch(() => null);
  if (!body || typeof body !== "object") return text("Invalid JSON", 400);
  const existing = await store().get(DATA_KEY, { type: "json" }) || { version: 0, state: {}, assetCustomFields: [] };
  const expectedVersion = Number(body.version || 0);
  if (Number(existing.version || 0) !== expectedVersion) {
    return json({ error: "Version conflict", latest: existing }, { status: 409 });
  }
  const next = {
    version: expectedVersion + 1,
    state: body.state || {},
    assetCustomFields: Array.isArray(body.assetCustomFields) ? body.assetCustomFields : existing.assetCustomFields || [],
    updatedAt: new Date().toISOString(),
    updatedBy: user.username
  };
  await store().setJSON(DATA_KEY, next);
  return json(next);
}

function fieldId(label: string): string {
  return `custom_${label.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "") || crypto.randomUUID()}`;
}

async function getAssetFields(req: Request): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const fields = await store().get(ASSET_FIELDS_KEY, { type: "json" }) || [];
  return json({ assetCustomFields: fields });
}

async function postAssetField(req: Request): Promise<Response> {
  const user = await requireAdmin(req);
  if (user instanceof Response) return user;
  const body = await req.json().catch(() => ({}));
  const label = String(body.label || "").trim();
  const type = ["text", "number", "date"].includes(String(body.type)) ? String(body.type) : "text";
  const scopeSiteId = String(body.scopeSiteId || "").trim();
  if (!label) return text("Enter a data cell name", 400);
  const fields = await store().get(ASSET_FIELDS_KEY, { type: "json" }) || [];
  const idBase = fieldId(label);
  let id = idBase;
  let suffix = 2;
  while (fields.some((field: { id: string }) => field.id === id)) id = `${idBase}_${suffix++}`;
  fields.push({ id, label, type, scopeSiteId, createdAt: new Date().toISOString(), createdBy: user.username });
  await store().setJSON(ASSET_FIELDS_KEY, fields);

  const data = await store().get(DATA_KEY, { type: "json" }) || { version: 0, state: {}, assetCustomFields: [] };
  data.assetCustomFields = fields;
  data.version = Number(data.version || 0) + 1;
  data.updatedAt = new Date().toISOString();
  data.updatedBy = user.username;
  await store().setJSON(DATA_KEY, data);

  return json({ assetCustomFields: fields });
}

function removeAssetFieldValues(state: any, fieldId: string): void {
  const clients = Array.isArray(state?.clients) ? state.clients : [];
  for (const client of clients) {
    const sites = Array.isArray(client?.sites) ? client.sites : [];
    for (const site of sites) {
      const assets = Array.isArray(site?.assets) ? site.assets : [];
      for (const asset of assets) {
        if (asset?.customData && Object.prototype.hasOwnProperty.call(asset.customData, fieldId)) {
          delete asset.customData[fieldId];
        }
      }
    }
  }
}

async function deleteAssetField(req: Request, fieldId: string): Promise<Response> {
  const user = await requireAdmin(req);
  if (user instanceof Response) return user;
  const id = decodeURIComponent(fieldId || "").trim();
  if (!id) return text("Data cell id is required", 400);
  const fields = await store().get(ASSET_FIELDS_KEY, { type: "json" }) || [];
  const nextFields = fields.filter((field: { id: string }) => field.id !== id);
  if (nextFields.length === fields.length) return text("Data cell not found", 404);
  await store().setJSON(ASSET_FIELDS_KEY, nextFields);

  const data = await store().get(DATA_KEY, { type: "json" }) || { version: 0, state: {}, assetCustomFields: [] };
  data.assetCustomFields = Array.isArray(data.assetCustomFields)
    ? data.assetCustomFields.filter((field: { id: string }) => field.id !== id)
    : nextFields;
  removeAssetFieldValues(data.state, id);
  data.version = Number(data.version || 0) + 1;
  data.updatedAt = new Date().toISOString();
  data.updatedBy = user.username;
  await store().setJSON(DATA_KEY, data);

  return json({ assetCustomFields: nextFields });
}

async function listUsers(req: Request): Promise<Response> {
  const user = await requireAdmin(req);
  if (user instanceof Response) return user;
  const users = await readUsers();
  return json({ users: users.map(publicUser) });
}

async function postUser(req: Request): Promise<Response> {
  const admin = await requireAdmin(req);
  if (admin instanceof Response) return admin;
  const body = await req.json().catch(() => ({}));
  const username = String(body.username || "").trim();
  const password = String(body.password || "");
  const role = body.role === "admin" ? "admin" : "user";
  if (!username || !password) return text("Enter username and password", 400);
  const users = await readUsers();
  if (users.some((user) => user.usernameKey === username.toLowerCase())) return text("Username already exists", 409);
  users.push(await createUser(username, password, role));
  await writeUsers(users);
  return json({ users: users.map(publicUser) }, { status: 201 });
}

async function resetUserPassword(req: Request, userId: string): Promise<Response> {
  const admin = await requireAdmin(req);
  if (admin instanceof Response) return admin;
  const id = decodeURIComponent(userId || "").trim();
  const body = await req.json().catch(() => ({}));
  const password = String(body.password || "");
  if (!id || !password) return text("Choose a user and enter a new password", 400);
  if (password.length < 8) return text("Password must be at least 8 characters", 400);
  const users = await readUsers();
  const user = users.find((item) => item.id === id);
  if (!user) return text("User not found", 404);
  const { salt, passwordHash } = await hashPassword(password);
  user.salt = salt;
  user.passwordHash = passwordHash;
  await writeUsers(users);
  return json({ users: users.map(publicUser) });
}

async function sendReminderPayload(body: Record<string, unknown>): Promise<Response> {
  const apiKey = env("BREVO_API_KEY");
  const senderEmail = env("BREVO_SENDER_EMAIL") || env("REMINDER_FROM_EMAIL") || "hello@greenecs.co.uk";
  const senderName = env("BREVO_SENDER_NAME") || env("REMINDER_FROM_NAME") || "Greenacre Client Database";
  const reminderToEmail = env("REMINDER_TO_EMAIL") || "georgia@greenecs.co.uk";
  if (!apiKey) return json({ ok: true, configured: false });

  const dueDate = String(body.dueDate || "").trim();
  if (!dueDate) return text("Due date is required", 400);
  const reminderDate = sixWeeksBeforeIso(dueDate);
  const subject = `Asset test reminder - ${body.siteReference || body.siteName || "Site"} - ${body.assetReference || body.assetName || "Asset"}`;
  const summary = `Greenacre asset test reminder`;
  const description = [
    `Client: ${body.clientName || ""}`,
    `Site: ${body.siteReference || ""} ${body.siteName || ""}`,
    `Asset: ${body.assetReference || ""} ${body.assetName || ""}`,
    `Previous job/project numbers: ${body.projectNumbers || ""}`,
    `Due date: ${dueDate}`
  ].join("\\n");
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Greenacre//Client Database//EN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${crypto.randomUUID()}@greenecs.co.uk`,
    `DTSTAMP:${compactDate(new Date().toISOString().slice(0, 10))}T090000Z`,
    `DTSTART;VALUE=DATE:${compactDate(reminderDate)}`,
    `DTEND;VALUE=DATE:${compactDate(reminderDate)}`,
    `SUMMARY:${escapeIcs(summary)}`,
    `DESCRIPTION:${escapeIcs(description)}`,
    "END:VEVENT",
    "END:VCALENDAR"
  ].join("\r\n");

  const htmlContent = `
    <p>An asset test reminder has been created.</p>
    <p><strong>Client:</strong> ${htmlEscape(String(body.clientName || ""))}<br>
    <strong>Site:</strong> ${htmlEscape([body.siteReference, body.siteName].filter(Boolean).join(" - "))}<br>
    <strong>Asset:</strong> ${htmlEscape([body.assetReference, body.assetName].filter(Boolean).join(" - "))}<br>
    <strong>Previous job/project numbers:</strong> ${htmlEscape(String(body.projectNumbers || ""))}<br>
    <strong>Due date:</strong> ${htmlEscape(dueDate)}<br>
    <strong>Calendar reminder date:</strong> ${htmlEscape(reminderDate)}</p>
  `;

  const response = await fetch("https://api.brevo.com/v3/smtp/email", {
    method: "POST",
    headers: {
      "api-key": apiKey,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      sender: { name: senderName, email: senderEmail },
      to: [{ email: reminderToEmail, name: "Georgia" }],
      subject,
      htmlContent,
      attachment: [{
        name: "greenacre-asset-reminder.ics",
        content: btoa(ics)
      }]
    })
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => response.statusText);
    return text(`Brevo email failed: ${detail || response.statusText}`, 502);
  }
  return json({ ok: true, configured: true, reminderDate });
}

async function postReminderEmail(req: Request): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const body = await req.json().catch(() => ({}));
  return sendReminderPayload(body);
}

async function postReminderEmailTest(req: Request): Promise<Response> {
  const user = await requireAdmin(req);
  if (user instanceof Response) return user;
  const due = new Date();
  due.setMonth(due.getMonth() + 1);
  return sendReminderPayload({
    clientName: "Greenacre Brevo Test",
    siteReference: "TEST-SITE",
    siteName: "Brevo Email Test",
    assetReference: "TEST-ASSET",
    assetName: "Reminder Email Test",
    projectNumbers: "TEST-EMAIL",
    dueDate: due.toISOString().slice(0, 10)
  });
}

async function listAssetAttachments(req: Request, assetId: string): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const id = decodeURIComponent(assetId || "").trim();
  if (!id) return text("Asset id is required", 400);
  const { blobs } = await store().list({ prefix: attachmentPrefix(id) });
  const files = await Promise.all((blobs || []).map(async (blob: { key: string }) => {
    const fileId = decodeURIComponent(blob.key.slice(attachmentPrefix(id).length));
    const stored = await store().getWithMetadata(blob.key);
    const metadata = (stored?.metadata || {}) as Record<string, string | number>;
    return {
      id: fileId,
      name: metadata.name || fileId,
      type: metadata.contentType || "application/octet-stream",
      size: Number(metadata.size || 0),
      uploadedAt: metadata.uploadedAt || "",
      uploadedBy: metadata.uploadedBy || "",
      url: `/api/asset-attachments/${encodeURIComponent(id)}/${encodeURIComponent(fileId)}`
    };
  }));
  files.sort((a, b) => String(b.uploadedAt || "").localeCompare(String(a.uploadedAt || "")));
  return json({ files });
}

async function uploadAssetAttachment(req: Request, assetId: string): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const id = decodeURIComponent(assetId || "").trim();
  if (!id) return text("Asset id is required", 400);
  const form = await req.formData().catch(() => null);
  const file = form?.get("file");
  if (!(file instanceof File)) return text("Choose a file to upload", 400);
  if (!ALLOWED_ATTACHMENT_TYPES.has(file.type)) return text("Only photos and PDF files can be uploaded", 400);
  if (file.size > MAX_ATTACHMENT_BYTES) return text("File is too large. Maximum size is 12 MB.", 400);
  const fileId = `${Date.now()}-${crypto.randomUUID()}-${safeFileName(file.name)}`;
  const buffer = await file.arrayBuffer();
  await store().set(attachmentKey(id, fileId), buffer, {
    metadata: {
      name: safeFileName(file.name),
      contentType: file.type,
      size: file.size,
      uploadedAt: new Date().toISOString(),
      uploadedBy: user.username
    }
  });
  return json({ ok: true, fileId }, { status: 201 });
}

async function getAssetAttachment(req: Request, assetId: string, fileId: string): Promise<Response> {
  const user = await requireUser(req);
  if (user instanceof Response) return user;
  const id = decodeURIComponent(assetId || "").trim();
  const file = decodeURIComponent(fileId || "").trim();
  if (!id || !file) return text("Attachment not found", 404);
  const key = attachmentKey(id, file);
  const stored = await store().getWithMetadata(key, { type: "arrayBuffer" });
  if (!stored?.data) return text("Attachment not found", 404);
  const metadata = (stored.metadata || {}) as Record<string, string>;
  const name = safeFileName(metadata.name || file);
  return new Response(stored.data as BodyInit, {
    headers: {
      "Content-Type": metadata.contentType || "application/octet-stream",
      "Content-Disposition": `inline; filename="${name.replace(/"/g, "")}"`,
      "Cache-Control": "private, no-store"
    }
  });
}

async function deleteAssetAttachment(req: Request, assetId: string, fileId: string): Promise<Response> {
  const user = await requireAdmin(req);
  if (user instanceof Response) return user;
  const id = decodeURIComponent(assetId || "").trim();
  const file = decodeURIComponent(fileId || "").trim();
  if (!id || !file) return text("Attachment not found", 404);
  await store().delete(attachmentKey(id, file));
  return json({ ok: true });
}

function importHeaderKey(value = ""): string {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function canonicalAssetFieldLabel(label = ""): string {
  return String(label || "").trim().replace(/^extra\s*:\s*/i, "").trim();
}

function assetFieldLabelKey(label = ""): string {
  return canonicalAssetFieldLabel(label).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function standardAssetFieldForLabel(label = ""): any {
  const labelKey = assetFieldLabelKey(label);
  return STANDARD_ASSET_CUSTOM_FIELDS.find((standard) => {
    const aliases = STANDARD_ASSET_FIELD_ALIASES[standard.id] || [standard.label];
    return [standard.label, ...aliases].some((alias) => assetFieldLabelKey(alias) === labelKey);
  }) || null;
}

function ensureStandardAssetFields(fields: any[] = []): any[] {
  const merged = Array.isArray(fields) ? [...fields] : [];
  for (const standard of STANDARD_ASSET_CUSTOM_FIELDS.slice().reverse()) {
    const existing = merged.find((field) => field.id === standard.id && !field.scopeSiteId);
    if (existing) {
      existing.type ||= standard.type;
      existing.standard = true;
    } else {
      merged.unshift({ ...standard });
    }
  }
  return merged;
}

function importFieldId(label = ""): string {
  return `custom_${canonicalAssetFieldLabel(label).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "") || crypto.randomUUID()}`;
}

function findOrCreateField(fields: any[], label = ""): any {
  const clean = canonicalAssetFieldLabel(label);
  const standard = standardAssetFieldForLabel(clean);
  if (standard) return standard;
  const key = assetFieldLabelKey(clean);
  let field = fields.find((item) => !item.scopeSiteId && assetFieldLabelKey(item.label) === key);
  if (!field) {
    let id = importFieldId(clean);
    let suffix = 2;
    while (fields.some((item) => item.id === id)) id = `${importFieldId(clean)}_${suffix++}`;
    field = { id, label: clean, type: /date|due|last\s*test/i.test(clean) ? "date" : "text", imported: true };
    fields.push(field);
  }
  return field;
}

function excelCell(row: Record<string, any>, aliases: string[]): string {
  const aliasKeys = aliases.map(importHeaderKey);
  const key = Object.keys(row || {}).find((header) => aliasKeys.includes(importHeaderKey(header)));
  if (!key) return "";
  const value = row[key];
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  return String(value ?? "").trim();
}

function canFillExisting(value = ""): boolean {
  const textValue = String(value ?? "").trim();
  return !textValue || /^unnamed\b/i.test(textValue) || /^site\s+\d+$/i.test(textValue) || /^asset\s+\d+$/i.test(textValue) || /^imported client$/i.test(textValue);
}

function importMatchKey(value = ""): string {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function weakImportRef(value = ""): boolean {
  const key = importMatchKey(value);
  return !key || ["tbc", "tba", "n/a", "na", "none", "no ref", "noref", "-"].includes(key);
}

function applyImportedValue(target: any, key: string, value: string, stats?: ImportStats): void {
  const clean = String(value ?? "").trim();
  if (!clean) return;
  if (!canFillExisting(target?.[key])) {
    if (stats) stats.cellsKept += 1;
    return;
  }
  target[key] = clean;
  if (stats) stats.cellsFilled += 1;
}

function applyImportedCustomValue(asset: any, fieldId: string, value: string, stats?: ImportStats): void {
  const clean = String(value ?? "").trim();
  if (!clean) return;
  asset.customData ||= {};
  if (!canFillExisting(asset.customData[fieldId])) {
    if (stats) stats.cellsKept += 1;
    return;
  }
  asset.customData[fieldId] = clean;
  if (stats) stats.cellsFilled += 1;
}

function readWorkbookRows(workbook: XLSX.WorkBook, names: string[], fallbackToFirstSheet = false): Record<string, any>[] {
  const wanted = names.map(importHeaderKey);
  const sheetName = workbook.SheetNames.find((name) => wanted.includes(importHeaderKey(name))) || (fallbackToFirstSheet ? workbook.SheetNames[0] : "");
  return sheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", raw: false }) as Record<string, any>[] : [];
}

function clientNameFromImport(value = ""): string {
  const name = String(value || "").trim();
  if (/^sw$/i.test(name)) return "Scottish Water";
  if (/^yws$/i.test(name)) return "Yorkshire Water Services Ltd";
  return name;
}

function blankImportedClient(name: string): any {
  return { id: randomId("client"), name, address: "", contacts: [], sites: [] };
}

function blankImportedSite(index: number): any {
  return { id: randomId("site"), reference: "", name: `Site ${index}`, address: "", assets: [] };
}

function blankImportedAsset(index: number): any {
  return {
    id: randomId("asset"),
    reference: "",
    name: `Asset ${index}`,
    location: "",
    flowRate: "",
    fanModel: "",
    fanSerial: "",
    flowSensorModel: "",
    flowSensorSerial: "",
    flowSetpointPercent: "",
    carbonMediaType: "",
    carbonMediaSlNumber: "",
    carbonMediaVolume: "",
    carbonPressureDrop: "",
    carbonHighSetpoint: "",
    carbonHighHighSetpoint: "",
    carbonTempProbeModel: "",
    carbonTempProbeSerial: "",
    bioMediaType: "",
    bioMediaVolume: "",
    customData: {}
  };
}

function findOrCreateImportedClient(state: any, name: string, stats?: ImportStats): any {
  const clientName = clientNameFromImport(name) || "Imported Client";
  state.clients ||= [];
  let client = state.clients.find((item: any) => importMatchKey(item.name) === importMatchKey(clientName));
  if (!client) {
    client = blankImportedClient(clientName);
    state.clients.push(client);
    if (stats) stats.clientsCreated += 1;
  } else if (stats) {
    stats.clientsExisting += 1;
  }
  client.sites ||= [];
  client.contacts ||= [];
  return client;
}

function findOrCreateImportedSite(client: any, reference: string, name: string, address: string, stats?: ImportStats): any {
  const ref = String(reference || "").trim();
  const siteName = String(name || "").trim() || ref || `Site ${(client.sites || []).length + 1}`;
  client.sites ||= [];
  const refKey = importMatchKey(ref);
  const nameKey = importMatchKey(siteName);
  let site = client.sites.find((item: any) => {
    const itemRef = importMatchKey(item.reference);
    const itemName = importMatchKey(item.name);
    return (!weakImportRef(ref) && itemRef === refKey) || itemName === nameKey;
  });
  if (!site) {
    site = { ...blankImportedSite(client.sites.length + 1), reference: ref, name: siteName, address: String(address || "").trim() };
    client.sites.push(site);
    if (stats) stats.sitesCreated += 1;
  } else {
    if (stats) stats.sitesExisting += 1;
    applyImportedValue(site, "reference", ref, stats);
    applyImportedValue(site, "name", siteName, stats);
    applyImportedValue(site, "address", address, stats);
  }
  site.assets ||= [];
  return site;
}

function findOrCreateImportedAsset(site: any, reference: string, name: string, location: string, stats?: ImportStats): any {
  const ref = String(reference || "").trim();
  const assetName = String(name || "").trim() || ref || `Asset ${(site.assets || []).length + 1}`;
  const assetLocation = String(location || "").trim();
  site.assets ||= [];
  const refKey = importMatchKey(ref);
  const nameKey = importMatchKey(assetName);
  const locationKey = importMatchKey(assetLocation);
  let asset = site.assets.find((item: any) => {
    const itemRef = importMatchKey(item.reference);
    const itemName = importMatchKey(item.name);
    const itemLocation = importMatchKey(item.location);
    if (!weakImportRef(ref) && itemRef === refKey) return true;
    if (nameKey && itemName === nameKey && (!locationKey || !itemLocation || itemLocation === locationKey)) return true;
    return false;
  });
  if (!asset) {
    asset = { ...blankImportedAsset(site.assets.length + 1), reference: ref, name: assetName, location: assetLocation };
    site.assets.push(asset);
    if (stats) stats.assetsCreated += 1;
  } else {
    if (stats) stats.assetsExisting += 1;
    applyImportedValue(asset, "reference", ref, stats);
    applyImportedValue(asset, "name", assetName, stats);
    applyImportedValue(asset, "location", location, stats);
  }
  asset.customData ||= {};
  return asset;
}

function importedCustomAliasGroups(): Record<string, string[]> {
  return {
    "Job Numbers / Project Numbers": STANDARD_ASSET_FIELD_ALIASES.custom_job_numbers,
    "Recent Project Numbers": STANDARD_ASSET_FIELD_ALIASES.custom_recent_project_numbers,
    "Batch": STANDARD_ASSET_FIELD_ALIASES.custom_batch,
    "Frequency": STANDARD_ASSET_FIELD_ALIASES.custom_frequency,
    "System Type": STANDARD_ASSET_FIELD_ALIASES.custom_system_type,
    "Operating State": STANDARD_ASSET_FIELD_ALIASES.custom_operating_state,
    "Last Test": STANDARD_ASSET_FIELD_ALIASES.custom_last_test,
    "Due Date": STANDARD_ASSET_FIELD_ALIASES.custom_due_date,
    "W3W (Asset)": STANDARD_ASSET_FIELD_ALIASES.custom_w3w_asset,
    "W3W (Site Gate)": STANDARD_ASSET_FIELD_ALIASES.custom_w3w_site_gate,
    "DCO": STANDARD_ASSET_FIELD_ALIASES.custom_dco,
    "Status": STANDARD_ASSET_FIELD_ALIASES.custom_status,
    "Outstanding Reports": STANDARD_ASSET_FIELD_ALIASES.custom_outstanding_reports,
    "Odour Control": STANDARD_ASSET_FIELD_ALIASES.custom_odour_control,
    "Details": STANDARD_ASSET_FIELD_ALIASES.custom_details,
    "Recent Changes": STANDARD_ASSET_FIELD_ALIASES.custom_recent_changes,
    "Site Contacts": STANDARD_ASSET_FIELD_ALIASES.custom_site_contacts
  };
}

function applyImportedAssetFields(row: Record<string, any>, asset: any, fields: any[], stats?: ImportStats): void {
  const coreMap: Record<string, string[]> = {
    flowRate: ["Flow Rate", "Flow", "Flow m3/h", "Flow Rate m3/h"],
    fanModel: ["Fan Model"],
    fanSerial: ["Fan Serial", "Fan Serial Number"],
    flowSensorModel: ["Low Flow Sensor Model", "Flow Sensor Model"],
    flowSensorSerial: ["Low Flow Sensor Serial Number", "Flow Sensor Serial Number"],
    flowSetpointPercent: ["Low Flow Set Point %", "Low Flow Setpoint %", "Flow Set Point %", "Flow Setpoint %"],
    carbonMediaType: ["Carbon Media Type"],
    carbonMediaSlNumber: ["Carbon Media SL Number", "Carbon SL Number"],
    carbonMediaVolume: ["Carbon Media Volume", "Carbon Media Volume m3"],
    carbonPressureDrop: ["Carbon PT Pressure Drop", "Carbon Pressure Drop"],
    carbonHighSetpoint: ["Carbon PT High Setpoint", "Carbon High Setpoint"],
    carbonHighHighSetpoint: ["Carbon PT High-High Setpoint", "Carbon High High Setpoint"],
    carbonTempProbeModel: ["Carbon Temp Probe Model"],
    carbonTempProbeSerial: ["Carbon Temp Probe Serial"],
    bioMediaType: ["Biofilter Media Type", "Bio Media Type"],
    bioMediaVolume: ["Biofilter Media Volume", "Biofilter Media Volume m3", "Bio Media Volume"]
  };
  for (const [key, aliases] of Object.entries(coreMap)) applyImportedValue(asset, key, excelCell(row, aliases), stats);
  const aliasGroups = importedCustomAliasGroups();
  for (const standard of STANDARD_ASSET_CUSTOM_FIELDS) {
    const value = excelCell(row, [standard.label, ...(aliasGroups[standard.label] || []), `Extra: ${standard.label}`]);
    applyImportedCustomValue(asset, standard.id, value, stats);
  }
  const knownKeys = new Set([
    "Client", "Client Name", "Company", "Customer", "Client Code", "Site Reference", "Site Ref", "Site ID", "Site", "Site Name", "Site Address", "Address",
    "Asset Reference", "Asset Ref", "Asset ID", "Equipment Reference", "Asset", "Asset Name", "Equipment", "Equipment Name", "Asset Location", "Location",
    ...Object.values(coreMap).flat(), ...Object.entries(aliasGroups).flatMap(([label, aliases]) => [label, `Extra: ${label}`, ...aliases])
  ].map(importHeaderKey));
  for (const [header, raw] of Object.entries(row)) {
    if (!String(raw ?? "").trim() || knownKeys.has(importHeaderKey(header))) continue;
    const field = findOrCreateField(fields, header);
    applyImportedCustomValue(asset, field.id, String(raw ?? "").trim(), stats);
  }
}

function normalizeStandardCustomData(state: any, fields: any[]): void {
  fields = ensureStandardAssetFields(fields);
  for (const client of state.clients || []) for (const site of client.sites || []) for (const asset of site.assets || []) {
    asset.customData ||= {};
    for (const field of fields) {
      const standard = standardAssetFieldForLabel(field.label);
      if (!standard || field.id === standard.id || asset.customData[field.id] === undefined || asset.customData[field.id] === "") continue;
      if (canFillExisting(asset.customData[standard.id])) asset.customData[standard.id] = asset.customData[field.id];
      delete asset.customData[field.id];
    }
  }
}

function parseImportedDate(value: string): Date | null {
  const text = String(value || "").trim();
  if (!text) return null;
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  const uk = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (uk) {
    const year = Number(uk[3].length === 2 ? `20${uk[3]}` : uk[3]);
    return new Date(year, Number(uk[2]) - 1, Number(uk[1]));
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function frequencyMonths(value: string): number {
  const text = String(value || "").toLowerCase();
  const match = text.match(/\d+(?:\.\d+)?/);
  const number = match ? Number(match[0]) : 0;
  if (!number) return 0;
  if (text.includes("year")) return Math.round(number * 12);
  if (text.includes("month")) return Math.round(number);
  if (text.includes("week")) return Math.round(number / 4.345);
  return Math.round(number * 12);
}

function isoDate(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function autoPopulateImportedDueDates(state: any): void {
  for (const client of state.clients || []) for (const site of client.sites || []) for (const asset of site.assets || []) {
    asset.customData ||= {};
    const lastTest = parseImportedDate(asset.customData.custom_last_test);
    const months = frequencyMonths(asset.customData.custom_frequency);
    if (!lastTest || !months) continue;
    const due = new Date(lastTest);
    due.setMonth(due.getMonth() + months);
    asset.customData.custom_due_date = isoDate(due);
  }
}

function removeEmptyImportedPlaceholderAssets(state: any): number {
  let removed = 0;
  for (const client of state.clients || []) for (const site of client.sites || []) {
    const assets = Array.isArray(site.assets) ? site.assets : [];
    site.assets = assets.filter((asset: any) => {
      const placeholderName = /^asset\s+\d+$/i.test(String(asset?.name || "").trim());
      const hasIdentity = String(asset?.reference || "").trim() || String(asset?.location || "").trim();
      const remove = placeholderName && !hasIdentity;
      if (remove) removed += 1;
      return !remove;
    });
  }
  return removed;
}

async function importExcel(req: Request): Promise<Response> {
  const user = await requireAdmin(req);
  if (user instanceof Response) return user;
  const form = await req.formData().catch(() => null);
  const file = form?.get("file");
  if (!(file instanceof File)) return text("Choose an Excel file to import", 400);
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const existing = await store().get(DATA_KEY, { type: "json" }) || { version: 0, state: { clients: [] }, assetCustomFields: [] };
  const state = existing.state || { clients: [] };
  state.clients ||= [];
  let fields = ensureStandardAssetFields(Array.isArray(existing.assetCustomFields) ? existing.assetCustomFields : []);
  const stats: ImportStats = {
    clientsCreated: 0,
    clientsExisting: 0,
    sitesCreated: 0,
    sitesExisting: 0,
    assetsCreated: 0,
    assetsExisting: 0,
    cellsFilled: 0,
    cellsKept: 0,
    placeholderAssetsRemoved: 0
  };

  let clientsProcessed = 0;
  for (const row of readWorkbookRows(workbook, ["Clients"])) {
    const name = excelCell(row, ["Client Name", "Client", "Company", "Customer"]);
    if (!name) continue;
    const client = findOrCreateImportedClient(state, name, stats);
    applyImportedValue(client, "address", excelCell(row, ["Billing Address", "BillingAddress", "Address"]), stats);
    clientsProcessed += 1;
  }

  let sitesProcessed = 0;
  for (const row of readWorkbookRows(workbook, ["Sites"])) {
    const clientName = excelCell(row, ["Client Name", "Client", "Company", "Customer", "Client Code"]);
    const siteReference = excelCell(row, ["Site Reference", "Site Ref", "Site ID"]);
    const siteName = excelCell(row, ["Site Name", "Site"]);
    if (!clientName && !siteReference && !siteName) continue;
    const client = findOrCreateImportedClient(state, clientName, stats);
    findOrCreateImportedSite(client, siteReference, siteName, excelCell(row, ["Site Address", "Address"]), stats);
    sitesProcessed += 1;
  }

  for (const row of readWorkbookRows(workbook, ["Asset Data Cells"])) {
    const label = excelCell(row, ["Data Cell Name", "Name", "Field", "Field Name"]);
    if (label) findOrCreateField(fields, label);
  }

  let assetRowsProcessed = 0;
  let assetRowsSkipped = 0;
  for (const row of readWorkbookRows(workbook, ["Assets", "Import"], true)) {
    const clientName = excelCell(row, ["Client", "Client Name", "Company", "Customer", "Client Code"]);
    const siteReference = excelCell(row, ["Site Reference", "Site Ref", "Site ID"]);
    const siteName = excelCell(row, ["Site", "Site Name"]);
    const assetReference = excelCell(row, ["Asset Reference", "Asset Ref", "Asset ID", "Equipment Reference"]);
    const assetName = excelCell(row, ["Asset", "Asset Name", "Equipment", "Equipment Name"]);
    if (!clientName && !siteReference && !siteName && !assetReference && !assetName) continue;
    if (!assetReference && !assetName && !excelCell(row, ["Asset Location", "Location"])) {
      assetRowsSkipped += 1;
      continue;
    }
    const client = findOrCreateImportedClient(state, clientName, stats);
    const site = findOrCreateImportedSite(client, siteReference, siteName, excelCell(row, ["Site Address", "Address"]), stats);
    const asset = findOrCreateImportedAsset(site, assetReference, assetName, excelCell(row, ["Asset Location", "Location"]), stats);
    applyImportedAssetFields(row, asset, fields, stats);
    assetRowsProcessed += 1;
  }
  if (!assetRowsProcessed) return text("No client/site/asset rows found in the Assets or Import worksheet.", 400);

  normalizeStandardCustomData(state, fields);
  autoPopulateImportedDueDates(state);
  stats.placeholderAssetsRemoved = removeEmptyImportedPlaceholderAssets(state);
  fields = ensureStandardAssetFields(fields).filter((field) => field.standard || !standardAssetFieldForLabel(field.label));
  const next = {
    version: Number(existing.version || 0) + 1,
    state,
    assetCustomFields: fields,
    updatedAt: new Date().toISOString(),
    updatedBy: user.username
  };
  await store().setJSON(DATA_KEY, next);
  await store().setJSON(ASSET_FIELDS_KEY, fields);
  const siteCount = state.clients.reduce((sum: number, client: any) => sum + (client.sites || []).length, 0);
  const assetCount = state.clients.reduce((sum: number, client: any) => sum + (client.sites || []).reduce((siteSum: number, site: any) => siteSum + (site.assets || []).length, 0), 0);
  return json({
    ...next,
    importReport: {
      fileName: file.name,
      assetRowsProcessed,
      assetRowsSkipped,
      clientsProcessed,
      sitesProcessed,
      totals: { clients: state.clients.length, sites: siteCount, assets: assetCount },
      ...stats
    },
    summary: `${assetRowsProcessed} asset rows processed; ${assetRowsSkipped} site-only rows skipped. Added ${stats.clientsCreated} clients, ${stats.sitesCreated} sites, ${stats.assetsCreated} assets. Matched existing ${stats.clientsExisting} client rows, ${stats.sitesExisting} site rows, ${stats.assetsExisting} asset rows. Filled ${stats.cellsFilled} empty cells; kept ${stats.cellsKept} populated cells. Removed ${stats.placeholderAssetsRemoved} blank placeholder assets. Totals: clients ${state.clients.length}, sites ${siteCount}, assets ${assetCount}.`
  });
}

export default async (req: Request, context: Context): Promise<Response> => {
  const path = new URL(req.url).pathname.replace(/^\/\.netlify\/functions\/api/, "/api");

  if (req.method === "GET" && path === "/api/health") return json({ ok: true });
  if (req.method === "GET" && path === "/api/session") {
    const user = await currentUser(req);
    return user ? json({ user: publicUser(user) }) : text("Not signed in", 401);
  }
  if (req.method === "POST" && path === "/api/login") return login(req);
  if (req.method === "POST" && path === "/api/logout") return logout(req);
  if (req.method === "GET" && path === "/api/data") return getData(req);
  if (req.method === "PUT" && path === "/api/data") return putData(req);
  if (req.method === "GET" && path === "/api/asset-fields") return getAssetFields(req);
  if (req.method === "POST" && path === "/api/asset-fields") return postAssetField(req);
  if (req.method === "DELETE" && path.startsWith("/api/asset-fields/")) {
    return deleteAssetField(req, path.slice("/api/asset-fields/".length));
  }
  if (req.method === "GET" && path === "/api/users") return listUsers(req);
  if (req.method === "POST" && path === "/api/users") return postUser(req);
  if (req.method === "POST" && path.startsWith("/api/users/") && path.endsWith("/password")) {
    return resetUserPassword(req, path.slice("/api/users/".length, -"/password".length));
  }
  if (req.method === "POST" && path === "/api/reminder-email/test") return postReminderEmailTest(req);
  if (req.method === "POST" && path === "/api/reminder-email") return postReminderEmail(req);
  if (req.method === "POST" && path === "/api/import-excel") return importExcel(req);
  const assetAttachmentMatch = path.match(/^\/api\/assets\/([^/]+)\/attachments$/);
  if (assetAttachmentMatch && req.method === "GET") return listAssetAttachments(req, assetAttachmentMatch[1]);
  if (assetAttachmentMatch && req.method === "POST") return uploadAssetAttachment(req, assetAttachmentMatch[1]);
  const attachmentMatch = path.match(/^\/api\/asset-attachments\/([^/]+)\/([^/]+)$/);
  if (attachmentMatch && req.method === "GET") return getAssetAttachment(req, attachmentMatch[1], attachmentMatch[2]);
  if (attachmentMatch && req.method === "DELETE") return deleteAssetAttachment(req, attachmentMatch[1], attachmentMatch[2]);

  return text("API route not found", 404);
};

export const config: Config = {
  path: "/api/*"
};
