import type { Config } from "@netlify/functions";
import { getStore } from "@netlify/blobs";

declare const Netlify: {
  env?: { get?: (name: string) => string | undefined };
};

type ReminderCandidate = {
  clientName: string;
  siteReference: string;
  siteName: string;
  assetId: string;
  assetReference: string;
  assetName: string;
  projectNumbers: string;
  dueDate: string;
};

const STORE_NAME = "greenacre-client-db";
const DATA_KEY = "data";
const SENT_PREFIX = "reminder-sends/";

const STANDARD_DUE_FIELD_ID = "custom_due_date";
const PROJECT_NUMBER_FIELD_KEYS = new Set([
  "recentprojectnumbers",
  "projectnumbers",
  "jobnumbers",
  "jobnumbersprojectnumbers"
]);

function store() {
  return getStore(STORE_NAME, { consistency: "strong" });
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

function htmlEscape(value = ""): string {
  return String(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[char] || char));
}

function escapeIcs(value = ""): string {
  return String(value).replace(/\\/g, "\\\\").replace(/,/g, "\\,").replace(/;/g, "\\;").replace(/\r?\n/g, "\\n");
}

function cleanKey(value = ""): string {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function compactDate(value = ""): string {
  return String(value || "").replace(/-/g, "");
}

function isoDate(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function parseDate(value: unknown): Date | null {
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

function daysAfterIso(value: string, days: number): string {
  const date = parseDate(value);
  if (!date) return value;
  date.setDate(date.getDate() + days);
  return isoDate(date);
}

function todayIso(): string {
  return isoDate(new Date());
}

function reminderWindow(today = todayIso()): { start: string; end: string } {
  return {
    start: daysAfterIso(today, 42),
    end: daysAfterIso(today, 48)
  };
}

function dateInWindow(value: string, window: { start: string; end: string }): boolean {
  return value >= window.start && value <= window.end;
}

function sendKey(candidate: ReminderCandidate): string {
  const assetId = encodeURIComponent(candidate.assetId || `${candidate.siteReference}-${candidate.assetReference}`);
  return `${SENT_PREFIX}${assetId}-${candidate.dueDate}`;
}

function dueFieldIds(fields: any[] = []): string[] {
  const ids = new Set<string>([STANDARD_DUE_FIELD_ID]);
  for (const field of fields) {
    if (cleanKey(field?.label || field?.id) === "duedate") ids.add(String(field.id));
  }
  return [...ids];
}

function projectNumberFieldIds(fields: any[] = []): string[] {
  return fields
    .filter((field) => PROJECT_NUMBER_FIELD_KEYS.has(cleanKey(field?.label || field?.id)))
    .map((field) => String(field.id));
}

function fieldValue(data: Record<string, unknown> = {}, ids: string[]): string {
  for (const id of ids) {
    const value = data[id];
    if (value !== undefined && value !== null && String(value).trim() !== "") return String(value).trim();
  }
  return "";
}

function collectReminderCandidates(data: any, window: { start: string; end: string }): ReminderCandidate[] {
  const fields = Array.isArray(data?.assetCustomFields) ? data.assetCustomFields : [];
  const projectIds = projectNumberFieldIds(fields);
  const candidates: ReminderCandidate[] = [];
  for (const client of data?.state?.clients || []) {
    for (const site of client?.sites || []) {
      for (const asset of site?.assets || []) {
        const customFields = [...fields, ...(Array.isArray(asset?.localCustomFields) ? asset.localCustomFields : [])];
        const dueDateValue = fieldValue(asset?.customData || {}, dueFieldIds(customFields));
        const dueDate = parseDate(dueDateValue);
        const dueDateIso = dueDate ? isoDate(dueDate) : "";
        if (!dueDateIso || !dateInWindow(dueDateIso, window)) continue;
        const projectNumbers = fieldValue(asset?.customData || {}, projectNumberFieldIds(customFields)) || fieldValue(asset?.customData || {}, projectIds);
        candidates.push({
          clientName: String(client?.name || ""),
          siteReference: String(site?.reference || ""),
          siteName: String(site?.name || ""),
          assetId: String(asset?.id || ""),
          assetReference: String(asset?.reference || ""),
          assetName: String(asset?.name || ""),
          projectNumbers,
          dueDate: dueDateIso
        });
      }
    }
  }
  return candidates;
}

async function sendBrevoReminder(candidate: ReminderCandidate): Promise<void> {
  const apiKey = env("BREVO_API_KEY");
  if (!apiKey) throw new Error("BREVO_API_KEY is not configured");
  const senderEmail = env("BREVO_SENDER_EMAIL") || env("REMINDER_FROM_EMAIL") || "hello@greenecs.co.uk";
  const senderName = env("BREVO_SENDER_NAME") || env("REMINDER_FROM_NAME") || "Greenacre Client Database";
  const reminderToEmail = env("REMINDER_TO_EMAIL") || "georgia@greenecs.co.uk";
  const inviteEndDate = daysAfterIso(candidate.dueDate, 1);
  const subject = `Asset test reminder - ${candidate.siteReference || candidate.siteName || "Site"} - ${candidate.assetReference || candidate.assetName || "Asset"}`;
  const description = [
    `Client: ${candidate.clientName}`,
    `Site: ${candidate.siteReference} ${candidate.siteName}`,
    `Asset: ${candidate.assetReference} ${candidate.assetName}`,
    `Previous job/project numbers: ${candidate.projectNumbers}`,
    `Due date: ${candidate.dueDate}`
  ].join("\\n");
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Greenacre//Client Database//EN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${crypto.randomUUID()}@greenecs.co.uk`,
    `DTSTAMP:${compactDate(todayIso())}T090000Z`,
    `DTSTART;VALUE=DATE:${compactDate(candidate.dueDate)}`,
    `DTEND;VALUE=DATE:${compactDate(inviteEndDate)}`,
    "SUMMARY:Greenacre asset test due",
    `DESCRIPTION:${escapeIcs(description)}`,
    "END:VEVENT",
    "END:VCALENDAR"
  ].join("\r\n");
  const htmlContent = `
    <p>An asset test is due in six weeks.</p>
    <p><strong>Client:</strong> ${htmlEscape(candidate.clientName)}<br>
    <strong>Site:</strong> ${htmlEscape([candidate.siteReference, candidate.siteName].filter(Boolean).join(" - "))}<br>
    <strong>Asset:</strong> ${htmlEscape([candidate.assetReference, candidate.assetName].filter(Boolean).join(" - "))}<br>
    <strong>Previous job/project numbers:</strong> ${htmlEscape(candidate.projectNumbers)}<br>
    <strong>Due date:</strong> ${htmlEscape(candidate.dueDate)}</p>
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
        name: "greenacre-asset-due-date.ics",
        content: btoa(ics)
      }]
    })
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => response.statusText);
    throw new Error(`Brevo email failed: ${detail || response.statusText}`);
  }
}

async function alreadySent(candidate: ReminderCandidate): Promise<boolean> {
  return Boolean(await store().getMetadata(sendKey(candidate)));
}

async function markSent(candidate: ReminderCandidate): Promise<void> {
  await store().setJSON(sendKey(candidate), {
    ...candidate,
    sentAt: new Date().toISOString(),
    protocol: "weekly-monday-six-weeks-before-due"
  });
}

export default async () => {
  const window = reminderWindow();
  const data = await store().get(DATA_KEY, { type: "json" }) || { state: { clients: [] }, assetCustomFields: [] };
  const candidates = collectReminderCandidates(data, window);
  const result = {
    ok: true,
    reminderWindow: window,
    matched: candidates.length,
    sent: 0,
    skippedAlreadySent: 0,
    failed: [] as Array<{ assetId: string; dueDate: string; error: string }>
  };

  for (const candidate of candidates) {
    if (await alreadySent(candidate)) {
      result.skippedAlreadySent += 1;
      continue;
    }
    try {
      await sendBrevoReminder(candidate);
      await markSent(candidate);
      result.sent += 1;
    } catch (error) {
      result.failed.push({
        assetId: candidate.assetId,
        dueDate: candidate.dueDate,
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  return json(result, { status: result.failed.length ? 207 : 200 });
};

export const config: Config = {
  schedule: "0 7 * * 1"
};
