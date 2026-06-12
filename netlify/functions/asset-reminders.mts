import type { Config } from "@netlify/functions";
import { getDeployStore, getStore } from "@netlify/blobs";

declare const Netlify: {
  context?: { deploy?: { context?: string } };
  env?: { get?: (name: string) => string | undefined };
};

const STORE_NAME = "greenacre-client-db";
const DATA_KEY = "data";
const REMINDER_SENT_KEY = "sent-reminders";

function store() {
  const isProduction = Netlify?.context?.deploy?.context === "production";
  return isProduction ? getStore(STORE_NAME, { consistency: "strong" }) : getDeployStore({ name: STORE_NAME });
}

function env(name: string): string {
  return Netlify?.env?.get?.(name) || "";
}

function cleanHeader(value = ""): string {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function parseDate(value = ""): Date | null {
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

function isoDate(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function compactDate(value = ""): string {
  return String(value || "").replace(/-/g, "") || new Date().toISOString().slice(0, 10).replace(/-/g, "");
}

function escapeIcs(value = ""): string {
  return String(value).replace(/\\/g, "\\\\").replace(/,/g, "\\,").replace(/;/g, "\\;").replace(/\r?\n/g, "\\n");
}

function htmlEscape(value = ""): string {
  return String(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[char] || char));
}

function fieldByLabel(fields: any[] = [], wanted: string): any {
  const key = cleanHeader(wanted);
  return fields.find((field) => cleanHeader(field?.label || field?.id) === key || cleanHeader(field?.id) === key);
}

function customValue(asset: any, fields: any[], labels: string[]): string {
  for (const label of labels) {
    const directId = label.startsWith("custom_") ? label : "";
    if (directId && asset?.customData?.[directId]) return String(asset.customData[directId]).trim();
    const field = fieldByLabel(fields, label);
    if (field?.id && asset?.customData?.[field.id]) return String(asset.customData[field.id]).trim();
  }
  return "";
}

async function sendBrevoReminder(payload: Record<string, string>): Promise<void> {
  const apiKey = env("BREVO_API_KEY");
  if (!apiKey) throw new Error("BREVO_API_KEY is not configured.");
  const senderEmail = env("BREVO_SENDER_EMAIL") || env("REMINDER_FROM_EMAIL") || "hello@greenecs.co.uk";
  const senderName = env("BREVO_SENDER_NAME") || env("REMINDER_FROM_NAME") || "Greenacre Client Asset Database";
  const toEmail = env("REMINDER_TO_EMAIL") || "georgia@greenecs.co.uk";
  const subject = `Asset test reminder - ${payload.siteReference || payload.siteName || "Site"} - ${payload.assetReference || payload.assetName || "Asset"}`;
  const reminderDate = payload.reminderDate || isoDate(new Date());
  const description = [
    `Client: ${payload.clientName || ""}`,
    `Site: ${payload.siteReference || ""} ${payload.siteName || ""}`,
    `Asset: ${payload.assetReference || ""} ${payload.assetName || ""}`,
    `Previous job/project numbers: ${payload.projectNumbers || ""}`,
    `Due date: ${payload.dueDate || ""}`
  ].join("\\n");
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Greenacre//Client Asset Database//EN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${crypto.randomUUID()}@greenecs.co.uk`,
    `DTSTAMP:${compactDate(new Date().toISOString().slice(0, 10))}T090000Z`,
    `DTSTART;VALUE=DATE:${compactDate(reminderDate)}`,
    `DTEND;VALUE=DATE:${compactDate(reminderDate)}`,
    `SUMMARY:${escapeIcs(subject)}`,
    `DESCRIPTION:${escapeIcs(description)}`,
    "END:VEVENT",
    "END:VCALENDAR"
  ].join("\r\n");
  const htmlContent = `
    <p>An asset is due for test in six weeks.</p>
    <p><strong>Client:</strong> ${htmlEscape(payload.clientName)}<br>
    <strong>Site:</strong> ${htmlEscape([payload.siteReference, payload.siteName].filter(Boolean).join(" - "))}<br>
    <strong>Asset:</strong> ${htmlEscape([payload.assetReference, payload.assetName].filter(Boolean).join(" - "))}<br>
    <strong>Previous job/project numbers:</strong> ${htmlEscape(payload.projectNumbers)}<br>
    <strong>Due date:</strong> ${htmlEscape(payload.dueDate)}<br>
    <strong>Calendar reminder date:</strong> ${htmlEscape(reminderDate)}</p>
  `;
  const response = await fetch("https://api.brevo.com/v3/smtp/email", {
    method: "POST",
    headers: { "api-key": apiKey, "Content-Type": "application/json" },
    body: JSON.stringify({
      sender: { name: senderName, email: senderEmail },
      to: [{ email: toEmail, name: "Georgia" }],
      subject,
      htmlContent,
      attachment: [{ name: "greenacre-asset-reminder.ics", content: btoa(ics) }]
    })
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => response.statusText);
    throw new Error(detail || response.statusText);
  }
}

export default async () => {
  const data = await store().get(DATA_KEY, { type: "json" }) as any;
  const state = data?.state || {};
  const fields = data?.assetCustomFields || [];
  const sent = (await store().get(REMINDER_SENT_KEY, { type: "json" }) || {}) as Record<string, string>;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sixWeeks = new Date(today);
  sixWeeks.setDate(sixWeeks.getDate() + 42);
  let checked = 0;
  let sentCount = 0;
  let skipped = 0;

  for (const client of state.clients || []) for (const site of client.sites || []) for (const asset of site.assets || []) {
    checked += 1;
    const dueDateText = customValue(asset, fields, ["custom_due_date", "Due Date"]);
    const dueDate = parseDate(dueDateText);
    if (!dueDate) {
      skipped += 1;
      continue;
    }
    dueDate.setHours(0, 0, 0, 0);
    if (dueDate < today || dueDate > sixWeeks) continue;
    const dueIso = isoDate(dueDate);
    const reminderKey = `${asset.id || asset.reference || asset.name}|${dueIso}`;
    if (sent[reminderKey]) continue;
    const reminderDate = new Date(dueDate);
    reminderDate.setDate(reminderDate.getDate() - 42);
    await sendBrevoReminder({
      clientName: client.name || "",
      siteReference: site.reference || "",
      siteName: site.name || "",
      assetReference: asset.reference || "",
      assetName: asset.name || "",
      projectNumbers: customValue(asset, fields, ["custom_job_numbers", "custom_recent_project_numbers", "Job Numbers / Project Numbers", "Project Numbers"]),
      dueDate: dueIso,
      reminderDate: isoDate(reminderDate)
    });
    sent[reminderKey] = new Date().toISOString();
    sentCount += 1;
  }

  await store().setJSON(REMINDER_SENT_KEY, sent);
  console.log(JSON.stringify({ checked, sent: sentCount, skippedNoDueDate: skipped }));
};

export const config: Config = {
  schedule: "0 7 * * *"
};
