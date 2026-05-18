import type { Config, Context } from "@netlify/functions";
import { getDeployStore, getStore } from "@netlify/blobs";

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

const STORE_NAME = "greenacre-client-db";
const DATA_KEY = "data";
const USERS_KEY = "users";
const ASSET_FIELDS_KEY = "asset-fields";
const SESSION_COOKIE = "greenacre_session";
const SESSION_DAYS = 7;

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

function randomId(prefix: string): string {
  return `${prefix}-${crypto.randomUUID()}`;
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
  if (!label) return text("Enter a data cell name", 400);
  const fields = await store().get(ASSET_FIELDS_KEY, { type: "json" }) || [];
  const idBase = fieldId(label);
  let id = idBase;
  let suffix = 2;
  while (fields.some((field: { id: string }) => field.id === id)) id = `${idBase}_${suffix++}`;
  fields.push({ id, label, type, createdAt: new Date().toISOString(), createdBy: user.username });
  await store().setJSON(ASSET_FIELDS_KEY, fields);

  const data = await store().get(DATA_KEY, { type: "json" }) || { version: 0, state: {}, assetCustomFields: [] };
  data.assetCustomFields = fields;
  data.version = Number(data.version || 0) + 1;
  data.updatedAt = new Date().toISOString();
  data.updatedBy = user.username;
  await store().setJSON(DATA_KEY, data);

  return json({ assetCustomFields: fields });
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
  if (req.method === "GET" && path === "/api/users") return listUsers(req);
  if (req.method === "POST" && path === "/api/users") return postUser(req);

  return text("API route not found", 404);
};

export const config: Config = {
  path: "/api/*"
};
