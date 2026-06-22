const SHAREPOINT_LISTS = {
  clients: "GCE_Clients",
  contacts: "GCE_ClientContacts",
  sites: "GCE_Sites",
  assets: "GCE_Assets"
};
const LOCAL_STORAGE_KEY = "greenacre-client-database-netlify-local-fallback-v1";
const SERVER_STORAGE = {
  healthUrl: "/api/health",
  dataUrl: "/api/data",
  sessionUrl: "/api/session",
  loginUrl: "/api/login",
  logoutUrl: "/api/logout",
  usersUrl: "/api/users",
  assetFieldsUrl: "/api/asset-fields",
  importExcelUrl: "/api/import-excel",
  reminderEmailTestUrl: "/api/reminder-email/test"
};

const defaultProducts = [];

const defaultState = {
  project: {
    reference: "",
    customer: "",
    site: "",
    item: "",
    engineer: "",
    checked: "",
    revision: ""
  },
  safety: { flow: 0, carbon: 0, bio: 0 },
  gas: { density: 1.23, viscosity: 0.000018 },
  filters: {
    carbon: { count: 0, diameter: 0, depth: 1, pellet: 4, voidFraction: 0.4 },
    bio: { count: 0, diameter: 0, depth: 1, pellet: 3.7, voidFraction: 0.44 }
  },
  options: { internalBypass: false, steelBase: false, electrical: false },
  commercial: {
    fanUnitPrice: 0,
    fanQty: 0,
    contingency: 0,
    markup: 0,
    installHours: 0,
    managementHours: 0,
    labourSet: "other"
  },
  schematic: {
    title: "",
    drawingNumber: "",
    revision: "",
    showBomTags: "no",
    showPressures: "no"
  },
  fan: { rpmMeasured: 0, rpmBase: 0, urvPa: 0, quickCheckPa: 0 },
  zones: [
    { name: "Zone 1", shape: "Cylinder", length: 0, width: 0, height: 0, diameter: 0, volumeOverride: 0, ach: 0, designVelocity: 0, ductDiameter: 0 }
  ],
  legs: [
    { name: "L1", share: 0, material: "GRP Ductwork", length: 0, diameter: 0, roughness: 0, sr90: 0, lr90: 0, bend45: 0, tee: 0, reducer: 0, nrd: 0, vcd: 0, sb: 0, shoe: 0, entrance: "None", exit: "None", customK: 0 }
  ],
  products: [],
  clients: [],
  suppliers: [],
  enquiries: [],
  quotes: [],
  activeClientId: "",
  activeEnquiryId: "",
  activeQuoteId: ""
};

let state = clone(defaultState);
let currentResults = null;
let activeAssetHistory = null;
let enquiryEditorOpen = false;
let quoteEditorOpen = false;
let quoteDraft = null;
let clientEditorOpen = false;
let clientDraft = null;
let activeClientSiteId = "";
let activeClientAssetId = "";
let clientDrillLevel = "register";
let supplierEditorOpen = false;
let supplierDraft = null;
let expandedProductCode = "";
let sharePointItems = {
  clients: new Map(),
  contacts: new Map(),
  sites: new Map(),
  assets: new Map()
};
let sharePointFields = {
  clients: {},
  contacts: {},
  sites: {},
  assets: {}
};
let sharePointSaveTimer = null;
let sharePointLoaded = false;
let sharePointSaving = false;
let serverLoaded = false;
let serverSaving = false;
let serverSaveTimer = null;
let serverVersion = 0;
let currentUser = null;
let assetCustomFields = [];

const $ = (selector) => document.querySelector(selector);
const fmt = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 });
const moneyFmt = new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP", maximumFractionDigits: 0 });
const QUOTE_NUMBER_START = 1940;
const REMINDER_EMAIL = "georgia@greenecs.co.uk";
let excelLibraryPromise = null;

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadExcelLibrary() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (excelLibraryPromise) return excelLibraryPromise;
  const sources = [
    "xlsx.full.min.js",
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"
  ];
  excelLibraryPromise = new Promise((resolve, reject) => {
    const trySource = (index) => {
      if (window.XLSX) {
        resolve(window.XLSX);
        return;
      }
      if (index >= sources.length) {
        reject(new Error("Excel library could not be loaded. Check xlsx.full.min.js is uploaded or allow the CDN fallback."));
        return;
      }
      const script = document.createElement("script");
      script.src = sources[index];
      script.async = true;
      script.onload = () => window.XLSX ? resolve(window.XLSX) : trySource(index + 1);
      script.onerror = () => trySource(index + 1);
      document.head.appendChild(script);
    };
    trySource(0);
  });
  return excelLibraryPromise;
}


function isAdminUser() {
  return currentUser?.role === "admin";
}

function adminDeleteButton(attributes, label = "Remove") {
  return isAdminUser() ? `<button class="danger-btn" ${attributes} type="button">${label}</button>` : "";
}

function requireAdminRemoval() {
  if (isAdminUser()) return true;
  alert("Only admin users can remove clients, sites, assets, contacts or asset data cells.");
  return false;
}

async function apiFetch(url, options = {}) {
  const isFormData = options.body instanceof FormData;
  const response = await fetch(url, {
    credentials: "same-origin",
    cache: "no-store",
    headers: {
      ...(options.body && !isFormData ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {})
    },
    ...options
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => response.statusText);
    throw new Error(`Request failed (${response.status}): ${detail || response.statusText}`);
  }
  if (response.status === 204) return null;
  return response.json();
}

async function loadCurrentUser() {
  try {
    const session = await apiFetch(SERVER_STORAGE.sessionUrl);
    currentUser = session.user || null;
    return currentUser;
  } catch {
    currentUser = null;
    return null;
  }
}

function showLoginScreen(message = "") {
  document.body.innerHTML = `
    <main class="auth-screen">
      <form id="loginForm" class="auth-card">
        <img src="Picture1.png" alt="Greenacre Environmental Systems Ltd">
        <h1>Greenacre Client Database</h1>
        <label>Username <input id="loginUsername" autocomplete="username" required></label>
        <label>Password <input id="loginPassword" type="password" autocomplete="current-password" required></label>
        <button type="submit">Sign In</button>
        <p id="loginError" class="auth-error">${escapeHtml(message)}</p>
      </form>
    </main>
  `;
  $("#loginForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const username = $("#loginUsername").value.trim();
    const password = $("#loginPassword").value;
    try {
      await apiFetch(SERVER_STORAGE.loginUrl, {
        method: "POST",
        body: JSON.stringify({ username, password })
      });
      window.location.reload();
    } catch (error) {
      $("#loginError").textContent = error.message || "Sign in failed";
    }
  });
}

function renderUserControls() {
  const actions = document.querySelector(".header-actions");
  if (!actions || document.querySelector("#userControls")) return;
  document.querySelectorAll(".admin-export-control").forEach((element) => {
    element.hidden = !isAdminUser();
  });
  const wrapper = document.createElement("div");
  wrapper.id = "userControls";
  wrapper.className = "header-actions";
  wrapper.innerHTML = `
    <span class="user-pill">${escapeHtml(currentUser?.username || "User")} ${isAdminUser() ? "(admin)" : ""}</span>
    ${isAdminUser() ? '<button id="toggleAdminPanelBtn" type="button">Admin</button>' : ""}
    <button id="logoutBtn" type="button">Logout</button>
  `;
  actions.appendChild(wrapper);

  const adminPanel = document.createElement("section");
  adminPanel.id = "adminPanel";
  adminPanel.className = "admin-panel";
  adminPanel.hidden = true;
  adminPanel.innerHTML = `
    <div class="panel-head">
      <h2>Admin</h2>
      <button id="closeAdminPanelBtn" type="button">Close</button>
    </div>
    <div class="form-grid compact">
      <label>Username <input id="adminNewUsername"></label>
      <label>Password <input id="adminNewPassword" type="password"></label>
      <label>Role
        <select id="adminNewRole">
          <option value="user">User</option>
          <option value="admin">Admin</option>
        </select>
      </label>
    </div>
    <button id="adminCreateUserBtn" type="button">Create User</button>
    <button id="adminTestEmailBtn" type="button">Send Test Reminder Email</button>
    <p id="adminUserMessage" class="admin-message"></p>
    <div id="adminUserList"></div>
  `;
  document.body.appendChild(adminPanel);

  $("#logoutBtn")?.addEventListener("click", async () => {
    await apiFetch(SERVER_STORAGE.logoutUrl, { method: "POST" }).catch(() => null);
    window.location.reload();
  });
  $("#toggleAdminPanelBtn")?.addEventListener("click", async () => {
    adminPanel.hidden = !adminPanel.hidden;
    if (!adminPanel.hidden) await renderAdminUsers();
  });
  $("#closeAdminPanelBtn")?.addEventListener("click", () => {
    adminPanel.hidden = true;
  });
  $("#adminCreateUserBtn")?.addEventListener("click", createAdminUser);
  $("#adminTestEmailBtn")?.addEventListener("click", sendAdminTestEmail);
}

async function renderAdminUsers() {
  if (!isAdminUser()) return;
  const list = $("#adminUserList");
  if (!list) return;
  try {
    const data = await apiFetch(SERVER_STORAGE.usersUrl);
    list.innerHTML = `
      <h3>Users</h3>
      <div class="admin-user-list">
        ${(data.users || []).map((user) => `
          <article class="contact-row">
            <strong>${escapeHtml(user.username)}</strong>
            <span>${escapeHtml(user.role)}</span>
            <input data-reset-password="${escapeHtml(user.id)}" type="password" placeholder="New password">
            <button data-reset-user-password="${escapeHtml(user.id)}" type="button">Reset Password</button>
          </article>
        `).join("")}
      </div>
    `;
  } catch (error) {
    list.innerHTML = `<p class="admin-message">${escapeHtml(error.message || "Could not load users")}</p>`;
  }
}

async function createAdminUser() {
  const message = $("#adminUserMessage");
  const username = $("#adminNewUsername")?.value.trim();
  const password = $("#adminNewPassword")?.value;
  const role = $("#adminNewRole")?.value || "user";
  if (!username || !password) {
    if (message) message.textContent = "Enter a username and password.";
    return;
  }
  try {
    await apiFetch(SERVER_STORAGE.usersUrl, {
      method: "POST",
      body: JSON.stringify({ username, password, role })
    });
    if (message) message.textContent = "User created.";
    ["adminNewUsername", "adminNewPassword"].forEach((id) => {
      const input = $(`#${id}`);
      if (input) input.value = "";
    });
    await renderAdminUsers();
  } catch (error) {
    if (message) message.textContent = error.message || "Could not create user.";
  }
}

async function sendAdminTestEmail() {
  const message = $("#adminUserMessage");
  try {
    if (message) message.textContent = "Sending test email...";
    const data = await apiFetch(SERVER_STORAGE.reminderEmailTestUrl, { method: "POST", body: JSON.stringify({}) });
    if (message) {
      message.textContent = data.configured
        ? `Test email sent. Calendar reminder date: ${data.reminderDate || ""}`
        : "Brevo is not configured. Add BREVO_API_KEY in Netlify environment variables.";
    }
  } catch (error) {
    if (message) message.textContent = error.message || "Test email failed.";
  }
}

async function resetAdminUserPassword(userId) {
  const message = $("#adminUserMessage");
  const input = document.querySelector(`[data-reset-password="${CSS.escape(userId)}"]`);
  const password = input?.value || "";
  if (!password) {
    if (message) message.textContent = "Enter a new password first.";
    return;
  }
  try {
    await apiFetch(`${SERVER_STORAGE.usersUrl}/${encodeURIComponent(userId)}/password`, {
      method: "POST",
      body: JSON.stringify({ password })
    });
    if (message) message.textContent = "Password reset.";
    if (input) input.value = "";
    await renderAdminUsers();
  } catch (error) {
    if (message) message.textContent = error.message || "Password reset failed.";
  }
}

function showImportReport(report, fallbackSummary = "") {
  document.querySelectorAll(".modal-backdrop").forEach((modal) => modal.remove());
  const overlay = document.createElement("div");
  overlay.className = "modal-backdrop";
  const rows = report ? [
    ["File", report.fileName || ""],
    ["Asset rows processed", report.assetRowsProcessed ?? 0],
    ["Site-only rows skipped", report.assetRowsSkipped ?? 0],
    ["Clients added", report.clientsCreated ?? 0],
    ["Sites added", report.sitesCreated ?? 0],
    ["Assets added", report.assetsCreated ?? 0],
    ["Existing assets matched", report.assetsExisting ?? 0],
    ["Empty cells filled", report.cellsFilled ?? 0],
    ["Populated cells kept", report.cellsKept ?? 0],
    ["Blank placeholder assets removed", report.placeholderAssetsRemoved ?? 0],
    ["Total clients", report.totals?.clients ?? 0],
    ["Total sites", report.totals?.sites ?? 0],
    ["Total assets", report.totals?.assets ?? 0]
  ] : [["Import", fallbackSummary || "Excel import complete."]];
  overlay.innerHTML = `
    <section class="modal-card">
      <div class="panel-head">
        <h2>Import Report</h2>
        <button id="closeImportReportBtn" type="button">Close</button>
      </div>
      <div class="table-wrap">
        <table>
          <tbody>
            ${rows.map(([label, value]) => `<tr><th>${escapeHtml(label)}</th><td>${escapeHtml(value)}</td></tr>`).join("")}
          </tbody>
        </table>
      </div>
      ${fallbackSummary ? `<p>${escapeHtml(fallbackSummary)}</p>` : ""}
    </section>
  `;
  document.body.appendChild(overlay);
  const close = () => overlay.remove();
  overlay.querySelector("#closeImportReportBtn")?.addEventListener("click", close);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
}

function sharedStateSnapshot() {
  return clientDatabaseSnapshot();
}

function clientDatabaseSnapshot() {
  return {
    clients: clone(state.clients || []),
    activeClientId: state.activeClientId || "",
    assetCustomFields: clone(assetCustomFields || [])
  };
}

function isProjectNumberField(field) {
  const key = cleanHeader(field?.label || field?.id || "");
  return ["recentprojectnumbers", "recentprojectnumbersasset", "projectnumbers", "jobnumbers"].includes(key);
}

function isDueDateField(field) {
  return cleanHeader(field?.label || field?.id || "") === "duedate";
}

function consolidateProjectNumberFields() {
  const fields = assetCustomFields || [];
  const projectFields = fields.filter(isProjectNumberField);
  if (!projectFields.length) return;
  const keeper = projectFields.find((field) => cleanHeader(field.label) === "projectnumbers") || projectFields[0];
  keeper.label = "Project Numbers";
  keeper.type = "text";
  keeper.id ||= "custom_project_numbers";
  for (const client of state.clients || []) {
    for (const site of client.sites || []) {
      for (const asset of site.assets || []) {
        asset.customData ||= {};
        const values = projectFields
          .map((field) => asset.customData?.[field.id])
          .filter((value) => value !== undefined && value !== null && String(value).trim() !== "")
          .map((value) => String(value).trim());
        if (values.length) {
          asset.customData[keeper.id] = Array.from(new Set(values.join(", ").split(",").map((value) => value.trim()).filter(Boolean))).join(", ");
        }
        projectFields.forEach((field) => {
          if (field.id !== keeper.id) delete asset.customData[field.id];
        });
      }
    }
  }
  assetCustomFields = [
    ...fields.filter((field) => !isProjectNumberField(field)),
    keeper
  ];
}

function cleanHeader(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function fieldId(label) {
  return `custom_${String(label || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "") || newId("field")}`;
}

function rowValue(row, names) {
  const wanted = names.map(cleanHeader);
  for (const [key, value] of Object.entries(row || {})) {
    if (wanted.includes(cleanHeader(key))) return value == null ? "" : String(value).trim();
  }
  return "";
}

function rowNumber(row, names) {
  const value = rowValue(row, names);
  return value === "" ? 0 : n(value);
}

function applyRowNumber(target, key, row, names) {
  const value = rowValue(row, names);
  if (value !== "") target[key] = n(value);
}

function findOrCreateClientByName(name, address = "") {
  const safeName = name || "Unnamed client";
  let client = (state.clients || []).find((item) => String(item.name || "").trim().toLowerCase() === safeName.toLowerCase());
  if (!client) {
    client = { id: newId("client"), name: safeName, address: address || "", contacts: [], sites: [] };
    state.clients ||= [];
    state.clients.push(client);
  }
  if (address) client.address = address;
  client.contacts ||= [];
  client.sites ||= [];
  return client;
}

function findOrCreateSite(client, reference, name, address = "") {
  const safeRef = reference || "";
  const safeName = name || safeRef || "Unnamed site";
  let site = (client.sites || []).find((item) => (
    (safeRef && String(item.reference || "").trim().toLowerCase() === safeRef.toLowerCase()) ||
    String(item.name || "").trim().toLowerCase() === safeName.toLowerCase()
  ));
  if (!site) {
    site = { ...blankSite((client.sites || []).length + 1), reference: safeRef, name: safeName, address: address || "", assets: [] };
    client.sites ||= [];
    client.sites.push(site);
  }
  if (safeRef) site.reference = safeRef;
  if (safeName) site.name = safeName;
  if (address) site.address = address;
  site.assets ||= [];
  return site;
}

function findOrCreateAsset(site, reference, name) {
  const safeRef = reference || "";
  const safeName = name || safeRef || "Unnamed asset";
  let asset = (site.assets || []).find((item) => (
    (safeRef && String(item.reference || "").trim().toLowerCase() === safeRef.toLowerCase()) ||
    String(item.name || "").trim().toLowerCase() === safeName.toLowerCase()
  ));
  if (!asset) {
    asset = { id: newId("asset"), reference: safeRef, name: safeName, location: "" };
    site.assets ||= [];
    site.assets.push(asset);
  }
  if (safeRef) asset.reference = safeRef;
  if (safeName) asset.name = safeName;
  return asset;
}

async function exportClientDatabaseExcel() {
  if (!window.XLSX) {
    setSyncStatus("Loading Excel tools", "", "Preparing the Excel download tools.");
    await loadExcelLibrary();
  }
  setSyncStatus("Preparing Excel export", "", "Building workbook for download.");
  const clients = state.clients || [];
  const clientRows = clients.map((client) => ({
    "Client Name": client.name || "",
    "Billing Address": client.address || ""
  }));
  const contactRows = clients.flatMap((client) => (client.contacts || []).map((contact) => ({
    "Client Name": client.name || "",
    "Contact Name": contact.name || "",
    "Role": contact.role || "",
    "Email": contact.email || "",
    "Phone": contact.phone || ""
  })));
  const siteRows = clients.flatMap((client) => (client.sites || []).map((site) => ({
    "Client Name": client.name || "",
    "Site Reference": site.reference || "",
    "Site Name": site.name || "",
    "Site Address": site.address || ""
  })));
  const assetRows = clients.flatMap((client) => (client.sites || []).flatMap((site) => (site.assets || []).map((asset) => {
    const normalized = normalizeAsset(asset, 1);
    const row = {
      "Client Name": client.name || "",
      "Site Reference": site.reference || "",
      "Site Name": site.name || "",
      "Asset Reference": normalized.reference || "",
      "Asset Name": normalized.name || "",
      "Asset Location": normalized.location || "",
      "Flow Rate m3/h": normalized.flowRate || 0,
      "Fan Model": normalized.fanModel || "",
      "Fan Serial Number": normalized.fanSerial || "",
      "Flow Sensor Model": normalized.flowSensorModel || "",
      "Flow Sensor Serial Number": normalized.flowSensorSerial || "",
      "Flow Set Point %": normalized.flowSetpointPercent || 0,
      "Carbon Media Type": normalized.carbonMediaType || "",
      "Carbon Media SL Number": normalized.carbonMediaSlNumber || "",
      "Carbon Media Volume m3": normalized.carbonMediaVolume || 0,
      "Carbon PT Pressure Drop": normalized.carbonPressureDrop || 0,
      "Carbon PT High Setpoint": normalized.carbonHighSetpoint || 0,
      "Carbon PT High-High Setpoint": normalized.carbonHighHighSetpoint || 0,
      "Carbon Temp Probe Model": normalized.carbonTempProbeModel || "",
      "Carbon Temp Probe Serial": normalized.carbonTempProbeSerial || "",
      "Biofilter Media Type": normalized.bioMediaType || "",
      "Biofilter Media Volume m3": normalized.bioMediaVolume || 0
    };
    (assetCustomFields || []).forEach((field) => {
      row[`Extra: ${field.label}`] = normalized.customData?.[field.id] ?? "";
    });
    return row;
  })));
  const fieldRows = (assetCustomFields || []).map((field) => ({
    "Data Cell ID": field.id,
    "Data Cell Name": field.label,
    "Data Type": field.type || "text"
  }));
  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: "Greenacre Client Asset Database Export",
    Subject: "Client, site, asset and asset data export",
    Author: "Greenacre Client Asset Database",
    CreatedDate: new Date()
  };
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(assetRows.length ? assetRows : [{ "Client Name": "", "Site Reference": "", "Site Name": "", "Asset Reference": "", "Asset Name": "", "Asset Location": "" }]), "Assets");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(siteRows.length ? siteRows : [{ "Client Name": "", "Site Reference": "", "Site Name": "", "Site Address": "" }]), "Sites");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(clientRows.length ? clientRows : [{ "Client Name": "", "Billing Address": "" }]), "Clients");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(contactRows.length ? contactRows : [{ "Client Name": "", "Contact Name": "", "Role": "", "Email": "", "Phone": "" }]), "Contacts");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(fieldRows.length ? fieldRows : [{ "Data Cell ID": "", "Data Cell Name": "", "Data Type": "text" }]), "Asset Data Cells");
  const fileName = `greenacre-client-database-${todayIso()}.xlsx`;
  XLSX.writeFile(workbook, fileName, { compression: false });
  setSyncStatus("Excel downloaded", "ok", `${fileName} has been sent to your browser downloads.`);
}

async function importClientDatabaseExcel(file) {
  if (!window.XLSX) await loadExcelLibrary();
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
  const sheetRows = (name) => {
    const sheetName = workbook.SheetNames.find((item) => cleanHeader(item) === cleanHeader(name)) || workbook.SheetNames.find((item) => cleanHeader(item).includes(cleanHeader(name)));
    return sheetName ? XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" }) : [];
  };
  const clientsRows = sheetRows("Clients");
  const contactsRows = sheetRows("Contacts");
  const sitesRows = sheetRows("Sites");
  let assetsRows = sheetRows("Assets");
  if (!assetsRows.length && workbook.SheetNames.length) {
    assetsRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
  }
  const fieldRows = sheetRows("Asset Data Cells");

  fieldRows.forEach((row) => {
    const label = rowValue(row, ["Data Cell Name", "Field Name", "Name"]);
    if (!label) return;
    const id = rowValue(row, ["Data Cell ID", "Field ID", "ID"]) || fieldId(label);
    const type = ["number", "date"].includes(rowValue(row, ["Data Type", "Type"]).toLowerCase()) ? rowValue(row, ["Data Type", "Type"]).toLowerCase() : "text";
    if (!(assetCustomFields || []).some((field) => field.id === id || cleanHeader(field.label) === cleanHeader(label))) {
      assetCustomFields.push({ id, label, type, createdAt: new Date().toISOString(), createdBy: currentUser?.username || "excel-import" });
    }
  });

  clientsRows.forEach((row) => {
    findOrCreateClientByName(rowValue(row, ["Client Name", "Client", "Company"]), rowValue(row, ["Billing Address", "Client Address", "Address"]));
  });
  contactsRows.forEach((row) => {
    const client = findOrCreateClientByName(rowValue(row, ["Client Name", "Client", "Company"]));
    const name = rowValue(row, ["Contact Name", "Name"]);
    if (!name) return;
    const existing = (client.contacts || []).find((contact) => String(contact.name || "").trim().toLowerCase() === name.toLowerCase());
    const contact = existing || { id: newId("contact"), name, role: "", email: "", phone: "" };
    contact.role = rowValue(row, ["Role", "Contact Role"]) || contact.role || "";
    contact.email = rowValue(row, ["Email", "Contact Email"]) || contact.email || "";
    contact.phone = rowValue(row, ["Phone", "Telephone", "Contact Phone"]) || contact.phone || "";
    if (!existing) client.contacts.push(contact);
  });
  sitesRows.forEach((row) => {
    const client = findOrCreateClientByName(rowValue(row, ["Client Name", "Client", "Company"]));
    findOrCreateSite(client, rowValue(row, ["Site Reference", "Site Ref", "Site ID"]), rowValue(row, ["Site Name", "Site"]), rowValue(row, ["Site Address", "Address"]));
  });
  assetsRows.forEach((row) => {
    const client = findOrCreateClientByName(rowValue(row, ["Client Name", "Client", "Company"]));
    const site = findOrCreateSite(client, rowValue(row, ["Site Reference", "Site Ref", "Site ID"]), rowValue(row, ["Site Name", "Site"]), rowValue(row, ["Site Address", "Address"]));
    const asset = findOrCreateAsset(site, rowValue(row, ["Asset Reference", "Asset Ref", "Asset ID"]), rowValue(row, ["Asset Name", "Asset"]));
    asset.location = rowValue(row, ["Asset Location", "Location"]) || asset.location || "";
    applyRowNumber(asset, "flowRate", row, ["Flow Rate m3/h", "Flow Rate", "Flow"]);
    asset.fanModel = rowValue(row, ["Fan Model"]) || asset.fanModel || "";
    asset.fanSerial = rowValue(row, ["Fan Serial Number", "Fan Serial"]) || asset.fanSerial || "";
    asset.flowSensorModel = rowValue(row, ["Flow Sensor Model"]) || asset.flowSensorModel || "";
    asset.flowSensorSerial = rowValue(row, ["Flow Sensor Serial Number", "Flow Sensor Serial"]) || asset.flowSensorSerial || "";
    applyRowNumber(asset, "flowSetpointPercent", row, ["Flow Set Point %", "Flow Setpoint", "Flow Set Point"]);
    asset.carbonMediaType = rowValue(row, ["Carbon Media Type"]) || asset.carbonMediaType || "";
    asset.carbonMediaSlNumber = rowValue(row, ["Carbon Media SL Number", "Carbon Media SL"]) || asset.carbonMediaSlNumber || "";
    applyRowNumber(asset, "carbonMediaVolume", row, ["Carbon Media Volume m3", "Carbon Media Volume"]);
    applyRowNumber(asset, "carbonPressureDrop", row, ["Carbon PT Pressure Drop", "Carbon Pressure Drop"]);
    applyRowNumber(asset, "carbonHighSetpoint", row, ["Carbon PT High Setpoint", "Carbon High Setpoint"]);
    applyRowNumber(asset, "carbonHighHighSetpoint", row, ["Carbon PT High-High Setpoint", "Carbon High High Setpoint"]);
    asset.carbonTempProbeModel = rowValue(row, ["Carbon Temp Probe Model"]) || asset.carbonTempProbeModel || "";
    asset.carbonTempProbeSerial = rowValue(row, ["Carbon Temp Probe Serial"]) || asset.carbonTempProbeSerial || "";
    asset.bioMediaType = rowValue(row, ["Biofilter Media Type", "Bio Media Type"]) || asset.bioMediaType || "";
    applyRowNumber(asset, "bioMediaVolume", row, ["Biofilter Media Volume m3", "Bio Media Volume"]);
    asset.customData ||= {};
    Object.entries(row).forEach(([key, value]) => {
      if (!String(key).startsWith("Extra:")) return;
      const label = String(key).replace(/^Extra:\s*/, "").trim();
      const field = (assetCustomFields || []).find((item) => cleanHeader(item.label) === cleanHeader(label));
      if (field) asset.customData[field.id] = value;
    });
  });

  state.activeClientId = state.clients?.[0]?.id || "";
  consolidateProjectNumberFields();
  clientEditorOpen = false;
  clientDraft = null;
  clientDrillLevel = "register";
  activeClientSiteId = "";
  activeClientAssetId = "";
  update();
}

async function loadStateFromServer() {
  setSyncStatus("Server loading");
  await apiFetch(SERVER_STORAGE.healthUrl);
  const data = await apiFetch(SERVER_STORAGE.dataUrl);
  serverLoaded = true;
  serverVersion = Number(data.version || 0);
  assetCustomFields = Array.isArray(data.assetCustomFields) ? data.assetCustomFields : [];
  setSyncStatus("Server connected", "ok", "Netlify shared storage is active.");
  const loadedState = normalizeLockedLogic(mergeDefaults(clone(defaultState), data.state || {}));
  state = loadedState;
  consolidateProjectNumberFields();
  return state;
}

async function saveStateToServer() {
  if (serverSaving) {
    saveState();
    return;
  }
  serverSaving = true;
  setSyncStatus("Server saving");
  try {
    const data = await apiFetch(SERVER_STORAGE.dataUrl, {
      method: "PUT",
      body: JSON.stringify({ version: serverVersion, state: sharedStateSnapshot(), assetCustomFields })
    });
    serverVersion = Number(data.version || serverVersion + 1);
    assetCustomFields = Array.isArray(data.assetCustomFields) ? data.assetCustomFields : assetCustomFields;
    setSyncStatus("Server saved", "ok", "Netlify shared storage is active.");
  } catch (error) {
    if (String(error.message || "").includes("409")) {
      setSyncStatus("Save conflict - reloading latest server data", "error", "Another user saved first. The latest server data is being loaded.");
      state = await loadStateFromServer();
      update();
      return;
    }
    throw error;
  } finally {
    serverSaving = false;
  }
}

function renderCustomAssetFields(asset = {}) {
  const target = $("#assetCustomFieldsPanel");
  if (!target) return;
  const normalizedAsset = normalizeAsset(asset, 1);
  const standardFields = assetCustomFields || [];
  const localFields = normalizedAsset.localCustomFields || [];
  const fields = [...standardFields, ...localFields];
  target.innerHTML = `
    <div class="panel-head additional-data-head">
      <h3>Additional Asset Data</h3>
      ${isAdminUser() ? `<button id="addAssetFieldBtn" type="button">Add Data Cell</button>` : ""}
    </div>
    ${isAdminUser() ? `
      <div class="additional-data-admin">
        <label>Data cell name <input id="newAssetFieldLabel" placeholder="Example: Service interval"></label>
        <label>Data type
          <select id="newAssetFieldType">
            <option value="text">Text</option>
            <option value="number">Number</option>
            <option value="date">Date</option>
          </select>
        </label>
        <p id="assetFieldAdminMessage" class="admin-message"></p>
      </div>
    ` : ""}
    ${fields.length ? `
      <div class="custom-field-list">
        ${fields.map((field) => `
          <label class="custom-field-item">${escapeHtml(field.label)}
            <span class="custom-field-control">
              <input data-custom-asset-field="${escapeHtml(field.id)}" type="${field.type === "date" ? "date" : "text"}" value="${escapeHtml(normalizedAsset.customData?.[field.id] ?? "")}">
              ${isDueDateField(field) ? `<button data-create-reminder="${escapeHtml(field.id)}" type="button">Reminder</button>` : ""}
              ${isAdminUser() && field.local ? adminDeleteButton(`data-delete-local-asset-field="${escapeHtml(field.id)}"`, "Delete") : ""}
              ${isAdminUser() && !field.local ? adminDeleteButton(`data-clear-custom-asset-field="${escapeHtml(field.id)}"`, "Clear Value") : ""}
            </span>
          </label>
        `).join("")}
      </div>
    ` : `<p class="empty-additional-data">No additional asset data cells yet.</p>`}
  `;
  const adminPanel = $("#adminAssetFieldsPanel");
  if (adminPanel) {
    adminPanel.hidden = true;
  }
}

function updateCustomAssetField(target) {
  if (!target?.dataset?.customAssetField || !clientDraft) return;
  const site = clientDraft.sites?.find((item) => item.id === activeClientSiteId);
  const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
  if (!asset) return;
  asset.customData ||= {};
  asset.customData[target.dataset.customAssetField] = target.value;
  autoPopulateDueDateForAsset(asset);
}

function clearCustomAssetFieldValue(fieldIdToClear) {
  if (!clientDraft) return;
  const site = clientDraft.sites?.find((item) => item.id === activeClientSiteId);
  const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
  if (!asset?.customData) return;
  delete asset.customData[fieldIdToClear];
  renderCustomAssetFields(asset);
  saveState();
}

function deleteLocalAssetField(fieldIdToDelete) {
  if (!clientDraft) return;
  const site = clientDraft.sites?.find((item) => item.id === activeClientSiteId);
  const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
  if (!asset) return;
  asset.localCustomFields = (asset.localCustomFields || []).filter((field) => field.id !== fieldIdToDelete);
  if (asset.customData) delete asset.customData[fieldIdToDelete];
  renderCustomAssetFields(asset);
  saveState();
}

async function renderAssetAttachments(asset = {}) {
  const panel = $("#assetAttachmentsPanel");
  if (!panel) return;
  if (!asset?.id) {
    panel.innerHTML = "";
    return;
  }
  panel.innerHTML = `
    <div class="panel-head">
      <h3>Asset Photos & PDFs</h3>
      <div class="attachment-actions">
        <input id="assetAttachmentInput" type="file" accept="image/*,.pdf,application/pdf">
        <button id="uploadAssetAttachmentBtn" type="button">Upload File</button>
      </div>
    </div>
    <p class="attachment-message" id="assetAttachmentMessage">Loading files...</p>
    <div class="attachment-list" id="assetAttachmentList"></div>
  `;
  try {
    const data = await apiFetch(`/api/assets/${encodeURIComponent(asset.id)}/attachments`);
    const files = Array.isArray(data.files) ? data.files : [];
    const message = $("#assetAttachmentMessage");
    if (message) message.textContent = files.length ? "" : "No photos or PDFs attached.";
    const list = $("#assetAttachmentList");
    if (list) {
      list.innerHTML = files.map((file) => `
        <article class="attachment-row">
          <span>${escapeHtml(file.name || "Attachment")}</span>
          <span>${escapeHtml(file.uploadedAt ? new Date(file.uploadedAt).toLocaleString("en-GB") : "")}</span>
          <div class="attachment-row-actions">
            <a class="button-link" href="${escapeHtml(file.url)}" target="_blank" rel="noopener">Open</a>
            ${isAdminUser() ? adminDeleteButton(`data-delete-asset-attachment="${escapeHtml(file.id)}"`, "Delete") : ""}
          </div>
        </article>
      `).join("");
    }
  } catch (error) {
    const message = $("#assetAttachmentMessage");
    if (message) message.textContent = error.message || "Could not load files.";
  }
}

const ATTACHMENT_UPLOAD_TARGET_BYTES = 4.5 * 1024 * 1024;
const ATTACHMENT_IMAGE_MAX_EDGE = 2560;

function attachmentSizeLabel(bytes = 0) {
  return `${(Number(bytes || 0) / (1024 * 1024)).toFixed(1)} MB`;
}

async function imageDimensionsFromFile(file) {
  if (file.type === "image/jpeg") {
    const bytes = new Uint8Array(await file.slice(0, Math.min(file.size, 1024 * 1024)).arrayBuffer());
    if (bytes[0] === 0xff && bytes[1] === 0xd8) {
      let offset = 2;
      while (offset + 9 < bytes.length) {
        if (bytes[offset] !== 0xff) {
          offset += 1;
          continue;
        }
        const marker = bytes[offset + 1];
        const length = (bytes[offset + 2] << 8) + bytes[offset + 3];
        if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
          return {
            width: (bytes[offset + 7] << 8) + bytes[offset + 8],
            height: (bytes[offset + 5] << 8) + bytes[offset + 6]
          };
        }
        if (!length || length < 2) break;
        offset += length + 2;
      }
    }
  }

  return new Promise((resolve, reject) => {
    const image = new Image();
    const url = URL.createObjectURL(file);
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve({ image, width: image.naturalWidth, height: image.naturalHeight });
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("This phone image format cannot be resized. Save or share it as a JPEG, then upload it again."));
    };
    image.src = url;
  });
}

function canvasToJpeg(canvas, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("The photo could not be compressed on this device."));
    }, "image/jpeg", quality);
  });
}

async function prepareAssetAttachment(file, onProgress = () => {}) {
  if (!(file instanceof File)) throw new Error("Choose a photo or PDF to upload.");
  if (file.type === "application/pdf") {
    if (file.size > ATTACHMENT_UPLOAD_TARGET_BYTES) {
      throw new Error(`This PDF is ${attachmentSizeLabel(file.size)}. Please reduce it below 4.5 MB before uploading.`);
    }
    return file;
  }
  if (!String(file.type || "").startsWith("image/")) {
    throw new Error("Only photos and PDF files can be uploaded.");
  }
  if (file.size <= ATTACHMENT_UPLOAD_TARGET_BYTES) return file;

  onProgress(`Compressing ${attachmentSizeLabel(file.size)} photo...`);
  const source = await imageDimensionsFromFile(file);
  const scale = Math.min(1, ATTACHMENT_IMAGE_MAX_EDGE / Math.max(source.width, source.height));
  let width = Math.max(1, Math.round(source.width * scale));
  let height = Math.max(1, Math.round(source.height * scale));
  let drawable = source.image;
  if (typeof createImageBitmap === "function") {
    try {
      drawable = await createImageBitmap(file, {
        resizeWidth: width,
        resizeHeight: height,
        resizeQuality: "high",
        imageOrientation: "from-image"
      });
    } catch {
      drawable = source.image || await createImageBitmap(file);
    }
  }
  if (!drawable) {
    const decoded = await new Promise((resolve, reject) => {
      const image = new Image();
      const url = URL.createObjectURL(file);
      image.onload = () => {
        URL.revokeObjectURL(url);
        resolve(image);
      };
      image.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error("This phone image format cannot be resized. Save or share it as a JPEG, then upload it again."));
      };
      image.src = url;
    });
    drawable = decoded;
  }

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { alpha: false });
  if (!context) throw new Error("Photo compression is not supported by this browser.");
  context.fillStyle = "#fff";
  context.fillRect(0, 0, width, height);
  context.drawImage(drawable, 0, 0, width, height);
  drawable.close?.();

  let compressed = await canvasToJpeg(canvas, 0.82);
  if (compressed.size > ATTACHMENT_UPLOAD_TARGET_BYTES) compressed = await canvasToJpeg(canvas, 0.68);
  if (compressed.size > ATTACHMENT_UPLOAD_TARGET_BYTES) {
    width = Math.max(1, Math.round(width * 0.75));
    height = Math.max(1, Math.round(height * 0.75));
    const smaller = document.createElement("canvas");
    smaller.width = width;
    smaller.height = height;
    const smallerContext = smaller.getContext("2d", { alpha: false });
    smallerContext.fillStyle = "#fff";
    smallerContext.fillRect(0, 0, width, height);
    smallerContext.drawImage(canvas, 0, 0, width, height);
    compressed = await canvasToJpeg(smaller, 0.68);
  }
  if (compressed.size > ATTACHMENT_UPLOAD_TARGET_BYTES) {
    throw new Error("This photo is still too large after compression. Please use a lower-resolution copy.");
  }

  const baseName = String(file.name || "photo").replace(/\.[^.]+$/, "") || "photo";
  return new File([compressed], `${baseName}.jpg`, { type: "image/jpeg", lastModified: Date.now() });
}

function customFieldByLabel(asset, label) {
  const key = cleanHeader(label);
  return [...(assetCustomFields || []), ...(asset?.localCustomFields || [])].find((field) => cleanHeader(field.label) === key);
}

function parseFrequencyMonths(value) {
  const text = String(value || "").toLowerCase();
  const number = Number((text.match(/\d+(?:\.\d+)?/) || [0])[0]);
  if (!number) return 0;
  if (text.includes("year")) return Math.round(number * 12);
  if (text.includes("month")) return Math.round(number);
  if (text.includes("week")) return Math.round(number / 4.345);
  return Math.round(number * 12);
}

function isoDateInput(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function autoPopulateDueDateForAsset(asset) {
  if (!asset?.customData) return;
  const lastTestField = customFieldByLabel(asset, "Last Test");
  const dueDateField = customFieldByLabel(asset, "Due Date");
  const frequencyField = customFieldByLabel(asset, "Frequency");
  if (!lastTestField || !dueDateField || !frequencyField) return;
  const lastTest = parseAssetDate(asset.customData[lastTestField.id]);
  const months = parseFrequencyMonths(asset.customData[frequencyField.id]);
  if (!lastTest || !months) return;
  const dueDate = new Date(lastTest);
  dueDate.setMonth(dueDate.getMonth() + months);
  asset.customData[dueDateField.id] = isoDateInput(dueDate);
  const dueInput = document.querySelector(`[data-custom-asset-field="${CSS.escape(dueDateField.id)}"]`);
  if (dueInput) dueInput.value = asset.customData[dueDateField.id];
}

function parseAssetDate(value) {
  if (!value) return null;
  const text = String(value).trim();
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

function formatDateForEmail(date) {
  return date.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" });
}

function icsDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}`;
}

function escapeIcs(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/,/g, "\\,")
    .replace(/;/g, "\\;");
}

function projectNumbersForAsset(asset) {
  const field = (assetCustomFields || []).find(isProjectNumberField);
  return field ? String(asset.customData?.[field.id] || "").trim() : "";
}

function createDueDateReminder(fieldId) {
  updateClientDraftFromForm();
  const client = clientDraft || findClient(state.activeClientId);
  const site = client?.sites?.find((item) => item.id === activeClientSiteId);
  const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
  if (!client || !site || !asset) return;
  const dueDate = parseAssetDate(asset.customData?.[fieldId]);
  if (!dueDate) {
    alert("Enter a valid Due Date before creating a reminder.");
    return;
  }
  const reminderDate = new Date(dueDate);
  reminderDate.setDate(reminderDate.getDate() - 42);
  const projectNumbers = projectNumbersForAsset(asset) || "Not recorded";
  const subject = `Reminder: ${site.name || site.reference || "Site"} - ${asset.name || asset.reference || "Asset"} due ${formatDateForEmail(dueDate)}`;
  const bodyLines = [
    "Please arrange the upcoming asset review.",
    "",
    `Client: ${client.name || ""}`,
    `Site: ${site.name || ""}`,
    `Site reference: ${site.reference || ""}`,
    `Asset: ${asset.name || ""}`,
    `Asset reference: ${asset.reference || ""}`,
    `Asset location: ${asset.location || ""}`,
    `Due date: ${formatDateForEmail(dueDate)}`,
    `Reminder date: ${formatDateForEmail(reminderDate)}`,
    `Previous job/project numbers: ${projectNumbers}`
  ];
  const description = bodyLines.join("\\n");
  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Greenacre//Client Database//EN",
    "METHOD:REQUEST",
    "BEGIN:VEVENT",
    `UID:${newId("reminder")}@greenacre-client-database`,
    `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z")}`,
    `DTSTART;VALUE=DATE:${icsDate(dueDate)}`,
    `DTEND;VALUE=DATE:${icsDate(new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate() + 1))}`,
    `SUMMARY:${escapeIcs(subject)}`,
    `DESCRIPTION:${escapeIcs(description)}`,
    `LOCATION:${escapeIcs(site.name || site.reference || "")}`,
    `ORGANIZER;CN=Greenacre:mailto:${REMINDER_EMAIL}`,
    `ATTENDEE;CN=Georgia;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:${REMINDER_EMAIL}`,
    "END:VEVENT",
    "END:VCALENDAR"
  ].join("\r\n");
  const blob = new Blob([ics], { type: "text/calendar;charset=utf-8;method=REQUEST" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${(site.reference || site.name || "site").replace(/[^a-z0-9]+/gi, "-")}-${(asset.reference || asset.name || "asset").replace(/[^a-z0-9]+/gi, "-")}-due-reminder.ics`;
  link.click();
  URL.revokeObjectURL(url);
  window.location.href = `mailto:${encodeURIComponent(REMINDER_EMAIL)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join("\n\n"))}`;
}

async function createAssetCustomField() {
  const label = $("#newAssetFieldLabel")?.value.trim();
  const type = $("#newAssetFieldType")?.value || "text";
  const message = $("#assetFieldAdminMessage");
  if (!label) {
    if (message) message.textContent = "Enter a data cell name.";
    return;
  }
  updateClientDraftFromForm();
  const site = clientDraft?.sites?.find((item) => item.id === activeClientSiteId);
  const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
  if (!asset) {
    if (message) message.textContent = "Open an asset before adding a data cell.";
    return;
  }
  asset.localCustomFields ||= [];
  const existing = [...(assetCustomFields || []), ...asset.localCustomFields].some((field) => cleanHeader(field.label) === cleanHeader(label));
  if (existing) {
    if (message) message.textContent = "That data cell already exists on this asset.";
    return;
  }
  const field = { id: `${fieldId(label)}_${Date.now()}`, label, type: type === "date" ? "date" : "text", local: true, createdAt: new Date().toISOString(), createdBy: currentUser?.username || "admin" };
  asset.localCustomFields.push(field);
  asset.customData ||= {};
  asset.customData[field.id] = "";
  if ($("#newAssetFieldLabel")) $("#newAssetFieldLabel").value = "";
  if (message) message.textContent = "Data cell added to this asset only.";
  renderCustomAssetFields(asset);
  saveState();
}

async function removeAssetCustomField(fieldIdToRemove) {
  const message = $("#assetFieldAdminMessage");
  try {
    const data = await apiFetch(`${SERVER_STORAGE.assetFieldsUrl}/${encodeURIComponent(fieldIdToRemove)}`, {
      method: "DELETE"
    });
    assetCustomFields = data.assetCustomFields || [];
    for (const client of state.clients || []) {
      for (const site of client.sites || []) {
        for (const asset of site.assets || []) {
          if (asset.customData) delete asset.customData[fieldIdToRemove];
        }
      }
    }
    if (message) message.textContent = "Data cell removed from all assets.";
    updateClientDraftFromForm();
    renderCustomAssetFields(findAsset(clientDraft?.sites?.find((site) => site.id === activeClientSiteId), activeClientAssetId) || {});
    saveState();
  } catch (error) {
    if (message) message.textContent = error.message || "Could not remove data cell.";
  }
}
function sharePointBaseUrl() {
  const context = window._spPageContextInfo;
  if (context?.webAbsoluteUrl) return context.webAbsoluteUrl.replace(/\/$/, "");
  const marker = "/SiteAssets/";
  const href = window.location.href;
  const markerIndex = href.toLowerCase().indexOf(marker.toLowerCase());
  if (markerIndex > -1) return href.slice(0, markerIndex);
  const pathParts = window.location.pathname.split("/").filter(Boolean);
  const managedPathIndex = pathParts.findIndex((part) => ["sites", "teams"].includes(part.toLowerCase()));
  if (managedPathIndex > -1 && pathParts[managedPathIndex + 1]) {
    return `${window.location.origin}/${pathParts[managedPathIndex]}/${pathParts[managedPathIndex + 1]}`;
  }
  return window.location.origin;
}

function sharePointListUrl(listKey) {
  const title = encodeURIComponent(SHAREPOINT_LISTS[listKey].replace(/'/g, "''"));
  return `${sharePointBaseUrl()}/_api/web/lists/getbytitle('${title}')/items`;
}

function sharePointFieldsUrl(listKey) {
  const title = encodeURIComponent(SHAREPOINT_LISTS[listKey].replace(/'/g, "''"));
  return `${sharePointBaseUrl()}/_api/web/lists/getbytitle('${title}')/fields`;
}

function setSyncStatus(message, status = "", detail = "") {
  const statusEl = $("#syncStatus");
  if (statusEl) {
    statusEl.textContent = message;
    statusEl.classList.toggle("ok", status === "ok");
    statusEl.classList.toggle("error", status === "error");
    if (detail) statusEl.title = detail;
  }
  const detailEl = $("#syncStatusDetails");
  if (detailEl) detailEl.textContent = detail;
}

async function sharePointRequest(url, options = {}) {
  const response = await fetch(url, {
    credentials: "same-origin",
    headers: {
      Accept: "application/json;odata=nometadata",
      ...(options.body ? { "Content-Type": "application/json;odata=nometadata" } : {}),
      ...(options.headers || {})
    },
    ...options
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => response.statusText);
    throw new Error(`SharePoint request failed (${response.status}) at ${url}: ${detail}`);
  }
  if (response.status === 204) return null;
  return response.json();
}

async function sharePointDigest() {
  const data = await sharePointRequest(`${sharePointBaseUrl()}/_api/contextinfo`, {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  return data?.FormDigestValue;
}

async function loadSharePointFieldsForList(listKey, requiredTitles, optionalTitles = []) {
  const titles = [...requiredTitles, ...optionalTitles];
  const filter = encodeURIComponent(titles.map((title) => `Title eq '${title.replace(/'/g, "''")}'`).join(" or "));
  const data = await sharePointRequest(`${sharePointFieldsUrl(listKey)}?$select=Title,InternalName&$filter=${filter}`);
  const fields = Object.fromEntries((data?.value || []).map((field) => [field.Title, field.InternalName]));
  const missing = requiredTitles.filter((title) => !fields[title]);
  if (missing.length) {
    throw new Error(`The ${SHAREPOINT_LISTS[listKey]} list needs columns named ${missing.join(", ")}.`);
  }
  sharePointFields[listKey] = fields;
  return fields;
}

async function loadSharePointFields() {
  await loadSharePointFieldsForList("clients", ["Title"], ["BillingAddress", "ClientNotes"]);
  await loadSharePointFieldsForList("contacts", ["Title", "Client"], ["Role", "Email", "Phone"]);
  await loadSharePointFieldsForList("sites", ["Title", "Client"], ["SiteReference", "SiteAddress"]);
  await loadSharePointFieldsForList("assets", ["Title", "Site"], [
    "AssetReference", "AssetLocation", "FlowRate", "FanModel", "FanSerial", "FlowSensorModel", "FlowSensorSerial", "FlowSetpointPercent",
    "CarbonMediaType", "CarbonMediaSlNumber", "CarbonMediaVolume", "CarbonPressureDrop", "CarbonHighSetpoint", "CarbonHighHighSetpoint",
    "CarbonTempProbeModel", "CarbonTempProbeSerial", "BioMediaType", "BioMediaVolume"
  ]);
}

function fieldValue(item, fields, title, fallback = "") {
  const internal = fields[title];
  return internal ? item[internal] ?? fallback : fallback;
}

function numberFieldValue(item, fields, title) {
  const value = fieldValue(item, fields, title, 0);
  return value === null || value === undefined || value === "" ? 0 : Number(value);
}

function lookupId(item, fields, title) {
  const internal = fields[title];
  return internal ? item[internal]?.Id || item[`${internal}Id`] || null : null;
}

function optionalSelect(fields, titles) {
  return titles.map((title) => fields[title]).filter(Boolean);
}

async function loadListItems(listKey, selectFields, expandFields = []) {
  const select = encodeURIComponent(selectFields.join(","));
  const expand = expandFields.length ? `&$expand=${encodeURIComponent(expandFields.join(","))}` : "";
  return sharePointRequest(`${sharePointListUrl(listKey)}?$select=${select}${expand}&$top=5000`);
}

async function loadStateFromSharePoint() {
  setSyncStatus("SharePoint loading");
  await loadSharePointFields();

  const clientFields = sharePointFields.clients;
  const contactFields = sharePointFields.contacts;
  const siteFields = sharePointFields.sites;
  const assetFields = sharePointFields.assets;
  const clientsBySpId = new Map();
  const sitesBySpId = new Map();
  sharePointItems = { clients: new Map(), contacts: new Map(), sites: new Map(), assets: new Map() };

  const clientData = await loadListItems("clients", ["Id", "Title", ...optionalSelect(clientFields, ["BillingAddress", "ClientNotes"])]);
  const clients = (clientData?.value || []).map((item) => {
    const id = `sp-client-${item.Id}`;
    const client = {
      id,
      name: item.Title || "",
      address: fieldValue(item, clientFields, "BillingAddress", ""),
      notes: fieldValue(item, clientFields, "ClientNotes", ""),
      contacts: [],
      sites: []
    };
    clientsBySpId.set(item.Id, client);
    sharePointItems.clients.set(id, { id: item.Id });
    return client;
  });

  const contactLookup = contactFields.Client;
  const contactData = await loadListItems("contacts", ["Id", "Title", `${contactLookup}/Id`, ...optionalSelect(contactFields, ["Role", "Email", "Phone"])], [contactLookup]);
  (contactData?.value || []).forEach((item) => {
    const client = clientsBySpId.get(lookupId(item, contactFields, "Client"));
    if (!client) return;
    const id = `sp-contact-${item.Id}`;
    client.contacts.push({
      id,
      name: item.Title || "",
      role: fieldValue(item, contactFields, "Role", ""),
      email: fieldValue(item, contactFields, "Email", ""),
      phone: fieldValue(item, contactFields, "Phone", "")
    });
    sharePointItems.contacts.set(id, { id: item.Id });
  });

  const siteLookup = siteFields.Client;
  const siteData = await loadListItems("sites", ["Id", "Title", `${siteLookup}/Id`, ...optionalSelect(siteFields, ["SiteReference", "SiteAddress"])], [siteLookup]);
  (siteData?.value || []).forEach((item) => {
    const client = clientsBySpId.get(lookupId(item, siteFields, "Client"));
    if (!client) return;
    const id = `sp-site-${item.Id}`;
    const site = {
      id,
      reference: fieldValue(item, siteFields, "SiteReference", ""),
      name: item.Title || "",
      address: fieldValue(item, siteFields, "SiteAddress", ""),
      assets: []
    };
    client.sites.push(site);
    sitesBySpId.set(item.Id, site);
    sharePointItems.sites.set(id, { id: item.Id });
  });

  const assetLookup = assetFields.Site;
  const assetData = await loadListItems("assets", ["Id", "Title", `${assetLookup}/Id`, ...optionalSelect(assetFields, [
    "AssetReference", "AssetLocation", "FlowRate", "FanModel", "FanSerial", "FlowSensorModel", "FlowSensorSerial", "FlowSetpointPercent",
    "CarbonMediaType", "CarbonMediaSlNumber", "CarbonMediaVolume", "CarbonPressureDrop", "CarbonHighSetpoint", "CarbonHighHighSetpoint",
    "CarbonTempProbeModel", "CarbonTempProbeSerial", "BioMediaType", "BioMediaVolume"
  ])], [assetLookup]);
  (assetData?.value || []).forEach((item) => {
    const site = sitesBySpId.get(lookupId(item, assetFields, "Site"));
    if (!site) return;
    const id = `sp-asset-${item.Id}`;
    site.assets.push(normalizeAsset({
      id,
      reference: fieldValue(item, assetFields, "AssetReference", ""),
      name: item.Title || "",
      location: fieldValue(item, assetFields, "AssetLocation", ""),
      flowRate: numberFieldValue(item, assetFields, "FlowRate"),
      fanModel: fieldValue(item, assetFields, "FanModel", ""),
      fanSerial: fieldValue(item, assetFields, "FanSerial", ""),
      flowSensorModel: fieldValue(item, assetFields, "FlowSensorModel", ""),
      flowSensorSerial: fieldValue(item, assetFields, "FlowSensorSerial", ""),
      flowSetpointPercent: numberFieldValue(item, assetFields, "FlowSetpointPercent"),
      carbonMediaType: fieldValue(item, assetFields, "CarbonMediaType", ""),
      carbonMediaSlNumber: fieldValue(item, assetFields, "CarbonMediaSlNumber", ""),
      carbonMediaVolume: numberFieldValue(item, assetFields, "CarbonMediaVolume"),
      carbonPressureDrop: numberFieldValue(item, assetFields, "CarbonPressureDrop"),
      carbonHighSetpoint: numberFieldValue(item, assetFields, "CarbonHighSetpoint"),
        carbonHighHighSetpoint: fieldValue(item, assetFields, "CarbonHighHighSetpoint", ""),
      carbonTempProbeModel: fieldValue(item, assetFields, "CarbonTempProbeModel", ""),
      carbonTempProbeSerial: fieldValue(item, assetFields, "CarbonTempProbeSerial", ""),
      bioMediaType: fieldValue(item, assetFields, "BioMediaType", ""),
      bioMediaVolume: numberFieldValue(item, assetFields, "BioMediaVolume")
    }, site.assets.length + 1));
    sharePointItems.assets.set(id, { id: item.Id });
  });

  const sharedState = mergeDefaults(clone(defaultState), { clients });
  sharePointLoaded = true;
  setSyncStatus(clients.length ? "SharePoint loaded" : "SharePoint ready", "ok", "Connected to GCE_Clients, GCE_ClientContacts, GCE_Sites and GCE_Assets.");
  return normalizeLockedLogic(sharedState);
}

function loadStateFromLocalStorage(error) {
  try {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
    const fallbackState = saved ? mergeDefaults(clone(defaultState), JSON.parse(saved)) : clone(defaultState);
    setSyncStatus(
      "Using local browser storage",
      "error",
      error ? `${error.message || error}. Netlify storage unavailable; changes are saved only in this browser.` : "Changes are saved only in this browser."
    );
    return normalizeLockedLogic(fallbackState);
  } catch (localError) {
    console.error(localError);
    setSyncStatus("Using temporary local data", "error", "Local browser storage could not be read. Export your data before closing.");
    return normalizeLockedLogic(clone(defaultState));
  }
}

function localStorageSnapshot() {
  const {
    products,
    suppliers,
    ...workingState
  } = state;
  return workingState;
}

function saveStateToLocalStorage() {
  try {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(localStorageSnapshot()));
    if (!serverLoaded) {
      setSyncStatus("Saved locally only", "error", "Netlify storage is not connected, so changes are saved only in this browser.");
    }
  } catch (error) {
    console.error(error);
    setSyncStatus("Local save failed", "error", "Browser local storage could not save. Use Export to keep a copy.");
  }
}
function mergeDefaults(base, incoming) {
  for (const [key, value] of Object.entries(incoming || {})) {
    if (value && typeof value === "object" && !Array.isArray(value) && base[key]) {
      mergeDefaults(base[key], value);
    } else {
      base[key] = value;
    }
  }
  return base;
}

function saveState() {
  saveStateToLocalStorage();
  if (serverLoaded) {
    clearTimeout(serverSaveTimer);
    serverSaveTimer = setTimeout(() => {
      saveStateToServer().catch((error) => {
        console.error(error);
        setSyncStatus("Server save failed", "error", `${error.message || error}. Your changes have not been saved to Netlify.`);
      });
    }, 500);
    return;
  }
}

function lookupFieldBody(fields, title, lookupItemId) {
  const internal = fields[title];
  return internal && lookupItemId ? { [`${internal}Id`]: lookupItemId } : {};
}

function optionalBody(fields, title, value) {
  const internal = fields[title];
  return internal ? { [internal]: value ?? "" } : {};
}

function optionalNumberBody(fields, title, value) {
  const internal = fields[title];
  const numeric = value === "" || value === null || value === undefined ? null : Number(value);
  return internal ? { [internal]: Number.isFinite(numeric) ? numeric : null } : {};
}

async function upsertSharePointItem(listKey, localId, body, digest) {
  const item = sharePointItems[listKey].get(localId);
  if (item?.id) {
    await sharePointRequest(`${sharePointListUrl(listKey)}(${item.id})`, {
      method: "POST",
      headers: {
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
        "X-RequestDigest": digest
      },
      body: JSON.stringify(body)
    });
    return item.id;
  }
  const created = await sharePointRequest(sharePointListUrl(listKey), {
    method: "POST",
    headers: { "X-RequestDigest": digest },
    body: JSON.stringify(body)
  });
  sharePointItems[listKey].set(localId, { id: created.Id });
  return created.Id;
}

async function deleteRemovedSharePointItems(listKey, activeIds, digest) {
  for (const [localId, item] of [...sharePointItems[listKey].entries()]) {
    if (activeIds.has(localId)) continue;
    await sharePointRequest(`${sharePointListUrl(listKey)}(${item.id})`, {
      method: "POST",
      headers: {
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
        "X-RequestDigest": digest
      }
    });
    sharePointItems[listKey].delete(localId);
  }
}

async function saveStateToSharePoint() {
  if (sharePointSaving) {
    saveState();
    return;
  }
  sharePointSaving = true;
  setSyncStatus("SharePoint saving");
  try {
    const digest = await sharePointDigest();
    const active = {
      clients: new Set(),
      contacts: new Set(),
      sites: new Set(),
      assets: new Set()
    };

    for (const client of state.clients || []) {
      if (!client.id) continue;
      active.clients.add(client.id);
      await upsertSharePointItem("clients", client.id, {
        Title: client.name || client.company || client.id,
        ...optionalBody(sharePointFields.clients, "BillingAddress", client.address || ""),
        ...optionalBody(sharePointFields.clients, "ClientNotes", client.notes || "")
      }, digest);
    }

    for (const client of state.clients || []) {
      const clientItem = sharePointItems.clients.get(client.id);
      if (!clientItem?.id) continue;
      for (const contact of client.contacts || []) {
        if (!contact.id) continue;
        active.contacts.add(contact.id);
        await upsertSharePointItem("contacts", contact.id, {
          Title: contact.name || contact.email || contact.phone || contact.id,
          ...lookupFieldBody(sharePointFields.contacts, "Client", clientItem.id),
          ...optionalBody(sharePointFields.contacts, "Role", contact.role || ""),
          ...optionalBody(sharePointFields.contacts, "Email", contact.email || ""),
          ...optionalBody(sharePointFields.contacts, "Phone", contact.phone || "")
        }, digest);
      }
    }

    for (const client of state.clients || []) {
      const clientItem = sharePointItems.clients.get(client.id);
      if (!clientItem?.id) continue;
      for (const site of client.sites || []) {
        if (!site.id) continue;
        active.sites.add(site.id);
        await upsertSharePointItem("sites", site.id, {
          Title: site.name || site.reference || site.id,
          ...lookupFieldBody(sharePointFields.sites, "Client", clientItem.id),
          ...optionalBody(sharePointFields.sites, "SiteReference", site.reference || ""),
          ...optionalBody(sharePointFields.sites, "SiteAddress", site.address || "")
        }, digest);
      }
    }

    for (const client of state.clients || []) {
      for (const site of client.sites || []) {
        const siteItem = sharePointItems.sites.get(site.id);
        if (!siteItem?.id) continue;
        for (const assetRaw of site.assets || []) {
          const asset = normalizeAsset(assetRaw, 1);
          if (!asset.id) continue;
          active.assets.add(asset.id);
          await upsertSharePointItem("assets", asset.id, {
            Title: asset.name || asset.reference || asset.id,
            ...lookupFieldBody(sharePointFields.assets, "Site", siteItem.id),
            ...optionalBody(sharePointFields.assets, "AssetReference", asset.reference || ""),
            ...optionalBody(sharePointFields.assets, "AssetLocation", asset.location || ""),
            ...optionalNumberBody(sharePointFields.assets, "FlowRate", asset.flowRate),
            ...optionalBody(sharePointFields.assets, "FanModel", asset.fanModel || ""),
            ...optionalBody(sharePointFields.assets, "FanSerial", asset.fanSerial || ""),
            ...optionalBody(sharePointFields.assets, "FlowSensorModel", asset.flowSensorModel || ""),
            ...optionalBody(sharePointFields.assets, "FlowSensorSerial", asset.flowSensorSerial || ""),
            ...optionalNumberBody(sharePointFields.assets, "FlowSetpointPercent", asset.flowSetpointPercent),
            ...optionalBody(sharePointFields.assets, "CarbonMediaType", asset.carbonMediaType || ""),
            ...optionalBody(sharePointFields.assets, "CarbonMediaSlNumber", asset.carbonMediaSlNumber || ""),
            ...optionalNumberBody(sharePointFields.assets, "CarbonMediaVolume", asset.carbonMediaVolume),
            ...optionalNumberBody(sharePointFields.assets, "CarbonPressureDrop", asset.carbonPressureDrop),
            ...optionalNumberBody(sharePointFields.assets, "CarbonHighSetpoint", asset.carbonHighSetpoint),
            ...optionalBody(sharePointFields.assets, "CarbonHighHighSetpoint", asset.carbonHighHighSetpoint || ""),
            ...optionalBody(sharePointFields.assets, "CarbonTempProbeModel", asset.carbonTempProbeModel || ""),
            ...optionalBody(sharePointFields.assets, "CarbonTempProbeSerial", asset.carbonTempProbeSerial || ""),
            ...optionalBody(sharePointFields.assets, "BioMediaType", asset.bioMediaType || ""),
            ...optionalNumberBody(sharePointFields.assets, "BioMediaVolume", asset.bioMediaVolume)
          }, digest);
        }
      }
    }

    await deleteRemovedSharePointItems("assets", active.assets, digest);
    await deleteRemovedSharePointItems("contacts", active.contacts, digest);
    await deleteRemovedSharePointItems("sites", active.sites, digest);
    await deleteRemovedSharePointItems("clients", active.clients, digest);
    setSyncStatus("SharePoint saved", "ok", "Saved to GCE_Clients, GCE_ClientContacts, GCE_Sites and GCE_Assets.");
  } finally {
    sharePointSaving = false;
  }
}
function normalizeLockedLogic(target = state) {
  target.clients ||= [];
  target.enquiries ||= [];
  target.quotes ||= [];
  target.quotes.forEach((quote) => {
    quote.offerText ||= "";
    quote.exclusions ||= "";
    quote.lines ||= [];
  });
  target.products = [];
  target.suppliers ||= [];
  target.suppliers.forEach((supplier) => {
    supplier.contacts ||= [];
    supplier.sites ||= [];
  });
  target.clients.forEach((client) => {
    client.contacts ||= [];
    client.sites ||= [];
    client.sites.forEach((site) => {
      site.assets ||= [];
      site.assets = site.assets.map((asset, index) => normalizeAsset(asset, index + 1));
    });
  });
  target.gas = clone(defaultState.gas);
  target.filters.carbon.depth = defaultState.filters.carbon.depth;
  target.filters.carbon.pellet = defaultState.filters.carbon.pellet;
  target.filters.carbon.voidFraction = defaultState.filters.carbon.voidFraction;
  target.filters.bio.pellet = defaultState.filters.bio.pellet;
  target.filters.bio.voidFraction = defaultState.filters.bio.voidFraction;
  return target;
}

function productKey(product) {
  const code = String(product.productCode || "").trim().toLowerCase();
  if (code) return `code:${code}`;
  return [
    product.description,
    product.size,
    product.category,
    product.supplier,
    product.material
  ].map((value) => String(value || "").trim().toLowerCase()).join("|");
}

function mergeProductDatabase(existingProducts = []) {
  const existingByKey = new Map((existingProducts || []).map((product) => [productKey(product), product]));
  const merged = defaultProducts.map((product) => {
    const existing = existingByKey.get(productKey(product));
    return existing ? { ...product, ...existing } : clone(product);
  });
  const mergedKeys = new Set(merged.map(productKey));
  (existingProducts || []).forEach((product) => {
    if (!mergedKeys.has(productKey(product))) merged.push(product);
  });
  return merged;
}

function ensureProductSuppliers(target) {
  const existingKeys = new Set((target.suppliers || []).map((supplier) => supplierIdFromName(supplier.name)));
  buildSuppliersFromProducts(target.products || defaultProducts).forEach((supplier) => {
    if (!existingKeys.has(supplier.id)) {
      target.suppliers.push(supplier);
      existingKeys.add(supplier.id);
    }
  });
}

function ensureTestQuoteData(target) {
  const testClient = target.clients.find((client) => client.id === "test-client-001");
  if (!testClient) return;

  const defaultTestClient = defaultState.clients.find((client) => client.id === "test-client-001");
  const testContacts = defaultTestClient?.contacts || [];
  testContacts.forEach((contact) => {
    if (!testClient.contacts.some((item) => item.id === contact.id)) {
      testClient.contacts.push(clone(contact));
    }
  });
  (defaultTestClient?.sites || []).forEach((site) => {
    const existingSite = testClient.sites.find((item) => item.id === site.id);
    if (!existingSite) {
      testClient.sites.push(clone(site));
      return;
    }
    existingSite.assets ||= [];
    (site.assets || []).forEach((asset) => {
      if (!existingSite.assets.some((item) => item.id === asset.id)) {
        existingSite.assets.push(clone(asset));
      }
    });
  });

  const testEnquiries = defaultState.enquiries.filter((enquiry) => enquiry.id.startsWith("test-enquiry-"));
  testEnquiries.forEach((enquiry) => {
    if (!target.enquiries.some((item) => item.id === enquiry.id)) {
      target.enquiries.push(clone(enquiry));
    }
  });

  const testQuotes = defaultState.quotes.filter((quote) => quote.id.startsWith("test-quote-"));
  testQuotes.forEach((quote) => {
    if (!target.quotes.some((item) => item.id === quote.id)) {
      target.quotes.push(clone(quote));
    }
  });
}

function clearUiOnlyInputs() {
  ["siteSearch", "clientSearchSite", "assetRegisterSearchSiteRef", "assetRegisterSearchSite", "productSearch", "importInput", "productFilterSize", "productFilterSupplier", "productFilterMaterial", "productFilterCategory", "productFilterMinPrice", "productFilterMaxPrice", "priceMatchSupplier", "priceMatchMaterial", "priceMatchCategory", "pricePercent", "newProductCode", "newProductDescription", "newProductSize", "newProductPrice", "newProductCategory", "newProductSupplier", "newProductMaterial"].forEach((id) => {
    const input = $(`#${id}`);
    if (input) input.value = "";
  });
}

function setActiveTab(tabName) {
  tabName = "clients";
  document.querySelectorAll(".tabs button").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === tabName);
  });
  const summary = $(".summary-strip");
  if (summary) summary.classList.remove("visible");
}

function showClientRegisterHome() {
  if (clientDraft) {
    updateClientDraftFromForm();
    const existingClient = (state.clients || []).some((client) => client.id === clientDraft.id);
    const hasUsefulDraft = Boolean(
      String(clientDraft.name || "").trim() ||
      String(clientDraft.address || "").trim() ||
      (clientDraft.contacts || []).length ||
      (clientDraft.sites || []).length
    );
    if (existingClient || hasUsefulDraft) {
      saveClientDraft();
    }
  }
  clientDraft = null;
  clientEditorOpen = false;
  clientDrillLevel = "register";
  state.activeClientId = "";
  activeClientSiteId = "";
  activeClientAssetId = "";
  document.querySelectorAll(".modal-backdrop").forEach((modal) => modal.remove());
}

function activeTabName() {
  const activePanel = document.querySelector(".tab-panel.active");
  if (activePanel?.id === "design-workflow") {
    return document.querySelector(".design-panel.active")?.id || "zones";
  }
  return activePanel?.id || "design";
}

function blankZone(index = 1) {
  return {
    name: `Zone ${index}`,
    shape: "Cylinder",
    length: 0,
    width: 0,
    height: 0,
    diameter: 0,
    volumeOverride: 0,
    ach: 0,
    designVelocity: 0,
    ductDiameter: 0
  };
}

function blankLeg(index = 1) {
  return {
    name: `L${index}`,
    share: 0,
    material: "GRP Ductwork",
    length: 0,
    diameter: 0,
    roughness: 0,
    sr90: 0,
    lr90: 0,
    bend45: 0,
    tee: 0,
    reducer: 0,
    nrd: 0,
    vcd: 0,
    sb: 0,
    shoe: 0,
    entrance: "None",
    exit: "None",
    customK: 0
  };
}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function newId(prefix) {
  const uuid = globalThis.crypto?.randomUUID?.();
  return uuid || `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function nextSequence(prefix, records, field) {
  const year = new Date().getFullYear().toString().slice(-2);
  const base = `${prefix}${year}`;
  const max = records.reduce((highest, record) => {
    const value = String(record[field] || "");
    if (!value.startsWith(base)) return highest;
    const numeric = Number(value.slice(base.length).replace(/\D/g, ""));
    return Number.isFinite(numeric) ? Math.max(highest, numeric) : highest;
  }, 0);
  return `${base}${String(max + 1).padStart(4, "0")}`;
}

function quoteYearSuffix(date = todayIso()) {
  const year = new Date(date).getFullYear();
  return Number.isFinite(year) ? String(year).slice(-2) : new Date().getFullYear().toString().slice(-2);
}

function nextQuoteNumber(date = todayIso()) {
  const suffix = quoteYearSuffix(date);
  const max = state.quotes.reduce((highest, quote) => {
    const match = String(quote.number || "").match(/^W(\d+)-(\d{2})$/i);
    if (!match || match[2] !== suffix) return highest;
    return Math.max(highest, Number(match[1]) || highest);
  }, QUOTE_NUMBER_START);
  return `W${max + 1}-${suffix}`;
}

function nextParentProductCode(date = todayIso()) {
  const suffix = quoteYearSuffix(date);
  const pattern = new RegExp(`^GRE-ASM-${suffix}-(\\d{6})$`, "i");
  const max = (state.products || []).reduce((highest, product) => {
    const match = String(product.productCode || "").match(pattern);
    return match ? Math.max(highest, Number(match[1]) || 0) : highest;
  }, 0);
  return `GRE-ASM-${suffix}-${String(max + 1).padStart(6, "0")}`;
}

function blankClient() {
  const id = newId("client");
  return {
    id,
    name: "",
    address: "",
    contacts: [],
    sites: []
  };
}

function blankClientDraft() {
  return blankClient();
}

function blankSupplier() {
  return {
    id: newId("supplier"),
    name: "",
    invoiceAddress: "",
    phone: "",
    email: "",
    notes: "",
    contacts: [],
    sites: []
  };
}

function blankSupplierSite(index = 1) {
  return { id: newId("supplier-site"), name: `Site ${index}`, address: "", phone: "", email: "" };
}

function supplierIdFromName(name) {
  return `supplier-${String(name || "unnamed").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "unnamed"}`;
}

function blankSupplierFromName(name) {
  return {
    id: supplierIdFromName(name),
    name,
    invoiceAddress: "",
    phone: "",
    email: "",
    notes: "",
    contacts: [],
    sites: []
  };
}

function buildSuppliersFromProducts(products = []) {
  return Array.from(new Set(products.map((product) => String(product.supplier || "").trim()).filter(Boolean)))
    .sort((a, b) => a.localeCompare(b))
    .map(blankSupplierFromName);
}

function blankSite(index = 1) {
  return { id: newId("site"), reference: "", name: `Site ${index}`, address: "", assets: [] };
}

function blankAsset(index = 1) {
  return {
    id: newId("asset"),
    reference: "",
    name: `Asset ${index}`,
    location: "",
    flowRate: 0,
    fanModel: "",
    fanSerial: "",
    flowSensorModel: "",
    flowSensorSerial: "",
    flowSetpointPercent: 0,
    carbonMediaType: "",
    carbonMediaSlNumber: "",
    carbonMediaVolume: 0,
    carbonPressureDrop: 0,
    carbonHighSetpoint: 0,
    carbonHighHighSetpoint: 0,
    carbonTempProbeModel: "",
    carbonTempProbeSerial: "",
    bioMediaType: "",
    bioMediaVolume: 0,
    customData: {}
  };
}


function normalizeAsset(asset, index = 1) {
  const normalized = { ...blankAsset(index), ...(asset || {}) };
  normalized.customData ||= {};
  (assetCustomFields || []).forEach((field) => {
    if (normalized.customData[field.id] === undefined) normalized.customData[field.id] = "";
  });
  return normalized;
}

function blankEnquiry() {
  const id = newId("enquiry");
  return {
    id,
    number: nextSequence("ENQ", state.enquiries || [], "number"),
    status: "New",
    date: todayIso(),
    clientId: "",
    contactId: "",
    siteId: "",
    assetId: "",
    siteRef: "",
    assetRef: "",
    notes: ""
  };
}

function blankQuote(enquiry) {
  const related = state.quotes.filter((quote) => quote.enquiryId === enquiry?.id);
  const revision = related.length;
  const number = related[0]?.number || nextQuoteNumber(todayIso());
  return {
    id: newId("quote"),
    number,
    revision,
    status: "Draft",
    date: todayIso(),
    enquiryId: enquiry?.id || "",
    offerText: "",
    notes: "",
    exclusions: "",
    lines: [],
    total: currentResults?.bom?.sell || 0
  };
}
function activeEnquiry() {
  return state.enquiries.find((enquiry) => enquiry.id === state.activeEnquiryId) || null;
}

function activeQuote() {
  return state.quotes.find((quote) => quote.id === state.activeQuoteId) || null;
}

function findClient(id) {
  return state.clients.find((client) => client.id === id) || null;
}

function findContact(client, id) {
  return client?.contacts?.find((contact) => contact.id === id) || null;
}

function findSite(client, id) {
  return client?.sites?.find((site) => site.id === id) || null;
}

function findAsset(site, id) {
  return site?.assets?.find((asset) => asset.id === id) || null;
}

function editableClient() {
  const clientId = state.activeClientId || newId("client");
  let client = findClient(clientId);
  if (!client) {
    client = { id: clientId, name: "", address: "", contacts: [], sites: [] };
    state.clients.push(client);
    state.activeClientId = clientId;
  }
  client.contacts ||= [];
  client.sites ||= [];
  return client;
}

function syncProjectFromEnquiry() {
  const enquiry = activeEnquiry();
  if (!enquiry) return;
  const client = findClient(enquiry.clientId);
  const site = findSite(client, enquiry.siteId);
  const asset = findAsset(site, enquiry.assetId);
  state.project.reference = enquiry.number || "";
  state.project.customer = client?.name || "";
  state.project.site = site?.name || enquiry.siteRef || "";
  state.project.item = asset?.name || enquiry.assetRef || "";
}

function resetOpenTab() {
  const tabName = activeTabName();

  if (tabName === "design") {
    state.safety = { flow: 0, carbon: 0, bio: 0 };
    state.gas = clone(defaultState.gas);
    state.filters = {
      carbon: { ...clone(defaultState.filters.carbon), count: 0, diameter: 0 },
      bio: { ...clone(defaultState.filters.bio), count: 0, diameter: 0 }
    };
  } else if (tabName === "enquiry") {
    const enquiry = activeEnquiry();
    if (enquiry) {
      Object.assign(enquiry, { status: "New", clientId: "", contactId: "", siteId: "", assetId: "", siteRef: "", assetRef: "", notes: "" });
    }
  } else if (tabName === "clients") {
    state.activeClientId = "";
    ["clientName", "clientAddress", "clientContactName", "clientContactRole", "clientContactEmail", "clientContactPhone", "clientSiteRef", "clientSiteName", "clientSiteAddress", "clientAssetRef", "clientAssetName", "clientAssetLocation"].forEach((id) => {
      const input = $(`#${id}`);
      if (input) input.value = "";
    });
  } else if (tabName === "zones") {
    state.zones = [blankZone(1)];
  } else if (tabName === "ductwork") {
    state.legs = [blankLeg(1)];
  } else if (tabName === "schematic") {
    state.schematic = { title: "", drawingNumber: "", revision: "", showBomTags: "no", showPressures: "no" };
  } else if (tabName === "bom") {
    state.options = { internalBypass: false, steelBase: false, electrical: false };
    state.commercial = {
      fanUnitPrice: 0,
      fanQty: 0,
      contingency: 0,
      markup: 0,
      installHours: 0,
      managementHours: 0,
      labourSet: "other"
    };
  } else if (tabName === "fan") {
    state.fan = { rpmMeasured: 0, rpmBase: 0, urvPa: 0, quickCheckPa: 0 };
  } else if (tabName === "products") {
    clearUiOnlyInputs();
  } else if (tabName === "quote") {
    const quote = activeQuote();
    if (quote) {
      quote.notes = "";
      quote.status = "Draft";
    }
  }

  update();
  setActiveTab(tabName);
}

function getPath(path) {
  return path.split(".").reduce((acc, key) => acc?.[key], state);
}

function setPath(path, value) {
  const keys = path.split(".");
  const last = keys.pop();
  const target = keys.reduce((acc, key) => acc[key], state);
  target[last] = value;
}

function n(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function round(value, places = 2) {
  const factor = 10 ** places;
  return Math.round((n(value) + Number.EPSILON) * factor) / factor;
}

function findProduct(description, size = "") {
  const products = state.products || defaultProducts;
  const text = description.toLowerCase();
  const exact = products.find((p) => p.description.toLowerCase() === text && String(p.size) === String(size));
  if (exact) return exact;
  return products.find((p) => p.description.toLowerCase() === text) ||
    products.find((p) => text.includes(p.description.toLowerCase()) || p.description.toLowerCase().includes(text));
}

function line(qty, size, description, unit, source, category, supplier, material) {
  const product = findProduct(description, size) || {};
  return {
    qty: round(qty, 4),
    size,
    description,
    unit: unit ?? product.price ?? 0,
    source,
    category: category ?? product.category ?? "",
    supplier: supplier ?? product.supplier ?? "",
    material: material ?? product.material ?? "",
    get total() {
      return n(this.qty) * n(this.unit);
    }
  };
}

function calculateZones() {
  return state.zones.map((z) => {
    const volume = n(z.volumeOverride) > 0
      ? n(z.volumeOverride)
      : z.shape === "Cylinder"
        ? Math.PI * (n(z.diameter) / 2) ** 2 * n(z.height)
        : n(z.length) * n(z.width) * n(z.height);
    const requiredFlow = volume * n(z.ach);
    const designedFlow = requiredFlow * (1 + n(state.safety.flow));
    const actualAch = requiredFlow > 0 ? n(z.ach) * (designedFlow / requiredFlow) : 0;
    const minDuct = designedFlow && n(z.designVelocity)
      ? 2 * Math.sqrt(designedFlow / (Math.PI * 3600 * n(z.designVelocity))) * 1000
      : 0;
    const ductArea = Math.PI * (n(z.ductDiameter) / 2000) ** 2;
    const actualVelocity = ductArea ? (designedFlow / 3600) / ductArea : 0;
    const gap = designedFlow - requiredFlow;
    const status = actualVelocity === 0 ? "" : actualVelocity >= 7 && actualVelocity <= 13 ? "PASS" : "FAIL";
    return { ...z, volume, requiredFlow, designedFlow, actualAch, minDuct, ductArea, actualVelocity, gap, status };
  });
}

function filterCalc(filter, flowM3h, safety) {
  const count = Math.max(1, n(filter.count));
  const flowM3s = flowM3h / 3600 / count;
  const diameterM = n(filter.diameter) / 1000;
  const area = Math.PI * diameterM ** 2 / 4;
  const velocity = area ? flowM3s / area : 0;
  const pelletM = n(filter.pellet) / 1000;
  const eps = n(filter.voidFraction);
  const depth = n(filter.depth);
  let pressure = 0;
  if (eps > 0 && pelletM > 0 && velocity > 0) {
    const viscous = 150 * ((1 - eps) ** 2 / eps ** 3) * state.gas.viscosity * velocity * (depth / pelletM ** 2);
    const inertial = 1.75 * ((1 - eps) / eps ** 3) * state.gas.density * velocity ** 2 * (depth / pelletM);
    pressure = (viscous + inertial) * (1 + n(safety));
  }
  const bedVolume = Math.PI * (diameterM / 2) ** 2 * depth * n(filter.count);
  const ebct = flowM3s ? (Math.PI * (diameterM / 2) ** 2 * depth) / flowM3s : 0;
  return { flowM3s, area, velocity, pressure, bedVolume, ebct };
}

function calculateDuctLegs(totalFlow) {
  const entranceK = { None: 0, "Flanged inlet": 0.49 };
  const exitK = { None: 0, "Into vessel": 0, Stack: 1 };
  return state.legs.map((leg) => {
    const diameterM = n(leg.diameter) / 1000;
    const area = Math.PI * (diameterM / 2) ** 2;
    const flow = totalFlow * n(leg.share) / 100;
    const velocity = area ? (flow / 3600) / area : 0;
    const reynolds = state.gas.viscosity ? n(state.gas.density) * velocity * diameterM / n(state.gas.viscosity) : 0;
    const relRough = diameterM ? (n(leg.roughness) / 1000) / diameterM : 0;
    const friction = reynolds === 0 ? 0 : reynolds < 2300 ? 64 / reynolds : 0.25 / (Math.log10(relRough / 3.7 + 5.74 / reynolds ** 0.9) ** 2);
    const dynamicPressure = n(state.gas.density) * velocity ** 2 / 2;
    const k = n(leg.sr90) * 1.5 + n(leg.lr90) * 0.5 + n(leg.bend45) * 0.4 + n(leg.tee) * 0.6 +
      n(leg.reducer) * 0.3 + n(leg.nrd) * 2 + n(leg.vcd) * 0.25 + n(leg.sb) * 0.5 + n(leg.shoe) * 0.35 +
      (entranceK[leg.entrance] ?? 0) + (exitK[leg.exit] ?? 0) + n(leg.customK);
    const straightDp = diameterM ? friction * (n(leg.length) / diameterM) * dynamicPressure : 0;
    const fittingDp = k * dynamicPressure;
    return { ...leg, area, flow, velocity, reynolds, friction, k, straightDp, fittingDp, pressure: straightDp + fittingDp };
  });
}

function buildBom(results) {
  const rows = [];
  const add = (...args) => {
    const row = line(...args);
    if (n(row.qty) > 0) rows.push(row);
  };

  results.legs.forEach((leg) => {
    const size = n(leg.diameter);
    add(Math.ceil(n(leg.length) / 3), size, "CIRC. VENT 2LGRP", null, `Leg ${leg.name} - straight duct`);
    add(leg.lr90, size, "CIRC. 90 BEND (LR)", null, `Leg ${leg.name}`);
    add(leg.bend45, size, "CIRC. 45 BEND", null, `Leg ${leg.name}`);
    add(leg.tee, size, "CIRC. 90 TEE", null, `Leg ${leg.name}`);
    add(leg.reducer, "", "CIRC. REDUCER", null, `Leg ${leg.name}`);
    add(leg.vcd, size, "CIRC. VCD MLEAF", null, `Leg ${leg.name}`);
    add(leg.sb, size, "CIRC. SB DAMPER VENT 2LGRP", null, `Leg ${leg.name}`);
    add(leg.shoe, size, "CIRC. SHOE OFF CIRC DUCT", null, `Leg ${leg.name}`);
    add(leg.nrd * 2, size, "CIRC. NON RETURN DAMPER", null, `Leg ${leg.name}`);
  });

  add(state.filters.carbon.count, state.filters.carbon.diameter, `Carbon filter vessel ${state.filters.carbon.diameter}mm x ${state.filters.carbon.depth.toFixed(2)}m`, 11784.97, "System design + settings", "FILTERS", "RIPCO", "PP");
  add(results.carbon.bedVolume, "", "Carbon media - SA70 (Puragen)", 1500, "Carbon bed volume", "MEDIA", "PUREGEN", "SA70");
  add(state.commercial.fanQty, "", "Fans", state.commercial.fanUnitPrice, "Fan unit price");
  add(state.filters.carbon.count, "", "VegaDIF 85", null, "Instrumentation per filter");
  add(state.filters.carbon.count, "", "6mm UV resistant tube - Blue 15m", null, "Instrumentation per filter");
  add(state.filters.carbon.count, "", "6mm UV resistant tube - Red 15m", null, "Instrumentation per filter");
  add(state.filters.carbon.count, "", "Unistrut 41mm x 41mm HDG -1.5m", null, "Instrumentation support");
  add(state.filters.carbon.count, "", "Gusseted bracket 41mm x 41mm HDG", null, "Instrumentation support");
  add(state.commercial.fanQty, "", "SF620A Flow sensor for connection to evaluation unit", null, "Add-on per fan");
  add(state.commercial.fanQty, "", "SR2301 Control monitor evaluation unit", null, "Add-on per fan");
  add(state.commercial.fanQty, "", "E40160 Clamp fitting", null, "Add-on per fan");
  add(state.commercial.fanQty, "", "ENC08A Connecting cable with socket - 5m", null, "Add-on per fan");

  if (state.options.steelBase) {
    add(Math.ceil(n(state.commercial.fanQty) / 2), "", "Standard package fan base", null, "Steel base required");
    add(state.filters.carbon.count, "", "Standard package Carbon filter base", null, "Steel base required");
    const supports = Math.max(1, Math.ceil(results.legs.reduce((sum, leg) => sum + n(leg.length), 0) / 4));
    add(supports, "", "Brackets", null, "Duct supports");
    add(supports, 250, "split ring pair", null, "Duct supports");
    add(1, "", "Stack support lower 3m", null, "Stack supports");
    add(1, "", "Stack support upper 3m", null, "Stack supports");
  }

  if (state.options.electrical) {
    add(22, "2.5", "CB3325 6944LSH 4 CORE SWA LSOH 2.5MM BLK - per m", null, "Electrical add-on");
    add(22, "1.5", "CB3316 6943LSH 3 CORE SWA LSOH 1.5MM BLK - per m", null, "Electrical add-on");
    add(49, "0.75", "BS5308 PT1 TY2 0.75MM 1PR LSF BLK - per m", null, "Electrical add-on");
    add(3, "20s16", "CMP20s16 cable gland pk 2", null, "Electrical add-on");
    add(2, "20s", "CMP20s cable gland pk 2", null, "Electrical add-on");
    add(10, "20s", "Hawke 501/453/UNIV/OS/M20 Cable gland pk 1", null, "Electrical add-on");
    add(4, "20", "Hawke 501/453/UNIV/O/M20 Cable gland pk 1", null, "Electrical add-on");
    add(14, "20", "KITM/BRS/M20 GLAND KIT", null, "Electrical add-on");
    add(6, "150x3m", "LEGD SRFL150G CABLE TRAY 150MMX3M", null, "Electrical add-on");
    add(6, "150x3m", "LEGD SRFCC150G CLOSED COVER 150MMX3M", null, "Electrical add-on");
    add(6, "3m", "LEGD SRFDVG DIVIDER 3M", null, "Electrical add-on");
    add(1, "5m", "Cable tray edge protection 5m", null, "Electrical add-on");
    add(4, "", "Labour to fit electrics per day plus travel if on site", null, "Electrical add-on");
    add(1, "M6 x 20", "Galvanised Roofing nuts and bolts m6 x 20mm - per 100", null, "Electrical add-on");
  }

  const materialsTotal = rows.reduce((sum, row) => sum + row.total, 0);
  const labourRates = state.commercial.labourSet === "yw"
    ? { supervisor: 49.2, technician: 41, pm: 50, design: 45, admin: 35 }
    : { supervisor: 62.5, technician: 48, pm: 70, design: 62.5, admin: 45 };
  const install = (n(state.commercial.installHours) * 0.5 * labourRates.supervisor) + (n(state.commercial.installHours) * 0.5 * labourRates.technician);
  const management = (n(state.commercial.managementHours) * 0.5 * labourRates.pm) + (n(state.commercial.managementHours) * 0.35 * labourRates.design) + (n(state.commercial.managementHours) * 0.15 * labourRates.admin);
  const contingency = materialsTotal * n(state.commercial.contingency);
  const sell = (materialsTotal + contingency) * (1 + n(state.commercial.markup)) + install + management;
  return { rows, materialsTotal, contingency, install, management, sell };
}

function fanResults(systemPressure) {
  const curvePressure = [0, 500, 700, 850, 1050, 1150, 1280, 1450, 1500, 1550, 1560];
  const curveFlow = [4000, 3800, 3600, 3400, 3200, 3000, 2800, 2600, 2400, 2200, 2000];
  const weighted = curvePressure.reduce((acc, pressure, i) => {
    if (pressure <= 0) return acc;
    acc.top += curveFlow[i] * Math.sqrt(pressure);
    acc.bottom += pressure;
    return acc;
  }, { top: 0, bottom: 0 });
  const kFit = weighted.bottom ? weighted.top / weighted.bottom : 0;
  const qCurve = kFit * Math.sqrt(Math.max(systemPressure, 0));
  const qActual = qCurve * (n(state.fan.rpmMeasured) / Math.max(n(state.fan.rpmBase), 1));
  const qMax = kFit * Math.sqrt(Math.max(n(state.fan.urvPa), 0));
  const qLow = qMax * 0.85;
  const qHigh = qMax * 0.7;
  const dpLow = kFit ? (qLow / kFit) ** 2 : 0;
  const dpHigh = kFit ? (qHigh / kFit) ** 2 : 0;
  const maLow = n(state.fan.urvPa) ? 4 + 16 * Math.sqrt(dpLow / n(state.fan.urvPa)) : 0;
  const maHigh = n(state.fan.urvPa) ? 4 + 16 * Math.sqrt(dpHigh / n(state.fan.urvPa)) : 0;
  const quickMa = n(state.fan.urvPa) ? 4 + 16 * Math.sqrt(n(state.fan.quickCheckPa) / n(state.fan.urvPa)) : 0;
  const quickQ = kFit * Math.sqrt(Math.max(n(state.fan.quickCheckPa), 0));
  return { kFit, qCurve, qActual, qMax, qLow, qHigh, dpLow, dpHigh, maLow, maHigh, quickMa, quickQ };
}

function calculate() {
  const zones = calculateZones();
  const requiredFlow = zones.reduce((sum, zone) => sum + zone.requiredFlow, 0);
  const designedFlow = zones.reduce((sum, zone) => sum + zone.designedFlow, 0);
  const carbon = filterCalc(state.filters.carbon, designedFlow, state.safety.carbon);
  const bio = n(state.filters.bio.count) > 0 ? filterCalc(state.filters.bio, designedFlow, state.safety.bio) : { pressure: 0, bedVolume: 0, ebct: 0, velocity: 0 };
  const legs = calculateDuctLegs(designedFlow);
  const maxLegPressure = Math.max(0, ...legs.map((leg) => leg.pressure));
  const outletPressure = legs.slice(-3).reduce((sum, leg) => sum + leg.pressure, 0);
  const newSystemPressure = maxLegPressure + carbon.pressure + bio.pressure + outletPressure;
  const spentSystemPressure = newSystemPressure + carbon.pressure;
  const fan = fanResults(newSystemPressure);
  const results = { zones, requiredFlow, designedFlow, carbon, bio, legs, maxLegPressure, outletPressure, newSystemPressure, spentSystemPressure, fan };
  results.bom = buildBom(results);
  return results;
}

function renderInputs() {
  document.querySelectorAll("[data-path]").forEach((input) => {
    const value = getPath(input.dataset.path);
    if (input.type === "checkbox") input.checked = Boolean(value);
    else input.value = value ?? "";
  });
}

function bindInputs() {
  document.querySelectorAll("[data-path]").forEach((input) => {
    input.addEventListener("change", () => {
      const value = input.type === "checkbox" ? input.checked : input.type === "number" ? n(input.value) : input.value;
      setPath(input.dataset.path, value);
      update();
    });
  });
}

function renderKpis(results) {
  $("#kpiRequired").textContent = `${fmt.format(results.requiredFlow)} m3/h`;
  $("#kpiDesigned").textContent = `${fmt.format(results.designedFlow)} m3/h`;
  $("#kpiMaxLegPressure").textContent = `${fmt.format(results.maxLegPressure)} Pa`;
  $("#kpiPressure").textContent = `${fmt.format(results.newSystemPressure)} Pa`;
  $("#kpiSpent").textContent = `${fmt.format(results.spentSystemPressure)} Pa`;
  $("#kpiSell").textContent = moneyFmt.format(results.bom.sell);
}

function renderCards(container, cards) {
  container.innerHTML = cards.map((card) => `
    <article class="metric">
      <span>${card.label}</span>
      <strong>${card.value}</strong>
    </article>
  `).join("");
}

function renderPressure(results) {
  renderCards($("#pressureCards"), [
    { label: "Ductwork max leg pressure", value: `${round(results.maxLegPressure, 1)} Pa` },
    { label: "Carbon dP", value: `${round(results.carbon.pressure, 1)} Pa` },
    { label: "Biofilter dP", value: `${round(results.bio.pressure, 1)} Pa` },
    { label: "Outlet legs dP", value: `${round(results.outletPressure, 1)} Pa` },
    { label: "Carbon velocity", value: `${round(results.carbon.velocity, 3)} m/s` },
    { label: "Carbon bed volume", value: `${round(results.carbon.bedVolume, 2)} m3` },
    { label: "Carbon EBCT", value: `${round(results.carbon.ebct, 1)} sec` },
    { label: "Bio EBCT", value: `${round(results.bio.ebct, 1)} sec` }
  ]);
}

function field(value, type, handler, options = {}) {
  if (options.select) {
    return `<select data-edit="${handler}">${options.select.map((opt) => `<option ${value === opt ? "selected" : ""}>${opt}</option>`).join("")}</select>`;
  }
  return `<input data-edit="${handler}" type="${type}" value="${value ?? ""}" ${options.step ? `step="${options.step}"` : ""}>`;
}

function renderZones(results) {
  $("#zonesBody").innerHTML = results.zones.map((z, i) => `
    <tr>
      <td>${field(z.name, "text", `zone.${i}.name`)}</td>
      <td>${field(z.shape, "text", `zone.${i}.shape`, { select: ["Cylinder", "Rectangle"] })}</td>
      <td>${field(z.length, "number", `zone.${i}.length`, { step: "0.1" })}</td>
      <td>${field(z.width, "number", `zone.${i}.width`, { step: "0.1" })}</td>
      <td>${field(z.height, "number", `zone.${i}.height`, { step: "0.1" })}</td>
      <td>${field(z.diameter, "number", `zone.${i}.diameter`, { step: "0.1" })}</td>
      <td>${field(z.volumeOverride, "number", `zone.${i}.volumeOverride`, { step: "1" })}</td>
      <td>${field(z.ach, "number", `zone.${i}.ach`, { step: "0.1" })}</td>
      <td>${field(z.designVelocity, "number", `zone.${i}.designVelocity`, { step: "0.1" })}</td>
      <td class="num">${round(z.minDuct, 0)}</td>
      <td>${field(z.ductDiameter, "number", `zone.${i}.ductDiameter`, { step: "10" })}</td>
      <td class="num">${round(z.requiredFlow, 1)}</td>
      <td class="num">${round(z.designedFlow, 1)}</td>
      <td class="num">${round(z.actualVelocity, 2)}</td>
      <td><span class="status ${z.status === "PASS" ? "ok" : z.status ? "fail" : "blank"}">${z.status || "-"}</span></td>
      <td><button class="danger-btn" data-remove-zone="${i}" type="button">Remove</button></td>
    </tr>
  `).join("");
}

function renderLegs(results) {
  $("#legsBody").innerHTML = results.legs.map((leg, i) => `
    <tr>
      <td>${field(leg.name, "text", `leg.${i}.name`)}</td>
      <td>${field(leg.share, "number", `leg.${i}.share`, { step: "1" })}</td>
      <td>${field(leg.material, "text", `leg.${i}.material`, { select: ["GRP Ductwork", "PVC", "Spiral"] })}</td>
      <td>${field(leg.length, "number", `leg.${i}.length`, { step: "0.1" })}</td>
      <td>${field(leg.diameter, "number", `leg.${i}.diameter`, { step: "10" })}</td>
      <td>${field(leg.roughness, "number", `leg.${i}.roughness`, { step: "0.01" })}</td>
      <td>${field(leg.sr90, "number", `leg.${i}.sr90`)}</td>
      <td>${field(leg.lr90, "number", `leg.${i}.lr90`)}</td>
      <td>${field(leg.bend45, "number", `leg.${i}.bend45`)}</td>
      <td>${field(leg.tee, "number", `leg.${i}.tee`)}</td>
      <td>${field(leg.reducer, "number", `leg.${i}.reducer`)}</td>
      <td>${field(leg.nrd, "number", `leg.${i}.nrd`)}</td>
      <td>${field(leg.vcd, "number", `leg.${i}.vcd`)}</td>
      <td>${field(leg.sb, "number", `leg.${i}.sb`)}</td>
      <td>${field(leg.shoe, "number", `leg.${i}.shoe`)}</td>
      <td>${field(leg.entrance, "text", `leg.${i}.entrance`, { select: ["None", "Flanged inlet"] })}</td>
      <td>${field(leg.exit, "text", `leg.${i}.exit`, { select: ["None", "Into vessel", "Stack"] })}</td>
      <td>${field(leg.customK, "number", `leg.${i}.customK`, { step: "0.1" })}</td>
      <td class="num">${round(leg.pressure, 1)}</td>
      <td><button class="danger-btn" data-remove-leg="${i}" type="button">Remove</button></td>
    </tr>
  `).join("");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function svgText(x, y, text, cls = "", anchor = "middle") {
  return `<text x="${x}" y="${y}" text-anchor="${anchor}" class="${cls}">${escapeHtml(text)}</text>`;
}

function pipe(x1, y1, x2, y2, cls = "pipe") {
  return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" class="${cls}" marker-end="url(#arrow)"></line>`;
}

function valve(x, y, label) {
  return `
    <g class="symbol valve" transform="translate(${x} ${y})">
      <path d="M-14 -10 L0 0 L-14 10 Z"></path>
      <path d="M14 -10 L0 0 L14 10 Z"></path>
      <line x1="0" y1="-16" x2="0" y2="-2"></line>
      ${svgText(0, 29, label, "tiny")}
    </g>`;
}

function damper(x, y, label) {
  return `
    <g class="symbol damper" transform="translate(${x} ${y})">
      <rect x="-16" y="-12" width="32" height="24" rx="3"></rect>
      <line x1="-12" y1="8" x2="12" y2="-8"></line>
      ${svgText(0, 31, label, "tiny")}
    </g>`;
}

function fanSymbol(x, y, label) {
  return `
    <g class="equipment fan" transform="translate(${x} ${y})">
      <circle r="34"></circle>
      <path d="M0 -25 C18 -18 22 -3 5 0 C23 6 16 24 0 25 C-6 9 -23 11 -22 -7 C-10 -3 -7 -19 0 -25Z"></path>
      ${svgText(0, 56, label, "equip-label")}
    </g>`;
}

function filterVessel(x, y, label, subLabel, cls = "carbon") {
  return `
    <g class="equipment ${cls}" transform="translate(${x} ${y})">
      <rect x="-54" y="-72" width="108" height="144" rx="16"></rect>
      <line x1="-42" y1="-34" x2="42" y2="-34"></line>
      <line x1="-42" y1="34" x2="42" y2="34"></line>
      ${svgText(0, -6, label, "equip-label")}
      ${svgText(0, 14, subLabel, "tiny")}
    </g>`;
}

function instrumentBubble(x, y, label) {
  return `
    <g class="instrument" transform="translate(${x} ${y})">
      <circle r="18"></circle>
      ${svgText(0, 5, label, "tiny")}
    </g>`;
}

function renderSchematic(results) {
  const showPressures = state.schematic.showPressures === "yes";
  const showBomTags = state.schematic.showBomTags === "yes";
  const projectTitle = [state.project.reference, state.project.customer, state.project.site, state.project.item].filter(Boolean).join(" ");
  $("#schematicProject").textContent = projectTitle;
  $("#schematicMeta").textContent = `${state.schematic.drawingNumber} Rev ${state.schematic.revision}`;

  const carbonCount = Math.max(1, n(state.filters.carbon.count));
  const bioEnabled = n(state.filters.bio.count) > 0;
  const fanCount = Math.max(1, n(state.commercial.fanQty));
  const midY = 285;
  const x = {
    zones: 95,
    manifold: 245,
    bio: bioEnabled ? 420 : 0,
    carbon: bioEnabled ? 590 : 420,
    fans: bioEnabled ? 760 : 650,
    stack: bioEnabled ? 960 : 850,
    discharge: bioEnabled ? 1110 : 1000
  };
  const fanSpacing = fanCount > 1 ? 82 : 0;
  const fanStart = x.fans - ((fanCount - 1) * fanSpacing) / 2;
  const zoneRows = results.zones.length ? results.zones : [{ name: "Asset", designedFlow: results.designedFlow, ductDiameter: 0 }];

  const zoneSvg = zoneRows.map((zone, i) => {
    const spread = Math.min(72, 360 / Math.max(zoneRows.length, 1));
    const y = 120 + i * spread;
    const ductDia = n(zone.ductDiameter) || n(state.legs[0]?.diameter) || 0;
    return `
      <g class="asset">
        <rect x="${x.zones - 58}" y="${y - 24}" width="116" height="48" rx="6"></rect>
        ${svgText(x.zones, y - 4, zone.name || `Zone ${i + 1}`, "asset-label")}
        ${svgText(x.zones, y + 14, `${round(zone.designedFlow, 0)} m3/h`, "tiny")}
      </g>
      ${pipe(x.zones + 58, y, x.manifold - 20, midY, "branch-pipe")}
      ${svgText((x.zones + x.manifold) / 2, (y + midY) / 2 - 8, `${round(ductDia, 0)}mm`, "tiny")}
    `;
  }).join("");

  const legLabels = results.legs.slice(0, 5).map((leg, i) => {
    const lx = [280, 430, 590, 760, 920][i] || (280 + i * 140);
    return `
      <g class="leg-callout">
        <rect x="${lx - 70}" y="455" width="140" height="58" rx="6"></rect>
        ${svgText(lx, 476, `${leg.name} (${round(leg.share, 0)}%)`, "small")}
        ${svgText(lx, 494, `${round(leg.diameter, 0)}mm  ${round(leg.length, 1)}m`, "tiny")}
        ${showPressures ? svgText(lx, 512, `${round(leg.pressure, 1)} Pa`, "tiny") : ""}
      </g>
    `;
  }).join("");

  const fanSvg = Array.from({ length: fanCount }, (_, i) => {
    const fy = fanCount === 1 ? midY : midY - Math.min(58, (fanCount - 1) * 24) + i * Math.min(72, 140 / Math.max(fanCount - 1, 1));
    const fx = fanStart + i * fanSpacing;
    return `
      ${fanSymbol(fx, fy, `FAN ${String.fromCharCode(65 + i)}`)}
      ${pipe(fx - 42, midY, fx - 42, fy, "branch-pipe")}
      ${pipe(fx + 36, fy, x.stack - 42, midY, "pipe")}
    `;
  }).join("");

  const fanInlet = fanCount === 1
    ? pipe(x.carbon + 58, midY, fanStart - 42, midY)
    : pipe(x.carbon + 58, midY, fanStart - 42, midY, "pipe");

  const bioSvg = bioEnabled
    ? `${filterVessel(x.bio, midY, "BIOFILTER", `${round(results.bio.bedVolume, 2)} m3`, "bio")}
       ${showPressures ? svgText(x.bio, midY + 102, `${round(results.bio.pressure, 1)} Pa`, "pressure") : ""}
       ${pipe(x.bio + 58, midY, x.carbon - 58, midY)}`
    : "";

  const carbonSvg = `
       ${filterVessel(x.carbon, midY, `CARBON FILTER x${carbonCount}`, `${round(results.carbon.bedVolume, 2)} m3`, "carbon")}
       ${showPressures ? svgText(x.carbon, midY + 102, `${round(results.carbon.pressure, 1)} Pa`, "pressure") : ""}`;

  const firstFilterPipe = bioEnabled
    ? pipe(x.manifold, midY, x.bio - 62, midY)
    : pipe(x.manifold, midY, x.carbon - 62, midY);

  const nrdCount = results.legs.reduce((sum, leg) => sum + n(leg.nrd), 0);
  const vcdCount = results.legs.reduce((sum, leg) => sum + n(leg.vcd), 0);
  const sbCount = results.legs.reduce((sum, leg) => sum + n(leg.sb), 0);
  const symbolSvg = `
    ${vcdCount ? damper(x.manifold + 42, midY, `VCD x${vcdCount}`) : ""}
    ${sbCount ? damper(x.carbon - 92, midY, `SB x${sbCount}`) : ""}
    ${nrdCount ? valve(x.stack - 84, midY, `NRD x${nrdCount}`) : ""}
    ${instrumentBubble(x.carbon - 16, midY - 105, "P")}
    ${instrumentBubble(fanStart - 36, midY - 105, "L")}
  `;

  const svg = `
    <svg id="pidSvg" viewBox="0 0 1180 620" role="img" aria-label="Generated system schematic and P&ID">
      <defs>
        <marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
          <path d="M 0 0 L 10 5 L 0 10 z" class="arrow-head"></path>
        </marker>
        <style>
          #pidSvg text { font-family: Arial, Helvetica, sans-serif; fill: #17242b; }
          #pidSvg .title { font-size: 22px; font-weight: 700; }
          #pidSvg .small { font-size: 13px; font-weight: 700; }
          #pidSvg .tiny { font-size: 11px; fill: #53636d; }
          #pidSvg .asset-label, #pidSvg .equip-label { font-size: 13px; font-weight: 700; }
          #pidSvg .pressure { font-size: 12px; font-weight: 700; fill: #1f6b56; }
          #pidSvg .pipe, #pidSvg .branch-pipe { stroke: #263640; stroke-width: 7; fill: none; stroke-linecap: round; stroke-linejoin: round; }
          #pidSvg .branch-pipe { stroke-width: 4; stroke: #60727c; }
          #pidSvg .arrow-head { fill: #263640; }
          #pidSvg .equipment rect, #pidSvg .asset rect, #pidSvg .leg-callout rect { fill: #f8fbfa; stroke: #263640; stroke-width: 2; }
          #pidSvg .equipment line { stroke: #6b7d85; stroke-width: 2; }
          #pidSvg .fan circle { fill: #eef4f1; stroke: #263640; stroke-width: 2; }
          #pidSvg .fan path { fill: #1f6b56; opacity: .92; }
          #pidSvg .symbol path, #pidSvg .symbol rect { fill: #fff7ef; stroke: #263640; stroke-width: 2; }
          #pidSvg .symbol line { stroke: #263640; stroke-width: 2; }
          #pidSvg .instrument circle { fill: #fff; stroke: #c77332; stroke-width: 2; }
          #pidSvg .frame { fill: none; stroke: #9baab1; stroke-width: 1; }
          #pidSvg .flow-label { fill: #1f6b56; font-size: 14px; font-weight: 700; }
        </style>
      </defs>
      <rect class="frame" x="18" y="18" width="1144" height="584"></rect>
      ${svgText(590, 52, state.schematic.title || "SYSTEM SCHEMATIC - LEFT TO RIGHT", "title")}
      ${svgText(590, 78, projectTitle, "tiny")}
      ${svgText(1038, 50, `${state.schematic.drawingNumber}  Rev ${state.schematic.revision}`, "small")}
      ${svgText(x.manifold, midY - 70, "FLOW", "flow-label")}
      ${zoneSvg}
      ${firstFilterPipe}
      ${bioSvg}
      ${carbonSvg}
      ${fanInlet}
      ${fanSvg}
      ${symbolSvg}
      <g class="equipment stack" transform="translate(${x.stack} ${midY})">
        <rect x="-28" y="-96" width="56" height="192" rx="8"></rect>
        <line x1="-28" y1="-62" x2="28" y2="-62"></line>
        ${svgText(0, 124, "STACK", "equip-label")}
      </g>
      ${pipe(x.stack + 34, midY, x.discharge, midY)}
      ${svgText(x.discharge + 4, midY - 18, `${round(results.designedFlow, 0)} m3/h`, "small", "start")}
      ${showPressures ? svgText(x.discharge + 4, midY + 2, `${round(results.spentSystemPressure, 0)} Pa spent`, "pressure", "start") : ""}
      ${legLabels}
      ${showBomTags ? svgText(102, 574, `BOM lines: ${results.bom.rows.length}`, "small", "start") : ""}
      ${svgText(1078, 574, "Generated from live app inputs", "tiny", "end")}
    </svg>`;

  $("#schematicCanvas").innerHTML = svg;
  $("#schematicLegend").innerHTML = [
    "P = pressure transmitter",
    "L = low-flow switch",
    "VCD = volume control damper",
    "SB = shut-off balancing damper",
    "NRD = non-return damper"
  ].map((item) => `<span>${item}</span>`).join("");
}

function renderBom(results) {
  $("#bomBody").innerHTML = results.bom.rows.map((row) => `
    <tr>
      <td class="num">${round(row.qty, 3)}</td>
      <td>${row.size ?? ""}</td>
      <td>${row.description}</td>
      <td class="num">${moneyFmt.format(row.unit)}</td>
      <td class="num">${moneyFmt.format(row.total)}</td>
      <td>${row.source}</td>
      <td>${row.category}</td>
      <td>${row.supplier}</td>
    </tr>
  `).join("");
  $("#bomFoot").innerHTML = `
    <tr><th colspan="4">Materials subtotal</th><th class="num">${moneyFmt.format(results.bom.materialsTotal)}</th><th colspan="3"></th></tr>
    <tr><th colspan="4">Contingency</th><th class="num">${moneyFmt.format(results.bom.contingency)}</th><th colspan="3"></th></tr>
    <tr><th colspan="4">Install costs</th><th class="num">${moneyFmt.format(results.bom.install)}</th><th colspan="3"></th></tr>
    <tr><th colspan="4">Management costs</th><th class="num">${moneyFmt.format(results.bom.management)}</th><th colspan="3"></th></tr>
    <tr><th colspan="4">Total sell</th><th class="num">${moneyFmt.format(results.bom.sell)}</th><th colspan="3"></th></tr>
  `;
}

function renderFan(results) {
  renderCards($("#fanCards"), [
    { label: "K fit", value: round(results.fan.kFit, 3) },
    { label: "Q from curve", value: `${round(results.fan.qCurve, 1)} m3/h` },
    { label: "Q actual", value: `${round(results.fan.qActual, 1)} m3/h` },
    { label: "Q max at URV", value: `${round(results.fan.qMax, 1)} m3/h` },
    { label: "Low alarm flow", value: `${round(results.fan.qLow, 1)} m3/h` },
    { label: "High alarm flow", value: `${round(results.fan.qHigh, 1)} m3/h` },
    { label: "Low alarm mA", value: `${round(results.fan.maLow, 2)} mA` },
    { label: "High alarm mA", value: `${round(results.fan.maHigh, 2)} mA` },
    { label: "Quick check mA", value: `${round(results.fan.quickMa, 2)} mA` },
    { label: "Quick check Q", value: `${round(results.fan.quickQ, 1)} m3/h` }
  ]);

  const displayDecimals = results.newSystemPressure < 1000 ? 1 : 0;
  const rows = [
    ["Input type", "Current"],
    ["Input lower", "4 mA"],
    ["Input upper", "20 mA"],
    ["Display lower", "0 Pa"],
    ["Display upper", `${round(results.newSystemPressure, displayDecimals)} Pa`],
    ["Display decimals", displayDecimals],
    ["Display units", "Pa"],
    ["Relay 1 mode", "HI"],
    ["Relay 1 setpoint", `${round(results.newSystemPressure * 0.85, displayDecimals)} Pa`],
    ["Relay 2 mode", "LO"],
    ["Relay 2 setpoint", `${round(results.newSystemPressure * 0.7, displayDecimals)} Pa`]
  ];
  $("#commissioningList").innerHTML = rows.map(([label, value]) => `
    <div class="commissioning-row"><span>${label}</span><strong>${value}</strong></div>
  `).join("");
}

function renderProducts() {
  const q = ($("#productSearch").value || "").toLowerCase();
  renderProductFilterOptions();
  const supplierFilter = $("#productFilterSupplier")?.value || "";
  const materialFilter = $("#productFilterMaterial")?.value || "";
  const categoryFilter = $("#productFilterCategory")?.value || "";
  const sizeFilter = $("#productFilterSize")?.value || "";
  const minPriceRaw = $("#productFilterMinPrice")?.value || "";
  const maxPriceRaw = $("#productFilterMaxPrice")?.value || "";
  const minPrice = minPriceRaw === "" ? null : n(minPriceRaw);
  const maxPrice = maxPriceRaw === "" ? null : n(maxPriceRaw);
  const rows = (state.products || defaultProducts)
    .map((product, index) => ({ ...product, index }))
    .filter((p) => !q || [p.productCode, p.description, p.category, p.supplier, p.material, p.notes, String(p.size)].join(" ").toLowerCase().includes(q))
    .filter((p) => !supplierFilter || p.supplier === supplierFilter)
    .filter((p) => !materialFilter || p.material === materialFilter)
    .filter((p) => !categoryFilter || p.category === categoryFilter)
    .filter((p) => !sizeFilter || String(p.size ?? "") === sizeFilter)
    .filter((p) => minPrice === null || n(p.price) >= minPrice)
    .filter((p) => maxPrice === null || n(p.price) <= maxPrice)
    .slice(0, 120);
  $("#productsBody").innerHTML = rows.map((p) => `
    <tr>
      <td><input data-product-edit="${p.index}.productCode" value="${escapeHtml(p.productCode || "")}"></td>
      <td><input data-product-edit="${p.index}.description" value="${escapeHtml(p.description)}">${p.parentProduct ? `<small>${(p.childParts || []).length} child part(s)</small>` : ""}</td>
      <td><input data-product-edit="${p.index}.size" value="${escapeHtml(p.size ?? "")}"></td>
      <td><input class="num" data-product-edit="${p.index}.price" type="number" step="0.01" value="${p.price ?? 0}"></td>
      <td><input data-product-edit="${p.index}.category" value="${escapeHtml(p.category)}"></td>
      <td><input data-product-edit="${p.index}.supplier" value="${escapeHtml(p.supplier)}"></td>
      <td><input data-product-edit="${p.index}.material" value="${escapeHtml(p.material)}"></td>
      <td>
        ${p.parentProduct ? `<button data-toggle-product-children="${escapeHtml(p.productCode)}" type="button">${expandedProductCode === p.productCode ? "Hide" : "Expand"}</button>` : ""}
        <button class="danger-btn" data-remove-product="${p.index}" type="button">Remove</button>
      </td>
    </tr>
    ${p.parentProduct && expandedProductCode === p.productCode ? renderProductChildRows(p) : ""}
  `).join("");
}

function renderProductChildRows(product) {
  const rows = product.childParts || [];
  return `
    <tr class="child-row">
      <td colspan="8">
        <strong>Child BOM parts for ${escapeHtml(product.productCode || "")}</strong>
        <table class="nested-table">
          <thead><tr><th>Description</th><th>Size</th><th>Qty</th><th>Unit</th><th>Total</th><th>Supplier</th></tr></thead>
          <tbody>${rows.length ? rows.map((row) => `
            <tr><td>${escapeHtml(row.description || "")}</td><td>${escapeHtml(row.size || "")}</td><td class="num">${round(row.qty, 3)}</td><td class="num">${moneyFmt.format(row.unit || 0)}</td><td class="num">${moneyFmt.format(row.total || 0)}</td><td>${escapeHtml(row.supplier || "")}</td></tr>
          `).join("") : `<tr><td colspan="6">No child parts saved.</td></tr>`}</tbody>
        </table>
      </td>
    </tr>
  `;
}

function renderProductFilterOptions() {
  const products = state.products || defaultProducts;
  [
    ["productFilterSize", "size", "All sizes"],
    ["productFilterSupplier", "supplier", "All suppliers"],
    ["productFilterMaterial", "material", "All materials"],
    ["productFilterCategory", "category", "All categories"]
  ].forEach(([id, key, allLabel]) => {
    const select = $(`#${id}`);
    if (!select) return;
    const current = select.value;
    const values = Array.from(new Set(products.map((product) => product[key]).filter(Boolean))).sort((a, b) => String(a).localeCompare(String(b)));
    select.innerHTML = optionHtml("", allLabel, current) + values.map((value) => optionHtml(value, value, current)).join("");
  });
}

function optionHtml(value, label, selectedValue) {
  return `<option value="${escapeHtml(value)}" ${value === selectedValue ? "selected" : ""}>${escapeHtml(label)}</option>`;
}

function setBlock(selector, visible) {
  const element = $(selector);
  if (element) element.style.display = visible ? "" : "none";
}

function enquiryRecord(enquiry) {
  const client = findClient(enquiry?.clientId);
  const contact = findContact(client, enquiry?.contactId);
  const site = findSite(client, enquiry?.siteId);
  const asset = findAsset(site, enquiry?.assetId);
  const quotes = latestQuoteRevisions(state.quotes.filter((quote) => quote.enquiryId === enquiry?.id));
  return { enquiry, client, contact, site, asset, quotes };
}

function renderEnquiryRegister() {
  const body = $("#enquirySearchBody");
  if (!body) return;
  const numberQuery = ($("#enquirySearchNumber")?.value || "").trim().toLowerCase();
  const clientQuery = ($("#enquirySearchClient")?.value || "").trim().toLowerCase();
  const siteQuery = ($("#enquirySearchSite")?.value || "").trim().toLowerCase();
  const assetQuery = ($("#enquirySearchAsset")?.value || "").trim().toLowerCase();
  const rows = state.enquiries
    .map(enquiryRecord)
    .filter(({ enquiry, client, site, asset }) => (
      includesText(enquiry.number, numberQuery) &&
      includesText(client?.name, clientQuery) &&
      includesText(site?.name, siteQuery) &&
      includesText(asset?.name, assetQuery)
    ))
    .sort((a, b) => String(b.enquiry.date || "").localeCompare(String(a.enquiry.date || "")) || String(b.enquiry.number || "").localeCompare(String(a.enquiry.number || "")));

  body.innerHTML = rows.length ? rows.map(({ enquiry, client, site, asset, quotes }) => `
    <tr>
      <td><strong>${escapeHtml(enquiry.number || "")}</strong></td>
      <td>${escapeHtml(enquiry.status || "")}</td>
      <td>${escapeHtml(enquiry.date || "")}</td>
      <td>${escapeHtml(client?.name || "")}</td>
      <td>${escapeHtml(site?.name || enquiry.siteRef || "")}</td>
      <td>${escapeHtml(asset?.name || enquiry.assetRef || "")}</td>
      <td>${quotes.map((quote) => `${escapeHtml(quote.number)} Rev ${escapeHtml(quote.revision)}`).join(", ") || "None"}</td>
      <td><button data-select-enquiry="${escapeHtml(enquiry.id)}" type="button">Open</button></td>
    </tr>
  `).join("") : `<tr><td colspan="8">No matching enquiries found.</td></tr>`;
}

function renderEnquiry() {
  if (!state.enquiries.length) {
    const enquiry = blankEnquiry();
    state.enquiries.push(enquiry);
    state.activeEnquiryId = enquiry.id;
  }
  renderEnquiryRegister();
  setBlock("#enquiryRegisterPanel", !enquiryEditorOpen);
  setBlock("#enquiryEditor", enquiryEditorOpen);
  if (!enquiryEditorOpen) return;
  const enquiry = activeEnquiry();
  const client = findClient(enquiry.clientId);
  const site = findSite(client, enquiry.siteId);

  $("#enquiryNumber").value = enquiry.number;
  $("#enquiryStatus").value = enquiry.status;
  $("#enquiryDate").value = enquiry.date;
  $("#enquirySiteRef").value = enquiry.siteRef;
  $("#enquiryAssetRef").value = enquiry.assetRef;
  $("#enquiryNotes").value = enquiry.notes;

  $("#enquiryClient").innerHTML = optionHtml("", "Select client", enquiry.clientId) +
    state.clients.map((item) => optionHtml(item.id, item.name || "Unnamed client", enquiry.clientId)).join("");
  $("#enquiryContact").innerHTML = optionHtml("", "Select contact", enquiry.contactId) +
    (client?.contacts || []).map((item) => optionHtml(item.id, item.name || item.email || "Unnamed contact", enquiry.contactId)).join("");
  $("#enquirySite").innerHTML = optionHtml("", "Select site", enquiry.siteId) +
    (client?.sites || []).map((item) => optionHtml(item.id, `${item.reference || "No ref"} - ${item.name || "Unnamed site"}`, enquiry.siteId)).join("");
  $("#enquiryAsset").innerHTML = optionHtml("", "Select asset", enquiry.assetId) +
    (site?.assets || []).map((item) => optionHtml(item.id, `${item.reference || "No ref"} - ${item.name || "Unnamed asset"}`, enquiry.assetId)).join("");

  renderSearchResults();
}

function renderSearchResults() {
  if (activeAssetHistory) {
    renderAssetQuoteHistory();
    return;
  }
  const siteQuery = ($("#siteSearch")?.value || "").trim().toLowerCase();
  const assetQuery = ($("#assetSearch")?.value || "").trim().toLowerCase();
  const searchControls = $("#searchControls");
  const searchResults = $("#searchResults");
  const assetHistory = $("#assetQuoteHistory");
  if (searchControls) searchControls.style.display = "";
  if (searchResults) searchResults.style.display = "";
  if (assetHistory) {
    assetHistory.style.display = "none";
    assetHistory.innerHTML = "";
  }
  if (!siteQuery && !assetQuery) {
    $("#searchResults").innerHTML = `<article class="record-card"><p>Enter a site name or asset name to search.</p></article>`;
    return;
  }
  const matches = [];
  state.clients.forEach((client) => {
    (client.sites || []).forEach((site) => {
      if (siteQuery && !String(site.name || "").toLowerCase().includes(siteQuery)) return;
      (site.assets || [{ id: "", reference: "", name: "" }]).forEach((asset) => {
        if (assetQuery && !String(asset.name || "").toLowerCase().includes(assetQuery)) return;
        const enquiries = state.enquiries.filter((enquiry) => enquiry.siteId === site.id && (!asset.id || enquiry.assetId === asset.id));
        const quotes = state.quotes.filter((quote) => enquiries.some((enquiry) => enquiry.id === quote.enquiryId));
        matches.push({ client, site, asset, enquiries, quotes });
      });
    });
  });
  $("#searchResults").innerHTML = matches.length ? matches.map(({ client, site, asset, enquiries, quotes }) => `
    <article class="record-card">
      <strong>${escapeHtml(site.reference || "No site ref")} - ${escapeHtml(site.name || "Unnamed site")}</strong>
      <p>${escapeHtml(client.name || "Unnamed client")} | Asset: ${escapeHtml(asset.reference || "No asset ref")} ${escapeHtml(asset.name || "")}</p>
      <p>Enquiries: ${enquiries.map((item) => item.number).join(", ") || "None"} | Quotes: ${quotes.map((item) => `${item.number} Rev ${item.revision}`).join(", ") || "None"}</p>
      ${asset.id ? `<button data-asset-history-client="${escapeHtml(client.id)}" data-asset-history-site="${escapeHtml(site.id)}" data-asset-history-asset="${escapeHtml(asset.id)}" type="button">View asset quote history</button>` : ""}
      ${enquiries.map((item) => `<button data-select-enquiry="${item.id}" type="button">Open ${escapeHtml(item.number)}</button>`).join(" ")}
    </article>
  `).join("") : `<article class="record-card"><p>No matching site or asset records yet.</p></article>`;
}

function renderAssetQuoteHistory() {
  const target = $("#assetQuoteHistory");
  if (!target) return;
  if (!activeAssetHistory) {
    const searchControls = $("#searchControls");
    const searchResults = $("#searchResults");
    if (searchControls) searchControls.style.display = "";
    if (searchResults) searchResults.style.display = "";
    target.style.display = "none";
    target.innerHTML = "";
    return;
  }
  const searchControls = $("#searchControls");
  const searchResults = $("#searchResults");
  if (searchControls) searchControls.style.display = "none";
  if (searchResults) searchResults.style.display = "none";
  target.style.display = "";

  const client = findClient(activeAssetHistory.clientId);
  const site = findSite(client, activeAssetHistory.siteId);
  const asset = findAsset(site, activeAssetHistory.assetId);
  if (!client || !site || !asset) {
    activeAssetHistory = null;
    renderSearchResults();
    return;
  }

  const enquiries = state.enquiries.filter((enquiry) => enquiry.siteId === site.id && enquiry.assetId === asset.id);
  const rows = enquiries.flatMap((enquiry) => {
    const quotes = state.quotes.filter((quote) => quote.enquiryId === enquiry.id);
    if (!quotes.length) {
      return [{ enquiry, quote: null }];
    }
    return quotes.map((quote) => ({ enquiry, quote }));
  });

  target.innerHTML = `
    <article class="record-card">
      <strong>Quote history: ${escapeHtml(asset.reference || "No asset ref")} ${escapeHtml(asset.name || "Unnamed asset")}</strong>
      <p>${escapeHtml(client.name || "Unnamed client")} | ${escapeHtml(site.reference || "No site ref")} ${escapeHtml(site.name || "Unnamed site")}</p>
      <button data-back-to-search type="button">Back to search results</button>
    </article>
    ${rows.length ? rows.map(({ enquiry, quote }) => `
      <article class="record-card">
        <strong>${quote ? `Quote: ${escapeHtml(quote.number)} Rev ${escapeHtml(quote.revision)}` : `Enquiry: ${escapeHtml(enquiry.number)}`}</strong>
        <p>${quote ? `Enquiry: ${escapeHtml(enquiry.number)}` : "Quote: Not created yet"} | ${escapeHtml(enquiry.date || "")} | ${escapeHtml(enquiry.status || "")}</p>
        ${quote ? `<p>${escapeHtml(quote.date || "")} | ${escapeHtml(quote.status || "")} | ${moneyFmt.format(quote.total || 0)}</p><button data-select-quote="${escapeHtml(quote.id)}" type="button">Open quote</button>` : `<button data-select-enquiry="${escapeHtml(enquiry.id)}" type="button">Open enquiry</button>`}
      </article>
    `).join("") : `<article class="record-card"><p>No enquiries or quotes have been saved against this asset yet.</p></article>`}
  `;
}

function renderClients() {
  renderClientRegister();
  const panelByLevel = {
    "client-details": "#clientDetailsPanel",
    client: "#clientSitesPanel",
    site: "#clientAssetRegisterPanel",
    asset: "#clientAssetDetailPanel"
  };
  if (!panelByLevel[clientDrillLevel]) {
    clientDrillLevel = "register";
    clientEditorOpen = false;
  }
  setBlock("#clientRegisterPanel", clientDrillLevel === "register");
  setBlock("#clientEditor", clientDrillLevel !== "register");
  ["#clientDetailsPanel", "#clientSitesPanel", "#clientAssetRegisterPanel", "#clientAssetDetailPanel"].forEach((selector) => {
    setBlock(selector, panelByLevel[clientDrillLevel] === selector);
  });
  if (clientDrillLevel === "register") return;

  const client = clientDraft || blankClientDraft();
  client.contacts ||= [];
  client.sites ||= [];
  if (activeClientSiteId && !client.sites.some((site) => site.id === activeClientSiteId)) {
    activeClientSiteId = "";
  }
  const site = client.sites.find((item) => item.id === activeClientSiteId) || {};
  if (activeClientAssetId && !site.assets?.some((asset) => asset.id === activeClientAssetId)) {
    activeClientAssetId = "";
  }
  const selectedAsset = site.assets?.find((item) => item.id === activeClientAssetId);
  const asset = selectedAsset ? normalizeAsset(selectedAsset, 1) : {};
  $("#clientName").value = client.name || "";
  $("#clientAddress").value = client.address || "";
  if ($("#clientSiteRef")) $("#clientSiteRef").value = site.reference || "";
  if ($("#clientSiteName")) $("#clientSiteName").value = site.name || "";
  if ($("#clientSiteAddress")) $("#clientSiteAddress").value = site.address || "";
  $("#clientAssetRef").value = asset.reference || "";
  $("#clientAssetName").value = asset.name || "";
  $("#clientAssetLocation").value = asset.location || "";
  $("#assetFlowRate").value = asset.flowRate || 0;
  $("#assetFanModel").value = asset.fanModel || "";
  $("#assetFanSerial").value = asset.fanSerial || "";
  $("#assetFlowSensorModel").value = asset.flowSensorModel || "";
  $("#assetFlowSensorSerial").value = asset.flowSensorSerial || "";
  $("#assetFlowSetpoint").value = asset.flowSetpointPercent || 0;
  $("#assetCarbonMediaType").value = asset.carbonMediaType || "";
  $("#assetCarbonMediaSl").value = asset.carbonMediaSlNumber || "";
  $("#assetCarbonMediaVolume").value = asset.carbonMediaVolume || 0;
  $("#assetCarbonPressureDrop").value = asset.carbonPressureDrop || 0;
  $("#assetCarbonHighSetpoint").value = asset.carbonHighSetpoint || 0;
  $("#assetCarbonHighHighSetpoint").value = asset.carbonHighHighSetpoint || 0;
  $("#assetCarbonTempProbeModel").value = asset.carbonTempProbeModel || "";
  $("#assetCarbonTempProbeSerial").value = asset.carbonTempProbeSerial || "";
  $("#assetBioMediaType").value = asset.bioMediaType || "";
  $("#assetBioMediaVolume").value = asset.bioMediaVolume || 0;
  renderClientContacts();
  renderClientSites();
  renderClientAssets();
  renderAssetRegister();
  renderAssetDatalists();
  renderCustomAssetFields(asset);
  renderAssetAttachments(asset);
}

function clientSearchText(client, key) {
  if (key === "site") return (client.sites || []).map((site) => site.name).join(" ");
  if (key === "asset") return (client.sites || []).flatMap((site) => site.assets || []).map((asset) => asset.name).join(" ");
  if (key === "contact") return (client.contacts || []).map((contact) => [contact.name, contact.role, contact.email, contact.phone].join(" ")).join(" ");
  return "";
}

function renderClientRegister() {
  const body = $("#clientList");
  if (!body) return;
  const nameQuery = ($("#clientSearchName")?.value || "").trim().toLowerCase();
  const addressQuery = ($("#clientSearchAddress")?.value || "").trim().toLowerCase();
  const contactQuery = ($("#clientSearchContact")?.value || "").trim().toLowerCase();
  const siteQuery = ($("#clientSearchSite")?.value || "").trim().toLowerCase();
  const assetQuery = ($("#clientSearchAsset")?.value || "").trim().toLowerCase();
  const rows = state.clients.filter((client) => (
    includesText(client.name, nameQuery) &&
    includesText(client.address, addressQuery) &&
    includesText(clientSearchText(client, "contact"), contactQuery) &&
    includesText(clientSearchText(client, "site"), siteQuery) &&
    includesText(clientSearchText(client, "asset"), assetQuery)
  ));

  body.innerHTML = rows.length ? rows.map((client) => {
    const assetCount = (client.sites || []).reduce((total, site) => total + (site.assets || []).length, 0);
    return `
      <tr class="clickable-row" data-row-open-client="${escapeHtml(client.id)}">
        <td><strong>${escapeHtml(client.name || "Unnamed client")}</strong></td>
        <td>${escapeHtml(client.address || "No address")}</td>
        <td>${escapeHtml((client.contacts || []).length)}</td>
        <td>${escapeHtml((client.sites || []).length)}</td>
        <td>${escapeHtml(assetCount)}</td>
        <td>
          <button data-select-client="${escapeHtml(client.id)}" type="button">Open Client</button>
          ${adminDeleteButton(`data-remove-client="${escapeHtml(client.id)}"`)}
        </td>
      </tr>
    `;
  }).join("") : `<tr><td colspan="6">No matching clients found.</td></tr>`;
}

function renderAssetRegister() {
  const body = $("#assetRegisterList");
  if (!body) return;
  const siteRefQuery = ($("#assetRegisterSearchSiteRef")?.value || "").trim().toLowerCase();
  const siteQuery = ($("#assetRegisterSearchSite")?.value || "").trim().toLowerCase();
  const assetRefQuery = ($("#assetRegisterSearchAssetRef")?.value || "").trim().toLowerCase();
  const assetQuery = ($("#assetRegisterSearchAsset")?.value || "").trim().toLowerCase();
  const locationQuery = ($("#assetRegisterSearchLocation")?.value || "").trim().toLowerCase();
  const fanQuery = ($("#assetRegisterSearchFan")?.value || "").trim().toLowerCase();
  const client = clientDraft || findClient(state.activeClientId);
  const sourceSites = activeClientSiteId
    ? (client?.sites || []).filter((site) => site.id === activeClientSiteId)
    : (client?.sites || []);
  const rows = sourceSites.flatMap((site) =>
    (site.assets || []).map((asset) => ({ client, site, asset: normalizeAsset(asset, 1) }))
  ).filter(({ site, asset }) => (
    includesText(site.reference, siteRefQuery) &&
    includesText(site.name, siteQuery) &&
    includesText(asset.reference, assetRefQuery) &&
    includesText(asset.name, assetQuery) &&
    includesText(asset.location, locationQuery) &&
    includesText([asset.fanModel, asset.fanSerial].join(" "), fanQuery)
  ));

  body.innerHTML = rows.length ? rows.map(({ site, asset }) => `
    <tr class="clickable-row" data-row-open-asset-site="${escapeHtml(site.id)}" data-row-open-asset="${escapeHtml(asset.id)}">
      <td>${escapeHtml(site.reference || "")}</td>
      <td>${escapeHtml(site.name || "Unnamed site")}</td>
      <td>${escapeHtml(asset.reference || "No ref")}</td>
      <td><strong>${escapeHtml(asset.name || "Unnamed asset")}</strong></td>
      <td>${escapeHtml(asset.location || "")}</td>
      <td class="num">${fmt.format(asset.flowRate || 0)}</td>
      <td>${escapeHtml([asset.fanModel, asset.fanSerial].filter(Boolean).join(" / "))}</td>
      <td>
        <button data-open-client-asset-site="${escapeHtml(site.id)}" data-open-client-asset="${escapeHtml(asset.id)}" type="button">Open Asset Info</button>
        ${adminDeleteButton(`data-remove-client-asset-site="${escapeHtml(site.id)}" data-remove-client-asset="${escapeHtml(asset.id)}"`)}
      </td>
    </tr>
  `).join("") : `<tr><td colspan="8">No matching assets found for this client.</td></tr>`;
}

function renderAssetSavedDataRow(site, asset) {
  const field = (label, key, type = "text", step = "") => `
    <label>${label}
      <input data-asset-detail-field="${key}" data-asset-detail-site="${escapeHtml(site.id)}" data-asset-detail-asset="${escapeHtml(asset.id)}" ${type === "number" ? `type="number" step="${step || "0.01"}"` : ""} value="${escapeHtml(asset[key] ?? "")}">
    </label>
  `;
  return `
    <tr class="asset-detail-row">
      <td colspan="8">
        <div class="asset-detail-panel">
          <h4>Saved data for ${escapeHtml(asset.reference || "No ref")} ${escapeHtml(asset.name || "Unnamed asset")}</h4>
          <div class="form-grid compact">
            ${field("Asset reference", "reference")}
            ${field("Asset name", "name")}
            ${field("Asset location", "location")}
            ${field("Flow rate m3/h", "flowRate")}
            ${field("Fan model", "fanModel")}
            ${field("Fan serial number", "fanSerial")}
            ${field("Flow sensor model", "flowSensorModel")}
            ${field("Flow sensor serial number", "flowSensorSerial")}
            ${field("Flow set point %", "flowSetpointPercent", "number")}
            ${field("Carbon media type", "carbonMediaType")}
            ${field("Carbon media SL number", "carbonMediaSlNumber")}
            ${field("Carbon media volume m3", "carbonMediaVolume")}
            ${field("Carbon PT pressure drop", "carbonPressureDrop")}
            ${field("Carbon PT high setpoint", "carbonHighSetpoint")}
              ${field("Carbon PT high-high setpoint", "carbonHighHighSetpoint")}
            ${field("Carbon temp probe model", "carbonTempProbeModel")}
            ${field("Carbon temp probe serial", "carbonTempProbeSerial")}
            ${field("Biofilter media type", "bioMediaType")}
            ${field("Biofilter media volume m3", "bioMediaVolume")}
          </div>
        </div>
      </td>
    </tr>
  `;
}

function assetDetailInputId(key) {
  return {
    reference: "clientAssetRef",
    name: "clientAssetName",
    location: "clientAssetLocation",
    flowRate: "assetFlowRate",
    fanModel: "assetFanModel",
    fanSerial: "assetFanSerial",
    flowSensorModel: "assetFlowSensorModel",
    flowSensorSerial: "assetFlowSensorSerial",
    flowSetpointPercent: "assetFlowSetpoint",
    carbonMediaType: "assetCarbonMediaType",
    carbonMediaSlNumber: "assetCarbonMediaSl",
    carbonMediaVolume: "assetCarbonMediaVolume",
    carbonPressureDrop: "assetCarbonPressureDrop",
    carbonHighSetpoint: "assetCarbonHighSetpoint",
    carbonHighHighSetpoint: "assetCarbonHighHighSetpoint",
    carbonTempProbeModel: "assetCarbonTempProbeModel",
    carbonTempProbeSerial: "assetCarbonTempProbeSerial",
    bioMediaType: "assetBioMediaType",
    bioMediaVolume: "assetBioMediaVolume"
  }[key] || "";
}

function updateOpenedAssetField(target, shouldRender = false) {
  const key = target.dataset.assetDetailField;
  if (!key || !clientDraft) return;
  const siteId = target.dataset.assetDetailSite || activeClientSiteId;
  const assetId = target.dataset.assetDetailAsset || activeClientAssetId;
  const site = clientDraft.sites?.find((item) => item.id === siteId);
  const asset = site?.assets?.find((item) => item.id === assetId);
  if (!site || !asset) return;
  activeClientSiteId = siteId;
  activeClientAssetId = assetId;
  asset[key] = target.type === "number" ? n(target.value) : target.value.trim();
  const hiddenInputId = assetDetailInputId(key);
  if (hiddenInputId && $(`#${hiddenInputId}`)) {
    $(`#${hiddenInputId}`).value = asset[key] ?? "";
  }
  if (shouldRender) renderAssetRegister();
}

function renderClientContacts() {
  const client = clientDraft || findClient(state.activeClientId);
  const list = $("#clientContactsList");
  if (!list) return;
  const query = ($("#clientContactSearch")?.value || "").trim().toLowerCase();
  const contacts = (client?.contacts || []).filter((contact) => !query || [
    contact.name,
    contact.role,
    contact.email,
    contact.phone
  ].join(" ").toLowerCase().includes(query));
  list.innerHTML = contacts.length ? contacts.map((contact) => `
    <article class="contact-row">
      <strong>${escapeHtml(contact.name || "Unnamed contact")}</strong>
      <span>${escapeHtml(contact.role || "-")}</span>
      <span>${escapeHtml(contact.email || "-")}</span>
      <span>${escapeHtml(contact.phone || "-")}</span>
      ${adminDeleteButton(`data-remove-contact="${escapeHtml(contact.id)}"`)}
    </article>
  `).join("") : `<article class="record-card"><p>No matching contacts found.</p></article>`;
}

function renderClientSites() {
  const body = $("#clientSitesBody");
  if (!body) return;
  const client = clientDraft || findClient(state.activeClientId);
  const refQuery = ($("#clientSiteRegisterSearchRef")?.value || "").trim().toLowerCase();
  const nameQuery = ($("#clientSiteRegisterSearchName")?.value || "").trim().toLowerCase();
  const sites = (client?.sites || []).filter((site) => (
    includesText(site.reference, refQuery) &&
    includesText([site.name, site.address].join(" "), nameQuery)
  ));
  body.innerHTML = sites.length ? sites.map((site) => `
    <tr class="clickable-row" data-row-open-site="${escapeHtml(site.id)}">
      <td>${escapeHtml(site.reference || "No ref")}</td>
      <td><strong>${escapeHtml(site.name || "Unnamed site")}</strong></td>
      <td>${escapeHtml((site.assets || []).length)}</td>
      <td>
        <button data-select-client-site="${escapeHtml(site.id)}" type="button">Open Site</button>
        ${adminDeleteButton(`data-remove-client-site="${escapeHtml(site.id)}"`)}
      </td>
    </tr>
  `).join("") : `<tr><td colspan="4">No matching sites found.</td></tr>`;
}

function renderClientAssets() {
  const body = $("#clientAssetsBody");
  if (!body) return;
  const client = clientDraft || findClient(state.activeClientId);
  const site = client?.sites?.find((item) => item.id === activeClientSiteId);
  const assets = site?.assets || [];
  body.innerHTML = assets.length ? assets.map((asset) => `
    <tr class="clickable-row" data-row-open-asset-site="${escapeHtml(site?.id || "")}" data-row-open-asset="${escapeHtml(asset.id)}">
      <td>${escapeHtml(asset.reference || "No ref")}</td>
      <td><strong>${escapeHtml(asset.name || "Unnamed asset")}</strong></td>
      <td>${escapeHtml(asset.location || "")}</td>
      <td class="num">${fmt.format(asset.flowRate || 0)}</td>
      <td>
        <button data-select-client-asset="${escapeHtml(asset.id)}" type="button">${asset.id === activeClientAssetId ? "Selected" : "Open"}</button>
        ${adminDeleteButton(`data-remove-client-asset="${escapeHtml(asset.id)}"`)}
      </td>
    </tr>
  `).join("") : `<tr><td colspan="5">No assets added for this site yet.</td></tr>`;
}

function renderAssetDatalists() {
  const referenceList = $("#assetReferenceOptions");
  const locationList = $("#assetLocationOptions");
  if (!referenceList || !locationList) return;
  const assets = (state.clients || []).flatMap((client) =>
    (client.sites || []).flatMap((site) => site.assets || [])
  );
  const references = [...new Set(assets.map((asset) => asset.reference).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b)));
  const locations = [...new Set(assets.map((asset) => asset.location).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b)));
  referenceList.innerHTML = references.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  locationList.innerHTML = locations.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
}

function supplierSearchText(supplier, key) {
  if (key === "site") return (supplier.sites || []).map((site) => [site.name, site.address, site.phone, site.email].join(" ")).join(" ");
  if (key === "contact") return (supplier.contacts || []).map((contact) => [contact.name, contact.role, contact.email, contact.phone].join(" ")).join(" ");
  return "";
}

function renderSuppliers() {
  renderSupplierRegister();
  setBlock("#supplierRegisterPanel", !supplierEditorOpen);
  setBlock("#supplierEditor", supplierEditorOpen);
  if (!supplierEditorOpen) return;

  const supplier = supplierDraft || blankSupplier();
  supplier.contacts ||= [];
  supplier.sites ||= [];
  $("#supplierName").value = supplier.name || "";
  $("#supplierPhone").value = supplier.phone || "";
  $("#supplierEmail").value = supplier.email || "";
  $("#supplierInvoiceAddress").value = supplier.invoiceAddress || "";
  $("#supplierNotes").value = supplier.notes || "";
  renderSupplierContacts();
  renderSupplierSites();
}

function renderSupplierRegister() {
  const body = $("#supplierList");
  if (!body) return;
  const nameQuery = ($("#supplierSearchName")?.value || "").trim().toLowerCase();
  const phoneQuery = ($("#supplierSearchPhone")?.value || "").trim().toLowerCase();
  const emailQuery = ($("#supplierSearchEmail")?.value || "").trim().toLowerCase();
  const siteQuery = ($("#supplierSearchSite")?.value || "").trim().toLowerCase();
  const contactQuery = ($("#supplierSearchContact")?.value || "").trim().toLowerCase();
  const rows = (state.suppliers || []).filter((supplier) => (
    includesText(supplier.name, nameQuery) &&
    includesText(supplier.phone, phoneQuery) &&
    includesText(supplier.email, emailQuery) &&
    includesText(supplierSearchText(supplier, "site"), siteQuery) &&
    includesText(supplierSearchText(supplier, "contact"), contactQuery)
  ));

  body.innerHTML = rows.length ? rows.map((supplier) => `
    <tr>
      <td><strong>${escapeHtml(supplier.name || "Unnamed supplier")}</strong></td>
      <td>${escapeHtml(supplier.phone || "")}</td>
      <td>${escapeHtml(supplier.email || "")}</td>
      <td>${escapeHtml((supplier.sites || []).length)}</td>
      <td>${escapeHtml((supplier.contacts || []).length)}</td>
      <td><button data-select-supplier="${escapeHtml(supplier.id)}" type="button">Open</button></td>
    </tr>
  `).join("") : `<tr><td colspan="6">No matching suppliers found.</td></tr>`;
}

function renderSupplierContacts() {
  const list = $("#supplierContactsList");
  if (!list) return;
  const query = ($("#supplierContactSearch")?.value || "").trim().toLowerCase();
  const contacts = (supplierDraft?.contacts || []).filter((contact) => !query || [
    contact.name,
    contact.role,
    contact.email,
    contact.phone
  ].join(" ").toLowerCase().includes(query));
  list.innerHTML = contacts.length ? contacts.map((contact) => `
    <article class="contact-row">
      <strong>${escapeHtml(contact.name || "Unnamed contact")}</strong>
      <span>${escapeHtml(contact.role || "-")}</span>
      <span>${escapeHtml(contact.email || "-")}</span>
      <span>${escapeHtml(contact.phone || "-")}</span>
      <button class="danger-btn" data-remove-supplier-contact="${escapeHtml(contact.id)}" type="button">Remove</button>
    </article>
  `).join("") : `<article class="record-card"><p>No matching contacts found.</p></article>`;
}

function renderSupplierSites() {
  const body = $("#supplierSitesBody");
  if (!body) return;
  const sites = supplierDraft?.sites || [];
  body.innerHTML = sites.length ? sites.map((site) => `
    <tr>
      <td><strong>${escapeHtml(site.name || "Unnamed site")}</strong></td>
      <td>${escapeHtml(site.address || "")}</td>
      <td>${escapeHtml(site.phone || "")}</td>
      <td>${escapeHtml(site.email || "")}</td>
      <td><button class="danger-btn" data-remove-supplier-site="${escapeHtml(site.id)}" type="button">Remove</button></td>
    </tr>
  `).join("") : `<tr><td colspan="5">No site addresses added yet.</td></tr>`;
}

function renderActiveEnquirySummary() {
  const summary = $("#activeEnquirySummary");
  if (!summary) return;
  const enquiry = activeEnquiry();
  const client = findClient(enquiry?.clientId);
  const site = findSite(client, enquiry?.siteId);
  const asset = findAsset(site, enquiry?.assetId);
  summary.innerHTML = enquiry ? `
    <article class="record-card">
      <strong>${escapeHtml(enquiry.number)} - ${escapeHtml(enquiry.status)}</strong>
      <p>${escapeHtml(client?.name || "No client selected")}</p>
      <p>${escapeHtml(site?.reference || enquiry.siteRef || "No site ref")} | ${escapeHtml(asset?.reference || enquiry.assetRef || "No asset ref")}</p>
    </article>
  ` : `<article class="record-card"><p>No active enquiry.</p></article>`;
}

function quoteContext() {
  const enquiry = activeEnquiry();
  const client = findClient(enquiry?.clientId);
  const contact = findContact(client, enquiry?.contactId);
  const site = findSite(client, enquiry?.siteId);
  const asset = findAsset(site, enquiry?.assetId);
  return { enquiry, client, contact, site, asset };
}

function quoteRecord(quote) {
  const enquiry = state.enquiries.find((item) => item.id === quote.enquiryId) || null;
  const client = findClient(enquiry?.clientId);
  const contact = findContact(client, enquiry?.contactId);
  const site = findSite(client, enquiry?.siteId);
  const asset = findAsset(site, enquiry?.assetId);
  return { quote, enquiry, client, contact, site, asset };
}

function latestQuoteRevisions(quotes) {
  const latest = new Map();
  quotes.forEach((quote) => {
    const key = String(quote.number || quote.id);
    const existing = latest.get(key);
    if (!existing || Number(quote.revision || 0) > Number(existing.revision || 0)) {
      latest.set(key, quote);
    }
  });
  return Array.from(latest.values());
}

function buildQuoteHtml(quote = activeQuote()) {
  const { enquiry, client, contact, site, asset } = quoteContext();
  const results = currentResults || calculate();
  const quoteLines = quote?.lines?.length ? quote.lines : results.bom.rows.slice(0, 18).map((row) => ({
    description: row.description,
    qty: row.qty,
    price: row.unit
  }));
  const quoteRows = quoteLines.map((row) => `
    <tr><td>${escapeHtml([row.description, row.freeText].filter(Boolean).join(" - "))}</td><td class="num">${round(row.qty, 2)}</td><td class="num">${moneyFmt.format(row.price ?? row.unit)}</td><td class="num">${moneyFmt.format(n(row.qty) * n(row.price ?? row.unit))}</td></tr>
  `).join("");
  const quoteTotal = quoteLines.reduce((sum, row) => sum + n(row.qty) * n(row.price ?? row.unit), 0);
  return `
    <h2>Quotation ${escapeHtml(quote?.number || "Draft")} Rev ${escapeHtml(quote?.revision ?? 0)}</h2>
    <p><strong>Date:</strong> ${escapeHtml(quote?.date || todayIso())}</p>
    <p><strong>Client:</strong> ${escapeHtml(client?.name || "")}<br>
    <strong>Contact:</strong> ${escapeHtml(contact?.name || "")} ${escapeHtml(contact?.email || "")}<br>
    <strong>Site:</strong> ${escapeHtml(site?.reference || enquiry?.siteRef || "")} ${escapeHtml(site?.name || "")}<br>
    <strong>Asset:</strong> ${escapeHtml(asset?.reference || enquiry?.assetRef || "")} ${escapeHtml(asset?.name || "")}</p>
    <h3>Design Summary</h3>
    <p>Required flow: ${fmt.format(results.requiredFlow)} m3/h<br>
    Designed flow: ${fmt.format(results.designedFlow)} m3/h<br>
    Estimated new system pressure: ${fmt.format(results.newSystemPressure)} Pa<br>
    Estimated spent system pressure: ${fmt.format(results.spentSystemPressure)} Pa</p>
    <h3>Offer</h3>
    <p>${escapeHtml(quote?.offerText || "Please find our offer for the odour control system described below.")}</p>
    <h3>Commercial Summary</h3>
    <table><thead><tr><th>Description</th><th>Qty</th><th>Unit</th><th>Total</th></tr></thead><tbody>${quoteRows}</tbody></table>
    <p><strong>Total sell:</strong> ${moneyFmt.format(quoteTotal || quote?.total || results.bom.sell)}</p>
    <h3>Notes</h3>
    <p>${escapeHtml(quote?.notes || "Draft quotation generated from the current app design.")}</p>
    <h3>Exclusions</h3>
    <p>${escapeHtml(quote?.exclusions || "No exclusions entered.")}</p>
  `;
}

function pdfEscape(value) {
  return String(value ?? "").replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

function wrapPdfText(text, maxChars = 88) {
  const words = String(text || "").replace(/\s+/g, " ").trim().split(" ").filter(Boolean);
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (next.length > maxChars && current) {
      lines.push(current);
      current = word;
    } else {
      current = next;
    }
  });
  if (current) lines.push(current);
  return lines.length ? lines : [""];
}

function pdfMoney(value) {
  return `GBP ${round(n(value), 2)}`;
}

function pdfDate(value = todayIso()) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value || "";
  return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
}

function pdfAddressLines(value, fallback = []) {
  const lines = String(value || "")
    .split(/\r?\n|,/)
    .map((line) => line.trim())
    .filter(Boolean);
  return lines.length ? lines : fallback;
}

function quoteLinesForPdf(quote, results) {
  return quote?.lines?.length ? quote.lines : results.bom.rows.slice(0, 18).map((row) => ({
    description: row.description,
    qty: row.qty,
    price: row.unit
  }));
}

function quotePdfLines(quote = activeQuote()) {
  const { enquiry, client, contact, site, asset } = quoteContext();
  const results = currentResults || calculate();
  const quoteLines = quoteLinesForPdf(quote, results);
  const total = quoteLines.reduce((sum, row) => sum + n(row.qty) * n(row.price ?? row.unit), 0);
  const lines = [
    `Quotation ${quote?.number || "Draft"} Rev ${quote?.revision ?? 0}`,
    `Date: ${quote?.date || todayIso()}`,
    "",
    `Client: ${client?.name || ""}`,
    `Contact: ${contact?.name || ""} ${contact?.email || ""}`,
    `Site: ${site?.reference || enquiry?.siteRef || ""} ${site?.name || ""}`,
    `Asset: ${asset?.reference || enquiry?.assetRef || ""} ${asset?.name || ""}`,
    "",
    "Design Summary",
    `Required flow: ${fmt.format(results.requiredFlow)} m3/h`,
    `Designed flow: ${fmt.format(results.designedFlow)} m3/h`,
    `Estimated new system pressure: ${fmt.format(results.newSystemPressure)} Pa`,
    `Estimated spent system pressure: ${fmt.format(results.spentSystemPressure)} Pa`,
    "",
    "Offer"
  ];
  wrapPdfText(quote?.offerText || "Please find our offer for the odour control system described below.").forEach((line) => lines.push(line));
  lines.push("", "Commercial Summary");
  quoteLines.forEach((row) => {
    const description = [row.description, row.freeText].filter(Boolean).join(" - ");
    lines.push(`${round(row.qty, 2)} x ${description} @ ${pdfMoney(row.price ?? row.unit)} = ${pdfMoney(n(row.qty) * n(row.price ?? row.unit))}`);
  });
  lines.push(`Total sell: ${pdfMoney(total || quote?.total || results.bom.sell)}`, "", "Notes");
  wrapPdfText(quote?.notes || "Draft quotation generated from the current app design.").forEach((line) => lines.push(line));
  lines.push("", "Exclusions");
  wrapPdfText(quote?.exclusions || "No exclusions entered.").forEach((line) => lines.push(line));
  return lines;
}

function createQuotePdfBlob(quote = activeQuote()) {
  const pageWidth = 595;
  const pageHeight = 842;
  const margin = 54;
  const { enquiry, client, contact, site, asset } = quoteContext();
  const results = currentResults || calculate();
  const quoteLines = quoteLinesForPdf(quote, results);
  const total = quoteLines.reduce((sum, row) => sum + n(row.qty) * n(row.price ?? row.unit), 0) || quote?.total || results.bom.sell;
  const clientAddress = pdfAddressLines(client?.address, [client?.name || "Client address"]);
  const subject = [site?.name, asset?.name].filter(Boolean).join(" - ") || enquiry?.notes || "Odour control works";
  const attention = contact?.name || "Sir / Madam";
  const offerLines = wrapPdfText(quote?.offerText || `Following our recent discussions in connection with the above we now have pleasure in confirming our costs to carry out the following works at ${site?.name || "site"}.`, 82);
  const noteLines = wrapPdfText(quote?.notes || "We trust that this is in line with your immediate requirements, however, should you need any additional information please do not hesitate to call.", 82);
  const exclusionLines = wrapPdfText(quote?.exclusions || "No exclusions entered.", 82);
  const companyFooter = [
    "Greenacre Environmental Systems Limited is registered in England and Wales with company number 09117078.",
    "The company is also registered for VAT number 191 4495 86.",
    "Unit 17, Riverside Way, Ravensthorpe Industrial Estate, Dewsbury, West Yorkshire, WF13 3LG"
  ];
  const companyAddress = [
    "Unit 17,",
    "Riverside Way,",
    "Ravensthorpe Industrial Estate,",
    "Dewsbury, West Yorkshire,",
    "WF13 3LG",
    "E-mail: hello@greenecs.co.uk",
    "Tel: 01924 494005"
  ];

  const objects = [];
  const addObject = (body) => {
    objects.push(body);
    return objects.length;
  };
  const fontId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const boldFontId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");
  const pageCommands = [];
  const textAt = (commands, x, y, text, size = 10, bold = false) => {
    commands.push(`BT /${bold ? "F2" : "F1"} ${size} Tf ${x} ${y} Td (${pdfEscape(text)}) Tj ET`);
  };
  const lineAt = (commands, x1, y1, x2, y2) => {
    commands.push(`0.55 w ${x1} ${y1} m ${x2} ${y2} l S`);
  };
  const footerAt = (commands) => {
    let y = 56;
    companyFooter.forEach((line) => {
      textAt(commands, margin, y, line, 7);
      y -= 10;
    });
  };

  {
    const commands = [];
    textAt(commands, margin, 742, "GREENACRE", 28, true);
    textAt(commands, margin + 4, 718, "ENVIRONMENTAL SYSTEMS LTD", 12);
    lineAt(commands, margin, 695, pageWidth - margin, 695);
    textAt(commands, margin, 646, "Leaders in Industrial Ventilation Equipment", 16, true);
    textAt(commands, margin, 536, "https://greenecs.co.uk/", 10);
    textAt(commands, margin, 518, "https://oda-bloc.com/", 10);
    wrapPdfText("Compact extraction and odour control modular units for hire and sale across the UK. ATEX-rated.", 45).forEach((line, index) => {
      textAt(commands, margin, 452 - index * 14, line, 11);
    });
    wrapPdfText("Ensure compliance with HSE and EEA requirements including DSEAR, COSHH and odour control.", 45).forEach((line, index) => {
      textAt(commands, margin, 378 - index * 14, line, 11);
    });
    footerAt(commands);
    pageCommands.push(commands.join("\n"));
  }

  {
    const commands = [];
    let y = 764;
    textAt(commands, pageWidth - 190, y, `Date: ${pdfDate(quote?.date || todayIso())}`, 10);
    y -= 18;
    textAt(commands, pageWidth - 190, y, "Reference: PN/APP", 10);
    y -= 18;
    textAt(commands, pageWidth - 190, y, `Quote Number: ${quote?.number || "Draft"}`, 10, true);
    y = 742;
    textAt(commands, margin, y, client?.name || "Client", 10, true);
    y -= 15;
    clientAddress.slice(0, 6).forEach((line) => {
      textAt(commands, margin, y, line, 10);
      y -= 15;
    });
    y -= 18;
    textAt(commands, margin, y, `For the Attention of ${attention}.`, 10);
    y -= 34;
    textAt(commands, margin, y, `Dear ${attention.split(" ")[0] || "Sir / Madam"},`, 10);
    y -= 32;
    textAt(commands, margin, y, `RE: ${subject}.`, 10, true);
    y -= 30;
    offerLines.forEach((line) => {
      textAt(commands, margin, y, line, 10);
      y -= 14;
    });
    y -= 18;
    textAt(commands, margin, y, "Works included", 10, true);
    y -= 18;
    quoteLines.slice(0, 10).forEach((row) => {
      const description = [row.description, row.freeText].filter(Boolean).join(" - ");
      wrapPdfText(`- ${round(row.qty, 2)} x ${description}`, 82).forEach((line) => {
        textAt(commands, margin + 12, y, line, 10);
        y -= 14;
      });
    });
    y -= 18;
    textAt(commands, margin, y, "Our budget cost for the above works is:", 10, true);
    y -= 28;
    textAt(commands, margin, y, `${pdfMoney(total)} Ex VAT`, 16, true);
    y -= 38;
    noteLines.forEach((line) => {
      textAt(commands, margin, y, line, 10);
      y -= 14;
    });
    y -= 22;
    textAt(commands, margin, y, "Yours sincerely,", 10);
    y -= 52;
    textAt(commands, margin, y, "Greenacre Environmental Systems Ltd", 10, true);
    if (exclusionLines.length && exclusionLines[0] !== "No exclusions entered.") {
      y -= 28;
      textAt(commands, margin, y, "Exclusions", 10, true);
      y -= 16;
      exclusionLines.slice(0, 5).forEach((line) => {
        textAt(commands, margin, y, line, 9);
        y -= 12;
      });
    }
    pageCommands.push(commands.join("\n"));
  }

  {
    const commands = [];
    let y = 730;
    companyAddress.forEach((line) => {
      textAt(commands, margin, y, line, 12);
      y -= 20;
    });
    y -= 30;
    textAt(commands, margin, y, "Quotation Summary", 12, true);
    y -= 20;
    textAt(commands, margin, y, `Quote: ${quote?.number || "Draft"} Rev ${quote?.revision ?? 0}`, 10);
    y -= 16;
    textAt(commands, margin, y, `Client: ${client?.name || ""}`, 10);
    y -= 16;
    textAt(commands, margin, y, `Site: ${site?.name || ""}`, 10);
    y -= 16;
    textAt(commands, margin, y, `Asset: ${asset?.name || ""}`, 10);
    y -= 16;
    textAt(commands, margin, y, `Total: ${pdfMoney(total)} Ex VAT`, 10, true);
    footerAt(commands);
    pageCommands.push(commands.join("\n"));
  }

  const pageIds = [];
  pageCommands.forEach((text) => {
    const contentId = addObject(`<< /Length ${text.length} >>\nstream\n${text}\nendstream`);
    const pageId = addObject(`<< /Type /Page /Parent 0 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontId} 0 R /F2 ${boldFontId} 0 R >> >> /Contents ${contentId} 0 R >>`);
    pageIds.push(pageId);
  });
  const pagesId = addObject(`<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`);
  pageIds.forEach((id) => {
    objects[id - 1] = objects[id - 1].replace("/Parent 0 0 R", `/Parent ${pagesId} 0 R`);
  });
  const catalogId = addObject(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);

  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((body, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return new Blob([pdf], { type: "application/pdf" });
}

function quotePrintDocumentHtml(quote = activeQuote()) {
  const { enquiry, client, contact, site, asset } = quoteContext();
  const results = currentResults || calculate();
  const quoteLines = quoteLinesForPdf(quote, results);
  const total = quoteLines.reduce((sum, row) => sum + n(row.qty) * n(row.price ?? row.unit), 0) || quote?.total || results.bom.sell;
  const clientAddress = pdfAddressLines(client?.address, []);
  const attention = contact?.name || "Sir / Madam";
  const firstName = attention.includes(" ") ? attention.split(" ")[0] : attention;
  const subject = [site?.name, asset?.name].filter(Boolean).join(" - ") || enquiry?.notes || "Odour control works";
  const baseHref = location.href.slice(0, location.href.lastIndexOf("/") + 1);
  const offerText = quote?.offerText || "Following our recent discussions in connection with the above we now have pleasure in confirming our costs to carry out the following tests and checks.";
  const notesText = quote?.notes || "We trust that this is in line with your immediate requirements, however, should you need any additional information please do not hesitate to call.";
  const lineItems = quoteLines.map((row) => {
    const description = [row.description, row.freeText].filter(Boolean).join(" - ");
    return `<li>${escapeHtml(round(row.qty, 2))} x ${escapeHtml(description)}</li>`;
  }).join("");
  const clientAddressHtml = clientAddress.map((line) => `${escapeHtml(line)}<br>`).join("");
  const exclusionsHtml = quote?.exclusions ? `
    <section class="quote-extra">
      <h3>Exclusions</h3>
      <p>${escapeHtml(quote.exclusions)}</p>
    </section>
  ` : "";
  return `<!doctype html>
<html>
<head>
  <base href="${escapeHtml(baseHref)}">
  <title>${escapeHtml(quote?.number || "Draft quotation")}</title>
  <style>
    @page { size: A4; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; background: #e9ece8; color: #1f2c27; font-family: Arial, Helvetica, sans-serif; }
    .toolbar { position: sticky; top: 0; z-index: 5; display: flex; justify-content: flex-end; gap: 8px; padding: 10px 16px; background: #f7faf6; border-bottom: 1px solid #d8ded8; }
    .toolbar button { border: 0; border-radius: 4px; background: #477c28; color: white; padding: 8px 12px; font: 600 13px Arial, sans-serif; cursor: pointer; }
    .page { width: 210mm; min-height: 297mm; margin: 12px auto; padding: 20mm 17mm 18mm; background: white; position: relative; box-shadow: 0 8px 26px rgba(0,0,0,.16); page-break-after: always; overflow: hidden; }
    .page:last-child { page-break-after: auto; }
    .logo { display: block; object-fit: contain; }
    .cover-logo { position: absolute; left: 74mm; top: 13mm; width: 72mm; height: 33mm; }
    .cover-title { position: absolute; left: 12mm; top: 54mm; width: 186mm; height: 9mm; display: flex; align-items: center; justify-content: center; background: white; font-size: 24px; font-weight: 700; color: #538135; }
    .cover-image { position: absolute; object-fit: cover; display: block; }
    .cover-image-oda { left: 98mm; top: 68mm; width: 66mm; height: 68mm; }
    .cover-image-filter { left: 47mm; top: 140mm; width: 75mm; height: 83mm; }
    .cover-certificates { position: absolute; left: 14mm; top: 231mm; width: 116mm; height: 20mm; object-fit: contain; display: block; }
    .cover-links { position: absolute; left: 126mm; top: 232mm; width: 52mm; line-height: 1.85; font-size: 13px; color: #0563c1; text-align: center; }
    .cover-copy { position: absolute; width: 67mm; min-height: 31mm; padding: 2mm 3mm; background: white; font-size: 14.5px; font-weight: 700; line-height: 1.28; color: #111; }
    .cover-copy-hire { left: 17mm; top: 84mm; }
    .cover-copy-compliance { left: 130mm; top: 162mm; }
    .company-note { position: absolute; left: 15mm; right: 15mm; bottom: 22mm; font-size: 8.5px; line-height: 1.32; color: #538135; text-align: center; }
    .doc-header { position: absolute; left: 17mm; right: 17mm; top: 9mm; height: 30mm; border-bottom: 1px solid #d8ded8; }
    .doc-header .logo { width: 70mm; height: 26mm; margin: 0 auto; object-fit: contain; }
    .doc-footer { position: absolute; left: 15mm; right: 15mm; bottom: 10mm; border-top: 1px solid #d8ded8; padding-top: 3mm; font-size: 8.5px; line-height: 1.32; color: #538135; text-align: center; }
    .letter-head { position: relative; min-height: 38mm; margin-top: 36mm; font-size: 13px; line-height: 1.5; }
    .address { position: absolute; right: 0; top: 0; width: 78mm; text-align: right; }
    .meta { width: 70mm; margin-left: 0; }
    .letter-body { margin-top: 15mm; padding-bottom: 26mm; font-size: 13px; line-height: 1.5; }
    .letter-body p { margin: 0 0 14px; }
    .re-line { margin-top: 14px; font-weight: 700; }
    .works-title { margin: 22px 0 8px; font-weight: 700; }
    ul { margin: 0 0 18px 19px; padding: 0; }
    li { margin: 0 0 6px; }
    .cost-label { margin-top: 17px; font-weight: 700; }
    .cost { margin-top: 10px; font-size: 17px; font-weight: 700; }
    .signoff { margin-top: 20px; }
    .signature-space { height: 7mm; }
    .extra-page-title { margin-top: 45mm; font-size: 13px; font-weight: 700; }
    .contact-block { margin-top: 52mm; font-size: 10.5px; line-height: 1.55; color: #538135; }
    .quote-extra { margin-top: 10mm; font-size: 12px; line-height: 1.4; }
    .quote-extra h3 { margin: 0 0 6px; font-size: 13px; }
    @media print {
      body { background: white; }
      .toolbar { display: none; }
      .page { margin: 0; box-shadow: none; }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <button onclick="window.print()">Print / Save PDF</button>
    <button onclick="window.close()">Close</button>
  </div>
  <section class="page cover">
    <img class="logo cover-logo" src="greenacre-logo.png" alt="Greenacre Environmental Systems Ltd">
    <div class="cover-title">Leaders in Industrial Ventilation Equipment</div>
    <img class="cover-image cover-image-oda" src="Picture1.png" alt="Oda Bloc odour control unit">
    <img class="cover-image cover-image-filter" src="Picture2.png" alt="Industrial odour control filter">
    <img class="cover-certificates" src="Picture3.png" alt="Greenacre accreditations">
    <div class="cover-copy cover-copy-hire">
      Compact extraction and odour control modular unit for hire and sale across the UK. ATEX-rated.
    </div>
    <div class="cover-copy cover-copy-compliance">
      Ensure compliance with HSE &amp; EEA requirements including DSEAR, COSHH and odour control.
    </div>
    <div class="cover-links">
      https://greenecs.co.uk/<br>
      https://oda-bloc.com/
    </div>
    <div class="company-note">
      Greenacre Environmental Systems Limited is registered in England and Wales with company number 09117078. The company is also registered for VAT number 191 4495 86.<br>
      Unit 17, Riverside Way, Ravensthorpe Industrial Estate, Dewsbury, West Yorkshire, WF13 3LG
    </div>
  </section>
  <section class="page letter">
    <header class="doc-header">
      <img class="logo" src="greenacre-logo.png" alt="Greenacre Environmental Systems Ltd">
    </header>
    <div class="letter-head">
      <div class="meta">
        Date: ${escapeHtml(pdfDate(quote?.date || todayIso()))}<br>
        Reference: PN/APP<br>
        Quote Number<br>
        <strong>${escapeHtml(quote?.number || "Draft")}</strong>
      </div>
      <div class="address">
        <strong>${escapeHtml(client?.name || "Client")}</strong><br>
        ${clientAddressHtml}
      </div>
    </div>
    <div class="letter-body">
      <p>For the Attention of ${escapeHtml(attention)}.</p>
      <p>Dear ${escapeHtml(firstName)},</p>
      <p class="re-line">RE: ${escapeHtml(subject)}.</p>
      <p>${escapeHtml(offerText)}</p>
      <p class="works-title">${escapeHtml(asset?.name || subject)}</p>
      <ul>${lineItems || `<li>${escapeHtml(subject)}</li>`}</ul>
      <p class="cost-label">Our budget cost for the above works is:</p>
      <p class="cost">${escapeHtml(moneyFmt.format(total))} Ex VAT</p>
      <p>${escapeHtml(notesText)}</p>
      ${exclusionsHtml}
      <p class="signoff">Yours sincerely,</p>
      <div class="signature-space"></div>
      <p><strong>Georgia Hobson</strong><br>
      Project Engineer<br>
      Greenacre Environmental Systems Ltd</p>
    </div>
    <footer class="doc-footer">
      Greenacre Environmental Systems Limited is registered in England and Wales with company number 09117078. The company is also registered for VAT number 191 4495 86.<br>
      Unit 17, Riverside Way, Ravensthorpe Industrial Estate, Dewsbury, West Yorkshire, WF13 3LG
    </footer>
  </section>
  <section class="page contact">
    <header class="doc-header">
      <img class="logo" src="greenacre-logo.png" alt="Greenacre Environmental Systems Ltd">
    </header>
    <div class="contact-block">
      Unit 17,<br>
      Riverside Way,<br>
      Ravensthorpe Industrial Estate,<br>
      Dewsbury, West Yorkshire,<br>
      WF13 3LG<br>
      E-mail: hello@greenecs.co.uk<br>
      Tel: 01924 494005
    </div>
    <footer class="doc-footer">
      Greenacre Environmental Systems Limited is registered in England and Wales with company number 09117078. The company is also registered for VAT number 191 4495 86.<br>
      Unit 17, Riverside Way, Ravensthorpe Industrial Estate, Dewsbury, West Yorkshire, WF13 3LG
    </footer>
  </section>
</body>
</html>`;
}

function openQuotePdfPreview(quote = activeQuote()) {
  const html = quotePrintDocumentHtml(quote);
  const popup = window.open("", "_blank");
  if (popup?.document) {
    popup.document.open();
    popup.document.write(html);
    popup.document.close();
    popup.focus();
    return;
  }
  const blob = new Blob([html], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  window.open(url, "_blank");
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function renderQuote() {
  const enquiry = activeEnquiry();
  const related = state.quotes.filter((quote) => quote.enquiryId === enquiry?.id);
  if (!state.activeQuoteId && related.length) state.activeQuoteId = related[related.length - 1].id;
  const quote = activeQuote();
  const displayQuote = quoteDraft || quote;
  renderQuoteSearch();
  setBlock("#quoteRegisterPanel", !quoteEditorOpen);
  setBlock("#quoteEditor", quoteEditorOpen);
  setBlock("#quoteRevisionDiffPanel", quoteEditorOpen);
  setBlock("#quotePreviewPanel", quoteEditorOpen);
  if (!quoteEditorOpen) return;
  $("#quoteNumber").value = displayQuote?.number || "";
  $("#quoteRevision").value = displayQuote?.revision ?? "";
  $("#quoteStatus").value = displayQuote?.status || "Draft";
  $("#quoteOfferText").value = displayQuote?.offerText || "";
  $("#quoteNotes").value = displayQuote?.notes || "";
  $("#quoteExclusions").value = displayQuote?.exclusions || "";
  $("#quoteHistory").innerHTML = related.length ? related.map((item) => `
    <article class="record-card">
      <strong>${escapeHtml(item.number)} Rev ${item.revision}</strong>
      <p>${escapeHtml(item.date)} | ${escapeHtml(item.status)} | ${moneyFmt.format(item.total || 0)}</p>
      <button data-select-quote="${item.id}" type="button">Open</button>
    </article>
  `).join("") : `<article class="record-card"><p>No quotes created for this enquiry yet.</p></article>`;
  renderQuoteProductOptions();
  renderQuoteLines(displayQuote);
  $("#quotePreview").innerHTML = displayQuote ? buildQuoteHtml(displayQuote) : `<p>Create a quote revision to preview the quotation.</p>`;
  renderQuoteRevisionDiff(related);
}

function renderQuoteLines(quote) {
  const body = $("#quoteLinesBody");
  if (!body) return;
  const lines = quote?.lines || [];
  body.innerHTML = lines.length ? lines.map((line, index) => `
    <tr>
      <td><input data-quote-line="${index}.description" value="${escapeHtml(line.description || "")}"></td>
      <td>${escapeHtml(line.supplier || "")}</td>
      <td><input class="num" data-quote-line="${index}.qty" type="number" step="0.01" value="${line.qty || 0}"></td>
      <td class="num">${line.basePrice === undefined ? "-" : moneyFmt.format(line.basePrice || 0)}</td>
      <td><input class="num" data-quote-line="${index}.price" type="number" step="0.01" value="${line.price || 0}"></td>
      <td class="num">${moneyFmt.format(n(line.qty) * n(line.price))}</td>
      <td><input data-quote-line="${index}.freeText" value="${escapeHtml(line.freeText || "")}"></td>
      <td>
        ${line.childParts?.length ? `<button data-toggle-quote-line-children="${index}" type="button">${line.showChildren ? "Hide BOM" : "Show BOM"}</button>` : ""}
        <button class="danger-btn" data-remove-quote-line="${index}" type="button">Remove</button>
      </td>
    </tr>
    ${line.showChildren ? renderQuoteLineChildRows(line) : ""}
  `).join("") : `<tr><td colspan="8">No quote lines added yet.</td></tr>`;
}

function renderQuoteLineChildRows(line) {
  return `
    <tr class="child-row">
      <td colspan="8">
        <strong>Child BOM reference only - quoted as parent line above</strong>
        <table class="nested-table">
          <thead><tr><th>Description</th><th>Size</th><th>Qty</th><th>Unit</th><th>Total</th><th>Supplier</th></tr></thead>
          <tbody>${(line.childParts || []).map((part) => `
            <tr><td>${escapeHtml(part.description || "")}</td><td>${escapeHtml(part.size || "")}</td><td class="num">${round(part.qty, 3)}</td><td class="num">${moneyFmt.format(part.unit || 0)}</td><td class="num">${moneyFmt.format(part.total || 0)}</td><td>${escapeHtml(part.supplier || "")}</td></tr>
          `).join("")}</tbody>
        </table>
      </td>
    </tr>
  `;
}

function renderQuoteProductOptions() {
  const select = $("#quoteProductSelect");
  if (!select) return;
  const current = select.value;
  const products = (state.products || []).filter((product) => product.parentProduct);
  select.innerHTML = optionHtml("", "Select parent product", current) + products.map((product) => (
    optionHtml(product.productCode, `${product.productCode} - ${product.description} (${moneyFmt.format(product.price || 0)})`, current)
  )).join("");
}

function quoteComparable(quote) {
  return {
    status: quote?.status || "",
    notes: quote?.notes || "",
    total: moneyFmt.format(quote?.total || 0)
  };
}

function renderQuoteRevisionDiff(relatedQuotes) {
  const target = $("#quoteRevisionDiff");
  if (!target) return;
  const sorted = [...relatedQuotes].sort((a, b) => Number(a.revision) - Number(b.revision));
  if (sorted.length < 2) {
    target.innerHTML = `<article class="record-card"><p>No previous revision to compare yet.</p></article>`;
    return;
  }

  target.innerHTML = sorted.slice(1).map((quote, index) => {
    const previous = sorted[index];
    const current = quoteComparable(quote);
    const prior = quoteComparable(previous);
    const changes = Object.keys(current).filter((key) => current[key] !== prior[key]);
    return `
      <article class="record-card">
        <strong>${escapeHtml(quote.number)} Rev ${escapeHtml(previous.revision)} to Rev ${escapeHtml(quote.revision)}</strong>
        ${changes.length ? changes.map((key) => `
          <p><strong>${escapeHtml(key)}:</strong> ${escapeHtml(prior[key] || "Blank")} -> ${escapeHtml(current[key] || "Blank")}</p>
        `).join("") : `<p>No tracked changes between these revisions.</p>`}
      </article>
    `;
  }).join("");
}

function includesText(value, query) {
  return !query || String(value || "").toLowerCase().includes(query);
}

function renderQuoteSearch() {
  const body = $("#quoteSearchBody");
  if (!body) return;
  const numberQuery = ($("#quoteSearchNumber")?.value || "").trim().toLowerCase();
  const clientQuery = ($("#quoteSearchClient")?.value || "").trim().toLowerCase();
  const siteQuery = ($("#quoteSearchSite")?.value || "").trim().toLowerCase();
  const assetQuery = ($("#quoteSearchAsset")?.value || "").trim().toLowerCase();
  const rows = latestQuoteRevisions(state.quotes)
    .map(quoteRecord)
    .filter(({ quote, client, site, asset }) => (
      includesText(quote.number, numberQuery) &&
      includesText(client?.name, clientQuery) &&
      includesText(site?.name, siteQuery) &&
      includesText(asset?.name, assetQuery)
    ))
    .sort((a, b) => String(b.quote.date || "").localeCompare(String(a.quote.date || "")) || String(b.quote.number || "").localeCompare(String(a.quote.number || "")));

  body.innerHTML = rows.length ? rows.map(({ quote, enquiry, client, site, asset }) => `
    <tr>
      <td><strong>${escapeHtml(quote.number || "")}</strong></td>
      <td>${escapeHtml(quote.revision ?? "")}</td>
      <td>${escapeHtml(quote.status || "")}</td>
      <td>${escapeHtml(quote.date || "")}</td>
      <td>${escapeHtml(client?.name || "")}</td>
      <td>${escapeHtml(site?.name || "")}</td>
      <td>${escapeHtml(asset?.name || "")}</td>
      <td>${escapeHtml(enquiry?.number || "")}</td>
      <td class="num">${moneyFmt.format(quote.total || 0)}</td>
      <td><button data-select-quote="${escapeHtml(quote.id)}" type="button">Open</button></td>
    </tr>
  `).join("") : `<tr><td colspan="10">No matching quotes found.</td></tr>`;
}

function bindDynamicTableEvents() {
  document.body.addEventListener("input", (event) => {
    if (event.target.dataset.assetDetailField) {
      updateOpenedAssetField(event.target);
    }
  });

  document.body.addEventListener("change", (event) => {
    const edit = event.target.dataset.edit;
    const productEdit = event.target.dataset.productEdit;

    if (event.target.dataset.assetDetailField) {
      updateOpenedAssetField(event.target, true);
      return;
    }

    if (edit) {
      const [kind, index, key] = edit.split(".");
      const collection = kind === "zone" ? state.zones : state.legs;
      const value = event.target.type === "number" ? n(event.target.value) : event.target.value;
      collection[Number(index)][key] = value;
      update();
    }

    if (productEdit) {
      const [index, key] = productEdit.split(".");
      state.products[Number(index)][key] = key === "price" ? n(event.target.value) : event.target.value;
      update();
    }

    const quoteLineEdit = event.target.dataset.quoteLine;
    if (quoteLineEdit && quoteDraft) {
      const [index, key] = quoteLineEdit.split(".");
      quoteDraft.lines ||= [];
      if (!quoteDraft.lines[Number(index)]) return;
      quoteDraft.lines[Number(index)][key] = ["description", "freeText"].includes(key) ? event.target.value : n(event.target.value);
      quoteDraft.total = quoteDraft.lines.reduce((sum, line) => sum + n(line.qty) * n(line.price), 0);
      renderQuoteLines(quoteDraft);
      $("#quotePreview").innerHTML = buildQuoteHtml(quoteDraft);
    }
  });

  document.body.addEventListener("click", async (event) => {
    if (event.target?.id === "addAssetFieldBtn") {
      createAssetCustomField();
      return;
    }
    const interactiveTarget = event.target?.closest?.("button, a, input, select, textarea, label");
    const clickableRow = event.target?.closest?.("tr.clickable-row");
    if (!interactiveTarget && clickableRow) {
      if (clickableRow.dataset.rowOpenClient !== undefined) {
        state.activeClientId = clickableRow.dataset.rowOpenClient;
        const client = findClient(state.activeClientId);
        clientDraft = client ? clone(client) : blankClientDraft();
        clientEditorOpen = true;
        clientDrillLevel = "client";
        activeClientSiteId = "";
        activeClientAssetId = "";
        update();
        return;
      }
      if (clickableRow.dataset.rowOpenSite !== undefined) {
        updateClientDraftFromForm();
        activeClientSiteId = clickableRow.dataset.rowOpenSite;
        activeClientAssetId = "";
        clientDrillLevel = "site";
        update();
        return;
      }
      if (clickableRow.dataset.rowOpenAsset !== undefined) {
        updateClientDraftFromForm();
        activeClientSiteId = clickableRow.dataset.rowOpenAssetSite || activeClientSiteId;
        activeClientAssetId = clickableRow.dataset.rowOpenAsset;
        clientDrillLevel = "asset";
        update();
        return;
      }
    }
    const zoneIndex = event.target.dataset.removeZone;
    const legIndex = event.target.dataset.removeLeg;
    const productIndex = event.target.dataset.removeProduct;
    const contactId = event.target.dataset.removeContact;
    if (zoneIndex !== undefined) {
      state.zones.splice(Number(zoneIndex), 1);
      update();
    }
    if (legIndex !== undefined) {
      state.legs.splice(Number(legIndex), 1);
      update();
    }
    if (productIndex !== undefined) {
      state.products.splice(Number(productIndex), 1);
      update();
    }
    const productChildrenCode = event.target.dataset.toggleProductChildren;
    if (productChildrenCode !== undefined) {
      expandedProductCode = expandedProductCode === productChildrenCode ? "" : productChildrenCode;
      renderProducts();
    }
    const quoteLineIndex = event.target.dataset.removeQuoteLine;
    if (quoteLineIndex !== undefined && quoteDraft) {
      quoteDraft.lines = (quoteDraft.lines || []).filter((_, index) => index !== Number(quoteLineIndex));
      quoteDraft.total = quoteDraft.lines.reduce((sum, line) => sum + n(line.qty) * n(line.price), 0);
      update();
    }
    const quoteLineChildrenIndex = event.target.dataset.toggleQuoteLineChildren;
    if (quoteLineChildrenIndex !== undefined && quoteDraft) {
      const line = quoteDraft.lines?.[Number(quoteLineChildrenIndex)];
      if (line) {
        line.showChildren = !line.showChildren;
        renderQuoteLines(quoteDraft);
      }
    }
    if (contactId !== undefined) {
      if (!requireAdminRemoval()) return;
      const client = clientDraft || findClient(state.activeClientId);
      if (client) {
        client.contacts = (client.contacts || []).filter((contact) => contact.id !== contactId);
        const enquiry = activeEnquiry();
        if (enquiry?.contactId === contactId) enquiry.contactId = "";
        update();
      }
    }
    const removeClientId = event.target.dataset.removeClient;
    if (removeClientId !== undefined) {
      if (!requireAdminRemoval()) return;
      state.clients = (state.clients || []).filter((client) => client.id !== removeClientId);
      if (state.activeClientId === removeClientId) state.activeClientId = state.clients[0]?.id || "";
      clientEditorOpen = false;
      clientDraft = null;
      clientDrillLevel = "register";
      activeClientSiteId = "";
      activeClientAssetId = "";
      update();
    }
    const siteId = event.target.dataset.selectClientSite;
    if (siteId !== undefined) {
      updateClientDraftFromForm();
      activeClientSiteId = siteId;
      activeClientAssetId = "";
      clientDrillLevel = "site";
      update();
    }
    const removeSiteId = event.target.dataset.removeClientSite;
    if (removeSiteId !== undefined && clientDraft) {
      if (!requireAdminRemoval()) return;
      clientDraft.sites = (clientDraft.sites || []).filter((site) => site.id !== removeSiteId);
      if (activeClientSiteId === removeSiteId) {
        activeClientSiteId = clientDraft.sites[0]?.id || "";
        activeClientAssetId = clientDraft.sites[0]?.assets?.[0]?.id || "";
        clientDrillLevel = activeClientSiteId ? "client" : "client";
      }
      update();
    }
    const assetId = event.target.dataset.selectClientAsset;
    if (assetId !== undefined) {
      updateClientDraftFromForm();
      activeClientAssetId = assetId;
      clientDrillLevel = "asset";
      update();
    }
    const removeAssetId = event.target.dataset.removeClientAsset;
    if (removeAssetId !== undefined && clientDraft) {
      if (!requireAdminRemoval()) return;
      const removeAssetSiteId = event.target.dataset.removeClientAssetSite || activeClientSiteId;
      const site = clientDraft.sites?.find((item) => item.id === removeAssetSiteId);
      if (site) {
        site.assets = (site.assets || []).filter((asset) => asset.id !== removeAssetId);
        if (activeClientAssetId === removeAssetId) activeClientAssetId = site.assets[0]?.id || "";
        if (!activeClientAssetId) clientDrillLevel = "site";
        update();
      }
    }
    const removeAssetFieldId = event.target.dataset.removeAssetField;
    if (removeAssetFieldId !== undefined) {
      if (!requireAdminRemoval()) return;
      alert("This action has been disabled on the asset screen. Use Clear Value to remove data from this asset only.");
      return;
    }
    const clearCustomFieldId = event.target.dataset.clearCustomAssetField;
    if (clearCustomFieldId !== undefined) {
      if (!requireAdminRemoval()) return;
      clearCustomAssetFieldValue(clearCustomFieldId);
    }
    const deleteLocalAssetFieldId = event.target.dataset.deleteLocalAssetField;
    if (deleteLocalAssetFieldId !== undefined) {
      if (!requireAdminRemoval()) return;
      deleteLocalAssetField(deleteLocalAssetFieldId);
    }
    if (event.target.id === "uploadAssetAttachmentBtn") {
      const file = $("#assetAttachmentInput")?.files?.[0];
      const site = clientDraft?.sites?.find((item) => item.id === activeClientSiteId);
      const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
      if (!asset || !file) return;
      const message = $("#assetAttachmentMessage");
      try {
        const uploadFile = await prepareAssetAttachment(file, (text) => {
          if (message) message.textContent = text;
        });
        if (message) {
          const compressed = uploadFile !== file;
          message.textContent = compressed
            ? `Uploading compressed photo (${attachmentSizeLabel(uploadFile.size)})...`
            : `Uploading ${attachmentSizeLabel(uploadFile.size)}...`;
        }
        const form = new FormData();
        form.append("file", uploadFile);
        await apiFetch(`/api/assets/${encodeURIComponent(asset.id)}/attachments`, { method: "POST", body: form });
        if ($("#assetAttachmentInput")) $("#assetAttachmentInput").value = "";
        await renderAssetAttachments(asset);
        const refreshedMessage = $("#assetAttachmentMessage");
        if (refreshedMessage) refreshedMessage.textContent = "Upload complete.";
      } catch (error) {
        if (message) message.textContent = error.message || "Upload failed.";
      }
    }
    const deleteAttachmentId = event.target.dataset.deleteAssetAttachment;
    if (deleteAttachmentId !== undefined) {
      if (!requireAdminRemoval()) return;
      const site = clientDraft?.sites?.find((item) => item.id === activeClientSiteId);
      const asset = site?.assets?.find((item) => item.id === activeClientAssetId);
      if (!asset) return;
      await apiFetch(`/api/asset-attachments/${encodeURIComponent(asset.id)}/${encodeURIComponent(deleteAttachmentId)}`, { method: "DELETE" });
      await renderAssetAttachments(asset);
    }
    const reminderFieldId = event.target.dataset.createReminder;
    if (reminderFieldId !== undefined) {
      createDueDateReminder(reminderFieldId);
    }
    const resetUserPasswordId = event.target.dataset.resetUserPassword;
    if (resetUserPasswordId !== undefined) {
      await resetAdminUserPassword(resetUserPasswordId);
    }
    const clientId = event.target.dataset.selectClient;
    if (clientId !== undefined) {
      state.activeClientId = clientId;
      const client = findClient(clientId);
      clientDraft = client ? clone(client) : blankClientDraft();
      clientEditorOpen = true;
      clientDrillLevel = "client";
      activeClientSiteId = "";
      activeClientAssetId = "";
      $("#clientContactName").value = "";
      $("#clientContactRole").value = "";
      $("#clientContactEmail").value = "";
      $("#clientContactPhone").value = "";
      update();
      setActiveTab("clients");
    }
    const openClientAssetId = event.target.dataset.openClientAsset;
    if (openClientAssetId !== undefined) {
      updateClientDraftFromForm();
      activeClientSiteId = event.target.dataset.openClientAssetSite || activeClientSiteId;
      activeClientAssetId = openClientAssetId;
      clientDrillLevel = "asset";
      update();
    }
    const supplierId = event.target.dataset.selectSupplier;
    if (supplierId !== undefined) {
      const supplier = state.suppliers.find((item) => item.id === supplierId);
      supplierDraft = supplier ? clone(supplier) : blankSupplier();
      supplierEditorOpen = true;
      ["supplierContactName", "supplierContactRole", "supplierContactEmail", "supplierContactPhone", "supplierSiteName", "supplierSiteAddress", "supplierSitePhone", "supplierSiteEmail"].forEach((id) => {
        const input = $(`#${id}`);
        if (input) input.value = "";
      });
      update();
      setActiveTab("suppliers");
    }
    const removeSupplierContactId = event.target.dataset.removeSupplierContact;
    if (removeSupplierContactId !== undefined && supplierDraft) {
      supplierDraft.contacts = (supplierDraft.contacts || []).filter((contact) => contact.id !== removeSupplierContactId);
      update();
    }
    const removeSupplierSiteId = event.target.dataset.removeSupplierSite;
    if (removeSupplierSiteId !== undefined && supplierDraft) {
      supplierDraft.sites = (supplierDraft.sites || []).filter((site) => site.id !== removeSupplierSiteId);
      update();
    }
    const enquiryId = event.target.dataset.selectEnquiry;
    if (enquiryId !== undefined) {
      state.activeEnquiryId = enquiryId;
      enquiryEditorOpen = true;
      syncProjectFromEnquiry();
      update();
      setActiveTab("enquiry");
    }
    const quoteId = event.target.dataset.selectQuote;
    if (quoteId !== undefined) {
      const quote = state.quotes.find((item) => item.id === quoteId);
      state.activeQuoteId = quoteId;
      if (quote?.enquiryId) {
        state.activeEnquiryId = quote.enquiryId;
        syncProjectFromEnquiry();
      }
      quoteDraft = quote ? clone(quote) : null;
      quoteEditorOpen = true;
      update();
      setActiveTab("quote");
    }
    if (event.target.dataset.backToSearch !== undefined) {
      activeAssetHistory = null;
      renderSearchResults();
    }
    const assetHistoryAssetId = event.target.dataset.assetHistoryAsset;
    if (assetHistoryAssetId !== undefined) {
      activeAssetHistory = {
        clientId: event.target.dataset.assetHistoryClient || "",
        siteId: event.target.dataset.assetHistorySite || "",
        assetId: assetHistoryAssetId
      };
      renderSearchResults();
    }
  });
}

function bindTabs() {
  document.querySelectorAll(".tabs button").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.tab === "clients") {
        showClientRegisterHome();
      }
      setActiveTab(button.dataset.tab);
      update();
    });
  });
}

function bindDesignTabs() {
  document.querySelectorAll(".design-tabs button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".design-tabs button").forEach((item) => {
        item.classList.toggle("active", item === button);
      });
      document.querySelectorAll(".design-panel").forEach((panel) => {
        panel.classList.toggle("active", panel.id === button.dataset.designSection);
      });
    });
  });
}

function saveEnquiryFromForm() {
  const enquiry = activeEnquiry();
  if (!enquiry || !$("#enquiryStatus")) return;
  enquiry.status = $("#enquiryStatus").value;
  enquiry.date = $("#enquiryDate").value;
  enquiry.clientId = $("#enquiryClient").value;
  enquiry.contactId = $("#enquiryContact").value;
  enquiry.siteId = $("#enquirySite").value;
  enquiry.assetId = $("#enquiryAsset").value;
  enquiry.siteRef = $("#enquirySiteRef").value;
  enquiry.assetRef = $("#enquiryAssetRef").value;
  enquiry.notes = $("#enquiryNotes").value;
  const client = findClient(enquiry.clientId);
  const site = findSite(client, enquiry.siteId);
  const asset = findAsset(site, enquiry.assetId);
  if (site && !enquiry.siteRef) enquiry.siteRef = site.reference || "";
  if (asset && !enquiry.assetRef) enquiry.assetRef = asset.reference || "";
  syncProjectFromEnquiry();
}

function saveQuoteFromForm() {
  if (!quoteDraft || !$("#quoteStatus")) return;
  quoteDraft.status = $("#quoteStatus").value;
  quoteDraft.offerText = $("#quoteOfferText").value;
  quoteDraft.notes = $("#quoteNotes").value;
  quoteDraft.exclusions = $("#quoteExclusions").value;
  quoteDraft.lines ||= [];
  quoteDraft.total = quoteDraft.lines.length ? quoteDraft.lines.reduce((sum, line) => sum + n(line.qty) * n(line.price), 0) : currentResults?.bom?.sell || quoteDraft.total || 0;
  const index = state.quotes.findIndex((quote) => quote.id === quoteDraft.id);
  if (index >= 0) {
    state.quotes[index] = clone(quoteDraft);
  } else {
    state.quotes.push(clone(quoteDraft));
  }
  state.activeQuoteId = quoteDraft.id;
}

function updateQuoteDraftFromForm() {
  if (!quoteDraft || !$("#quoteStatus")) return;
  quoteDraft.status = $("#quoteStatus").value;
  quoteDraft.offerText = $("#quoteOfferText").value;
  quoteDraft.notes = $("#quoteNotes").value;
  quoteDraft.exclusions = $("#quoteExclusions").value;
  quoteDraft.lines ||= [];
  quoteDraft.total = quoteDraft.lines.length ? quoteDraft.lines.reduce((sum, line) => sum + n(line.qty) * n(line.price), 0) : currentResults?.bom?.sell || quoteDraft.total || 0;
  $("#quotePreview").innerHTML = buildQuoteHtml(quoteDraft);
}

function updateClientDraftFromForm() {
  if (!clientDraft) return;
  if ($("#clientName")) clientDraft.name = $("#clientName").value.trim();
  if ($("#clientAddress")) clientDraft.address = $("#clientAddress").value.trim();
  clientDraft.contacts ||= [];
  clientDraft.sites ||= [];
  let site = clientDraft.sites.find((item) => item.id === activeClientSiteId);
  if (site && $("#clientSiteRef") && $("#clientSiteName") && $("#clientSiteAddress")) {
    Object.assign(site, {
      reference: $("#clientSiteRef").value.trim(),
      name: $("#clientSiteName").value.trim(),
      address: $("#clientSiteAddress").value.trim(),
      assets: site.assets || []
    });
  }
  if (!site) return;
  let asset = site.assets.find((item) => item.id === activeClientAssetId);
  const assetFieldIds = [
    "clientAssetRef", "clientAssetName", "clientAssetLocation", "assetFlowRate", "assetFanModel", "assetFanSerial",
    "assetFlowSensorModel", "assetFlowSensorSerial", "assetFlowSetpoint", "assetCarbonMediaType", "assetCarbonMediaSl",
    "assetCarbonMediaVolume", "assetCarbonPressureDrop", "assetCarbonHighSetpoint", "assetCarbonHighHighSetpoint",
    "assetCarbonTempProbeModel", "assetCarbonTempProbeSerial", "assetBioMediaType", "assetBioMediaVolume"
  ];
  const hasAssetValues = assetFieldIds.some((id) => $(`#${id}`) && String($(`#${id}`).value || "").trim() && Number($(`#${id}`).value) !== 0);
  if (!asset && hasAssetValues) {
    asset = blankAsset(site.assets.length + 1);
    site.assets.push(asset);
    activeClientAssetId = asset.id;
  }
  if (!asset) return;
  Object.assign(asset, {
    reference: $("#clientAssetRef").value.trim(),
    name: $("#clientAssetName").value.trim(),
    location: $("#clientAssetLocation").value.trim(),
    flowRate: $("#assetFlowRate").value.trim(),
    fanModel: $("#assetFanModel").value.trim(),
    fanSerial: $("#assetFanSerial").value.trim(),
    flowSensorModel: $("#assetFlowSensorModel").value.trim(),
    flowSensorSerial: $("#assetFlowSensorSerial").value.trim(),
    flowSetpointPercent: $("#assetFlowSetpoint").value.trim(),
    carbonMediaType: $("#assetCarbonMediaType").value.trim(),
    carbonMediaSlNumber: $("#assetCarbonMediaSl").value.trim(),
    carbonMediaVolume: $("#assetCarbonMediaVolume").value.trim(),
    carbonPressureDrop: $("#assetCarbonPressureDrop").value.trim(),
    carbonHighSetpoint: $("#assetCarbonHighSetpoint").value.trim(),
    carbonHighHighSetpoint: $("#assetCarbonHighHighSetpoint").value.trim(),
    carbonTempProbeModel: $("#assetCarbonTempProbeModel").value.trim(),
    carbonTempProbeSerial: $("#assetCarbonTempProbeSerial").value.trim(),
    bioMediaType: $("#assetBioMediaType").value.trim(),
    bioMediaVolume: $("#assetBioMediaVolume").value.trim()
  });
  asset.customData ||= {};
  [...(assetCustomFields || []), ...(asset.localCustomFields || [])].forEach((field) => {
    const input = document.querySelector(`[data-custom-asset-field="${CSS.escape(field.id)}"]`);
    if (input) asset.customData[field.id] = input.value;
  });
  renderAssetRegister();
}

function saveClientDraft() {
  updateClientDraftFromForm();
  if (!clientDraft) return;
  const index = state.clients.findIndex((client) => client.id === clientDraft.id);
  if (index >= 0) {
    state.clients[index] = clone(clientDraft);
  } else {
    state.clients.push(clone(clientDraft));
  }
  state.activeClientId = clientDraft.id;
  saveState();
}

function openAddSiteDialog() {
  if (!clientDraft) clientDraft = blankClientDraft();
  document.querySelectorAll(".modal-backdrop").forEach((modal) => modal.remove());
  const existing = clientDraft.sites || [];
  const overlay = document.createElement("div");
  overlay.className = "modal-backdrop";
  overlay.innerHTML = `
    <form class="modal-card" id="quickAddSiteForm">
      <div class="panel-head">
        <h2>Add Site</h2>
        <button id="cancelQuickAddSite" class="secondary-btn" type="button">Cancel</button>
      </div>
      <div class="form-grid">
        <label>Site reference <input id="quickSiteReference" required></label>
        <label>Site name <input id="quickSiteName" required></label>
        <label>Site address <input id="quickSiteAddress"></label>
      </div>
      <p id="quickAddSiteError" class="auth-error"></p>
      <div class="header-actions">
        <button type="submit">Save and Open</button>
      </div>
    </form>
  `;
  document.body.appendChild(overlay);
  const form = overlay.querySelector("#quickAddSiteForm");
  const referenceInput = overlay.querySelector("#quickSiteReference");
  const nameInput = overlay.querySelector("#quickSiteName");
  const addressInput = overlay.querySelector("#quickSiteAddress");
  const error = overlay.querySelector("#quickAddSiteError");
  const close = () => overlay.remove();
  referenceInput?.focus();
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
  overlay.querySelector("#cancelQuickAddSite")?.addEventListener("click", close);
  form?.addEventListener("submit", (event) => {
    event.preventDefault();
    const reference = referenceInput?.value.trim() || "";
    const name = nameInput?.value.trim() || "";
    if (!reference || !name) {
      if (error) error.textContent = "Enter a site reference and site name.";
      return;
    }
    const site = blankSite(existing.length + 1);
    site.reference = reference;
    site.name = name;
    site.address = addressInput?.value.trim() || "";
    site.assets = [];
    clientDraft.sites ||= [];
    clientDraft.sites.push(site);
    activeClientSiteId = site.id;
    activeClientAssetId = "";
    clientDrillLevel = "site";
    saveClientDraft();
    close();
    update();
  });
}

function openAddAssetDialog() {
  if (!clientDraft) clientDraft = blankClientDraft();
  updateClientDraftFromForm();
  const site = clientDraft.sites?.find((item) => item.id === activeClientSiteId);
  if (!site) {
    alert("Open a site before adding an asset.");
    return;
  }
  document.querySelectorAll(".modal-backdrop").forEach((modal) => modal.remove());
  const existing = site.assets || [];
  const overlay = document.createElement("div");
  overlay.className = "modal-backdrop";
  overlay.innerHTML = `
    <form class="modal-card" id="quickAddAssetForm">
      <div class="panel-head">
        <h2>Add Asset</h2>
        <button id="cancelQuickAddAsset" class="secondary-btn" type="button">Cancel</button>
      </div>
      <div class="form-grid">
        <label>Asset reference <input id="quickAssetReference" required></label>
        <label>Asset name <input id="quickAssetName" required></label>
        <label>Asset location <input id="quickAssetLocation"></label>
        <label>Flow rate m3/h <input id="quickAssetFlowRate"></label>
        <label>Fan <input id="quickAssetFan"></label>
      </div>
      <p id="quickAddAssetError" class="auth-error"></p>
      <div class="header-actions">
        <button type="submit">Save and Open</button>
      </div>
    </form>
  `;
  document.body.appendChild(overlay);
  const close = () => overlay.remove();
  overlay.querySelector("#quickAssetReference")?.focus();
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) close();
  });
  overlay.querySelector("#cancelQuickAddAsset")?.addEventListener("click", close);
  overlay.querySelector("#quickAddAssetForm")?.addEventListener("submit", (event) => {
    event.preventDefault();
    const reference = overlay.querySelector("#quickAssetReference")?.value.trim() || "";
    const name = overlay.querySelector("#quickAssetName")?.value.trim() || "";
    if (!reference || !name) {
      const error = overlay.querySelector("#quickAddAssetError");
      if (error) error.textContent = "Enter an asset reference and asset name.";
      return;
    }
    const asset = blankAsset(existing.length + 1);
    asset.reference = reference;
    asset.name = name;
    asset.location = overlay.querySelector("#quickAssetLocation")?.value.trim() || "";
    asset.flowRate = overlay.querySelector("#quickAssetFlowRate")?.value.trim() || "";
    asset.fanModel = overlay.querySelector("#quickAssetFan")?.value.trim() || "";
    site.assets ||= [];
    site.assets.push(asset);
    activeClientAssetId = asset.id;
    clientDrillLevel = "asset";
    saveClientDraft();
    close();
    update();
  });
}

function updateSupplierDraftFromForm() {
  if (!supplierDraft || !$("#supplierName")) return;
  supplierDraft.name = $("#supplierName").value.trim();
  supplierDraft.phone = $("#supplierPhone").value.trim();
  supplierDraft.email = $("#supplierEmail").value.trim();
  supplierDraft.invoiceAddress = $("#supplierInvoiceAddress").value.trim();
  supplierDraft.notes = $("#supplierNotes").value.trim();
  supplierDraft.contacts ||= [];
  supplierDraft.sites ||= [];
}

function saveSupplierDraft() {
  updateSupplierDraftFromForm();
  if (!supplierDraft) return;
  const index = state.suppliers.findIndex((supplier) => supplier.id === supplierDraft.id);
  if (index >= 0) {
    state.suppliers[index] = clone(supplierDraft);
  } else {
    state.suppliers.push(clone(supplierDraft));
  }
}

function createParentProductFromDesign() {
  const results = currentResults || calculate();
  const enquiry = activeEnquiry();
  const nameParts = [state.project.reference, state.project.site, state.project.item].filter(Boolean);
  const description = nameParts.length ? nameParts.join(" - ") : `Greenacre designed system ${todayIso()}`;
  const costPrice = (results.bom.materialsTotal || 0) + (results.bom.install || 0) + (results.bom.management || 0);
  const product = {
    productCode: nextParentProductCode(todayIso()),
    quantity: "1",
    description,
    size: `${round(results.designedFlow, 0)} m3/h`,
    costPrice,
    salePrice: results.bom.sell || 0,
    listPrice: costPrice,
    price: results.bom.sell || 0,
    category: "PARENT PRODUCT",
    discount: 0,
    material: "DESIGNED SYSTEM",
    supplier: "GREENACRE",
    notes: `Created from enquiry ${enquiry?.number || ""}`,
    parentProduct: true,
    childParts: results.bom.rows.map((row) => ({
      description: row.description,
      size: row.size,
      qty: row.qty,
      unit: row.unit,
      total: row.total,
      category: row.category,
      supplier: row.supplier,
      source: row.source
    }))
  };
  state.products.push(product);
  return product;
}

function bindActions() {
  $("#newEnquiryBtn").addEventListener("click", () => {
    const enquiry = blankEnquiry();
    state.enquiries.push(enquiry);
    state.activeEnquiryId = enquiry.id;
    enquiryEditorOpen = true;
    update();
  });

  $("#saveEnquiryBtn").addEventListener("click", () => {
    saveEnquiryFromForm();
    enquiryEditorOpen = false;
    update();
  });

  $("#cancelEnquiryBtn").addEventListener("click", () => {
    enquiryEditorOpen = false;
    update();
  });

  ["enquiryStatus", "enquiryDate", "enquiryClient", "enquiryContact", "enquirySite", "enquiryAsset", "enquirySiteRef", "enquiryAssetRef", "enquiryNotes"].forEach((id) => {
    $(`#${id}`).addEventListener("change", () => {
      saveEnquiryFromForm();
      update();
    });
  });

  $("#siteSearch").addEventListener("input", () => {
    activeAssetHistory = null;
    renderSearchResults();
  });
  $("#assetSearch").addEventListener("input", () => {
    activeAssetHistory = null;
    renderSearchResults();
  });
  ["quoteSearchNumber", "quoteSearchClient", "quoteSearchSite", "quoteSearchAsset"].forEach((id) => {
    $(`#${id}`).addEventListener("input", renderQuoteSearch);
  });
  ["enquirySearchNumber", "enquirySearchClient", "enquirySearchSite", "enquirySearchAsset"].forEach((id) => {
    $(`#${id}`).addEventListener("input", renderEnquiryRegister);
  });
  ["clientSearchName", "clientSearchAddress", "clientSearchContact", "clientSearchSite", "clientSearchAsset"].forEach((id) => {
    $(`#${id}`)?.addEventListener("input", renderClientRegister);
  });
  ["clientSiteRegisterSearchRef", "clientSiteRegisterSearchName"].forEach((id) => {
    $(`#${id}`)?.addEventListener("input", renderClientSites);
  });
  ["assetRegisterSearchSiteRef", "assetRegisterSearchSite", "assetRegisterSearchAssetRef", "assetRegisterSearchAsset", "assetRegisterSearchLocation", "assetRegisterSearchFan"].forEach((id) => {
    $(`#${id}`)?.addEventListener("input", renderAssetRegister);
  });
  ["supplierSearchName", "supplierSearchPhone", "supplierSearchEmail", "supplierSearchSite", "supplierSearchContact"].forEach((id) => {
    $(`#${id}`).addEventListener("input", renderSupplierRegister);
  });

  $("#newClientBtn").addEventListener("click", () => {
    clientDraft = blankClientDraft();
    state.activeClientId = clientDraft.id;
    activeClientSiteId = "";
    activeClientAssetId = "";
    clientDrillLevel = "client-details";
    clientEditorOpen = true;
    ["clientContactName", "clientContactRole", "clientContactEmail", "clientContactPhone"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    update();
  });

  $("#saveClientBtn").addEventListener("click", () => {
    saveClientDraft();
    clientDraft = null;
    clientDrillLevel = "register";
    clientEditorOpen = false;
    update();
  });

  $("#cancelClientBtn").addEventListener("click", () => {
    clientDraft = null;
    activeClientSiteId = "";
    activeClientAssetId = "";
    clientDrillLevel = "register";
    clientEditorOpen = false;
    update();
  });

  $("#addSiteBtn").addEventListener("click", () => {
    openAddSiteDialog();
  });

  $("#saveSiteBtn")?.addEventListener("click", () => {
    if (!clientDraft) clientDraft = blankClientDraft();
    updateClientDraftFromForm();
    saveClientDraft();
    clientDrillLevel = "client";
    update();
  });

  $("#addAssetBtn").addEventListener("click", () => {
    openAddAssetDialog();
  });

  $("#saveAssetBtn")?.addEventListener("click", () => {
    if (!clientDraft) clientDraft = blankClientDraft();
    updateClientDraftFromForm();
    saveClientDraft();
    clientDrillLevel = "asset";
    update();
  });

  $("#backToClientBtn")?.addEventListener("click", () => {
    updateClientDraftFromForm();
    saveClientDraft();
    clientDraft = null;
    activeClientSiteId = "";
    activeClientAssetId = "";
    clientDrillLevel = "register";
    clientEditorOpen = false;
    update();
  });

  $("#editClientDetailsBtn")?.addEventListener("click", () => {
    updateClientDraftFromForm();
    clientDrillLevel = "client-details";
    update();
  });

  $("#backToSitesBtn")?.addEventListener("click", () => {
    updateClientDraftFromForm();
    clientDrillLevel = "client";
    activeClientAssetId = "";
    update();
  });

  $("#backToAssetsBtn")?.addEventListener("click", () => {
    updateClientDraftFromForm();
    clientDrillLevel = "site";
    update();
  });

  [
    "clientName", "clientAddress", "clientSiteRef", "clientSiteName", "clientSiteAddress", "clientAssetRef", "clientAssetName", "clientAssetLocation",
    "assetFlowRate", "assetFanModel", "assetFanSerial", "assetFlowSensorModel", "assetFlowSensorSerial", "assetFlowSetpoint",
    "assetCarbonMediaType", "assetCarbonMediaSl", "assetCarbonMediaVolume", "assetCarbonPressureDrop", "assetCarbonHighSetpoint",
    "assetCarbonHighHighSetpoint", "assetCarbonTempProbeModel", "assetCarbonTempProbeSerial", "assetBioMediaType", "assetBioMediaVolume"
  ].forEach((id) => {
    $(`#${id}`)?.addEventListener("input", saveClientDraft);
    $(`#${id}`)?.addEventListener("change", saveClientDraft);
  });
  document.addEventListener("input", (event) => {
    if (event.target?.dataset?.customAssetField) {
      updateCustomAssetField(event.target);
      saveClientDraft();
    }
  });
  document.addEventListener("change", (event) => {
    if (event.target?.dataset?.customAssetField) {
      updateCustomAssetField(event.target);
      saveClientDraft();
    }
  });
  $("#clientContactSearch")?.addEventListener("input", renderClientContacts);

  $("#addContactBtn").addEventListener("click", () => {
    const name = $("#clientContactName").value.trim();
    const email = $("#clientContactEmail").value.trim();
    const phone = $("#clientContactPhone").value.trim();
    const role = $("#clientContactRole").value.trim();
    if (!name && !email && !phone) return;
    if (!clientDraft) clientDraft = blankClientDraft();
    updateClientDraftFromForm();
    clientDraft.contacts.push({ id: newId("contact"), name, role, email, phone });
    $("#clientContactName").value = "";
    $("#clientContactRole").value = "";
    $("#clientContactEmail").value = "";
    $("#clientContactPhone").value = "";
    update();
  });

  $("#newSupplierBtn").addEventListener("click", () => {
    supplierDraft = blankSupplier();
    supplierEditorOpen = true;
    ["supplierContactName", "supplierContactRole", "supplierContactEmail", "supplierContactPhone", "supplierSiteName", "supplierSiteAddress", "supplierSitePhone", "supplierSiteEmail"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    update();
  });

  $("#saveSupplierBtn").addEventListener("click", () => {
    saveSupplierDraft();
    supplierDraft = null;
    supplierEditorOpen = false;
    update();
  });

  $("#cancelSupplierBtn").addEventListener("click", () => {
    supplierDraft = null;
    supplierEditorOpen = false;
    update();
  });

  ["supplierName", "supplierPhone", "supplierEmail", "supplierInvoiceAddress", "supplierNotes"].forEach((id) => {
    $(`#${id}`).addEventListener("change", updateSupplierDraftFromForm);
  });
  $("#supplierContactSearch").addEventListener("input", renderSupplierContacts);

  $("#addSupplierContactBtn").addEventListener("click", () => {
    if (!supplierDraft) supplierDraft = blankSupplier();
    updateSupplierDraftFromForm();
    const name = $("#supplierContactName").value.trim();
    const role = $("#supplierContactRole").value.trim();
    const email = $("#supplierContactEmail").value.trim();
    const phone = $("#supplierContactPhone").value.trim();
    if (!name && !email && !phone) return;
    supplierDraft.contacts.push({ id: newId("supplier-contact"), name, role, email, phone });
    ["supplierContactName", "supplierContactRole", "supplierContactEmail", "supplierContactPhone"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    update();
  });

  $("#addSupplierSiteBtn").addEventListener("click", () => {
    if (!supplierDraft) supplierDraft = blankSupplier();
    updateSupplierDraftFromForm();
    const site = {
      ...blankSupplierSite((supplierDraft.sites || []).length + 1),
      name: $("#supplierSiteName").value.trim(),
      address: $("#supplierSiteAddress").value.trim(),
      phone: $("#supplierSitePhone").value.trim(),
      email: $("#supplierSiteEmail").value.trim()
    };
    if (!site.name && !site.address && !site.phone && !site.email) return;
    supplierDraft.sites.push(site);
    ["supplierSiteName", "supplierSiteAddress", "supplierSitePhone", "supplierSiteEmail"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    update();
  });

  $("#addZoneBtn").addEventListener("click", () => {
    state.zones.push(blankZone(state.zones.length + 1));
    update();
  });

  $("#addLegBtn").addEventListener("click", () => {
    state.legs.push(blankLeg(state.legs.length + 1));
    update();
  });

  $("#resetBtn")?.addEventListener("click", () => {
    resetOpenTab();
  });

  $("#exportBtn")?.addEventListener("click", () => {
    if (!isAdminUser()) {
      alert("Only admin users can export Excel files.");
      return;
    }
    setSyncStatus("Preparing Excel export", "", "Starting download.");
    window.setTimeout(() => {
      exportClientDatabaseExcel().catch((error) => {
        setSyncStatus("Excel export failed", "error", error.message || "The Excel file could not be downloaded.");
      });
    }, 25);
  });

  $("#importInput")?.addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) return;
    try {
      if (!isAdminUser()) throw new Error("Only admin users can import Excel files.");
      const form = new FormData();
      form.append("file", file);
      setSyncStatus("Importing Excel", "", "Uploading workbook to Netlify storage.");
      const data = await apiFetch(SERVER_STORAGE.importExcelUrl, {
        method: "POST",
        body: form
      });
      serverLoaded = true;
      serverVersion = Number(data.version || serverVersion);
      assetCustomFields = Array.isArray(data.assetCustomFields) ? data.assetCustomFields : assetCustomFields;
      state = normalizeLockedLogic(mergeDefaults(clone(defaultState), data.state || {}));
      consolidateProjectNumberFields();
      state.activeClientId = "";
      clientEditorOpen = false;
      clientDraft = null;
      clientDrillLevel = "register";
      activeClientSiteId = "";
      activeClientAssetId = "";
      setActiveTab("clients");
      update();
      setSyncStatus("Import complete", "ok", data.summary || "Excel import complete.");
      showImportReport(data.importReport, data.summary || "");
    } catch (error) {
      setSyncStatus("Import failed", "error", error.message || "Could not import Excel file.");
    } finally {
      event.target.value = "";
    }
  });

  $("#productSearch").addEventListener("input", renderProducts);
  ["productFilterSize", "productFilterSupplier", "productFilterMaterial", "productFilterCategory", "productFilterMinPrice", "productFilterMaxPrice"].forEach((id) => {
    $(`#${id}`).addEventListener("input", renderProducts);
  });
  $("#clearProductFiltersBtn").addEventListener("click", () => {
    ["productFilterSize", "productFilterSupplier", "productFilterMaterial", "productFilterCategory", "productFilterMinPrice", "productFilterMaxPrice"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    renderProducts();
  });

  $("#applyPriceUpdateBtn").addEventListener("click", () => {
    const supplier = ($("#priceMatchSupplier").value || "").trim().toLowerCase();
    const material = ($("#priceMatchMaterial").value || "").trim().toLowerCase();
    const category = ($("#priceMatchCategory").value || "").trim().toLowerCase();
    const percent = n($("#pricePercent").value);
    if ((!supplier && !material && !category) || percent === 0) return;
    state.products = state.products.map((product) => {
      const supplierOk = !supplier || String(product.supplier || "").toLowerCase() === supplier;
      const materialOk = !material || String(product.material || "").toLowerCase() === material;
      const categoryOk = !category || String(product.category || "").toLowerCase() === category;
      if (!(supplierOk && materialOk && categoryOk)) return product;
      return { ...product, price: round(n(product.price) * (1 + percent / 100), 2) };
    });
    update();
  });

  $("#addProductBtn").addEventListener("click", () => {
    const description = $("#newProductDescription").value.trim();
    if (!description) return;
    state.products.push({
      productCode: $("#newProductCode").value.trim(),
      description,
      size: $("#newProductSize").value.trim(),
      price: n($("#newProductPrice").value),
      category: $("#newProductCategory").value.trim(),
      supplier: $("#newProductSupplier").value.trim(),
      material: $("#newProductMaterial").value.trim()
    });
    ["newProductCode", "newProductDescription", "newProductSize", "newProductPrice", "newProductCategory", "newProductSupplier", "newProductMaterial"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    update();
  });

  $("#createQuoteBtn").addEventListener("click", () => {
    const enquiry = activeEnquiry();
    if (!enquiry) return;
    const quote = blankQuote(enquiry);
    quote.total = currentResults?.bom?.sell || 0;
    state.activeQuoteId = quote.id;
    quoteDraft = quote;
    quoteEditorOpen = true;
    update();
  });

  $("#saveQuoteBtn").addEventListener("click", () => {
    saveQuoteFromForm();
    quoteDraft = null;
    quoteEditorOpen = false;
    update();
  });

  $("#cancelQuoteBtn").addEventListener("click", () => {
    quoteDraft = null;
    quoteEditorOpen = false;
    update();
  });

  ["quoteStatus", "quoteOfferText", "quoteNotes", "quoteExclusions"].forEach((id) => {
    $(`#${id}`).addEventListener("input", updateQuoteDraftFromForm);
  });

  $("#addQuoteLineBtn").addEventListener("click", () => {
    if (!quoteDraft) return;
    const description = $("#quoteLineDescription").value.trim();
    const qty = n($("#quoteLineQty").value) || 1;
    const price = n($("#quoteLinePrice").value);
    if (!description) return;
    quoteDraft.lines ||= [];
    quoteDraft.lines.push({ description, qty, price, source: "Free text" });
    quoteDraft.total = quoteDraft.lines.reduce((sum, line) => sum + n(line.qty) * n(line.price), 0);
    ["quoteLineDescription", "quoteLineQty", "quoteLinePrice"].forEach((id) => {
      $(`#${id}`).value = "";
    });
    update();
  });

  $("#addProductToQuoteBtn").addEventListener("click", () => {
    if (!quoteDraft) return;
    const code = $("#quoteProductSelect").value;
    const product = (state.products || []).find((item) => item.productCode === code);
    if (!product) return;
    quoteDraft.lines ||= [];
    quoteDraft.lines.push({
      productCode: product.productCode,
      description: product.description,
      supplier: product.supplier || "GREENACRE",
      qty: 1,
      basePrice: n(product.price),
      price: n(product.price),
      freeText: "",
      childParts: clone(product.childParts || []),
      source: "Product"
    });
    quoteDraft.total = quoteDraft.lines.reduce((sum, line) => sum + n(line.qty) * n(line.price), 0);
    $("#quoteProductSelect").value = "";
    update();
  });

  $("#viewQuotePdfBtn").addEventListener("click", () => {
    const quote = quoteDraft || activeQuote();
    if (!quote) return;
    if (quoteDraft) updateQuoteDraftFromForm();
    openQuotePdfPreview(quoteDraft || quote);
  });

  $("#emailQuoteBtn").addEventListener("click", () => {
    const quote = quoteDraft || activeQuote();
    if (!quote) return;
    const { contact, client } = quoteContext();
    const subject = encodeURIComponent(`Quotation ${quote.number} Rev ${quote.revision}`);
    const body = encodeURIComponent(`Dear ${contact?.name || ""},\n\nPlease find quotation ${quote.number} Rev ${quote.revision} for ${client?.name || "your enquiry"}.\n\nTotal: ${moneyFmt.format(currentResults?.bom?.sell || quote.total || 0)}\n\nKind regards,\nGreencare Environmental`);
    window.location.href = `mailto:${encodeURIComponent(contact?.email || "")}?subject=${subject}&body=${body}`;
  });

  $("#createParentProductBtn").addEventListener("click", () => {
    createParentProductFromDesign();
    update();
    setActiveTab("products");
  });

  $("#downloadSvgBtn").addEventListener("click", () => {
    const svg = $("#pidSvg")?.outerHTML;
    if (!svg) return;
    const blob = new Blob([svg], { type: "image/svg+xml" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${state.project.reference || "flow-system"}-schematic.svg`;
    link.click();
    URL.revokeObjectURL(url);
  });

  $("#copySvgBtn").addEventListener("click", async () => {
    const svg = $("#pidSvg")?.outerHTML;
    if (!svg || !navigator.clipboard) return;
    await navigator.clipboard.writeText(svg);
  });
}

function update() {
  normalizeLockedLogic();
  syncProjectFromEnquiry();
  currentResults = calculate();
  renderInputs();
  renderEnquiry();
  renderClients();
  renderSuppliers();
  renderActiveEnquirySummary();
  renderKpis(currentResults);
  renderPressure(currentResults);
  renderZones(currentResults);
  renderLegs(currentResults);
  renderSchematic(currentResults);
  renderBom(currentResults);
  renderFan(currentResults);
  renderProducts();
  renderQuote();
  saveState();
}

async function bootstrap() {
  try {
    currentUser = await loadCurrentUser();
    if (!currentUser) {
      showLoginScreen();
      return;
    }
  } catch (error) {
    showLoginScreen(error.message || "Please sign in.");
    return;
  }

  bindInputs();
  bindTabs();
  bindDesignTabs();
  bindActions();
  bindDynamicTableEvents();
  renderUserControls();
  setActiveTab("clients");
  try {
    state = await loadStateFromServer();
  } catch (error) {
    console.error(error);
    serverLoaded = false;
    state = loadStateFromLocalStorage(error);
  }
  update();
}

window.addEventListener("beforeunload", () => {
  try {
    if (clientDraft) saveClientDraft();
    saveStateToLocalStorage();
  } catch (error) {
    console.error(error);
  }
});

bootstrap();












