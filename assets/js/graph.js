import { state } from "./state.js";
import { joinSharePointPath, rowsToCsv } from "./utils.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function authHeaders() {
  if (!state.graphToken) {
    throw new Error("Microsoft no esta conectado.");
  }
  return { Authorization: `Bearer ${state.graphToken}` };
}

async function graphJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      ...authHeaders(),
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.status === 204 ? {} : response.json();
}

async function graphBytes(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      ...authHeaders(),
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.arrayBuffer();
}

export async function resolveDriveId() {
  if (state.driveId) return state.driveId;
  const { spHostname, spSitePath, spDriveName } = state.config;
  const site = await graphJson(`${GRAPH_BASE}/sites/${spHostname}:${spSitePath}`);
  const drives = (await graphJson(`${GRAPH_BASE}/sites/${site.id}/drives`)).value || [];
  const drive = drives.find((item) => item.name === spDriveName) || drives[0];
  if (!drive) {
    throw new Error("No se pudo resolver el drive de SharePoint.");
  }
  state.driveId = drive.id;
  return drive.id;
}

export async function listChildrenByPath(path) {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  const url = `${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/children?$top=200&$select=id,name,folder,file,webUrl`;
  const data = await graphJson(url);
  return data.value || [];
}

export async function downloadSharePointFile(path) {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  const buffer = await graphBytes(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`);
  return new Uint8Array(buffer);
}

export async function uploadSharePointFile(path, content, contentType = "application/octet-stream") {
  const driveId = await resolveDriveId();
  const encodedPath = encodeURIComponent(path).replaceAll("%2F", "/");
  return graphJson(`${GRAPH_BASE}/drives/${driveId}/root:/${encodedPath}:/content`, {
    method: "PUT",
    headers: { "Content-Type": contentType },
    body: content,
  });
}

export async function loadDatabaseRows() {
  const csvPath = joinSharePointPath(state.config.receiptsDatabaseDir, state.config.receiptsDatabaseCsv);
  try {
    const bytes = await downloadSharePointFile(csvPath);
    const text = new TextDecoder("utf-8").decode(bytes);
    return parseCsv(text);
  } catch {
    return [];
  }
}

export async function saveDatabaseRows(rows) {
  const csvPath = joinSharePointPath(state.config.receiptsDatabaseDir, state.config.receiptsDatabaseCsv);
  const csvText = rowsToCsv(rows);
  await uploadSharePointFile(csvPath, new TextEncoder().encode(csvText), "text/csv;charset=utf-8");
}

function parseCsv(text) {
  const workbook = XLSX.read(text, { type: "string" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((row) => {
    const legacyChecked = String(row.checked).toLowerCase() === "true" || row.checked === true;
    const normalizedStatus = normalizeStatus(row.status || (legacyChecked ? "Paid" : "Pending"));
    const normalizedReceiptType = normalizeReceiptType(row.receipt_type);
    const { checked, ...rest } = row;
    return {
      ...rest,
      receipt_type: normalizedReceiptType,
      status: normalizedStatus,
    };
  });
}

function normalizeStatus(value) {
  return String(value).trim().toLowerCase() === "paid" ? "Paid" : "Pending";
}

function normalizeReceiptType(value) {
  return String(value).trim().toLowerCase() === "reimbursement" ? "reimbursement" : "bank_transaction";
}
