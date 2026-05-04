import { loadConfig } from "./config.js";
import { state } from "./state.js";
import { clearFlash, collectRowColumns, formatColumnLabel, showFlash, $ } from "./utils.js";
import { resolveDriveId } from "./graph.js";
import {
  applyFilters,
  downloadFilteredExcel,
  downloadFilteredPdfs,
  goToNextPage,
  goToPreviousPage,
  refreshDatabase,
  saveDatabaseEdits,
} from "./table.js";
import { keepProcessedResults, processUploadedPdf } from "./process.js";

const SUMMARY_PREFERRED_COLUMNS = [
  "source_page_number",
  "receipt_type",
  "company",
  "gpt_payment_date",
  "gpt_total_amount",
  "gpt_taxes_total",
  "gpt_gst",
  "gpt_hst",
  "gpt_pst",
  "gpt_qst",
  "gpt_tps",
  "gpt_iva",
  "gpt_vat",
  "gpt_retention",
  "gpt_category",
  "gpt_merchant_name",
  "gpt_ticket_number",
  "gpt_city",
  "gpt_province",
  "gpt_description",
  "notes",
  "file_name",
];

const elements = {
  appShell: $(".app-shell"),
  microsoftAuthBtn: $("#microsoft-auth-btn"),
  mainPanel: $("#main-panel"),
  refreshDbBtn: $("#refresh-db-btn"),
  saveDbBtn: $("#save-db-btn"),
  downloadExcelBtn: $("#download-excel-btn"),
  downloadPdfsBtn: $("#download-pdfs-btn"),
  filtersAccordionPanel: $("#filters-accordion-panel"),
  showFiltersBtn: $("#show-filters-btn"),
  hideFiltersBtn: $("#hide-filters-btn"),
  sidebarBackdrop: $("#sidebar-backdrop"),
  clearFiltersBtn: $("#clear-filters-btn"),
  databaseLayout: $("#database-layout"),
  filtersContainer: $("#filters-container"),
  dbCaption: $("#page-indicator"),
  dbTableHead: $("#db-table thead"),
  dbTableBody: $("#db-table tbody"),
  prevPageBtn: $("#prev-page-btn"),
  nextPageBtn: $("#next-page-btn"),
  pageIndicator: $("#page-indicator"),
  pdfInput: $("#pdf-input"),
  receiptTypeInput: $("#receipt-type-input"),
  companyInput: $("#company-input"),
  bankInput: $("#bank-input"),
  bankField: $("#bank-field"),
  cardTypeInput: $("#card-type-input"),
  cardTypeField: $("#card-type-field"),
  cardLast4Input: $("#card-last4-input"),
  cardLast4Field: $("#card-last4-field"),
  descriptionInput: $("#description-input"),
  descriptionField: $("#description-field"),
  processBtn: $("#process-btn"),
  keepResultsBtn: $("#keep-results-btn"),
  dropResultsBtn: $("#drop-results-btn"),
  downloadSummaryBtn: $("#download-summary-btn"),
  progressBar: $("#progress-bar"),
  progressLabel: $("#progress-label"),
  uploadCaption: $("#upload-caption"),
  summaryTableHead: $("#summary-table thead"),
  summaryTableBody: $("#summary-table tbody"),
  rawOutput: $("#raw-output"),
};

boot().catch((error) => showFlash(error.message, "error"));

async function boot() {
  wireTabs();
  wireButtons();
  await loadConfig();
  populateSelects();
  syncReceiptTypeFields();
  await completeMicrosoftRedirectIfNeeded();
  renderAuthState();
  if (state.graphToken) {
    await resolveDriveId();
    await refreshDatabase(elements);
  }
}

function wireTabs() {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab-button").forEach((item) => item.classList.remove("is-active"));
      document.querySelectorAll(".tab-panel").forEach((item) => item.classList.remove("is-active"));
      button.classList.add("is-active");
      $(`#${button.dataset.tabTarget}`).classList.add("is-active");
      if (button.dataset.tabTarget !== "database-tab") {
        setFiltersSidebar(false);
      }
      syncHeaderControls();
    });
  });
}

function wireButtons() {
  elements.microsoftAuthBtn.addEventListener("click", () => {
    if (state.graphToken) {
      disconnectMicrosoft();
      return;
    }
    runGuarded(connectMicrosoft);
  });
  elements.refreshDbBtn.addEventListener("click", async () => runGuarded(() => refreshDatabase(elements)));
  elements.showFiltersBtn.addEventListener("click", () => setFiltersSidebar(true));
  elements.hideFiltersBtn.addEventListener("click", () => setFiltersSidebar(false));
  elements.sidebarBackdrop.addEventListener("click", () => setFiltersSidebar(false));
  elements.saveDbBtn.addEventListener("click", async () => runGuarded(async () => {
    await saveDatabaseEdits();
    showFlash("Cambios guardados en el CSV de SharePoint.");
  }));
  elements.downloadExcelBtn.addEventListener("click", () => runGuarded(() => downloadFilteredExcel()));
  elements.downloadPdfsBtn.addEventListener("click", () => runGuarded(() => downloadFilteredPdfs()));
  elements.prevPageBtn.addEventListener("click", () => goToPreviousPage(elements));
  elements.nextPageBtn.addEventListener("click", () => goToNextPage(elements));
  elements.receiptTypeInput.addEventListener("change", syncReceiptTypeFields);
  elements.clearFiltersBtn.addEventListener("click", () => {
    document.querySelectorAll("#filters-container input, #filters-container select").forEach((input) => {
      if (input.type === "range") {
        input.value = input.id.endsWith("-min") ? input.min : input.max;
      } else {
        input.value = input.id === "filter-status" ? "all" : "";
      }
    });
    document.querySelectorAll(".range-filter-values span").forEach((label) => {
      const relatedInput = document.querySelector(`#${label.id.replace("-label", "")}`);
      if (relatedInput) {
        label.textContent = Number(relatedInput.value).toFixed(2).replace(/\.00$/, "");
      }
    });
    applyFilters(elements);
  });
  elements.filtersContainer.addEventListener("input", () => applyFilters(elements));
  elements.filtersContainer.addEventListener("change", () => applyFilters(elements));
  elements.processBtn.addEventListener("click", async () => runGuarded(handleProcess));
  elements.keepResultsBtn.addEventListener("click", async () => runGuarded(handleKeepResults));
  elements.dropResultsBtn.addEventListener("click", handleDropResults);
  elements.downloadSummaryBtn.addEventListener("click", handleDownloadSummary);
}

function setFiltersSidebar(opened) {
  elements.databaseLayout.classList.toggle("is-collapsed", !opened);
  elements.databaseLayout.classList.toggle("is-sidebar-open", opened);
  elements.filtersAccordionPanel.classList.toggle("is-open", opened);
  elements.sidebarBackdrop.hidden = !opened;
  elements.appShell.classList.toggle("sidebar-open", opened);
  elements.appShell.classList.toggle("sidebar-collapsed", !opened);
  syncHeaderControls();
}

function populateSelects() {
  elements.companyInput.innerHTML = state.config.companyOptions.map((item) => `<option value="${item}">${item}</option>`).join("");
  elements.bankInput.innerHTML = state.config.bankOptions.map((item) => `<option value="${item}">${item}</option>`).join("");
}

async function connectMicrosoft() {
  clearFlash();
  const response = await fetch("/.netlify/functions/microsoft-login");
  if (!response.ok) {
    throw new Error(await response.text());
  }
  const payload = await response.json();
  localStorage.setItem("msalState", payload.state);
  window.location.href = payload.authUrl;
}

async function completeMicrosoftRedirectIfNeeded() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  if (!code) return;

  const stateParam = params.get("state");
  const expectedState = localStorage.getItem("msalState");
  if (!stateParam || !expectedState || stateParam !== expectedState) {
    throw new Error("El state de Microsoft no coincide.");
  }

  const response = await fetch(`/.netlify/functions/microsoft-callback?code=${encodeURIComponent(code)}&redirect_uri=${encodeURIComponent(window.location.origin + "/")}`);
  if (!response.ok) {
    throw new Error(await response.text());
  }
  const payload = await response.json();
  state.graphToken = payload.accessToken;
  sessionStorage.setItem("graphToken", state.graphToken);
  history.replaceState({}, document.title, window.location.pathname);
}

function disconnectMicrosoft() {
  state.graphToken = "";
  state.driveId = "";
  state.databaseRows = [];
  state.filteredRows = [];
  sessionStorage.removeItem("graphToken");
  setFiltersSidebar(false);
  renderAuthState();
  elements.mainPanel.hidden = true;
  elements.dbTableHead.innerHTML = "";
  elements.dbTableBody.innerHTML = "";
  elements.dbCaption.textContent = "Conecta Microsoft para cargar la base.";
}

function renderAuthState() {
  const connected = Boolean(state.graphToken);
  elements.microsoftAuthBtn.querySelector("span").textContent = connected ? "Desconectar" : "Conectar";
  elements.mainPanel.hidden = !connected;
  if (!connected) {
    setFiltersSidebar(false);
    return;
  }
  setFiltersSidebar(false);
  syncHeaderControls();
}

function syncHeaderControls() {
  const connected = Boolean(state.graphToken);
  const databaseTabActive = $("#database-tab")?.classList.contains("is-active");
  const sidebarOpen = elements.databaseLayout.classList.contains("is-sidebar-open");
  elements.showFiltersBtn.hidden = !connected || !databaseTabActive || sidebarOpen;
}

function syncReceiptTypeFields() {
  const receiptType = elements.receiptTypeInput.value;
  const isReimbursement = receiptType === "reimbursement";
  elements.bankField.hidden = isReimbursement;
  elements.bankInput.disabled = isReimbursement;
  elements.cardTypeField.hidden = isReimbursement;
  elements.cardTypeInput.disabled = isReimbursement;
  elements.cardLast4Field.hidden = isReimbursement;
  elements.descriptionField.hidden = !isReimbursement;
  elements.cardLast4Input.disabled = isReimbursement;
  elements.descriptionInput.disabled = !isReimbursement;
  if (isReimbursement) {
    elements.bankInput.selectedIndex = 0;
    elements.cardTypeInput.selectedIndex = 0;
    elements.cardLast4Input.value = "";
  } else {
    elements.descriptionInput.value = "";
  }
}

async function handleProcess() {
  clearFlash();
  elements.progressBar.value = 0;
  elements.progressLabel.textContent = "Preparando procesamiento";
  const result = await processUploadedPdf(elements);
  renderSummary();
  elements.keepResultsBtn.disabled = !result.summaryRows.length;
  elements.dropResultsBtn.disabled = !result.summaryRows.length;
  elements.downloadSummaryBtn.disabled = !result.summaryRows.length;
  elements.uploadCaption.textContent = result.summaryRows.length
    ? `Se procesaron ${result.summaryRows.length} pagina(s).`
    : "No hubo paginas procesadas.";
  if (result.errors?.length) {
    showFlash(result.errors.join(" | "), "warning");
  } else {
    showFlash("Todas las paginas fueron procesadas.");
  }
}

async function handleKeepResults() {
  await keepProcessedResults();
  elements.keepResultsBtn.disabled = true;
  await refreshDatabase(elements);
  showFlash("Resultados persistidos en SharePoint.");
}

function handleDropResults() {
  state.processed = { summaryRows: [], rawRows: [], pendingUploads: [], saved: false };
  elements.summaryTableHead.innerHTML = "";
  elements.summaryTableBody.innerHTML = "";
  elements.rawOutput.textContent = "";
  elements.keepResultsBtn.disabled = true;
  elements.dropResultsBtn.disabled = true;
  elements.downloadSummaryBtn.disabled = true;
  elements.uploadCaption.textContent = "Resultados descartados.";
  showFlash("Resultados descartados.");
}

function handleDownloadSummary() {
  const rows = state.processed.summaryRows;
  if (!rows.length) return;
  const columns = getSummaryColumns(rows);
  const exportRows = rows.map((row) => Object.fromEntries(
    columns.map((column) => [formatColumnLabel(column), displaySummaryValue(column, row[column] ?? "")]),
  ));
  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  worksheet["!autofilter"] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: exportRows.length, c: Math.max(0, columns.length - 1) },
    }),
  };
  worksheet["!cols"] = columns.map((column) => ({ wch: getSummaryColumnWidth(column) }));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "summary");
  XLSX.writeFile(workbook, "invoice_summary_google_vision.xlsx");
}

function renderSummary() {
  const rows = state.processed.summaryRows;
  if (!rows.length) {
    elements.summaryTableHead.innerHTML = "";
    elements.summaryTableBody.innerHTML = "";
    elements.rawOutput.textContent = JSON.stringify(state.processed.rawRows, null, 2);
    return;
  }

  const columns = getSummaryColumns(rows);
  elements.summaryTableHead.innerHTML = `<tr>${columns.map((column) => `<th>${column}</th>`).join("")}</tr>`;
  elements.summaryTableBody.innerHTML = rows.map((row) => `
    <tr>
      ${columns.map((column) => `<td>${escapeHtml(displaySummaryValue(column, row[column] ?? ""))}</td>`).join("")}
    </tr>
  `).join("");

  elements.rawOutput.textContent = JSON.stringify(state.processed.rawRows, null, 2);
}

async function runGuarded(work) {
  try {
    clearFlash();
    await work();
  } catch (error) {
    showFlash(error.message, "error");
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function displaySummaryValue(column, value) {
  if (column === "receipt_type") {
    return String(value).trim().toLowerCase() === "reimbursement" ? "Reembolso" : "Transaccion bancaria";
  }
  return value;
}

function getSummaryColumns(rows) {
  const availableColumns = new Set(collectRowColumns(rows));
  return SUMMARY_PREFERRED_COLUMNS.filter((column) => availableColumns.has(column));
}

function getSummaryColumnWidth(column) {
  if (["source_page_number", "gpt_gst", "gpt_hst", "gpt_pst", "gpt_qst", "gpt_tps", "gpt_iva", "gpt_vat", "gpt_retention"].includes(column)) {
    return 12;
  }
  if (["gpt_total_amount", "gpt_taxes_total"].includes(column)) {
    return 14;
  }
  if (["receipt_type", "company", "gpt_category", "gpt_city", "gpt_province", "gpt_ticket_number"].includes(column)) {
    return 18;
  }
  if (["gpt_merchant_name", "gpt_description", "notes"].includes(column)) {
    return 26;
  }
  if (column === "file_name") {
    return 40;
  }
  return 18;
}
