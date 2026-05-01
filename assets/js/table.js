import { state } from "./state.js";
import { collectRowColumns, displayCardLast4, formatColumnLabel, showFlash, toFloat } from "./utils.js";
import { loadDatabaseRows, saveDatabaseRows, downloadSharePointFile } from "./graph.js";
import { mergePdfBlobs } from "./pdf.js";

const TAX_COLUMNS = [
  "gpt_gst",
  "gpt_hst",
  "gpt_pst",
  "gpt_qst",
  "gpt_tps",
  "gpt_iva",
  "gpt_vat",
  "gpt_retention",
];

const FILTER_COLUMNS = [
  "receipt_type",
  "company",
  "bank",
  "card_type",
  "card_last4",
  "gpt_category",
  "gpt_merchant_name",
  "gpt_city",
  "gpt_province",
];

const DATE_FILTER_COLUMNS = [
  { key: "processed_at", label: "Processed At" },
  { key: "gpt_payment_date", label: "Payment Date" },
];

const RANGE_FILTER_COLUMNS = [
  { key: "gpt_total_amount", label: "Total Amount" },
  { key: "gpt_taxes_total", label: "Taxes Total" },
  ...TAX_COLUMNS.map((key) => ({ key, label: formatColumnLabel(key) })),
];

export async function refreshDatabase(elements) {
  state.databaseRows = await loadDatabaseRows();
  renderFilters(elements.filtersContainer);
  applyFilters(elements);
}

export function renderFilters(container) {
  container.innerHTML = "";

  const checkedGroup = document.createElement("div");
  checkedGroup.className = "filter-group";
  checkedGroup.innerHTML = `
    <label for="filter-checked">Checked status</label>
    <select id="filter-checked">
      <option value="all">all</option>
      <option value="unchecked">unchecked</option>
      <option value="checked">checked</option>
    </select>
  `;
  container.append(checkedGroup);

  const searchGroup = document.createElement("div");
  searchGroup.className = "filter-group";
  searchGroup.innerHTML = `
    <label for="filter-text">Search in table</label>
    <input id="filter-text" type="text" placeholder="merchant, city, amount...">
  `;
  container.append(searchGroup);

  DATE_FILTER_COLUMNS.forEach(({ key, label }) => {
    const group = document.createElement("div");
    group.className = "filter-group";
    group.innerHTML = `
      <label>${label}</label>
      <div class="date-filter-row">
        <input id="filter-${key}-from" type="date" aria-label="${label} desde">
        <input id="filter-${key}-to" type="date" aria-label="${label} hasta">
      </div>
    `;
    container.append(group);
  });

  RANGE_FILTER_COLUMNS.forEach(({ key, label }) => {
    const values = state.databaseRows
      .map((row) => toFloat(row[key], NaN))
      .filter((value) => Number.isFinite(value));
    if (!values.length) return;
    const minValue = Math.floor(Math.min(...values));
    const maxValue = Math.ceil(Math.max(...values));
    const group = document.createElement("div");
    group.className = "filter-group";
    group.innerHTML = `
      <label>${label}</label>
      <div class="range-filter-values">
        <span id="filter-${key}-min-label">${formatRangeValue(minValue)}</span>
        <span id="filter-${key}-max-label">${formatRangeValue(maxValue)}</span>
      </div>
      <div class="range-filter-row">
        <input id="filter-${key}-min" type="range" min="${minValue}" max="${maxValue}" step="0.01" value="${minValue}" aria-label="${label} minimo">
        <input id="filter-${key}-max" type="range" min="${minValue}" max="${maxValue}" step="0.01" value="${maxValue}" aria-label="${label} maximo">
      </div>
    `;
    container.append(group);
    const minInput = group.querySelector(`#filter-${key}-min`);
    const maxInput = group.querySelector(`#filter-${key}-max`);
    const minLabel = group.querySelector(`#filter-${key}-min-label`);
    const maxLabel = group.querySelector(`#filter-${key}-max-label`);
    const syncRange = () => {
      if (Number(minInput.value) > Number(maxInput.value)) {
        if (document.activeElement === minInput) {
          maxInput.value = minInput.value;
        } else {
          minInput.value = maxInput.value;
        }
      }
      minLabel.textContent = formatRangeValue(minInput.value);
      maxLabel.textContent = formatRangeValue(maxInput.value);
    };
    minInput.addEventListener("input", syncRange);
    maxInput.addEventListener("input", syncRange);
    syncRange();
  });

  FILTER_COLUMNS.forEach((column) => {
    const values = [...new Set(state.databaseRows.map((row) => String(row[column] ?? "")).filter(Boolean))].sort();
    const group = document.createElement("div");
    group.className = "filter-group";
    const options = ['<option value="">all</option>', ...values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(displayValue(column, value))}</option>`)];
    group.innerHTML = `
      <label for="filter-${column}">${formatColumnLabel(column)}</label>
      <select id="filter-${column}">
        ${options.join("")}
      </select>
    `;
    container.append(group);
  });
}

export function applyFilters(elements) {
  const checkedStatus = document.querySelector("#filter-checked")?.value || "all";
  const textQuery = (document.querySelector("#filter-text")?.value || "").trim().toLowerCase();
  let rows = [...state.databaseRows];

  if (checkedStatus === "checked") {
    rows = rows.filter((row) => Boolean(row.checked));
  } else if (checkedStatus === "unchecked") {
    rows = rows.filter((row) => !row.checked);
  }

  DATE_FILTER_COLUMNS.forEach(({ key }) => {
    const from = document.querySelector(`#filter-${key}-from`)?.value || "";
    const to = document.querySelector(`#filter-${key}-to`)?.value || "";
    if (!from && !to) return;
    rows = rows.filter((row) => {
      const rowDate = extractDateOnly(row[key]);
      if (!rowDate) return false;
      if (from && rowDate < from) return false;
      if (to && rowDate > to) return false;
      return true;
    });
  });

  RANGE_FILTER_COLUMNS.forEach(({ key }) => {
    const min = document.querySelector(`#filter-${key}-min`)?.value;
    const max = document.querySelector(`#filter-${key}-max`)?.value;
    if (min == null || max == null) return;
    rows = rows.filter((row) => {
      const value = toFloat(row[key], NaN);
      if (!Number.isFinite(value)) return false;
      return value >= Number(min) && value <= Number(max);
    });
  });

  FILTER_COLUMNS.forEach((column) => {
    const selected = document.querySelector(`#filter-${column}`)?.value || "";
    if (selected) {
      rows = rows.filter((row) => String(row[column] ?? "") === selected);
    }
  });

  if (textQuery) {
    rows = rows.filter((row) => Object.values(row).some((value) => String(value ?? "").toLowerCase().includes(textQuery)));
  }

  state.filteredRows = rows;
  state.pagination.page = 1;
  renderDatabaseTable(elements);
}

function renderDatabaseTable(elements) {
  const { dbTableHead: head, dbTableBody: body } = elements;
  const rowColumns = collectRowColumns(state.databaseRows).filter((column) => column !== "checked");
  const columns = rowColumns.length ? ["checked", ...rowColumns] : [];
  if (!columns.length) {
    head.innerHTML = "";
    body.innerHTML = "";
    updatePagination(elements, 0);
    return;
  }
  head.innerHTML = `<tr>${columns.map((column) => `<th>${formatColumnLabel(column)}</th>`).join("")}</tr>`;
  const totalRows = state.filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / state.pagination.pageSize));
  state.pagination.page = Math.min(state.pagination.page, totalPages);
  const startIndex = (state.pagination.page - 1) * state.pagination.pageSize;
  const visibleRows = state.filteredRows.slice(startIndex, startIndex + state.pagination.pageSize);
  const shownCount = visibleRows.length;
  body.innerHTML = totalRows
    ? visibleRows.map((row) => `
      <tr data-file-path="${escapeHtml(row.file_path || "")}">
        ${columns.map((column) => renderCell(column, row[column] ?? "")).join("")}
      </tr>
    `).join("")
    : `<tr><td colspan="${columns.length}">No hay resultados para los filtros actuales.</td></tr>`;
  updatePagination(elements, totalPages, shownCount, totalRows, startIndex);
}

function renderCell(column, value) {
  if (column === "checked") {
    return `<td><input data-column="checked" type="checkbox" ${value ? "checked" : ""}></td>`;
  }
  if (column === "receipt_type") {
    return `<td>${escapeHtml(displayValue(column, value))}</td>`;
  }
  if (column === "gpt_description") {
    return `<td>${escapeHtml(value)}</td>`;
  }
  if (column === "card_last4") {
    return `<td>${escapeHtml(displayCardLast4(value))}</td>`;
  }
  if (isNumericColumn(column)) {
    return `<td>${escapeHtml(String(toFloat(value, 0)))}</td>`;
  }
  return `<td>${escapeHtml(value)}</td>`;
}

function isNumericColumn(column) {
  return [
    "gpt_total_amount",
    "gpt_taxes_total",
    "gpt_confidence",
    "vision_total_amount",
    "vision_taxes_total",
    ...TAX_COLUMNS,
    "vision_gst",
    "vision_hst",
    "vision_pst",
    "vision_qst",
    "vision_tps",
    "vision_iva",
    "vision_vat",
    "vision_retention",
  ].includes(column);
}

export function syncEditedRowsFromDom() {
  const rowElements = [...document.querySelectorAll(".db-table tbody tr")];
  rowElements.forEach((tr) => {
    const filePath = tr.dataset.filePath;
    const match = state.databaseRows.find((row) => row.file_path === filePath);
    if (!match) return;
    const checkedInput = tr.querySelector('input[data-column="checked"]');
    match.checked = Boolean(checkedInput?.checked);
  });
}

export async function saveDatabaseEdits() {
  syncEditedRowsFromDom();
  await saveDatabaseRows(state.databaseRows);
}

export function goToPreviousPage(elements) {
  if (state.pagination.page <= 1) return;
  state.pagination.page -= 1;
  renderDatabaseTable(elements);
}

export function goToNextPage(elements) {
  const totalPages = Math.max(1, Math.ceil(state.filteredRows.length / state.pagination.pageSize));
  if (state.pagination.page >= totalPages) return;
  state.pagination.page += 1;
  renderDatabaseTable(elements);
}

export function downloadFilteredExcel() {
  if (!state.filteredRows.length) {
    showFlash("No hay filas filtradas para exportar.", "warning");
    return;
  }
  const exportRows = state.filteredRows.map((row) => ({
    ...row,
    card_last4: displayCardLast4(row.card_last4),
  }));
  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "receipts");
  XLSX.writeFile(workbook, "filtered_receipts.xlsx");
}

export async function downloadFilteredPdfs() {
  const paths = state.filteredRows.map((row) => row.file_path).filter(Boolean);
  if (!paths.length) {
    throw new Error("No hay PDFs filtrados para descargar.");
  }
  const byteArrays = [];
  for (const path of paths) {
    try {
      const bytes = await downloadSharePointFile(path);
      byteArrays.push(bytes);
    } catch {
      // Skip missing files to match Streamlit behavior.
    }
  }
  if (!byteArrays.length) {
    throw new Error("No se pudo descargar ningun PDF filtrado.");
  }
  const mergedBytes = await mergePdfBlobs(byteArrays);
  const blob = new Blob([mergedBytes], { type: "application/pdf" });
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(blob);
  anchor.download = "filtered_invoices_merged.pdf";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(anchor.href);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function displayValue(column, value) {
  if (column === "receipt_type") {
    return value === "reimbursement" ? "Reembolso" : "Transaccion bancaria";
  }
  return column === "card_last4" ? displayCardLast4(value) : value;
}

function updatePagination(elements, totalPages, shownCount = 0, totalRows = state.filteredRows.length, startIndex = 0) {
  const hasRows = state.filteredRows.length > 0;
  if (hasRows) {
    const firstRow = startIndex + 1;
    const lastRow = startIndex + shownCount;
    elements.pageIndicator.textContent = `Mostrando ${firstRow}-${lastRow} de ${totalRows} recibos`;
  } else {
    elements.pageIndicator.textContent = "Mostrando 0 de 0 recibos";
  }
  elements.prevPageBtn.disabled = !hasRows || state.pagination.page <= 1;
  elements.nextPageBtn.disabled = !hasRows || state.pagination.page >= totalPages;
}

function extractDateOnly(value) {
  const text = String(value ?? "");
  const match = text.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : "";
}

function formatRangeValue(value) {
  return Number(value).toFixed(2).replace(/\.00$/, "");
}
