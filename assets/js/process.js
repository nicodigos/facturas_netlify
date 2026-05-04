import { state } from "./state.js";
import { buildSuggestedFileName, joinSharePointPath, normalizeCardLast4, toFloat } from "./utils.js";
import { listChildrenByPath, loadDatabaseRows, saveDatabaseRows, uploadSharePointFile } from "./graph.js";
import { splitPdfToPages } from "./pdf.js";

const TAX_FIELDS = ["gst", "hst", "pst", "qst", "tps", "iva", "vat", "retention"];

export async function processUploadedPdf(elements) {
  const file = elements.pdfInput.files[0];
  const receiptType = elements.receiptTypeInput.value;
  const isReimbursement = receiptType === "reimbursement";
  if (!state.graphToken) {
    throw new Error("Microsoft es obligatorio para guardar PDFs y CSV.");
  }
  if (!file) {
    throw new Error("Selecciona un PDF.");
  }

  const last4 = isReimbursement ? "" : normalizeCardLast4(elements.cardLast4Input.value);
  const description = isReimbursement ? elements.descriptionInput.value.trim() : "";
  if (!isReimbursement && !last4) {
    throw new Error("Card last 4 digits debe tener exactamente 4 numeros.");
  }
  if (isReimbursement && !description) {
    throw new Error("La descripcion es obligatoria para reembolsos.");
  }

  const pages = await splitPdfToPages(file);
  if (!pages.length) {
    throw new Error("No se detectaron paginas en el PDF.");
  }

  const existingNames = new Set((await listChildrenByPath(state.config.receiptsDatabaseDir)).map((item) => item.name));
  const summaryRows = [];
  const rawRows = [];
  const pendingUploads = [];
  const errors = [];

  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index];
    elements.progressBar.value = Math.round(((index + 1) / pages.length) * 100);
    elements.progressLabel.textContent = `Procesando pagina ${page.pageNumber} de ${pages.length}`;

    let result;
    try {
      result = await classifyPage({
        pageNumber: page.pageNumber,
        imageBase64: page.imageBase64,
        receiptType,
      });
    } catch (error) {
      errors.push(`Pagina ${page.pageNumber}: ${error.message}`);
      continue;
    }

    const baseName = buildSuggestedFileName(
      result.gpt.payment_date || result.compact.date,
      isReimbursement ? "reembolso" : elements.bankInput.value,
      isReimbursement ? "" : elements.cardTypeInput.value,
      result.gpt.merchant_name || result.compact.merchant,
      toFloat(result.gpt.total_amount, result.compact.total),
    );
    const uniqueName = makeUniquePdfName(baseName, existingNames);
    const remotePath = joinSharePointPath(state.config.receiptsDatabaseDir, uniqueName);

    pendingUploads.push({
      fileName: uniqueName,
      filePath: remotePath,
      content: page.pdfBytes,
    });

    const row = {
      status: "Pending",
      processed_at: new Date().toISOString().slice(0, 19),
      source_page_number: page.pageNumber,
      receipt_type: receiptType,
      company: elements.companyInput.value,
      bank: isReimbursement ? "" : elements.bankInput.value,
      card_type: isReimbursement ? "" : elements.cardTypeInput.value,
      card_last4: last4,
      gpt_payment_date: result.gpt.payment_date || result.compact.date || "",
      gpt_total_amount: toFloat(result.gpt.total_amount, result.compact.total),
      gpt_taxes_total: toFloat(result.gpt.taxes_total, result.compact.taxes_total),
      ...buildTaxColumns("gpt", result.gpt),
      gpt_category: result.gpt.category || "Diverse Expenses",
      gpt_merchant_name: result.gpt.merchant_name || result.compact.merchant || "",
      gpt_city: result.gpt.city || result.compact.city || "",
      gpt_province: result.gpt.province || result.compact.province || "",
      gpt_ticket_number: result.gpt.ticket_number || "",
      gpt_description: description,
      gpt_confidence: toFloat(result.gpt.confidence, 0),
      notes: result.gpt.notes || "",
      file_name: uniqueName,
      file_path: remotePath,
    };

    summaryRows.push(row);
    rawRows.push({
      ...row,
      raw_google_vision_json: JSON.stringify(result.vision),
      raw_gpt_json: JSON.stringify(result.gpt),
      vision_payment_date: result.compact.date || "",
      vision_total_amount: result.compact.total || 0,
      vision_taxes_total: result.compact.taxes_total || 0,
      ...buildTaxColumns("vision", result.compact),
      vision_merchant_name: result.compact.merchant || "",
      vision_city: result.compact.city || "",
      vision_province: result.compact.province || "",
    });
  }

  state.processed = {
    summaryRows,
    rawRows,
    pendingUploads,
    saved: false,
    errors,
  };
  elements.progressLabel.textContent = errors.length
    ? `Completado con ${errors.length} pagina(s) con error`
    : "Procesamiento completo";
  return state.processed;
}

async function classifyPage(payload) {
  const response = await fetch("/.netlify/functions/process-receipt", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    throw new Error(await response.text());
  }
  return response.json();
}

function makeUniquePdfName(baseName, existingNames) {
  const safeBase = String(baseName || "invoice").replace(/[^\w-]+/g, "_");
  let candidate = `${safeBase}.pdf`;
  let counter = 2;
  while (existingNames.has(candidate)) {
    candidate = `${safeBase}__${counter}.pdf`;
    counter += 1;
  }
  existingNames.add(candidate);
  return candidate;
}

export async function keepProcessedResults() {
  for (const pending of state.processed.pendingUploads) {
    await uploadSharePointFile(pending.filePath, pending.content, "application/pdf");
  }
  const database = await loadDatabaseRows();
  await saveDatabaseRows([...database.rows, ...state.processed.summaryRows], { expectedEtag: database.eTag });
  state.processed.saved = true;
}

function buildTaxColumns(prefix, source) {
  return Object.fromEntries(
    TAX_FIELDS.map((field) => [`${prefix}_${field}`, toFloat(source?.[field], 0)]),
  );
}
