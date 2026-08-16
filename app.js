import { TICKET_STATUSES, STUDIO_STATUSES, SALESPEOPLE } from "./config.js";
import * as auth from "./google-auth.js";
import { chooseDriveFolder, uploadFileToDrive, updateFileInDrive, trashDriveFile } from "./drive.js";
import {
  getOrCreateLedger, appendTicketRow, updateTicketRecord, listTickets, updateTicketStatus, updateTicketNotes, deleteTicketRow,
  getOrCreateStudioTab, appendStudioEntry, listStudioEntries, updateStudioStatus, markStudioConverted, deleteStudioEntry,
} from "./sheets.js";
import { buildDocxBlob } from "./docx.js";

const STORAGE_KEY = "alterationsTicketStateV2";

const ICONS = {
  print: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5.5 7V3.5h9V7"/><rect x="3.5" y="7" width="13" height="6.5" rx="1"/><path d="M5.5 12.5V16.5h9V12.5"/></svg>',
  edit: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M13.5 3.5l3 3L7 16H4v-3L13.5 3.5z"/></svg>',
  download: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10 3v9.5M6.5 9l3.5 3.5L13.5 9"/><path d="M4 15.5h12"/></svg>',
  delete: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4.5 6h11"/><path d="M8 6V4.3a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1V6"/><path d="M6 6l.6 9a1 1 0 0 0 1 1h4.8a1 1 0 0 0 1-1L14 6"/></svg>',
  notes: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="12" height="14" rx="1.2"/><path d="M7 7.5h6M7 10.5h6M7 13.5h3.5"/></svg>',
  convert: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3.5" y="3.5" width="13" height="13" rx="2"/><path d="M6.5 10h6.5M10 6.5l3.5 3.5-3.5 3.5"/></svg>',
};

/* ------------------------------- DOM refs -------------------------------- */
const el = (id) => document.getElementById(id);

const customerNameInput = el("customerName");
const tailorInput = el("tailor");
const salespersonInput = el("salesperson");
const dueDateInput = el("dueDate");
let dueDateHiddenInput = null;
let dueDateDisplayInput = null;

const garmentItemsEl = el("garmentItemsList");
const addGarmentBtn = el("addGarmentBtn");

const printArea = el("printArea");
const saveStatus = el("saveStatus");
const appShell = el("appShell");
const authGate = el("authGate");
const authLoginBtn = el("authLoginBtn");
const authStatus = el("authStatus");
const signOutBtn = el("signOutBtn");
const accountEmail = el("accountEmail");
const driveSaveBtn = el("driveSaveBtn");
const clearBtn = el("clearBtn");
const editingBanner = el("editingBanner");
const editingBannerText = el("editingBannerText");
const editingBannerExit = el("editingBannerExit");

const driveFolderDom = {
  modal: el("driveFolderModal"),
  path: el("driveFolderPath"),
  search: el("driveFolderSearch"),
  list: el("driveFolderList"),
  upBtn: el("driveFolderUpBtn"),
  cancelBtn: el("driveFolderCancelBtn"),
  chooseBtn: el("driveFolderChooseBtn"),
  newFolderInput: el("driveFolderNewName"),
  newFolderBtn: el("driveFolderCreateBtn"),
};

const notesModal = el("notesModal");
const notesModalSubtitle = el("notesModalSubtitle");
const notesModalTextarea = el("notesModalTextarea");
const notesModalCancelBtn = el("notesModalCancelBtn");
const notesModalSaveBtn = el("notesModalSaveBtn");

const navNewTicket = el("nav-newTicket");
const navTicketLog = el("nav-ticketLog");
const navStudio = el("nav-studio");
const viewNewTicket = el("view-newTicket");
const viewTicketLog = el("view-ticketLog");
const viewStudio = el("view-studio");

const ticketSearch = el("ticketSearch");
const statusFilterEls = {
  wrap: el("statusFilterWrap"), btn: el("statusFilterBtn"), panel: el("statusFilterPanel"),
  options: el("statusFilterOptions"), allBtn: el("statusFilterAll"), noneBtn: el("statusFilterNone"),
};
const tailorFilterEls = {
  wrap: el("tailorFilterWrap"), btn: el("tailorFilterBtn"), panel: el("tailorFilterPanel"),
  options: el("tailorFilterOptions"), allBtn: el("tailorFilterAll"), noneBtn: el("tailorFilterNone"),
};
const salespersonFilterEls = {
  wrap: el("salespersonFilterWrap"), btn: el("salespersonFilterBtn"), panel: el("salespersonFilterPanel"),
  options: el("salespersonFilterOptions"), allBtn: el("salespersonFilterAll"), noneBtn: el("salespersonFilterNone"),
};
const ticketRushFilter = el("ticketRushFilter");
const clearFiltersBtn = el("clearFiltersBtn");
const refreshTicketsBtn = el("refreshTicketsBtn");
const ticketLogStatus = el("ticketLogStatus");
const ticketLogBody = el("ticketLogBody");
const logStats = el("logStats");

let garmentItems = [];
let uidCounter = 1;
let collapsedItemIds = new Set(); // items the user has collapsed to a summary line
let ticketCache = [];
let editingTicket = null; // { id, rowNumber, driveFileId, customerName } when reopened from the log

function setEditingTicket(ticket) {
  editingTicket = ticket;
  if (ticket) {
    editingBanner.hidden = false;
    editingBannerText.textContent = `Editing existing ticket for ${ticket.customerName || "this client"} — saving will update it, not create a new one.`;
    driveSaveBtn.textContent = "✂ Update Ticket";
  } else {
    editingBanner.hidden = true;
    driveSaveBtn.textContent = "✂ Save Ticket";
  }
}

editingBannerExit.addEventListener("click", () => clearAllFields());

/* -------------------------------- helpers -------------------------------- */
function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function createEmptyItem(garmentType = "jacket") {
  return {
    _uid: uidCounter++,
    garmentType,
    sizeDescription: "", adjustments: "",
    halfBack: 0, halfWaist: 0, shortenBody: 0, sleeves: 0, sleeveWidth: 0, tightenCollar: 0, buttons: "",
    trouserWaist: 0, trouserSeat: 0, trouserThigh: 0, trouserKnee: 0, trouserLegOpening: 0, trouserInseam: 0,
    trouserTotalLength: "", trouserCuff: "",
    shirtSleeve: 0, shirtBody: 0, shirtSlimBody: 0,
  };
}

// Resets an item to a fresh set of fields for a new garment type, while
// keeping the fields that make sense across any type (size, description,
// notes) — so switching the picker doesn't silently keep stale trouser
// measurements around on what's now a jacket, say.
function changeItemGarmentType(item, newType) {
  const fresh = createEmptyItem(newType);
  fresh._uid = item._uid;
  fresh.sizeDescription = item.sizeDescription;
  fresh.adjustments = item.adjustments;
  return fresh;
}

function isNegativeOnlyField(field) {
  return ["tightenCollar", "shirtSleeve", "shirtBody", "shirtSlimBody"].includes(field);
}

function normalizeNegativeOnlyValue(value) {
  return Math.min(0, Number(value) || 0);
}

function hasGarmentData(item) {
  const measurementFields = ["halfBack", "halfWaist", "shortenBody", "sleeves", "sleeveWidth", "tightenCollar",
    "trouserWaist", "trouserSeat", "trouserThigh", "trouserKnee", "trouserLegOpening", "trouserInseam",
    "shirtSleeve", "shirtBody", "shirtSlimBody"];
  const hasMeasurements = measurementFields.some((field) => {
    const value = isNegativeOnlyField(field) ? normalizeNegativeOnlyValue(item?.[field]) : Number(item?.[field]) || 0;
    return value !== 0;
  });
  return Boolean(
    (item?.sizeDescription || "").trim() || (item?.adjustments || "").trim() ||
    (item?.buttons || "").trim() || (item?.trouserTotalLength || "").trim() || (item?.trouserCuff || "").trim() ||
    hasMeasurements,
  );
}

function formatSignedQuarter(value) {
  const numeric = Number(value) || 0;
  if (numeric === 0) return "0";
  const formatted = Number.isInteger(Math.abs(numeric)) ? `${Math.abs(numeric)}` : Math.abs(numeric).toFixed(1);
  return `${numeric > 0 ? "+ " : "– "}${formatted}`;
}

function formatFileBaseName(name) {
  const normalized = String(name || "ticket").trim().replace(/[^a-z0-9]+/gi, " ").replace(/\s+/g, " ").trim();
  if (!normalized) return "Ticket";
  return normalized.split(" ").map((part) => (part ? part[0].toUpperCase() + part.slice(1).toLowerCase() : part)).join(" ");
}

function formatFileTime(date = new Date()) {
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${hours}${minutes}`;
}

function normalizeTrouserCuff(value) {
  const cuff = String(value || "").trim();
  if (cuff === "1 3/4 in Cuff" || cuff === '1 3/4" Cuff') return "4.5 cm Cuff";
  if (cuff === "2 in Cuff" || cuff === '2" Cuff') return "5 cm Cuff";
  return cuff;
}


function buildOptions(options, selectedValue) {
  const base = '<option value="">Select</option>';
  const values = options.map((value) => {
    const selected = selectedValue === value ? " selected" : "";
    const label = value === "custom" ? "Custom" : value;
    return `<option value="${value}"${selected}>${label}</option>`;
  }).join("");
  return `${base}${values}`;
}

function buildOptionsFromPairs(options, selectedValue) {
  const base = '<option value="">Select</option>';
  const values = options.map(({ value, label }) => {
    const selected = selectedValue === value ? " selected" : "";
    return `<option value="${value}"${selected}>${label}</option>`;
  }).join("");
  return `${base}${values}`;
}

function buildButtonsOptions(selectedValue) { return buildOptions(["1", "2", "3", "4"], selectedValue); }

/* ---------------------------- measurement rows ---------------------------- */
function stepperRow(uid, field, labelText, item, { disablePlus = false } = {}) {
  const value = formatSignedQuarter(item?.[field] || 0);
  const plusDisabled = disablePlus && normalizeNegativeOnlyValue(item?.[field]) >= 0 ? "disabled" : "";
  return `
    <div class="measurement-row">
      <label for="${field}-${uid}">${labelText}</label>
      <div class="stepper" id="${field}-${uid}">
        <button type="button" class="stepper-btn" data-action="step" data-uid="${uid}" data-field="${field}" data-dir="-1">-</button>
        <input class="stepper-value-input" type="text" inputmode="none" value="${value}" data-action="clear-value" data-uid="${uid}" data-field="${field}" />
        <button type="button" class="stepper-btn" data-action="step" data-uid="${uid}" data-field="${field}" data-dir="1" ${plusDisabled}>+</button>
      </div>
    </div>`;
}

function renderJacketMeasurements(item) {
  const uid = item._uid;
  return `
    <div class="measurement-controls">
      ${stepperRow(uid, "halfBack", "1/2 Back", item)}
      ${stepperRow(uid, "halfWaist", "1/2 Waist", item)}
      ${stepperRow(uid, "shortenBody", "Body Length", item)}
      ${stepperRow(uid, "sleeves", "Sleeve Length", item)}
      ${stepperRow(uid, "sleeveWidth", "Sleeve Width", item)}
      ${stepperRow(uid, "tightenCollar", "Tighten Collar", item, { disablePlus: true })}
      <div class="measurement-row">
        <label for="buttons-${uid}">Sleeve Buttons</label>
        <div class="stepper">
          <select id="buttons-${uid}" class="button-select" data-uid="${uid}" data-field="buttons">
            ${buildButtonsOptions(item?.buttons || "")}
          </select>
        </div>
      </div>
    </div>`;
}

function renderTrouserMeasurements(item) {
  const uid = item._uid;
  return `
    <div class="measurement-controls">
      ${stepperRow(uid, "trouserWaist", "Waist", item)}
      ${stepperRow(uid, "trouserSeat", "Seat", item)}
      ${stepperRow(uid, "trouserThigh", "Thigh", item)}
      ${stepperRow(uid, "trouserKnee", "Knee", item)}
      ${stepperRow(uid, "trouserLegOpening", "Leg Opening", item)}
      ${stepperRow(uid, "trouserInseam", "Inseam", item)}
      <div class="measurement-row">
        <label for="totalLength-${uid}">Total Length</label>
        <div class="stepper">
          <input id="totalLength-${uid}" class="button-select" type="text" data-uid="${uid}" data-field="trouserTotalLength" value="${escapeHtml(item.trouserTotalLength || "")}" />
        </div>
      </div>
      <div class="measurement-row">
        <label for="cuff-${uid}">Cuff Style</label>
        <div class="stepper">
          <select id="cuff-${uid}" class="button-select" data-uid="${uid}" data-field="trouserCuff">
            ${buildOptionsFromPairs([
              { value: "No Cuff", label: "No Cuff" },
              { value: "4.5 cm Cuff", label: "4.5 cm Cuff" },
              { value: "5 cm Cuff", label: "5 cm Cuff" },
            ], normalizeTrouserCuff(item?.trouserCuff))}
          </select>
        </div>
      </div>
    </div>`;
}

function renderShirtMeasurements(item) {
  const uid = item._uid;
  return `
    <div class="measurement-controls">
      ${stepperRow(uid, "shirtSleeve", "Sleeve Length", item, { disablePlus: true })}
      ${stepperRow(uid, "shirtBody", "Body Length", item, { disablePlus: true })}
      ${stepperRow(uid, "shirtSlimBody", "Slim Body", item, { disablePlus: true })}
    </div>`;
}

function formatMultiline(text) {
  const lines = String(text).split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) return "";
  return lines.map((line) => `• ${escapeHtml(line)}`).join("<br>");
}

const GARMENT_TYPE_LABELS = { jacket: "Jacket", trouser: "Trouser", shirt: "Shirt" };
const SIZE_DESCRIPTION_PLACEHOLDERS = {
  jacket: "e.g. 40 Brown Tweed Jacket",
  trouser: "e.g. 33 White Denim",
  shirt: "e.g. 15 Blue Oxford",
};

function renderGarmentPicker(item) {
  return `<div class="garment-picker" role="group" aria-label="Garment type">${Object.entries(GARMENT_TYPE_LABELS)
    .map(
      ([value, label]) =>
        `<button type="button" class="garment-picker-btn${item.garmentType === value ? " is-active" : ""}" data-action="set-garment-type" data-uid="${item._uid}" data-garment-type="${value}">${label}</button>`,
    )
    .join("")}</div>`;
}

function summarizeSingleItem(item) {
  const text = (item.sizeDescription || "").trim();
  return text || "No details yet";
}

/* ------------------------------ item list UI ------------------------------ */
function renderGarmentItems() {
  garmentItemsEl.innerHTML = garmentItems
    .map((item) => {
      const canRemove = garmentItems.length > 1;
      const canCollapse = hasGarmentData(item);
      const isCollapsed = canCollapse && collapsedItemIds.has(String(item._uid));
      const typeLabel = GARMENT_TYPE_LABELS[item.garmentType] || "Garment";

      if (isCollapsed) {
        return `
          <section class="repeat-item is-collapsed" data-uid="${item._uid}">
            <button type="button" class="repeat-item-summary" data-action="toggle-collapse" data-uid="${item._uid}">
              <span class="repeat-item-summary-type">${typeLabel}</span>
              <span class="repeat-item-summary-detail">${escapeHtml(summarizeSingleItem(item))}</span>
              <span class="repeat-item-summary-chevron" aria-hidden="true">▸</span>
            </button>
            ${canRemove ? `<button type="button" class="remove-item-btn" data-action="remove" data-uid="${item._uid}">Remove</button>` : ""}
          </section>`;
      }

      const measurementBlock = item.garmentType === "jacket" ? renderJacketMeasurements(item)
        : item.garmentType === "trouser" ? renderTrouserMeasurements(item)
        : renderShirtMeasurements(item);

      return `
        <section class="repeat-item" data-uid="${item._uid}">
          <div class="repeat-item-header">
            ${renderGarmentPicker(item)}
            <div class="repeat-item-header-actions">
              ${canCollapse ? `<button type="button" class="link-action-btn" data-action="toggle-collapse" data-uid="${item._uid}">Collapse</button>` : ""}
              ${canRemove ? `<button type="button" class="remove-item-btn" data-action="remove" data-uid="${item._uid}">Remove</button>` : ""}
            </div>
          </div>
          <div class="garment-field size-desc-row">
            <label for="sizeDescription-${item._uid}">Garment Size &amp; Description</label>
            <input id="sizeDescription-${item._uid}" type="text" placeholder="${SIZE_DESCRIPTION_PLACEHOLDERS[item.garmentType] || "e.g. Size and description"}" data-uid="${item._uid}" data-field="sizeDescription" value="${escapeHtml(item.sizeDescription)}" />
          </div>
          ${measurementBlock}
          <label for="adjustments-${item._uid}">Additional Notes</label>
          <textarea id="adjustments-${item._uid}" rows="3" data-uid="${item._uid}" data-field="adjustments">${escapeHtml(item.adjustments)}</textarea>
        </section>`;
    })
    .join("");
}

function handleDynamicInput(event) {
  const target = event.target;
  if (!target?.dataset) return;
  const { field, uid, action } = target.dataset;
  if (!field || uid === undefined) return;
  const item = garmentItems.find((it) => String(it._uid) === uid);
  if (!item) return;

  if (action === "clear-value") {
    const rawValue = target.value.trim();
    if (!rawValue) {
      item[field] = 0;
      renderGarmentItems();
      onInputChange();
      return;
    }
    const formatted = formatSignedQuarter(item[field] || 0);
    if (target.value !== formatted) target.value = formatted;
    return;
  }

  item[field] = target.value;
  onInputChange();
}

function handleDynamicClick(event) {
  const target = event.target.closest("[data-action]");
  if (!target?.dataset) return;
  const { action, uid } = target.dataset;
  if (!action || uid === undefined) return;

  if (action === "toggle-collapse") {
    if (collapsedItemIds.has(uid)) collapsedItemIds.delete(uid);
    else collapsedItemIds.add(uid);
    renderGarmentItems();
    return;
  }

  if (action === "set-garment-type") {
    const newType = target.dataset.garmentType;
    const index = garmentItems.findIndex((it) => String(it._uid) === uid);
    if (index === -1 || !newType || garmentItems[index].garmentType === newType) return;
    garmentItems[index] = changeItemGarmentType(garmentItems[index], newType);
    renderGarmentItems();
    onInputChange();
    return;
  }

  if (action === "step") {
    const field = target.dataset.field;
    const dir = Number(target.dataset.dir);
    if (!field || Number.isNaN(dir)) return;
    const item = garmentItems.find((it) => String(it._uid) === uid);
    if (!item) return;
    const current = Number(item[field]) || 0;
    const next = Math.round((current + dir * 0.5) * 2) / 2;
    item[field] = isNegativeOnlyField(field) ? normalizeNegativeOnlyValue(next) : next;
    renderGarmentItems();
    onInputChange();
    return;
  }

  if (action === "remove") {
    garmentItems = garmentItems.filter((it) => String(it._uid) !== uid);
    collapsedItemIds.delete(uid);
    if (!garmentItems.length) garmentItems.push(createEmptyItem());
    renderGarmentItems();
    onInputChange();
  }
}

/* ------------------------------- misc form fns ------------------------------ */
function getBalanceDueValue() {
  const selected = document.querySelector('input[name="balanceDue"]:checked');
  return selected ? selected.value : "";
}

function normalizeBalanceValue(value) {
  if (value === "Yes") return "Due";
  if (value === "No") return "Paid";
  return value === "Paid" || value === "Due" ? value : "";
}

function formatBalanceDisplay(value) {
  const balance = normalizeBalanceValue(value);
  return balance ? `Balance ${balance}` : "";
}

function getTailorDisplayName() {
  return tailorInput.value || "";
}

function formatDueDate(dateValue) {
  if (!dateValue) return "";
  return new Date(`${dateValue}T00:00:00`).toLocaleDateString();
}

function isRushDueDate(dateValue, generatedAt) {
  if (!dateValue) return false;
  const due = new Date(`${dateValue}T00:00:00`);
  const current = new Date(generatedAt.getFullYear(), generatedAt.getMonth(), generatedAt.getDate());
  if (due < current) return false;
  let businessDays = 0;
  const cursor = new Date(current);
  cursor.setDate(cursor.getDate() + 1);
  while (cursor <= due) {
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) businessDays += 1;
    cursor.setDate(cursor.getDate() + 1);
  }
  return businessDays <= 5;
}

function getDueDateValue() { return dueDateHiddenInput ? dueDateHiddenInput.value : dueDateInput.value; }

function setDueDateValue(value) {
  if (dueDateHiddenInput) { dueDateHiddenInput.value = value || ""; updateDueDateDisplay(); return; }
  dueDateInput.value = value || "";
}

function updateDueDateDisplay() {
  if (!dueDateDisplayInput) return;
  const value = dueDateHiddenInput ? dueDateHiddenInput.value : "";
  dueDateDisplayInput.value = value ? new Date(`${value}T00:00:00`).toLocaleDateString() : "";
}

/* ---------------------------- ticket data builder --------------------------- */
function buildGarmentSections(items = garmentItems) {
  const sections = [];
  const filled = items.filter(hasGarmentData);

  const buildJacketSection = (entry, label) => {
    const measurements = [];
    if (Number(entry.halfBack)) measurements.push({ label: "1/2 Back", value: `${formatSignedQuarter(entry.halfBack)} cm` });
    if (Number(entry.halfWaist)) measurements.push({ label: "1/2 Waist", value: `${formatSignedQuarter(entry.halfWaist)} cm` });
    if (Number(entry.shortenBody)) measurements.push({ label: "Body Length", value: `${formatSignedQuarter(entry.shortenBody)} cm` });
    if (Number(entry.sleeves)) measurements.push({ label: "Sleeve Length", value: `${formatSignedQuarter(entry.sleeves)} cm` });
    if (Number(entry.sleeveWidth)) measurements.push({ label: "Sleeve Width", value: `${formatSignedQuarter(entry.sleeveWidth)} cm` });
    const tightenCollar = normalizeNegativeOnlyValue(entry.tightenCollar);
    if (Number(tightenCollar)) measurements.push({ label: "Tighten Collar", value: `${formatSignedQuarter(tightenCollar)} cm` });
    if ((entry.buttons || "").trim()) measurements.push({ label: "Sleeve Buttons", value: entry.buttons });
    const sizeDesc = (entry.sizeDescription || "").trim();
    return { label, sizeDesc, measurements, notes: (entry.adjustments || "").trim() };
  };

  const buildTrouserSection = (entry, label) => {
    const measurements = [];
    if (Number(entry.trouserWaist)) measurements.push({ label: "Waist", value: `${formatSignedQuarter(entry.trouserWaist)} cm` });
    if (Number(entry.trouserSeat)) measurements.push({ label: "Seat", value: `${formatSignedQuarter(entry.trouserSeat)} cm` });
    if (Number(entry.trouserThigh)) measurements.push({ label: "Thigh", value: `${formatSignedQuarter(entry.trouserThigh)} cm` });
    if (Number(entry.trouserKnee)) measurements.push({ label: "Knee", value: `${formatSignedQuarter(entry.trouserKnee)} cm` });
    if (Number(entry.trouserLegOpening)) measurements.push({ label: "Leg Opening", value: `${formatSignedQuarter(entry.trouserLegOpening)} cm` });
    if (Number(entry.trouserInseam)) measurements.push({ label: "Inseam", value: `${formatSignedQuarter(entry.trouserInseam)} cm` });
    if ((entry.trouserTotalLength || "").trim()) measurements.push({ label: "Total Length", value: `${entry.trouserTotalLength} cm` });
    if ((entry.trouserCuff || "").trim()) {
      const cuffLabel = normalizeTrouserCuff(entry.trouserCuff).replace(/\s*Cuff\b/i, "").trim();
      measurements.push({ label: "Cuff Style", value: cuffLabel });
    }
    const sizeDesc = (entry.sizeDescription || "").trim();
    return { label, sizeDesc, measurements, notes: (entry.adjustments || "").trim() };
  };

  const buildShirtSection = (entry, label) => {
    const measurements = [];
    const shirtSleeve = normalizeNegativeOnlyValue(entry.shirtSleeve);
    if (Number(shirtSleeve)) measurements.push({ label: "Sleeve Length", value: `${formatSignedQuarter(shirtSleeve)} cm` });
    const shirtBody = normalizeNegativeOnlyValue(entry.shirtBody);
    if (Number(shirtBody)) measurements.push({ label: "Body Length", value: `${formatSignedQuarter(shirtBody)} cm` });
    const shirtSlimBody = normalizeNegativeOnlyValue(entry.shirtSlimBody);
    if (Number(shirtSlimBody)) measurements.push({ label: "Slim Body", value: `${formatSignedQuarter(shirtSlimBody)} cm` });
    const sizeDesc = (entry.sizeDescription || "").trim();
    return { label, sizeDesc, measurements, notes: (entry.adjustments || "").trim() };
  };

  // Grouped by type on the printed ticket (all jackets together, then
  // trousers, then shirts) regardless of the order items were added in —
  // reads better on the physical document than the entry order would.
  ["jacket", "trouser", "shirt"].forEach((type) => {
    const typeLabel = GARMENT_TYPE_LABELS[type];
    const builder = type === "jacket" ? buildJacketSection : type === "trouser" ? buildTrouserSection : buildShirtSection;
    filled
      .filter((entry) => entry.garmentType === type)
      .forEach((entry, idx) => {
        sections.push(builder(entry, idx === 0 ? typeLabel : `${typeLabel} ${idx + 1}`));
      });
  });

  // Fold size/description into the measurement list so both the screen
  // preview and the docx show it as the first line of the section.
  return sections.map((section) => ({
    ...section,
    measurements: section.sizeDesc ? [{ label: "Garment", value: section.sizeDesc }, ...section.measurements] : section.measurements,
  }));
}

function renderTicketMarkup(ticketData) {
  const sectionHtml = ticketData.garmentSections.map((section) => {
    const measurementHtml = section.measurements.map((m) => `<p><strong>${escapeHtml(m.label)}:</strong> ${escapeHtml(m.value)}</p>`).join("");
    const notesHtml = section.notes ? `<p><strong>Notes:</strong><br>${formatMultiline(section.notes)}</p>` : "";
    return `<section class="garment-output-block"><p><strong>${escapeHtml(section.label)}</strong></p>${measurementHtml}${notesHtml}</section>`;
  }).join("");

  return `
    <p class="doc-title">J. Mueser</p>
    <p class="doc-subtitle">Alterations Ticket</p>
    ${ticketData.rush ? '<p class="rush-flag">★ Rush</p>' : ""}
    <p class="client-name"><strong>Name:</strong> ${escapeHtml(ticketData.customerName)}</p>
    <p class="ticket-detail"><strong>Tailor:</strong> ${escapeHtml(ticketData.tailor)}</p>
    <p class="ticket-detail"><strong>Salesperson:</strong> ${escapeHtml(ticketData.salesperson)}</p>
    <p class="ticket-detail"><strong>Due Date:</strong> ${escapeHtml(ticketData.dueDate)}</p>
    ${ticketData.balanceDisplay ? `<p class="ticket-detail"><strong>${escapeHtml(ticketData.balanceDisplay)}</strong></p>` : ""}
    <p class="meta">Created ${escapeHtml(ticketData.createdDisplay)}</p>
    ${sectionHtml}
  `;
}

function renderOutput() {
  printArea.innerHTML = renderTicketMarkup(buildTicketData());
}

function buildTicketData() {
  const now = new Date();
  const dueDateValue = getDueDateValue();
  return {
    customerName: customerNameInput.value.trim(),
    tailor: getTailorDisplayName(),
    salesperson: salespersonInput.value,
    dueDate: formatDueDate(dueDateValue),
    rush: isRushDueDate(dueDateValue, now),
    balanceDisplay: formatBalanceDisplay(getBalanceDueValue()),
    createdDisplay: now.toLocaleString(),
    createdAt: now.toISOString(),
    garmentSections: buildGarmentSections(),
  };
}

// Same shape as buildTicketData(), but sourced from a saved ticket's
// formState instead of the live form — used to Print/Download straight
// from the Ticket Log without touching what's currently in New Ticket.
// Reads garment data from a saved formState, handling both the current flat
// shape and the older {jackets, trousers, shirts} shape from before the
// single-list redesign — so historical tickets keep working unmodified.
// Every item gets a freshly generated uid, since any uid saved in old data
// could collide with this session's counter.
// Older saved tickets (from before Size and Description were merged into
// one field) still have separate item.size / item.description values.
// Combine them into sizeDescription so those tickets display correctly.
function migrateSizeDescription(item) {
  if (item.sizeDescription) return item;
  const combined = [item.size, item.description].map((v) => (v || "").trim()).filter(Boolean).join(" ");
  return combined ? { ...item, sizeDescription: combined } : item;
}

function normalizeGarmentItems(state) {
  let items;
  if (Array.isArray(state.garmentItems) && state.garmentItems.length) {
    items = state.garmentItems.map((item) => ({ ...createEmptyItem(item.garmentType || "jacket"), ...item }));
  } else {
    items = [
      ...(state.jackets || []).map((item) => ({ ...createEmptyItem("jacket"), ...item, garmentType: "jacket" })),
      ...(state.trousers || []).map((item) => ({ ...createEmptyItem("trouser"), ...item, garmentType: "trouser" })),
      ...(state.shirts || []).map((item) => ({ ...createEmptyItem("shirt"), ...item, garmentType: "shirt" })),
    ];
  }
  items = items.map((item) => ({ ...migrateSizeDescription(item), _uid: uidCounter++ }));
  return items.length ? items : [createEmptyItem()];
}

function buildTicketDataFromState(state) {
  const now = new Date();
  const dueDateValue = state.dueDate || "";
  return {
    customerName: (state.customerName || "").trim(),
    tailor: state.tailor || "",
    salesperson: state.salesperson || "",
    dueDate: formatDueDate(dueDateValue),
    rush: isRushDueDate(dueDateValue, now),
    balanceDisplay: formatBalanceDisplay(state.balanceDue),
    createdDisplay: now.toLocaleString(),
    createdAt: now.toISOString(),
    garmentSections: buildGarmentSections(normalizeGarmentItems(state)),
  };
}

/* --------------------------------- storage --------------------------------- */
function buildState(savedAt = new Date().toISOString()) {
  return {
    customerName: customerNameInput.value,
    tailor: tailorInput.value,
    salesperson: salespersonInput.value,
    dueDate: getDueDateValue(),
    balanceDue: getBalanceDueValue(),
    garmentItems,
    savedAt,
  };
}

function saveToStorage() {
  const state = buildState();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  saveStatus.textContent = `Autosaved at ${new Date(state.savedAt).toLocaleTimeString()}`;
}

function applyState(parsed) {
  customerNameInput.value = parsed.customerName || "";
  tailorInput.value = parsed.tailor || "";
  salespersonInput.value = parsed.salesperson || "";
  setDueDateValue(parsed.dueDate || "");

  garmentItems = normalizeGarmentItems(parsed);

  // Reopening a ticket with existing garments starts collapsed for a quick
  // scan; a brand-new blank item (nothing to summarize) stays expanded.
  collapsedItemIds = new Set(garmentItems.filter(hasGarmentData).map((item) => String(item._uid)));
  renderGarmentItems();

  const balanceDue = normalizeBalanceValue(parsed.balanceDue);
  const selectedBalance = document.querySelector('input[name="balanceDue"]:checked');
  if (selectedBalance) selectedBalance.checked = false;
  if (balanceDue) {
    const target = document.querySelector(`input[name="balanceDue"][value="${balanceDue}"]`);
    if (target) target.checked = true;
  }

  renderOutput();
}

function loadFromStorage() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    garmentItems = [createEmptyItem()];
    renderGarmentItems();
    renderOutput();
    return;
  }
  try {
    const parsed = JSON.parse(stored);
    applyState(parsed);
    if (parsed.savedAt) saveStatus.textContent = `Last saved ${new Date(parsed.savedAt).toLocaleString()}`;
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    garmentItems = [createEmptyItem()];
    renderGarmentItems();
    renderOutput();
  }
}

let saveTimer;
function onInputChange() {
  renderOutput();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveToStorage, 350);
}

function validateBeforePrint() {
  const missing = [];
  if (!customerNameInput.value.trim()) missing.push("Name");
  if (!tailorInput.value.trim()) missing.push("Tailor");
  if (!salespersonInput.value.trim()) missing.push("Salesperson");
  if (!getDueDateValue().trim()) missing.push("Due Date");
  return missing;
}

function clearAllFields() {
  clearTimeout(saveTimer);
  setEditingTicket(null);
  customerNameInput.value = "";
  tailorInput.value = "";
  salespersonInput.value = "";
  setDueDateValue("");
  const selectedBalance = document.querySelector('input[name="balanceDue"]:checked');
  if (selectedBalance) selectedBalance.checked = false;

  garmentItems = [createEmptyItem()];
  collapsedItemIds = new Set();
  renderGarmentItems();
  localStorage.removeItem(STORAGE_KEY);
  saveStatus.textContent = "Form cleared";
  renderOutput();
}

/* ----------------------------------- auth ---------------------------------- */
function updateAuthGate() {
  const isSignedIn = auth.hasValidSession();
  authGate.hidden = isSignedIn;
  appShell.hidden = !isSignedIn;
  authStatus.textContent = isSignedIn ? "Signed in with Google." : "J.Mueser accounts only.";
  accountEmail.textContent = isSignedIn ? auth.getSessionEmail() : "";
  if (isSignedIn) refreshTicketLog();
}

authLoginBtn.addEventListener("click", async () => {
  authStatus.textContent = "Opening Google sign-in…";
  try {
    await auth.signIn("consent");
    updateAuthGate();
  } catch (err) {
    authStatus.textContent = err.message;
  }
});

signOutBtn.addEventListener("click", () => {
  auth.signOut();
  updateAuthGate();
  saveStatus.textContent = "Signed out of Google";
});

/* --------------------------------- view nav --------------------------------- */
function setView(view) {
  const tabs = [
    { key: "newTicket", nav: navNewTicket, viewEl: viewNewTicket },
    { key: "ticketLog", nav: navTicketLog, viewEl: viewTicketLog },
    { key: "studio", nav: navStudio, viewEl: viewStudio },
  ];
  tabs.forEach(({ key, nav, viewEl }) => {
    const isActive = key === view;
    nav.classList.toggle("is-active", isActive);
    nav.setAttribute("aria-selected", isActive ? "true" : "false");
    viewEl.classList.toggle("is-active", isActive);
  });
  if (view === "ticketLog") refreshTicketLog();
  if (view === "studio") refreshStudioLog();
}

navNewTicket.addEventListener("click", () => setView("newTicket"));
navTicketLog.addEventListener("click", () => setView("ticketLog"));
navStudio.addEventListener("click", () => setView("studio"));

/* -------------------------------- ticket log --------------------------------- */
function createMultiSelectFilter({ wrap, btn, panel, options, allBtn, noneBtn }, values, allLabel, onChange) {
  let selected = new Set();

  const updateLabel = () => {
    if (selected.size === 0 || selected.size === values.length) btn.textContent = allLabel;
    else if (selected.size === 1) btn.textContent = [...selected][0];
    else btn.textContent = `${selected.size} selected`;
  };

  options.innerHTML = values
    .map(
      (value) => `
      <label class="filter-multiselect-option">
        <input type="checkbox" value="${escapeHtml(value)}" />
        <span>${escapeHtml(value)}</span>
      </label>`,
    )
    .join("");

  btn.addEventListener("click", () => {
    panel.hidden = !panel.hidden;
  });

  document.addEventListener("click", (event) => {
    if (!wrap.contains(event.target)) panel.hidden = true;
  });

  options.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[type="checkbox"]');
    if (!checkbox) return;
    if (checkbox.checked) selected.add(checkbox.value);
    else selected.delete(checkbox.value);
    updateLabel();
    onChange();
  });

  allBtn.addEventListener("click", () => {
    selected = new Set(values);
    options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = true; });
    updateLabel();
    onChange();
  });

  noneBtn.addEventListener("click", () => {
    selected = new Set();
    options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
    updateLabel();
    onChange();
  });

  updateLabel();
  return {
    matches: (value) => selected.size === 0 || selected.has(value),
    reset: () => {
      selected = new Set();
      options.querySelectorAll('input[type="checkbox"]').forEach((cb) => { cb.checked = false; });
      updateLabel();
    },
  };
}

const statusFilter = createMultiSelectFilter(statusFilterEls, TICKET_STATUSES, "Status", () => renderTicketLog());
const tailorFilter = createMultiSelectFilter(tailorFilterEls, ["Luis", "Jesus", "Sam"], "Tailor", () => renderTicketLog());
const salespersonFilter = createMultiSelectFilter(salespersonFilterEls, SALESPEOPLE, "Salesperson", () => renderTicketLog());

// Summarizes what's on a ticket (e.g. "Jacket ×2, Trouser") straight from
// its saved formState — no separate sheet column needed since the full
// garment data is already stored there.
/* --------------------------------- notes modal -------------------------------- */
let notesModalRowNumber = null;

function openNotesModal(ticket) {
  notesModalRowNumber = ticket.rowNumber;
  notesModalSubtitle.textContent = ticket.customerName || "This ticket";
  notesModalTextarea.value = ticket.notes || "";
  notesModal.hidden = false;
  notesModalTextarea.focus();
}

function closeNotesModal() {
  notesModal.hidden = true;
  notesModalRowNumber = null;
}

notesModalCancelBtn.addEventListener("click", closeNotesModal);

notesModalSaveBtn.addEventListener("click", async () => {
  if (notesModalRowNumber === null) return;
  const rowNumber = notesModalRowNumber;
  const newNotes = notesModalTextarea.value;
  const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
  if (!ticket) { closeNotesModal(); return; }

  notesModalSaveBtn.disabled = true;
  notesModalSaveBtn.textContent = "Saving…";
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    await updateTicketNotes(token, ledgerId, rowNumber, newNotes);
    ticket.notes = newNotes;
    ticketLogStatus.textContent = `Saved note for ${ticket.customerName || "ticket"}.`;
    closeNotesModal();
    renderTicketLog(); // refresh so the icon's filled/empty state updates
  } catch (err) {
    ticketLogStatus.textContent = `Could not save note: ${err.message}`;
  } finally {
    notesModalSaveBtn.disabled = false;
    notesModalSaveBtn.textContent = "Save Note";
  }
});

function summarizeGarments(formState) {
  if (!formState) return "—";
  const items = normalizeGarmentItems(formState).filter(hasGarmentData);
  const counts = [
    ["Jacket", items.filter((item) => item.garmentType === "jacket").length],
    ["Trouser", items.filter((item) => item.garmentType === "trouser").length],
    ["Shirt", items.filter((item) => item.garmentType === "shirt").length],
  ];
  const parts = counts.filter(([, count]) => count > 0).map(([label, count]) => (count > 1 ? `${label} ×${count}` : label));
  return parts.length ? parts.join(", ") : "—";
}

function renderTicketLog() {
  const search = ticketSearch.value.trim().toLowerCase();
  const rushOnly = ticketRushFilter.checked;

  const filtered = ticketCache.filter((ticket) => {
    if (search && !ticket.customerName.toLowerCase().includes(search)) return false;
    if (!statusFilter.matches(ticket.status)) return false;
    if (!tailorFilter.matches(ticket.tailor)) return false;
    if (!salespersonFilter.matches(ticket.salesperson)) return false;
    if (rushOnly && !ticket.rush) return false;
    return true;
  });

  ticketLogStatus.textContent =
    filtered.length === ticketCache.length
      ? `${ticketCache.length} ticket${ticketCache.length === 1 ? "" : "s"} logged.`
      : `Showing ${filtered.length} ticket${filtered.length === 1 ? "" : "s"} out of ${ticketCache.length} ticket${ticketCache.length === 1 ? "" : "s"} logged.`;

  if (!filtered.length) {
    ticketLogBody.innerHTML = `<tr><td colspan="10" class="log-status">No tickets match yet.</td></tr>`;
    return;
  }

  ticketLogBody.innerHTML = filtered
    .slice()
    .reverse()
    .map((ticket) => {
      const statusOptions = TICKET_STATUSES.map((status) => `<option value="${status}" ${status === ticket.status ? "selected" : ""}>${status}</option>`).join("");
      return `
        <tr data-row="${ticket.rowNumber}" data-ticket-id="${escapeHtml(ticket.id)}">
          <td data-label="Client"><span class="cell-value">${escapeHtml(ticket.customerName)}</span></td>
          <td data-label="Garments"><span class="cell-value garments-cell">${escapeHtml(summarizeGarments(ticket.formState))}</span></td>
          <td data-label="Tailor"><span class="cell-value">${escapeHtml(ticket.tailor)}</span></td>
          <td data-label="Salesperson"><span class="cell-value">${escapeHtml(ticket.salesperson)}</span></td>
          <td data-label="Due Date"><span class="cell-value${ticket.rush ? " due-date-rush" : ""}">${escapeHtml(ticket.dueDate)}</span></td>
          <td data-label="Balance"><span class="cell-value">${escapeHtml((ticket.balance || "—").replace(/^Balance\s+/i, ""))}</span></td>
          <td data-label="Status"><span class="cell-value"><select class="status-select" data-row="${ticket.rowNumber}" data-status="${escapeHtml(ticket.status)}">${statusOptions}</select></span></td>
          <td data-label="Status Date" class="status-date-cell"><span class="cell-value">${escapeHtml(ticket.statusDate || "—")}</span></td>
          <td data-label="Notes"><span class="cell-value"><button type="button" class="icon-btn notes-btn${(ticket.notes || "").trim() ? " has-notes" : ""}" data-row="${ticket.rowNumber}" title="${(ticket.notes || "").trim() ? "View/edit note" : "Add a note"}" aria-label="Notes">${ICONS.notes}</button></span></td>
          <td data-label="Ticket Actions">
            <span class="cell-value ticket-actions">
              <button type="button" class="icon-btn print-btn" data-row="${ticket.rowNumber}" title="Print" aria-label="Print">${ICONS.print}</button>
              <button type="button" class="icon-btn edit-btn" data-row="${ticket.rowNumber}" title="Edit" aria-label="Edit">${ICONS.edit}</button>
              <button type="button" class="icon-btn download-btn" data-row="${ticket.rowNumber}" title="Download" aria-label="Download">${ICONS.download}</button>
              <button type="button" class="icon-btn icon-btn-danger delete-btn" data-row="${ticket.rowNumber}" title="Delete" aria-label="Delete">${ICONS.delete}</button>
            </span>
          </td>
        </tr>`;
    })
    .join("");
}

async function refreshTicketLog() {
  if (!auth.hasValidSession()) {
    ticketLogStatus.textContent = "Sign in to load the ticket log.";
    return;
  }
  ticketLogStatus.textContent = "Loading tickets…";
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    ticketCache = await listTickets(token, ledgerId);
    const openCount = ticketCache.filter((t) => t.status === "Open" || t.status === "In Progress").length;
    const rushCount = ticketCache.filter((t) => t.rush && t.status !== "Completed").length;
    logStats.innerHTML = `<strong>${ticketCache.length}</strong> tickets · <span class="stat-amber">${openCount} open</span> · <span class="stat-red">${rushCount} rush</span>`;

    // Row numbers shift when any row is deleted — keep an in-progress edit
    // pointed at the right row (or drop out of edit mode if it's gone).
    if (editingTicket) {
      const match = ticketCache.find(
        (t) => (editingTicket.id && t.id === editingTicket.id) || (editingTicket.driveFileId && t.driveFileId === editingTicket.driveFileId),
      );
      if (match) editingTicket.rowNumber = match.rowNumber;
      else setEditingTicket(null);
    }

    renderTicketLog();
  } catch (err) {
    ticketLogStatus.textContent = `Could not load ticket log: ${err.message}`;
  }
}

refreshTicketsBtn.addEventListener("click", refreshTicketLog);
ticketSearch.addEventListener("input", renderTicketLog);
ticketRushFilter.addEventListener("change", renderTicketLog);

clearFiltersBtn.addEventListener("click", () => {
  ticketSearch.value = "";
  ticketRushFilter.checked = false;
  statusFilter.reset();
  tailorFilter.reset();
  salespersonFilter.reset();
  renderTicketLog();
});

ticketLogBody.addEventListener("change", async (event) => {
  const select = event.target.closest(".status-select");
  if (!select) return;
  const rowNumber = Number(select.dataset.row);
  const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
  if (!ticket) return;
  const previousStatus = ticket.status;
  const previousStatusDate = ticket.statusDate;
  ticket.status = select.value;
  select.dataset.status = select.value;
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    const now = await updateTicketStatus(token, ledgerId, rowNumber, select.value);
    ticket.statusDate = now;
    const dateCell = select.closest("tr")?.querySelector(".status-date-cell .cell-value");
    if (dateCell) dateCell.textContent = now;
    ticketLogStatus.textContent = `Updated ${ticket.customerName} to ${select.value}.`;
  } catch (err) {
    ticket.status = previousStatus;
    ticket.statusDate = previousStatusDate;
    select.value = previousStatus;
    select.dataset.status = previousStatus;
    ticketLogStatus.textContent = `Could not update status: ${err.message}`;
  }
});

ticketLogBody.addEventListener("click", async (event) => {
  const notesBtn = event.target.closest(".notes-btn");
  if (notesBtn) {
    const rowNumber = Number(notesBtn.dataset.row);
    const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
    if (!ticket) return;
    openNotesModal(ticket);
    return;
  }

  const editBtn = event.target.closest(".edit-btn");
  if (editBtn) {
    const rowNumber = Number(editBtn.dataset.row);
    const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
    if (!ticket || !ticket.formState || !Object.keys(ticket.formState).length) {
      ticketLogStatus.textContent = "This older ticket doesn't have editable form data.";
      return;
    }
    applyState(ticket.formState);
    setEditingTicket({ id: ticket.id, rowNumber: ticket.rowNumber, driveFileId: ticket.driveFileId, customerName: ticket.customerName });
    setView("newTicket");
    saveToStorage();
    return;
  }

  const printBtnEl = event.target.closest(".print-btn");
  if (printBtnEl) {
    const rowNumber = Number(printBtnEl.dataset.row);
    const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
    if (!ticket || !ticket.formState || !Object.keys(ticket.formState).length) {
      ticketLogStatus.textContent = "This older ticket doesn't have printable form data.";
      return;
    }
    printArea.innerHTML = renderTicketMarkup(buildTicketDataFromState(ticket.formState));
    const restoreLivePreview = () => {
      renderOutput(); // put the live form's own preview back afterward
      window.removeEventListener("afterprint", restoreLivePreview);
    };
    window.addEventListener("afterprint", restoreLivePreview);
    window.print();
    return;
  }

  const downloadBtnEl = event.target.closest(".download-btn");
  if (downloadBtnEl) {
    const rowNumber = Number(downloadBtnEl.dataset.row);
    const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
    if (!ticket || !ticket.formState || !Object.keys(ticket.formState).length) {
      ticketLogStatus.textContent = "This older ticket doesn't have downloadable form data.";
      return;
    }
    downloadBtnEl.disabled = true;
    try {
      const ticketData = buildTicketDataFromState(ticket.formState);
      const blob = await buildDocxBlob(ticketData);
      const filename = ticket.docxFilename || `JM_ALT_${formatFileBaseName(ticketData.customerName)}.docx`;
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url; link.download = filename;
      document.body.appendChild(link); link.click(); link.remove();
      URL.revokeObjectURL(url);
      ticketLogStatus.textContent = `Downloaded ${filename}`;
    } catch (err) {
      ticketLogStatus.textContent = `Download failed: ${err.message}`;
    } finally {
      downloadBtnEl.disabled = false;
    }
    return;
  }

  const deleteBtn = event.target.closest(".delete-btn");
  if (deleteBtn) {
    const rowNumber = Number(deleteBtn.dataset.row);
    const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
    if (!ticket) return;
    const fileNote = ticket.driveFileId
      ? "This removes it from the log and moves its .docx to Drive's trash (recoverable there for a while, not permanent)."
      : "This removes it from the log. No linked file was found to remove.";
    const confirmed = window.confirm(`Delete the ticket for ${ticket.customerName || "this client"}? ${fileNote}`);
    if (!confirmed) return;

    deleteBtn.disabled = true;
    try {
      const token = await auth.getValidToken();
      const ledgerId = await getOrCreateLedger(token);
      if (ticket.driveFileId) {
        try {
          await trashDriveFile(token, ticket.driveFileId);
        } catch {
          // Row deletion still proceeds even if the file couldn't be trashed
          // (e.g. already removed, or permissions changed).
        }
      }
      await deleteTicketRow(token, ledgerId, rowNumber);
      if (editingTicket && editingTicket.rowNumber === rowNumber) setEditingTicket(null);
      ticketLogStatus.textContent = `Deleted ${ticket.customerName || "ticket"} from the log.`;
      await refreshTicketLog(); // row numbers shift after a delete, so reload fresh
    } catch (err) {
      deleteBtn.disabled = false;
      ticketLogStatus.textContent = `Could not delete: ${err.message}`;
    }
  }
});

/* ---------------------------------- studio ----------------------------------- */
const studioClientNameInput = el("studioClientName");
const studioGarmentDescriptionInput = el("studioGarmentDescription");
const studioArrivalDateInput = el("studioArrivalDate");
const studioSalespersonInput = el("studioSalesperson");
const studioSaveBtn = el("studioSaveBtn");
const studioSaveStatus = el("studioSaveStatus");

const studioSearch = el("studioSearch");
const studioClearFiltersBtn = el("studioClearFiltersBtn");
const studioRefreshBtn = el("studioRefreshBtn");
const studioLogStatus = el("studioLogStatus");
const studioLogBody = el("studioLogBody");
const studioStats = el("studioStats");

const studioStatusFilterEls = {
  wrap: el("studioStatusFilterWrap"), btn: el("studioStatusFilterBtn"), panel: el("studioStatusFilterPanel"),
  options: el("studioStatusFilterOptions"), allBtn: el("studioStatusFilterAll"), noneBtn: el("studioStatusFilterNone"),
};
const studioSalespersonFilterEls = {
  wrap: el("studioSalespersonFilterWrap"), btn: el("studioSalespersonFilterBtn"), panel: el("studioSalespersonFilterPanel"),
  options: el("studioSalespersonFilterOptions"), allBtn: el("studioSalespersonFilterAll"), noneBtn: el("studioSalespersonFilterNone"),
};

let studioCache = [];
const studioStatusFilter = createMultiSelectFilter(studioStatusFilterEls, STUDIO_STATUSES, "Status", () => renderStudioLog());
const studioSalespersonFilter = createMultiSelectFilter(studioSalespersonFilterEls, SALESPEOPLE, "Salesperson", () => renderStudioLog());

function clearStudioForm() {
  studioClientNameInput.value = "";
  studioGarmentDescriptionInput.value = "";
  studioArrivalDateInput.value = "";
  studioSalespersonInput.value = "";
}

studioSaveBtn.addEventListener("click", async () => {
  const clientName = studioClientNameInput.value.trim();
  const garmentDescription = studioGarmentDescriptionInput.value.trim();
  if (!clientName || !garmentDescription) {
    alert("Please fill in at least Client Name and Garment Description.");
    return;
  }

  studioSaveBtn.disabled = true;
  studioSaveStatus.textContent = "Saving…";
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    await getOrCreateStudioTab(token, ledgerId);
    await appendStudioEntry(token, ledgerId, {
      id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      createdAt: new Date().toLocaleString(),
      clientName,
      garmentDescription,
      arrivalDate: studioArrivalDateInput.value ? new Date(`${studioArrivalDateInput.value}T00:00:00`).toLocaleDateString() : "",
      salesperson: studioSalespersonInput.value,
      status: STUDIO_STATUSES[0],
    });
    clearStudioForm();
    studioSaveStatus.textContent = "Added to studio.";
    refreshStudioLog();
  } catch (err) {
    studioSaveStatus.textContent = `Could not save: ${err.message}`;
    alert(`Add to Studio failed: ${err.message}`);
  } finally {
    studioSaveBtn.disabled = false;
  }
});

function renderStudioLog() {
  const search = studioSearch.value.trim().toLowerCase();

  const filtered = studioCache.filter((entry) => {
    if (search && !entry.clientName.toLowerCase().includes(search)) return false;
    if (!studioStatusFilter.matches(entry.status)) return false;
    if (!studioSalespersonFilter.matches(entry.salesperson)) return false;
    return true;
  });

  studioLogStatus.textContent =
    filtered.length === studioCache.length
      ? `${studioCache.length} item${studioCache.length === 1 ? "" : "s"} in studio.`
      : `Showing ${filtered.length} item${filtered.length === 1 ? "" : "s"} out of ${studioCache.length} logged.`;

  if (!filtered.length) {
    studioLogBody.innerHTML = `<tr><td colspan="7" class="log-status">No studio items match yet.</td></tr>`;
    return;
  }

  studioLogBody.innerHTML = filtered
    .slice()
    .reverse()
    .map((entry) => {
      const statusOptions = STUDIO_STATUSES.map((status) => `<option value="${status}" ${status === entry.status ? "selected" : ""}>${status}</option>`).join("");
      const convertedNote = entry.convertedAt ? `<span class="converted-note">Converted ${escapeHtml(entry.convertedAt)}</span>` : "";
      return `
        <tr data-row="${entry.rowNumber}">
          <td data-label="Client"><span class="cell-value">${escapeHtml(entry.clientName)}</span></td>
          <td data-label="Garment"><span class="cell-value">${escapeHtml(entry.garmentDescription)}</span></td>
          <td data-label="Arrival Date"><span class="cell-value">${escapeHtml(entry.arrivalDate || "—")}</span></td>
          <td data-label="Salesperson"><span class="cell-value">${escapeHtml(entry.salesperson || "—")}</span></td>
          <td data-label="Status"><span class="cell-value"><select class="status-select studio-status-select" data-row="${entry.rowNumber}" data-status="${escapeHtml(entry.status)}">${statusOptions}</select></span></td>
          <td data-label="Status Date" class="status-date-cell"><span class="cell-value">${escapeHtml(entry.statusDate || "—")}</span></td>
          <td data-label="Actions">
            <span class="cell-value ticket-actions">
              <button type="button" class="icon-btn convert-btn" data-row="${entry.rowNumber}" title="Convert to Alteration Ticket" aria-label="Convert to Alteration Ticket" ${entry.convertedAt ? "disabled" : ""}>${ICONS.convert}</button>
              <button type="button" class="icon-btn icon-btn-danger studio-delete-btn" data-row="${entry.rowNumber}" title="Delete" aria-label="Delete">${ICONS.delete}</button>
              ${convertedNote}
            </span>
          </td>
        </tr>`;
    })
    .join("");
}

async function refreshStudioLog() {
  if (!auth.hasValidSession()) {
    studioLogStatus.textContent = "Sign in to load the studio log.";
    return;
  }
  studioLogStatus.textContent = "Loading…";
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    await getOrCreateStudioTab(token, ledgerId);
    studioCache = await listStudioEntries(token, ledgerId);
    const inStudioCount = studioCache.filter((e) => e.status !== "Complete").length;
    studioStats.innerHTML = `<strong>${studioCache.length}</strong> total · <span class="stat-amber">${inStudioCount} active</span>`;
    renderStudioLog();
  } catch (err) {
    studioLogStatus.textContent = `Could not load studio log: ${err.message}`;
  }
}

studioRefreshBtn.addEventListener("click", refreshStudioLog);
studioSearch.addEventListener("input", renderStudioLog);
studioClearFiltersBtn.addEventListener("click", () => {
  studioSearch.value = "";
  studioStatusFilter.reset();
  studioSalespersonFilter.reset();
  renderStudioLog();
});

studioLogBody.addEventListener("change", async (event) => {
  const select = event.target.closest(".studio-status-select");
  if (!select) return;
  const rowNumber = Number(select.dataset.row);
  const entry = studioCache.find((item) => item.rowNumber === rowNumber);
  if (!entry) return;
  const previousStatus = entry.status;
  entry.status = select.value;
  select.dataset.status = select.value;
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    const now = await updateStudioStatus(token, ledgerId, rowNumber, select.value);
    entry.statusDate = now;
    const dateCell = select.closest("tr")?.querySelector(".status-date-cell .cell-value");
    if (dateCell) dateCell.textContent = now;
  } catch (err) {
    entry.status = previousStatus;
    select.value = previousStatus;
    select.dataset.status = previousStatus;
    studioLogStatus.textContent = `Could not update status: ${err.message}`;
  }
});

studioLogBody.addEventListener("click", async (event) => {
  const convertBtn = event.target.closest(".convert-btn");
  if (convertBtn) {
    const rowNumber = Number(convertBtn.dataset.row);
    const entry = studioCache.find((item) => item.rowNumber === rowNumber);
    if (!entry) return;

    clearAllFields();
    customerNameInput.value = entry.clientName;
    if (entry.salesperson) salespersonInput.value = entry.salesperson;
    renderOutput();
    saveToStorage();
    setView("newTicket");
    saveStatus.textContent = `Started from studio item for ${entry.clientName}. Fill in Tailor, Due Date, and garments to save.`;

    convertBtn.disabled = true;
    try {
      const token = await auth.getValidToken();
      const ledgerId = await getOrCreateLedger(token);
      await markStudioConverted(token, ledgerId, rowNumber);
    } catch (err) {
      studioLogStatus.textContent = `Ticket started, but couldn't mark the studio item converted: ${err.message}`;
    }
    return;
  }

  const deleteBtn = event.target.closest(".studio-delete-btn");
  if (deleteBtn) {
    const rowNumber = Number(deleteBtn.dataset.row);
    const entry = studioCache.find((item) => item.rowNumber === rowNumber);
    if (!entry) return;
    const confirmed = window.confirm(`Remove ${entry.clientName || "this item"} from the studio log? This can't be undone.`);
    if (!confirmed) return;

    deleteBtn.disabled = true;
    try {
      const token = await auth.getValidToken();
      const ledgerId = await getOrCreateLedger(token);
      await deleteStudioEntry(token, ledgerId, rowNumber);
      studioLogStatus.textContent = `Removed ${entry.clientName || "item"} from the studio log.`;
      await refreshStudioLog();
    } catch (err) {
      deleteBtn.disabled = false;
      studioLogStatus.textContent = `Could not delete: ${err.message}`;
    }
  }
});
driveSaveBtn.addEventListener("click", async () => {
  try {
    const missing = validateBeforePrint();
    if (missing.length) {
      alert(`Please complete these required fields before saving:\n- ${missing.join("\n- ")}`);
      return;
    }

    saveStatus.textContent = "Preparing…";
    renderOutput();
    const ticketData = buildTicketData();
    const formState = buildState();

    const safeName = formatFileBaseName(ticketData.customerName).slice(0, 40);
    const datePart = new Date(ticketData.createdAt).toISOString().slice(0, 10);
    const filename = `JM_ALT_${safeName}_${datePart}_${formatFileTime(new Date(ticketData.createdAt))}.docx`;

    const token = await auth.getValidToken();
    const docxBlob = await buildDocxBlob(ticketData);
    const mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    if (editingTicket && editingTicket.driveFileId) {
      // Overwrite the same file and the same log row — no folder picker needed.
      saveStatus.textContent = "Updating file in Drive…";
      const updated = await updateFileInDrive(token, editingTicket.driveFileId, { filename, blob: docxBlob, mimeType });

      saveStatus.textContent = "Updating ticket log…";
      const ledgerId = await getOrCreateLedger(token);
      await updateTicketRecord(token, ledgerId, editingTicket.rowNumber, {
        customerName: ticketData.customerName,
        tailor: ticketData.tailor,
        salesperson: ticketData.salesperson,
        dueDate: ticketData.dueDate,
        rush: ticketData.rush,
        balance: ticketData.balanceDisplay,
        docxFilename: filename,
        driveLink: updated.webViewLink || editingTicket.driveLink || `https://drive.google.com/file/d/${editingTicket.driveFileId}/view`,
        formState,
        driveFileId: editingTicket.driveFileId,
      });

      saveToStorage();
      saveStatus.textContent = `Updated ${filename}`;
      setView("ticketLog");
      return;
    }

    // Either a brand-new ticket, or a reopened ticket that predates file
    // tracking (no driveFileId) — ask where to save, then either append a
    // new row or, if we know the row, upgrade it in place.
    saveStatus.textContent = "Choose a Google Drive folder…";
    const selectedFolder = await chooseDriveFolder(token, driveFolderDom);
    if (!selectedFolder) {
      saveStatus.textContent = "Google save canceled";
      return;
    }

    saveStatus.textContent = "Uploading to Drive…";
    const uploaded = await uploadFileToDrive(token, { filename, blob: docxBlob, mimeType, folderId: selectedFolder.id });
    const driveLink = uploaded.webViewLink || `https://drive.google.com/file/d/${uploaded.id}/view`;

    saveStatus.textContent = "Logging ticket…";
    const ledgerId = await getOrCreateLedger(token);

    if (editingTicket) {
      // Known row, but no prior file reference — upgrade the existing row
      // rather than creating a duplicate.
      await updateTicketRecord(token, ledgerId, editingTicket.rowNumber, {
        customerName: ticketData.customerName,
        tailor: ticketData.tailor,
        salesperson: ticketData.salesperson,
        dueDate: ticketData.dueDate,
        rush: ticketData.rush,
        balance: ticketData.balanceDisplay,
        docxFilename: filename,
        driveLink,
        formState,
        driveFileId: uploaded.id,
      });
      setEditingTicket({ ...editingTicket, driveFileId: uploaded.id });
    } else {
      const created = await appendTicketRow(token, ledgerId, {
        id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        createdAt: ticketData.createdDisplay,
        customerName: ticketData.customerName,
        tailor: ticketData.tailor,
        salesperson: ticketData.salesperson,
        dueDate: ticketData.dueDate,
        rush: ticketData.rush,
        balance: ticketData.balanceDisplay,
        status: "Open",
        docxFilename: filename,
        driveLink,
        formState,
        driveFileId: uploaded.id,
      });
      if (created.rowNumber) {
        setEditingTicket({ id: null, rowNumber: created.rowNumber, driveFileId: uploaded.id, customerName: ticketData.customerName });
      }
    }

    saveToStorage();
    saveStatus.textContent = `Saved to ${selectedFolder.name}: ${filename}`;
    setView("ticketLog");
  } catch (err) {
    saveStatus.textContent = `Drive error: ${err.message}`;
    alert(`Save failed: ${err.message}`);
  }
});

clearBtn.addEventListener("click", clearAllFields);

/* ---------------------------------- wiring ----------------------------------- */
addGarmentBtn.addEventListener("click", () => {
  garmentItems.push(createEmptyItem("jacket"));
  renderGarmentItems();
  onInputChange();
});

garmentItemsEl.addEventListener("input", handleDynamicInput);
garmentItemsEl.addEventListener("change", handleDynamicInput);
garmentItemsEl.addEventListener("click", handleDynamicClick);

tailorInput.addEventListener("change", () => { onInputChange(); });

[customerNameInput, tailorInput, salespersonInput, ...document.querySelectorAll('input[name="balanceDue"]')]
  .forEach((elm) => elm.addEventListener("input", onInputChange));

if (!dueDateHiddenInput) dueDateInput.addEventListener("input", onInputChange);

function isIOSDevice() {
  const ua = navigator.userAgent || "";
  const isIOS = /iPad|iPhone|iPod/.test(ua);
  const isIpadOS = /Macintosh/.test(ua) && navigator.maxTouchPoints > 1;
  return isIOS || isIpadOS;
}

function setupIOSDateFallback() {
  if (!isIOSDevice()) return;
  const original = dueDateInput;
  const wrapper = document.createElement("div");
  wrapper.className = "ios-date-wrapper";
  const display = document.createElement("input");
  display.type = "text";
  display.id = "dueDateDisplay";
  display.placeholder = "MM/DD/YYYY";
  display.readOnly = true;
  display.inputMode = "none";
  display.autocomplete = "off";
  display.setAttribute("aria-label", "Due Date");
  original.id = "dueDateHidden";
  original.name = "dueDateHidden";
  Object.assign(original.style, { position: "absolute", opacity: "0", pointerEvents: "none", height: "0", width: "0", border: "0", padding: "0", margin: "0" });
  original.setAttribute("aria-hidden", "true");
  original.parentNode.insertBefore(wrapper, original);
  wrapper.appendChild(display);
  wrapper.appendChild(original);
  dueDateHiddenInput = original;
  dueDateDisplayInput = display;
  const openPicker = () => { dueDateHiddenInput.focus(); dueDateHiddenInput.click(); };
  display.addEventListener("click", openPicker);
  display.addEventListener("focus", openPicker);
  dueDateHiddenInput.addEventListener("change", () => { updateDueDateDisplay(); onInputChange(); });
  updateDueDateDisplay();
}

setupIOSDateFallback();
loadFromStorage();
updateAuthGate();
