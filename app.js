import { TICKET_STATUSES, STUDIO_STATUSES, SALESPEOPLE, JACKET_SIZE_SUGGESTIONS, TROUSER_SIZE_SUGGESTIONS, SHIRT_SIZE_SUGGESTIONS } from "./config.js";
import * as auth from "./google-auth.js";
import { chooseDriveFolder, uploadFileToDrive, updateFileInDrive, trashDriveFile } from "./drive.js";
import {
  getOrCreateLedger, appendTicketRow, updateTicketRecord, listTickets, updateTicketStatus, updateTicketNotes, deleteTicketRow, getNextTicketNumber,
  getOrCreateStudioTab, appendStudioEntry, listStudioEntries, updateStudioStatus, updateStudioEntry, markStudioConverted, deleteStudioEntry,
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
  duplicate: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="7" y="7" width="9.5" height="9.5" rx="1.3"/><path d="M13 7V4.8a1.3 1.3 0 0 0-1.3-1.3H4.3A1.3 1.3 0 0 0 3 4.8v7.4A1.3 1.3 0 0 0 4.3 13.5H7"/></svg>',
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
const addJacketBtn = el("addJacketBtn");
const addTrouserBtn = el("addTrouserBtn");
const addShirtBtn = el("addShirtBtn");

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

const studioEditModal = el("studioEditModal");
const studioEditClientName = el("studioEditClientName");
const studioEditGarmentDescription = el("studioEditGarmentDescription");
const studioEditArrivalDate = el("studioEditArrivalDate");
const studioEditSalesperson = el("studioEditSalesperson");
const studioEditCancelBtn = el("studioEditCancelBtn");
const studioEditSaveBtn = el("studioEditSaveBtn");

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
const salespersonFilterEls = {
  wrap: el("salespersonFilterWrap"), btn: el("salespersonFilterBtn"), panel: el("salespersonFilterPanel"),
  options: el("salespersonFilterOptions"), allBtn: el("salespersonFilterAll"), noneBtn: el("salespersonFilterNone"),
};
const ticketRushFilter = el("ticketRushFilter");
const showCompletedFilter = el("showCompletedFilter");
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
let convertingStudioRowNumber = null; // studio row a new ticket originated from, marked converted only once actually saved

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
    halfBack: 0, halfWaist: 0, shortenBody: 0, sleeves: 0, sleeveWidth: 0, tightenCollar: 0, buttons: "", buttonStyle: "",
    trouserWaist: 0, trouserSeat: 0, trouserThigh: 0, trouserKnee: 0, trouserLegOpening: 0, trouserInseam: 0,
    trouserTotalLength: "", trouserCuff: "",
    shirtSleeve: 0, shirtBody: 0, shirtSlimBody: 0,
  };
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
    (item?.buttons || "").trim() || (item?.buttonStyle || "").trim() || (item?.trouserTotalLength || "").trim() || (item?.trouserCuff || "").trim() ||
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
        <input class="stepper-value-input" type="text" inputmode="decimal" value="${value}" data-action="clear-value" data-uid="${uid}" data-field="${field}" />
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
      <div class="measurement-row">
        <label id="buttonStyleLabel-${uid}">Button Style</label>
        <div class="stepper">
          <div class="segmented-toggle" role="group" aria-labelledby="buttonStyleLabel-${uid}">
            <button type="button" class="segmented-toggle-btn${item.buttonStyle === "Working" ? " is-active" : ""}" data-action="set-button-style" data-uid="${uid}" data-button-style="Working">Working</button>
            <button type="button" class="segmented-toggle-btn${item.buttonStyle === "Tacked" ? " is-active" : ""}" data-action="set-button-style" data-uid="${uid}" data-button-style="Tacked">Tacked</button>
          </div>
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
      ${stepperRow(uid, "trouserThigh", "1/2 Thigh", item)}
      ${stepperRow(uid, "trouserKnee", "1/2 Knee", item)}
      ${stepperRow(uid, "trouserLegOpening", "1/2 Leg Opening", item)}
      ${stepperRow(uid, "trouserInseam", "Inseam", item)}
      <div class="measurement-row">
        <label for="totalLength-${uid}">Total Inseam</label>
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
              { value: "4 cm Cuff", label: "4 cm Cuff" },
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
const SIZE_SUGGESTION_SEEDS = { jacket: JACKET_SIZE_SUGGESTIONS, trouser: TROUSER_SIZE_SUGGESTIONS, shirt: SHIRT_SIZE_SUGGESTIONS };
const SIZE_DATALIST_IDS = { jacket: "jacketSizeSuggestions", trouser: "trouserSizeSuggestions", shirt: "shirtSizeSuggestions" };

// Blends the hardcoded seed list with real sizeDescription values already
// typed across saved tickets (once the Ticket Log has loaded this session)
// — starts useful on day one, gets more tailored to this shop over time.
function populateSizeDatalists() {
  const historicalByType = { jacket: new Set(), trouser: new Set(), shirt: new Set() };
  ticketCache.forEach((ticket) => {
    (ticket.formState?.garmentItems || []).forEach((item) => {
      const text = (item.sizeDescription || "").trim();
      if (text && historicalByType[item.garmentType]) historicalByType[item.garmentType].add(text);
    });
  });

  Object.entries(SIZE_DATALIST_IDS).forEach(([type, datalistId]) => {
    const datalistEl = document.getElementById(datalistId);
    if (!datalistEl) return;
    const combined = [...SIZE_SUGGESTION_SEEDS[type], ...historicalByType[type]];
    const deduped = [...new Set(combined)];
    datalistEl.innerHTML = deduped.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  });
}

function renderGarmentTypeBadge(item) {
  return `<div class="garment-picker garment-picker-locked"><span class="garment-picker-label">${GARMENT_TYPE_LABELS[item.garmentType] || "Garment"}</span></div>`;
}

function summarizeSingleItem(item) {
  const text = (item.sizeDescription || "").trim();
  return text || "No details yet";
}

/* ------------------------------ item list UI ------------------------------ */
function renderGarmentItems() {
  if (!garmentItems.length) {
    garmentItemsEl.innerHTML = `<p class="garment-empty-state">No garments added yet — use the buttons above to add one.</p>`;
    return;
  }

  garmentItemsEl.innerHTML = garmentItems
    .map((item) => {
      const isCollapsed = collapsedItemIds.has(String(item._uid));
      const typeLabel = GARMENT_TYPE_LABELS[item.garmentType] || "Garment";

      if (isCollapsed) {
        return `
          <section class="repeat-item is-collapsed" data-uid="${item._uid}">
            <button type="button" class="repeat-item-summary" data-action="toggle-collapse" data-uid="${item._uid}">
              <span class="repeat-item-summary-type">${typeLabel}</span>
              <span class="repeat-item-summary-detail">${escapeHtml(summarizeSingleItem(item))}</span>
              <span class="repeat-item-summary-chevron" aria-hidden="true">▸</span>
            </button>
            <button type="button" class="remove-item-btn" data-action="remove" data-uid="${item._uid}">Remove</button>
          </section>`;
      }

      const measurementBlock = item.garmentType === "jacket" ? renderJacketMeasurements(item)
        : item.garmentType === "trouser" ? renderTrouserMeasurements(item)
        : renderShirtMeasurements(item);

      return `
        <section class="repeat-item" data-uid="${item._uid}">
          <div class="repeat-item-header">
            ${renderGarmentTypeBadge(item)}
            <div class="repeat-item-header-actions">
              <button type="button" class="link-action-btn" data-action="toggle-collapse" data-uid="${item._uid}">Collapse</button>
              <button type="button" class="remove-item-btn" data-action="remove" data-uid="${item._uid}">Remove</button>
            </div>
          </div>
          <div class="garment-field size-desc-row">
            <label for="sizeDescription-${item._uid}" class="size-desc-label">Garment Size &amp; Description</label>
            <input id="sizeDescription-${item._uid}" type="text" placeholder="${SIZE_DESCRIPTION_PLACEHOLDERS[item.garmentType] || "e.g. Size and description"}" list="${SIZE_DATALIST_IDS[item.garmentType] || ""}" data-uid="${item._uid}" data-field="sizeDescription" value="${escapeHtml(item.sizeDescription)}" />
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
    if (event.type !== "change") return; // let typing happen freely; only commit once the field is done being edited (blur)
    const rawValue = target.value.trim();
    if (!rawValue) {
      item[field] = 0;
    } else {
      const parsed = parseFloat(rawValue.replace(/[^\d.+-]/g, ""));
      if (Number.isFinite(parsed)) {
        const rounded = Math.round(parsed * 2) / 2; // snap to the same 0.5 increments the +/- buttons use
        item[field] = isNegativeOnlyField(field) ? normalizeNegativeOnlyValue(rounded) : rounded;
      }
      // if it doesn't parse as a number at all, fall through and just
      // re-render — this naturally resets the field back to its last
      // valid value instead of silently accepting garbage input.
    }
    renderGarmentItems();
    onInputChange();
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

  if (action === "set-button-style") {
    const item = garmentItems.find((it) => String(it._uid) === uid);
    if (!item) return;
    const value = target.dataset.buttonStyle;
    item.buttonStyle = item.buttonStyle === value ? "" : value; // click active option again to clear
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
// Verb pairs per measurement, used to phrase the printed/on-screen output
// as a sentence (e.g. "Let out 1/2 Back + 2 cm") instead of a bare
// "Label: value" line. Fields with only a "negative" entry are the
// minus-only fields — their phrasing doesn't change with direction.
const JACKET_ADJUSTMENT_VERBS = {
  halfBack: { positive: "Let out 1/2 Back", negative: "Take in 1/2 Back" },
  halfWaist: { positive: "Let out 1/2 Waist", negative: "Take in 1/2 Waist" },
  shortenBody: { positive: "Lengthen Body", negative: "Shorten Body" },
  sleeves: { positive: "Lengthen Sleeve", negative: "Shorten Sleeve" },
  // Not explicitly specified — inferred to match the Back/Waist "let
  // out / take in" circumference pattern. Flag if this should read
  // differently.
  sleeveWidth: { positive: "Let out Sleeve Width", negative: "Take in Sleeve Width" },
  tightenCollar: { negative: "Tighten Collar" },
};

const TROUSER_ADJUSTMENT_VERBS = {
  trouserWaist: { positive: "Let out Waist", negative: "Take in Waist" },
  trouserSeat: { positive: "Let out Seat", negative: "Take in Seat" },
  trouserThigh: { positive: "Let out 1/2 Thigh", negative: "Take in 1/2 Thigh" },
  trouserKnee: { positive: "Let out 1/2 Knee", negative: "Take in 1/2 Knee" },
  trouserLegOpening: { positive: "Let out 1/2 Leg Opening", negative: "Take in 1/2 Leg Opening" },
  trouserInseam: { positive: "Lengthen Inseam", negative: "Shorten Inseam" },
};

const SHIRT_ADJUSTMENT_VERBS = {
  shirtSleeve: { negative: "Shorten Sleeve" },
  shirtBody: { negative: "Shorten Body" },
  shirtSlimBody: { negative: "Take in Body" },
};

// Turns a raw +/- value into a full sentence, e.g. "Let out 1/2 Back + 2
// cm" — keeps the +/- notation (per request) rather than dropping it now
// that the verb already implies direction.
function describeAdjustment(verbs, value) {
  const numeric = Number(value) || 0;
  if (!numeric) return null;
  const phrase = numeric > 0 ? verbs.positive || verbs.negative : verbs.negative || verbs.positive;
  if (!phrase) return null;
  return `${phrase} ${formatSignedQuarter(numeric)} cm`;
}

function buildGarmentSections(items = garmentItems) {
  const sections = [];
  const filled = items.filter(hasGarmentData);

  const buildJacketSection = (entry, label) => {
    const adjustments = [];
    const addAdjustment = (verbs, value) => {
      const line = describeAdjustment(verbs, value);
      if (line) adjustments.push(line);
    };
    addAdjustment(JACKET_ADJUSTMENT_VERBS.halfBack, entry.halfBack);
    addAdjustment(JACKET_ADJUSTMENT_VERBS.halfWaist, entry.halfWaist);
    addAdjustment(JACKET_ADJUSTMENT_VERBS.shortenBody, entry.shortenBody);
    addAdjustment(JACKET_ADJUSTMENT_VERBS.sleeves, entry.sleeves);
    addAdjustment(JACKET_ADJUSTMENT_VERBS.sleeveWidth, entry.sleeveWidth);
    addAdjustment(JACKET_ADJUSTMENT_VERBS.tightenCollar, normalizeNegativeOnlyValue(entry.tightenCollar));

    const attributes = [];
    const buttonParts = [(entry.buttons || "").trim(), (entry.buttonStyle || "").trim()].filter(Boolean);
    if (buttonParts.length) attributes.push({ label: "Sleeve Buttons", value: buttonParts.join(", ") });

    return { label, adjustments, attributes, notes: (entry.adjustments || "").trim() };
  };

  const buildTrouserSection = (entry, label) => {
    const adjustments = [];
    const addAdjustment = (verbs, value) => {
      const line = describeAdjustment(verbs, value);
      if (line) adjustments.push(line);
    };
    addAdjustment(TROUSER_ADJUSTMENT_VERBS.trouserWaist, entry.trouserWaist);
    addAdjustment(TROUSER_ADJUSTMENT_VERBS.trouserSeat, entry.trouserSeat);
    addAdjustment(TROUSER_ADJUSTMENT_VERBS.trouserThigh, entry.trouserThigh);
    addAdjustment(TROUSER_ADJUSTMENT_VERBS.trouserKnee, entry.trouserKnee);
    addAdjustment(TROUSER_ADJUSTMENT_VERBS.trouserLegOpening, entry.trouserLegOpening);
    addAdjustment(TROUSER_ADJUSTMENT_VERBS.trouserInseam, entry.trouserInseam);

    const attributes = [];
    if ((entry.trouserTotalLength || "").trim()) attributes.push({ label: "Total Inseam", value: `${entry.trouserTotalLength} cm` });
    if ((entry.trouserCuff || "").trim()) {
      const cuffLabel = normalizeTrouserCuff(entry.trouserCuff).replace(/\s*Cuff\b/i, "").trim();
      attributes.push({ label: "Cuff Style", value: cuffLabel });
    }

    return { label, adjustments, attributes, notes: (entry.adjustments || "").trim() };
  };

  const buildShirtSection = (entry, label) => {
    const adjustments = [];
    const addAdjustment = (verbs, value) => {
      const line = describeAdjustment(verbs, value);
      if (line) adjustments.push(line);
    };
    addAdjustment(SHIRT_ADJUSTMENT_VERBS.shirtSleeve, normalizeNegativeOnlyValue(entry.shirtSleeve));
    addAdjustment(SHIRT_ADJUSTMENT_VERBS.shirtBody, normalizeNegativeOnlyValue(entry.shirtBody));
    addAdjustment(SHIRT_ADJUSTMENT_VERBS.shirtSlimBody, normalizeNegativeOnlyValue(entry.shirtSlimBody));

    return { label, adjustments, attributes: [], notes: (entry.adjustments || "").trim() };
  };

  // Grouped by type on the printed ticket (all jackets together, then
  // trousers, then shirts) regardless of the order items were added in —
  // reads better on the physical document than the entry order would.
  // Section titles are "Type - Description" (e.g. "Jacket - 40 Brown
  // Tweed") rather than "Jacket 2" — more useful for telling multiple
  // items of the same type apart than a bare index number.
  ["jacket", "trouser", "shirt"].forEach((type) => {
    const typeLabel = GARMENT_TYPE_LABELS[type];
    const builder = type === "jacket" ? buildJacketSection : type === "trouser" ? buildTrouserSection : buildShirtSection;
    filled
      .filter((entry) => entry.garmentType === type)
      .forEach((entry) => {
        const description = (entry.sizeDescription || "").trim();
        const label = description ? `${typeLabel} - ${description}` : typeLabel;
        sections.push(builder(entry, label));
      });
  });

  return sections;
}

function renderTicketMarkup(ticketData) {
  const sectionHtml = ticketData.garmentSections.map((section) => {
    const adjustmentHtml = section.adjustments.map((line) => `<p class="adjustment-line">${escapeHtml(line)}</p>`).join("");
    const attributeHtml = section.attributes.map((a) => `<p class="adjustment-line">${escapeHtml(a.label)}: ${escapeHtml(a.value)}</p>`).join("");
    const notesHtml = section.notes ? `<p><strong>Notes:</strong><br>${formatMultiline(section.notes)}</p>` : "";
    return `<section class="garment-output-block"><p><strong>${escapeHtml(section.label)}</strong></p>${adjustmentHtml}${attributeHtml}${notesHtml}</section>`;
  }).join("");

  return `
    <p class="doc-title">J. Mueser</p>
    <p class="doc-subtitle">Alterations Ticket</p>
    <div class="ticket-hero">
      ${ticketData.rush ? '<p class="rush-flag rush-flag-hero">★ Rush</p>' : ""}
      <p class="hero-line"><span class="hero-label">Due</span>${escapeHtml(ticketData.dueDate)}</p>
      <p class="hero-client-name">${escapeHtml(ticketData.customerName)}</p>
      <p class="hero-line"><span class="hero-label">Tailor</span>${escapeHtml(ticketData.tailor)}</p>
      <p class="hero-line"><span class="hero-label">Salesperson</span>${escapeHtml(ticketData.salesperson)}</p>
    </div>
    ${sectionHtml}
    <div class="output-footer">
      ${ticketData.balanceDisplay ? `<p class="ticket-detail"><strong>${escapeHtml(ticketData.balanceDisplay)}</strong></p>` : ""}
      <p class="meta">Created ${escapeHtml(ticketData.createdDisplay)}</p>
    </div>
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
  return items;
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
    garmentItems = [];
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
    garmentItems = [];
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
  convertingStudioRowNumber = null;
  customerNameInput.value = "";
  tailorInput.value = "";
  salespersonInput.value = "";
  setDueDateValue("");
  const selectedBalance = document.querySelector('input[name="balanceDue"]:checked');
  if (selectedBalance) selectedBalance.checked = false;

  garmentItems = [];
  collapsedItemIds = new Set();
  renderGarmentItems();
  localStorage.removeItem(STORAGE_KEY);
  saveStatus.textContent = "Form cleared";
  renderOutput();
}

/* ----------------------------------- auth ---------------------------------- */
async function updateAuthGate() {
  let isSignedIn = auth.hasValidSession();

  // The cached token may have simply expired (not a true sign-out) — try a
  // silent background refresh before showing the blocking sign-in screen.
  // This is the moment someone's most likely to interpret as "getting
  // logged out," so it matters more here than anywhere else.
  if (!isSignedIn && auth.getSessionEmail()) {
    authStatus.textContent = "Checking your Google session…";
    try {
      await auth.signIn("");
      isSignedIn = auth.hasValidSession();
    } catch {
      // couldn't refresh silently (fully signed out of Google, access
      // revoked, or the browser is blocking it) — the gate below handles it
    }
  }

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

const statusFilter = createMultiSelectFilter(statusFilterEls, TICKET_STATUSES.filter((s) => s !== "Completed"), "Status", () => renderTicketLog());
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

let ticketSort = { column: null, direction: "asc" }; // null column = default: newest added first

function compareDates(aStr, bStr) {
  const aTime = aStr ? new Date(aStr).getTime() : NaN;
  const bTime = bStr ? new Date(bStr).getTime() : NaN;
  const aValid = !Number.isNaN(aTime);
  const bValid = !Number.isNaN(bTime);
  if (!aValid && !bValid) return 0;
  if (!aValid) return 1; // blank/unparseable dates sort last
  if (!bValid) return -1;
  return aTime - bTime;
}

function compareTickets(a, b, column) {
  if (column === "client") return (a.customerName || "").localeCompare(b.customerName || "");
  if (column === "dueDate") return compareDates(a.dueDate, b.dueDate);
  if (column === "statusDate") return compareDates(a.statusDate, b.statusDate);
  if (column === "ticketNumber") {
    const aNum = parseInt(a.ticketNumber, 10);
    const bNum = parseInt(b.ticketNumber, 10);
    const aValid = Number.isFinite(aNum);
    const bValid = Number.isFinite(bNum);
    if (!aValid && !bValid) return 0;
    if (!aValid) return 1; // unnumbered (pre-feature) tickets sort last
    if (!bValid) return -1;
    return aNum - bNum;
  }
  return 0;
}

const SORT_INDICATOR_IDS = {
  ticketNumber: "sortIndicatorTicketNumber", client: "sortIndicatorClient", dueDate: "sortIndicatorDueDate", statusDate: "sortIndicatorStatusDate",
};

function updateSortIndicators() {
  Object.entries(SORT_INDICATOR_IDS).forEach(([column, id]) => {
    const indicatorEl = document.getElementById(id);
    if (!indicatorEl) return;
    indicatorEl.textContent = ticketSort.column === column ? (ticketSort.direction === "asc" ? " ▲" : " ▼") : "";
  });
}

function handleSortClick(column) {
  // Click cycle: ascending -> descending -> back to default (newest added
  // first), so there's always a clear way back to "how it normally looks."
  if (ticketSort.column !== column) {
    ticketSort = { column, direction: "asc" };
  } else if (ticketSort.direction === "asc") {
    ticketSort = { column, direction: "desc" };
  } else {
    ticketSort = { column: null, direction: "asc" };
  }
  updateSortIndicators();
  renderTicketLog();
}

el("sortByClientBtn").addEventListener("click", () => handleSortClick("client"));
el("sortByDueDateBtn").addEventListener("click", () => handleSortClick("dueDate"));
el("sortByStatusDateBtn").addEventListener("click", () => handleSortClick("statusDate"));
el("sortByTicketNumberBtn").addEventListener("click", () => handleSortClick("ticketNumber"));

function updateTicketLogStats() {
  const activeTickets = ticketCache.filter((t) => t.status !== "Completed");
  const openCount = activeTickets.filter((t) => t.status === "Open" || t.status === "In Progress").length;
  const rushCount = activeTickets.filter((t) => t.rush).length;
  logStats.innerHTML = `<strong>${activeTickets.length}</strong> tickets · <span class="stat-amber">${openCount} open</span> · <span class="stat-red">${rushCount} rush</span>`;
}

function renderTicketLog() {
  const search = ticketSearch.value.trim().toLowerCase();
  const rushOnly = ticketRushFilter.checked;
  const showCompleted = showCompletedFilter.checked;

  const filtered = ticketCache.filter((ticket) => {
    if (!showCompleted && ticket.status === "Completed") return false;
    if (search && !ticket.customerName.toLowerCase().includes(search) && !(ticket.salesperson || "").toLowerCase().includes(search)) return false;
    if (!statusFilter.matches(ticket.status)) return false;
    if (!salespersonFilter.matches(ticket.salesperson)) return false;
    if (rushOnly && !ticket.rush) return false;
    return true;
  });

  ticketLogStatus.textContent =
    filtered.length === ticketCache.length
      ? `${ticketCache.length} ticket${ticketCache.length === 1 ? "" : "s"} logged.`
      : `Showing ${filtered.length} ticket${filtered.length === 1 ? "" : "s"} out of ${ticketCache.length} ticket${ticketCache.length === 1 ? "" : "s"} logged.`;

  if (!filtered.length) {
    ticketLogBody.innerHTML = `<tr><td colspan="11" class="log-status">No tickets match yet.</td></tr>`;
    return;
  }

  const orderedTickets = ticketSort.column
    ? filtered.slice().sort((a, b) => {
        const result = compareTickets(a, b, ticketSort.column);
        return ticketSort.direction === "desc" ? -result : result;
      })
    : filtered.slice().reverse(); // default: newest added first

  ticketLogBody.innerHTML = orderedTickets
    .map((ticket) => {
      const statusOptions = TICKET_STATUSES.map((status) => `<option value="${status}" ${status === ticket.status ? "selected" : ""}>${status}</option>`).join("");
      return `
        <tr data-row="${ticket.rowNumber}" data-ticket-id="${escapeHtml(ticket.id)}">
          <td data-label="#" class="ticket-number-cell"><span class="cell-value">${ticket.ticketNumber ? escapeHtml(String(ticket.ticketNumber)) : "—"}</span></td>
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
              <button type="button" class="icon-btn duplicate-btn" data-row="${ticket.rowNumber}" title="Duplicate" aria-label="Duplicate">${ICONS.duplicate}</button>
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
    updateTicketLogStats();
    populateSizeDatalists();

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
showCompletedFilter.addEventListener("change", renderTicketLog);

clearFiltersBtn.addEventListener("click", () => {
  ticketSearch.value = "";
  ticketRushFilter.checked = false;
  showCompletedFilter.checked = false;
  statusFilter.reset();
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
    updateTicketLogStats();
    renderTicketLog(); // reapply active filters — e.g. drop out of view if now Completed and hidden
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

  const duplicateBtn = event.target.closest(".duplicate-btn");
  if (duplicateBtn) {
    const rowNumber = Number(duplicateBtn.dataset.row);
    const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
    if (!ticket || !ticket.formState || !Object.keys(ticket.formState).length) {
      ticketLogStatus.textContent = "This older ticket doesn't have duplicable form data.";
      return;
    }
    setEditingTicket(null); // a duplicate is always a brand-new ticket on save, never an update
    convertingStudioRowNumber = null;
    applyState(ticket.formState);
    setDueDateValue(""); // a new job needs its own timeline, not the original's
    renderOutput();
    setView("newTicket");
    saveToStorage();
    saveStatus.textContent = `Duplicated ${ticket.customerName || "ticket"}'s measurements — confirm the Due Date, then Save Ticket.`;
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
    if (search && !entry.clientName.toLowerCase().includes(search) && !(entry.salesperson || "").toLowerCase().includes(search)) return false;
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
              <button type="button" class="icon-btn studio-edit-btn" data-row="${entry.rowNumber}" title="Edit" aria-label="Edit">${ICONS.edit}</button>
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
    const inStudioCount = studioCache.filter((e) => e.status !== "Completed").length;
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

/* ---------------------------- studio edit modal -------------------------------- */
let studioEditRowNumber = null;

// Arrival Date is stored as a locale display string (e.g. "8/16/2026"), not
// ISO — convert it to the yyyy-mm-dd shape a date input needs.
function localeDateToInputValue(localeDateStr) {
  if (!localeDateStr) return "";
  const parsed = new Date(localeDateStr);
  if (Number.isNaN(parsed.getTime())) return "";
  const yyyy = parsed.getFullYear();
  const mm = String(parsed.getMonth() + 1).padStart(2, "0");
  const dd = String(parsed.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function openStudioEditModal(entry) {
  studioEditRowNumber = entry.rowNumber;
  studioEditClientName.value = entry.clientName || "";
  studioEditGarmentDescription.value = entry.garmentDescription || "";
  studioEditArrivalDate.value = localeDateToInputValue(entry.arrivalDate);
  studioEditSalesperson.value = entry.salesperson || "";
  studioEditModal.hidden = false;
  studioEditClientName.focus();
}

function closeStudioEditModal() {
  studioEditModal.hidden = true;
  studioEditRowNumber = null;
}

studioEditCancelBtn.addEventListener("click", closeStudioEditModal);

studioEditSaveBtn.addEventListener("click", async () => {
  if (studioEditRowNumber === null) return;
  const rowNumber = studioEditRowNumber;
  const entry = studioCache.find((item) => item.rowNumber === rowNumber);
  if (!entry) { closeStudioEditModal(); return; }

  const clientName = studioEditClientName.value.trim();
  const garmentDescription = studioEditGarmentDescription.value.trim();
  if (!clientName || !garmentDescription) {
    alert("Please fill in at least Client Name and Garment Description.");
    return;
  }

  const updated = {
    clientName,
    garmentDescription,
    arrivalDate: studioEditArrivalDate.value ? new Date(`${studioEditArrivalDate.value}T00:00:00`).toLocaleDateString() : "",
    salesperson: studioEditSalesperson.value,
  };

  studioEditSaveBtn.disabled = true;
  studioEditSaveBtn.textContent = "Saving…";
  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    const now = await updateStudioEntry(token, ledgerId, rowNumber, updated);
    Object.assign(entry, updated, { statusDate: now });
    studioLogStatus.textContent = `Updated ${clientName}.`;
    closeStudioEditModal();
    renderStudioLog();
  } catch (err) {
    studioLogStatus.textContent = `Could not save changes: ${err.message}`;
  } finally {
    studioEditSaveBtn.disabled = false;
    studioEditSaveBtn.textContent = "Save Changes";
  }
});

studioLogBody.addEventListener("click", async (event) => {
  const editBtn = event.target.closest(".studio-edit-btn");
  if (editBtn) {
    const rowNumber = Number(editBtn.dataset.row);
    const entry = studioCache.find((item) => item.rowNumber === rowNumber);
    if (!entry) return;
    openStudioEditModal(entry);
    return;
  }

  const convertBtn = event.target.closest(".convert-btn");
  if (convertBtn) {
    const rowNumber = Number(convertBtn.dataset.row);
    const entry = studioCache.find((item) => item.rowNumber === rowNumber);
    if (!entry) return;

    clearAllFields();
    convertingStudioRowNumber = rowNumber; // marked converted only once this ticket is actually saved
    customerNameInput.value = entry.clientName;
    if (entry.salesperson) salespersonInput.value = entry.salesperson;
    renderOutput();
    saveToStorage();
    setView("newTicket");
    saveStatus.textContent = `Started from studio item for ${entry.clientName}. Fill in Tailor, Due Date, and garments, then Save Ticket — this studio item is marked converted once you save.`;
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
      const ticketNumber = await getNextTicketNumber(token, ledgerId);
      const created = await appendTicketRow(token, ledgerId, {
        id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`,
        createdAt: ticketData.createdDisplay,
        customerName: ticketData.customerName,
        tailor: ticketData.tailor,
        salesperson: ticketData.salesperson,
        dueDate: ticketData.dueDate,
        rush: ticketData.rush,
        balance: ticketData.balanceDisplay,
        status: "In Progress",
        docxFilename: filename,
        driveLink,
        formState,
        driveFileId: uploaded.id,
        ticketNumber,
      });
      if (created.rowNumber) {
        setEditingTicket({ id: null, rowNumber: created.rowNumber, driveFileId: uploaded.id, customerName: ticketData.customerName });
      }

      if (convertingStudioRowNumber !== null) {
        const studioRowToMark = convertingStudioRowNumber;
        convertingStudioRowNumber = null;
        try {
          await markStudioConverted(token, ledgerId, studioRowToMark);
        } catch (err) {
          saveStatus.textContent += ` (Saved, but couldn't mark the studio item converted: ${err.message})`;
        }
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
function addGarmentItem(garmentType) {
  // Minimize whatever tile you were just working on, regardless of whether
  // it has anything in it yet, so the newly added tile is the one clear
  // focal point.
  garmentItems.forEach((it) => collapsedItemIds.add(String(it._uid)));

  const item = createEmptyItem(garmentType);
  garmentItems.push(item);
  renderGarmentItems();
  onInputChange();
  const firstField = document.getElementById(`sizeDescription-${item._uid}`);
  if (firstField) {
    firstField.scrollIntoView({ behavior: "smooth", block: "center" });
    firstField.focus();
  }
}

addJacketBtn.addEventListener("click", () => addGarmentItem("jacket"));
addTrouserBtn.addEventListener("click", () => addGarmentItem("trouser"));
addShirtBtn.addEventListener("click", () => addGarmentItem("shirt"));

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
populateSizeDatalists();
