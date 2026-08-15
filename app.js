import { TICKET_STATUSES, SALESPEOPLE } from "./config.js";
import * as auth from "./google-auth.js";
import { chooseDriveFolder, uploadFileToDrive, updateFileInDrive, trashDriveFile } from "./drive.js";
import { getOrCreateLedger, appendTicketRow, updateTicketRecord, listTickets, updateTicketStatus, updateTicketNotes, deleteTicketRow } from "./sheets.js";
import { buildDocxBlob } from "./docx.js";

const STORAGE_KEY = "alterationsTicketStateV2";

const ICONS = {
  print: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5.5 7V3.5h9V7"/><rect x="3.5" y="7" width="13" height="6.5" rx="1"/><path d="M5.5 12.5V16.5h9V12.5"/></svg>',
  edit: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M13.5 3.5l3 3L7 16H4v-3L13.5 3.5z"/></svg>',
  download: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10 3v9.5M6.5 9l3.5 3.5L13.5 9"/><path d="M4 15.5h12"/></svg>',
  delete: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4.5 6h11"/><path d="M8 6V4.3a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1V6"/><path d="M6 6l.6 9a1 1 0 0 0 1 1h4.8a1 1 0 0 0 1-1L14 6"/></svg>',
};

/* ------------------------------- DOM refs -------------------------------- */
const el = (id) => document.getElementById(id);

const customerNameInput = el("customerName");
const tailorInput = el("tailor");
const salespersonInput = el("salesperson");
const dueDateInput = el("dueDate");
let dueDateHiddenInput = null;
let dueDateDisplayInput = null;

const jacketItemsEl = el("jacketItems");
const addJacketBtn = el("addJacketBtn");
const trouserItemsEl = el("trouserItems");
const addTrouserBtn = el("addTrouserBtn");
const shirtItemsEl = el("shirtItems");
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

const garmentTabs = Array.from(document.querySelectorAll(".garment-tab"));
const garmentPanels = Array.from(document.querySelectorAll(".garment-panel"));

const navNewTicket = el("nav-newTicket");
const navTicketLog = el("nav-ticketLog");
const viewNewTicket = el("view-newTicket");
const viewTicketLog = el("view-ticketLog");

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

let jackets = [];
let trousers = [];
let shirts = [];
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

function createEmptyItem() {
  return {
    size: "", description: "", adjustments: "",
    halfBack: 0, halfWaist: 0, shortenBody: 0, sleeves: 0, sleeveWidth: 0, tightenCollar: 0, buttons: "",
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
    (item?.size || "").trim() || (item?.description || "").trim() || (item?.adjustments || "").trim() ||
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

function formatSizeDisplay(value) { return value === "custom" ? "Custom" : value; }

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
function stepperRow(type, idx, field, labelText, item, { disablePlus = false } = {}) {
  const value = formatSignedQuarter(item?.[field] || 0);
  const plusDisabled = disablePlus && normalizeNegativeOnlyValue(item?.[field]) >= 0 ? "disabled" : "";
  return `
    <div class="measurement-row">
      <label for="${type}-${field}-${idx}">${labelText}</label>
      <div class="stepper" id="${type}-${field}-${idx}">
        <button type="button" class="stepper-btn" data-action="step" data-type="${type}" data-index="${idx}" data-field="${field}" data-dir="-1">-</button>
        <input class="stepper-value-input" type="text" inputmode="none" value="${value}" data-action="clear-value" data-type="${type}" data-index="${idx}" data-field="${field}" />
        <button type="button" class="stepper-btn" data-action="step" data-type="${type}" data-index="${idx}" data-field="${field}" data-dir="1" ${plusDisabled}>+</button>
      </div>
    </div>`;
}

function renderJacketMeasurements(item, idx) {
  return `
    <div class="measurement-controls">
      ${stepperRow("jacket", idx, "halfBack", "1/2 Back", item)}
      ${stepperRow("jacket", idx, "halfWaist", "1/2 Waist", item)}
      ${stepperRow("jacket", idx, "shortenBody", "Body Length", item)}
      ${stepperRow("jacket", idx, "sleeves", "Sleeve Length", item)}
      ${stepperRow("jacket", idx, "sleeveWidth", "Sleeve Width", item)}
      ${stepperRow("jacket", idx, "tightenCollar", "Tighten Collar", item, { disablePlus: true })}
      <div class="measurement-row">
        <label for="jacket-buttons-${idx}">Buttons</label>
        <div class="stepper">
          <select id="jacket-buttons-${idx}" class="button-select" data-type="jacket" data-index="${idx}" data-field="buttons">
            ${buildButtonsOptions(item?.buttons || "")}
          </select>
        </div>
      </div>
    </div>`;
}

function renderTrouserMeasurements(item, idx) {
  return `
    <div class="measurement-controls">
      ${stepperRow("trouser", idx, "trouserWaist", "Waist", item)}
      ${stepperRow("trouser", idx, "trouserSeat", "Seat", item)}
      ${stepperRow("trouser", idx, "trouserThigh", "Thigh", item)}
      ${stepperRow("trouser", idx, "trouserKnee", "Knee", item)}
      ${stepperRow("trouser", idx, "trouserLegOpening", "Leg Opening", item)}
      ${stepperRow("trouser", idx, "trouserInseam", "Inseam", item)}
      <div class="measurement-row">
        <label for="trouser-totalLength-${idx}">Total Length</label>
        <div class="stepper">
          <input id="trouser-totalLength-${idx}" class="button-select" type="text" data-type="trouser" data-index="${idx}" data-field="trouserTotalLength" value="${escapeHtml(item.trouserTotalLength || "")}" />
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-cuff-${idx}">Cuff Style</label>
        <div class="stepper">
          <select id="trouser-cuff-${idx}" class="button-select" data-type="trouser" data-index="${idx}" data-field="trouserCuff">
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

function renderShirtMeasurements(item, idx) {
  return `
    <div class="measurement-controls">
      ${stepperRow("shirt", idx, "shirtSleeve", "Sleeve Length", item, { disablePlus: true })}
      ${stepperRow("shirt", idx, "shirtBody", "Body Length", item, { disablePlus: true })}
      ${stepperRow("shirt", idx, "shirtSlimBody", "Slim Body", item, { disablePlus: true })}
    </div>`;
}

function formatMultiline(text) {
  const lines = String(text).split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) return "";
  return lines.map((line) => `• ${escapeHtml(line)}`).join("<br>");
}

/* ------------------------------ item list UI ------------------------------ */
function renderItemList(type) {
  const map = {
    jacket: { items: jackets, container: jacketItemsEl, label: "Jacket / Suit" },
    trouser: { items: trousers, container: trouserItemsEl, label: "Trouser" },
    shirt: { items: shirts, container: shirtItemsEl, label: "Shirt" },
  };
  const config = map[type];
  if (!config) return;
  const { items, container, label } = config;

  container.innerHTML = items.map((item, idx) => {
    const canRemove = items.length > 1;
    const showTitle = items.length > 1;
    const itemTitle = showTitle && idx > 0 ? `${label} ${idx + 1}` : "";
    const measurementBlock = type === "jacket" ? renderJacketMeasurements(item, idx)
      : type === "trouser" ? renderTrouserMeasurements(item, idx)
      : renderShirtMeasurements(item, idx);
    return `
      <section class="repeat-item" data-type="${type}" data-index="${idx}">
        ${itemTitle ? `<p class="repeat-item-title">${itemTitle}</p>` : ""}
        <div class="garment-field size-row">
          <label for="${type}-size-${idx}">Size</label>
          <input id="${type}-size-${idx}" type="text" placeholder="e.g. 40R" data-type="${type}" data-index="${idx}" data-field="size" value="${escapeHtml(item.size)}" />
        </div>
        <div class="garment-field">
          <label for="${type}-description-${idx}">Description</label>
          <input id="${type}-description-${idx}" type="text" data-type="${type}" data-index="${idx}" data-field="description" value="${escapeHtml(item.description)}" />
        </div>
        ${measurementBlock}
        <label for="${type}-adjustments-${idx}">Additional Notes</label>
        <textarea id="${type}-adjustments-${idx}" rows="3" data-type="${type}" data-index="${idx}" data-field="adjustments">${escapeHtml(item.adjustments)}</textarea>
        ${canRemove ? `<button type="button" class="remove-item-btn" data-action="remove" data-type="${type}" data-index="${idx}">Remove ${label}</button>` : ""}
      </section>`;
  }).join("");
}

function handleDynamicInput(event) {
  const target = event.target;
  if (!target?.dataset) return;
  const { type, field, index, action } = target.dataset;
  if (!type || !field || index === undefined) return;
  const idx = Number(index);
  if (Number.isNaN(idx)) return;
  const items = type === "jacket" ? jackets : type === "trouser" ? trousers : type === "shirt" ? shirts : null;
  if (!items || !items[idx]) return;

  if (action === "clear-value") {
    const rawValue = target.value.trim();
    if (!rawValue) {
      items[idx][field] = 0;
      renderItemList(type);
      onInputChange();
      return;
    }
    const formatted = formatSignedQuarter(items[idx][field] || 0);
    if (target.value !== formatted) target.value = formatted;
    return;
  }

  items[idx][field] = target.value;
  onInputChange();
}

function handleDynamicClick(event) {
  const target = event.target;
  if (!target?.dataset) return;
  const { action, type } = target.dataset;
  const idx = Number(target.dataset.index);
  if (!action || !type || Number.isNaN(idx)) return;

  if (action === "step") {
    const field = target.dataset.field;
    const dir = Number(target.dataset.dir);
    if (!field || Number.isNaN(dir)) return;
    const items = type === "jacket" ? jackets : type === "trouser" ? trousers : type === "shirt" ? shirts : null;
    if (!items || !items[idx]) return;
    const current = Number(items[idx][field]) || 0;
    const next = Math.round((current + dir * 0.5) * 2) / 2;
    items[idx][field] = isNegativeOnlyField(field) ? normalizeNegativeOnlyValue(next) : next;
    renderItemList(type);
    onInputChange();
    return;
  }

  if (action !== "remove") return;
  const listMap = { jacket: jackets, trouser: trousers, shirt: shirts };
  const list = listMap[type];
  if (!list) return;
  list.splice(idx, 1);
  if (!list.length) list.push(createEmptyItem());
  renderItemList(type);
  onInputChange();
}

/* --------------------------------- tabs ----------------------------------- */
function getActiveTab() {
  const active = garmentTabs.find((tab) => tab.classList.contains("is-active"));
  return active ? active.dataset.tab : "jacket";
}

function setActiveTab(tabName, shouldPersist = true) {
  const normalized = ["jacket", "trousers", "shirts"].includes(tabName) ? tabName : "jacket";
  garmentTabs.forEach((tab) => {
    const isActive = tab.dataset.tab === normalized;
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", isActive ? "true" : "false");
  });
  garmentPanels.forEach((panel) => {
    const isActive = panel.dataset.tabPanel === normalized;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });
  renderOutput();
  if (shouldPersist) saveToStorage();
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
function buildGarmentSections(jacketsArr = jackets, trousersArr = trousers, shirtsArr = shirts) {
  const sections = [];

  jacketsArr.filter(hasGarmentData).forEach((entry, idx) => {
    const measurements = [];
    if (Number(entry.halfBack)) measurements.push({ label: "1/2 Back", value: `${formatSignedQuarter(entry.halfBack)} cm` });
    if (Number(entry.halfWaist)) measurements.push({ label: "1/2 Waist", value: `${formatSignedQuarter(entry.halfWaist)} cm` });
    if (Number(entry.shortenBody)) measurements.push({ label: "Body Length", value: `${formatSignedQuarter(entry.shortenBody)} cm` });
    if (Number(entry.sleeves)) measurements.push({ label: "Sleeve Length", value: `${formatSignedQuarter(entry.sleeves)} cm` });
    if (Number(entry.sleeveWidth)) measurements.push({ label: "Sleeve Width", value: `${formatSignedQuarter(entry.sleeveWidth)} cm` });
    const tightenCollar = normalizeNegativeOnlyValue(entry.tightenCollar);
    if (Number(tightenCollar)) measurements.push({ label: "Tighten Collar", value: `${formatSignedQuarter(tightenCollar)} cm` });
    if ((entry.buttons || "").trim()) measurements.push({ label: "Buttons", value: entry.buttons });
    const sizeDesc = [formatSizeDisplay(entry.size), entry.description].filter((v) => (v || "").trim()).join(", ");
    sections.push({
      label: idx === 0 ? "Jacket" : `Jacket ${idx + 1}`,
      sizeDesc,
      measurements,
      notes: (entry.adjustments || "").trim(),
    });
  });

  trousersArr.filter(hasGarmentData).forEach((entry, idx) => {
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
    const sizeDesc = [formatSizeDisplay(entry.size), entry.description].filter((v) => (v || "").trim()).join(", ");
    sections.push({
      label: idx === 0 ? "Trouser" : `Trouser ${idx + 1}`,
      sizeDesc,
      measurements,
      notes: (entry.adjustments || "").trim(),
    });
  });

  shirtsArr.filter(hasGarmentData).forEach((entry, idx) => {
    const measurements = [];
    const shirtSleeve = normalizeNegativeOnlyValue(entry.shirtSleeve);
    if (Number(shirtSleeve)) measurements.push({ label: "Sleeve Length", value: `${formatSignedQuarter(shirtSleeve)} cm` });
    const shirtBody = normalizeNegativeOnlyValue(entry.shirtBody);
    if (Number(shirtBody)) measurements.push({ label: "Body Length", value: `${formatSignedQuarter(shirtBody)} cm` });
    const shirtSlimBody = normalizeNegativeOnlyValue(entry.shirtSlimBody);
    if (Number(shirtSlimBody)) measurements.push({ label: "Slim Body", value: `${formatSignedQuarter(shirtSlimBody)} cm` });
    const sizeDesc = [formatSizeDisplay(entry.size), entry.description].filter((v) => (v || "").trim()).join(", ");
    sections.push({
      label: idx === 0 ? "Shirt" : `Shirt ${idx + 1}`,
      sizeDesc,
      measurements,
      notes: (entry.adjustments || "").trim(),
    });
  });

  // Fold size/description into the measurement list so both the screen
  // preview and the docx show it as the first line of the section.
  return sections.map((section) => ({
    ...section,
    measurements: section.sizeDesc ? [{ label: "Size / Fit", value: section.sizeDesc }, ...section.measurements] : section.measurements,
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
function buildTicketDataFromState(state) {
  const now = new Date();
  const dueDateValue = state.dueDate || "";
  const jacketsArr = Array.isArray(state.jackets) && state.jackets.length ? state.jackets : [createEmptyItem()];
  const trousersArr = Array.isArray(state.trousers) && state.trousers.length ? state.trousers : [createEmptyItem()];
  const shirtsArr = Array.isArray(state.shirts) && state.shirts.length ? state.shirts : [createEmptyItem()];
  return {
    customerName: (state.customerName || "").trim(),
    tailor: state.tailor || "",
    salesperson: state.salesperson || "",
    dueDate: formatDueDate(dueDateValue),
    rush: isRushDueDate(dueDateValue, now),
    balanceDisplay: formatBalanceDisplay(state.balanceDue),
    createdDisplay: now.toLocaleString(),
    createdAt: now.toISOString(),
    garmentSections: buildGarmentSections(jacketsArr, trousersArr, shirtsArr),
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
    jackets, trousers, shirts,
    activeTab: getActiveTab(),
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

  jackets = Array.isArray(parsed.jackets) && parsed.jackets.length ? parsed.jackets.map((item) => ({ ...createEmptyItem(), ...item })) : [createEmptyItem()];
  trousers = Array.isArray(parsed.trousers) && parsed.trousers.length ? parsed.trousers.map((item) => ({ ...createEmptyItem(), ...item })) : [createEmptyItem()];
  shirts = Array.isArray(parsed.shirts) && parsed.shirts.length ? parsed.shirts.map((item) => ({ ...createEmptyItem(), ...item })) : [createEmptyItem()];

  renderItemList("jacket");
  renderItemList("trouser");
  renderItemList("shirt");
  setActiveTab(parsed.activeTab || "jacket", false);

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
    jackets = [createEmptyItem()];
    trousers = [createEmptyItem()];
    shirts = [createEmptyItem()];
    renderItemList("jacket"); renderItemList("trouser"); renderItemList("shirt");
    renderOutput();
    return;
  }
  try {
    const parsed = JSON.parse(stored);
    applyState(parsed);
    if (parsed.savedAt) saveStatus.textContent = `Last saved ${new Date(parsed.savedAt).toLocaleString()}`;
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    jackets = [createEmptyItem()]; trousers = [createEmptyItem()]; shirts = [createEmptyItem()];
    renderItemList("jacket"); renderItemList("trouser"); renderItemList("shirt");
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

  jackets = [createEmptyItem()]; trousers = [createEmptyItem()]; shirts = [createEmptyItem()];
  renderItemList("jacket"); renderItemList("trouser"); renderItemList("shirt");
  setActiveTab("jacket", false);
  localStorage.removeItem(STORAGE_KEY);
  saveStatus.textContent = "Form cleared";
  renderOutput();
}

/* ----------------------------------- auth ---------------------------------- */
function updateAuthGate() {
  const isSignedIn = auth.hasValidSession();
  authGate.hidden = isSignedIn;
  appShell.hidden = !isSignedIn;
  authStatus.textContent = isSignedIn ? "Signed in with Google." : "Google sign-in required.";
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
  const isNew = view === "newTicket";
  navNewTicket.classList.toggle("is-active", isNew);
  navNewTicket.setAttribute("aria-selected", isNew ? "true" : "false");
  navTicketLog.classList.toggle("is-active", !isNew);
  navTicketLog.setAttribute("aria-selected", !isNew ? "true" : "false");
  viewNewTicket.classList.toggle("is-active", isNew);
  viewTicketLog.classList.toggle("is-active", !isNew);
  if (!isNew) refreshTicketLog();
}

navNewTicket.addEventListener("click", () => setView("newTicket"));
navTicketLog.addEventListener("click", () => setView("ticketLog"));

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
function summarizeGarments(formState) {
  if (!formState) return "—";
  const counts = [
    ["Jacket", (formState.jackets || []).filter(hasGarmentData).length],
    ["Trouser", (formState.trousers || []).filter(hasGarmentData).length],
    ["Shirt", (formState.shirts || []).filter(hasGarmentData).length],
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
          <td data-label="Due Date"><span class="cell-value">${escapeHtml(ticket.dueDate)}${ticket.rush ? '<span class="rush-pill">Rush</span>' : ""}</span></td>
          <td data-label="Balance"><span class="cell-value">${escapeHtml((ticket.balance || "—").replace(/^Balance\s+/i, ""))}</span></td>
          <td data-label="Status"><span class="cell-value"><select class="status-select" data-row="${ticket.rowNumber}" data-status="${escapeHtml(ticket.status)}">${statusOptions}</select></span></td>
          <td data-label="Status Date" class="status-date-cell"><span class="cell-value">${escapeHtml(ticket.statusDate || "—")}</span></td>
          <td data-label="Notes"><span class="cell-value"><input type="text" class="notes-input" data-row="${ticket.rowNumber}" placeholder="Add a note…" value="${escapeHtml(ticket.notes || "")}" /></span></td>
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
    ticketLogStatus.textContent = `${ticketCache.length} ticket${ticketCache.length === 1 ? "" : "s"} logged.`;
    const openCount = ticketCache.filter((t) => t.status === "Open" || t.status === "In Progress").length;
    const rushCount = ticketCache.filter((t) => t.rush && t.status !== "Picked Up").length;
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

ticketLogBody.addEventListener("focusout", async (event) => {
  const input = event.target.closest(".notes-input");
  if (!input) return;
  const rowNumber = Number(input.dataset.row);
  const ticket = ticketCache.find((item) => item.rowNumber === rowNumber);
  if (!ticket) return;
  const newNotes = input.value;
  if (newNotes === (ticket.notes || "")) return; // unchanged — skip the write

  try {
    const token = await auth.getValidToken();
    const ledgerId = await getOrCreateLedger(token);
    await updateTicketNotes(token, ledgerId, rowNumber, newNotes);
    ticket.notes = newNotes;
    ticketLogStatus.textContent = `Saved note for ${ticket.customerName || "ticket"}.`;
  } catch (err) {
    input.value = ticket.notes || ""; // revert on failure
    ticketLogStatus.textContent = `Could not save note: ${err.message}`;
  }
});

ticketLogBody.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && event.target.closest(".notes-input")) {
    event.preventDefault();
    event.target.blur(); // triggers the focusout save handler above
  }
});

ticketLogBody.addEventListener("click", async (event) => {
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

/* ---------------------------------- saving ----------------------------------- */
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
      refreshTicketLog();
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
    refreshTicketLog();
  } catch (err) {
    saveStatus.textContent = `Drive error: ${err.message}`;
    alert(`Save failed: ${err.message}`);
  }
});

clearBtn.addEventListener("click", clearAllFields);

/* ---------------------------------- wiring ----------------------------------- */
addJacketBtn.addEventListener("click", () => { jackets.push(createEmptyItem()); renderItemList("jacket"); onInputChange(); });
addTrouserBtn.addEventListener("click", () => { trousers.push(createEmptyItem()); renderItemList("trouser"); onInputChange(); });
addShirtBtn.addEventListener("click", () => { shirts.push(createEmptyItem()); renderItemList("shirt"); onInputChange(); });

jacketItemsEl.addEventListener("input", handleDynamicInput);
trouserItemsEl.addEventListener("input", handleDynamicInput);
shirtItemsEl.addEventListener("input", handleDynamicInput);
jacketItemsEl.addEventListener("change", handleDynamicInput);
trouserItemsEl.addEventListener("change", handleDynamicInput);
shirtItemsEl.addEventListener("change", handleDynamicInput);
jacketItemsEl.addEventListener("click", handleDynamicClick);
trouserItemsEl.addEventListener("click", handleDynamicClick);
shirtItemsEl.addEventListener("click", handleDynamicClick);

tailorInput.addEventListener("change", () => { onInputChange(); });

garmentTabs.forEach((tab) => tab.addEventListener("click", () => setActiveTab(tab.dataset.tab || "jacket")));

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
