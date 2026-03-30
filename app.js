const STORAGE_KEY = "alterationsTicketStateV1";

const customerNameInput = document.getElementById("customerName");
const tailorInput = document.getElementById("tailor");
const tailorOtherInput = document.getElementById("tailorOther");
const tailorOtherWrap = document.getElementById("tailorOtherWrap");
const salespersonInput = document.getElementById("salesperson");
const dueDateInput = document.getElementById("dueDate");
let dueDateHiddenInput = null;
let dueDateDisplayInput = null;

const jacketItemsEl = document.getElementById("jacketItems");
const addJacketBtn = document.getElementById("addJacketBtn");

const trouserItemsEl = document.getElementById("trouserItems");
const addTrouserBtn = document.getElementById("addTrouserBtn");

const shirtItemsEl = document.getElementById("shirtItems");
const addShirtBtn = document.getElementById("addShirtBtn");

const printArea = document.getElementById("printArea");
const saveStatus = document.getElementById("saveStatus");
const saveBtn = document.getElementById("saveBtn");
const driveAuthBtn = document.getElementById("driveAuthBtn");
const driveSaveBtn = document.getElementById("driveSaveBtn");
const printBtn = document.getElementById("printBtn");
const clearBtn = document.getElementById("clearBtn");
const garmentTabs = Array.from(document.querySelectorAll(".garment-tab"));
const garmentPanels = Array.from(document.querySelectorAll(".garment-panel"));

const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";
const DRIVE_TOKEN_KEY = "driveAccessToken";
const DRIVE_TOKEN_EXP_KEY = "driveAccessTokenExp";
const GOOGLE_CLIENT_ID = "617892178220-84fg83gdjhjssb3et6e5ufjnkb8cn1v2.apps.googleusercontent.com";
const JACKET_SIZES = ["custom", "36", "38", "40", "42", "44", "46", "48"];
const TROUSER_SIZES = ["custom", "28", "30", "32", "34", "36", "38"];
const SHIRT_SIZES = ["custom", "15", "15.5", "15.75", "16", "16.5", "17", "17.5"];

let jackets = [];
let trousers = [];
let shirts = [];

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttr(text) {
  return escapeHtml(String(text));
}

function createEmptyItem() {
  return {
    size: "",
    description: "",
    adjustments: "",
    halfBack: 0,
    halfWaist: 0,
    shortenBody: 0,
    sleeves: 0,
    tightenCollar: 0,
    buttons: "",
    trouserWaist: 0,
    trouserSeat: 0,
    trouserThigh: 0,
    trouserKnee: 0,
    trouserLegOpening: 0,
    trouserInseam: 0,
    trouserTotalLength: "",
    trouserCuff: "",
    shirtSleeve: 0,
    shirtBody: 0,
  };
}

function hasGarmentData(item) {
  const measurementFields = [
    "halfBack",
    "halfWaist",
    "shortenBody",
    "sleeves",
    "tightenCollar",
    "trouserWaist",
    "trouserSeat",
    "trouserThigh",
    "trouserKnee",
    "trouserLegOpening",
    "trouserInseam",
    "shirtSleeve",
    "shirtBody",
  ];
  const hasMeasurements = measurementFields.some((field) => Number(item?.[field]) !== 0);
  return Boolean(
    (item?.size || "").trim() ||
      (item?.description || "").trim() ||
      (item?.adjustments || "").trim() ||
      (item?.buttons || "").trim() ||
      (item?.trouserTotalLength || "").trim() ||
      (item?.trouserCuff || "").trim() ||
      hasMeasurements,
  );
}

function formatQuarter(value) {
  const abs = Math.abs(Number(value) || 0);
  const whole = Math.floor(abs / 4);
  const remainder = abs % 4;
  const fraction =
    remainder === 0 ? "" : remainder === 1 ? "1/4" : remainder === 2 ? "1/2" : "3/4";

  if (whole && fraction) return `${whole} ${fraction}`;
  if (whole) return `${whole}`;
  return fraction || "0";
}

function formatSignedQuarter(value) {
  const numeric = Number(value) || 0;
  if (numeric === 0) return "0";
  const sign = numeric > 0 ? "+" : "-";
  return `${sign} ${formatQuarter(numeric)}`;
}

function formatFileBaseName(name) {
  const normalized = String(name || "ticket")
    .trim()
    .replace(/[^a-z0-9]+/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) return "Ticket";
  return normalized
    .split(" ")
    .map((part) => (part ? part[0].toUpperCase() + part.slice(1).toLowerCase() : part))
    .join(" ");
}

function buildButtonsOptions(selectedValue) {
  return buildOptions(["1", "2", "3", "4"], selectedValue);
}

function renderJacketMeasurements(item, idx) {
  const formatValue = (field) => formatSignedQuarter(item?.[field] || 0);
  return `
    <div class="measurement-controls">
      <div class="measurement-row">
        <label for="jacket-halfBack-${idx}">1/2 Back</label>
        <div class="stepper" id="jacket-halfBack-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="halfBack" data-dir="-1">-</button>
          <input
            class="stepper-value-input"
            type="text"
            inputmode="none"
            value="${formatValue("halfBack")}"
            data-action="clear-value"
            data-type="jacket"
            data-index="${idx}"
            data-field="halfBack"
          />
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="halfBack" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="jacket-halfWaist-${idx}">1/2 Waist</label>
        <div class="stepper" id="jacket-halfWaist-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="halfWaist" data-dir="-1">-</button>
          <input
            class="stepper-value-input"
            type="text"
            inputmode="none"
            value="${formatValue("halfWaist")}"
            data-action="clear-value"
            data-type="jacket"
            data-index="${idx}"
            data-field="halfWaist"
          />
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="halfWaist" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="jacket-shortenBody-${idx}">Shorten Body</label>
        <div class="stepper" id="jacket-shortenBody-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="shortenBody" data-dir="-1">-</button>
          <input
            class="stepper-value-input"
            type="text"
            inputmode="none"
            value="${formatValue("shortenBody")}"
            data-action="clear-value"
            data-type="jacket"
            data-index="${idx}"
            data-field="shortenBody"
          />
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="shortenBody" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="jacket-sleeves-${idx}">Sleeves</label>
        <div class="stepper" id="jacket-sleeves-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="sleeves" data-dir="-1">-</button>
          <input
            class="stepper-value-input"
            type="text"
            inputmode="none"
            value="${formatValue("sleeves")}"
            data-action="clear-value"
            data-type="jacket"
            data-index="${idx}"
            data-field="sleeves"
          />
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="sleeves" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="jacket-tightenCollar-${idx}">Tighten Collar</label>
        <div class="stepper" id="jacket-tightenCollar-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="tightenCollar" data-dir="-1">-</button>
          <input
            class="stepper-value-input"
            type="text"
            inputmode="none"
            value="${formatValue("tightenCollar")}"
            data-action="clear-value"
            data-type="jacket"
            data-index="${idx}"
            data-field="tightenCollar"
          />
          <button type="button" class="stepper-btn" data-action="step" data-type="jacket" data-index="${idx}" data-field="tightenCollar" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="jacket-buttons-${idx}">Buttons</label>
        <div class="stepper">
          <select id="jacket-buttons-${idx}" class="button-select" data-type="jacket" data-index="${idx}" data-field="buttons">
            ${buildButtonsOptions(item?.buttons || "")}
          </select>
        </div>
      </div>
    </div>
  `;
}

function renderTrouserMeasurements(item, idx) {
  const formatValue = (field) => formatSignedQuarter(item?.[field] || 0);
  return `
    <div class="measurement-controls">
      <div class="measurement-row">
        <label for="trouser-waist-${idx}">Waist</label>
        <div class="stepper" id="trouser-waist-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserWaist" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("trouserWaist")}" data-action="clear-value" data-type="trouser" data-index="${idx}" data-field="trouserWaist" />
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserWaist" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-seat-${idx}">Seat</label>
        <div class="stepper" id="trouser-seat-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserSeat" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("trouserSeat")}" data-action="clear-value" data-type="trouser" data-index="${idx}" data-field="trouserSeat" />
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserSeat" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-thigh-${idx}">Thigh</label>
        <div class="stepper" id="trouser-thigh-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserThigh" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("trouserThigh")}" data-action="clear-value" data-type="trouser" data-index="${idx}" data-field="trouserThigh" />
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserThigh" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-knee-${idx}">Knee</label>
        <div class="stepper" id="trouser-knee-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserKnee" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("trouserKnee")}" data-action="clear-value" data-type="trouser" data-index="${idx}" data-field="trouserKnee" />
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserKnee" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-legOpening-${idx}">Leg Opening</label>
        <div class="stepper" id="trouser-legOpening-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserLegOpening" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("trouserLegOpening")}" data-action="clear-value" data-type="trouser" data-index="${idx}" data-field="trouserLegOpening" />
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserLegOpening" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-inseam-${idx}">Inseam</label>
        <div class="stepper" id="trouser-inseam-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserInseam" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("trouserInseam")}" data-action="clear-value" data-type="trouser" data-index="${idx}" data-field="trouserInseam" />
          <button type="button" class="stepper-btn" data-action="step" data-type="trouser" data-index="${idx}" data-field="trouserInseam" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-totalLength-${idx}">Total Length</label>
        <div class="stepper">
          <input id="trouser-totalLength-${idx}" class="button-select" type="text" data-type="trouser" data-index="${idx}" data-field="trouserTotalLength" value="${escapeAttr(item.trouserTotalLength || "")}" />
        </div>
      </div>
      <div class="measurement-row">
        <label for="trouser-cuff-${idx}">Cuff Style</label>
        <div class="stepper">
          <select id="trouser-cuff-${idx}" class="button-select" data-type="trouser" data-index="${idx}" data-field="trouserCuff">
            ${buildOptions(["No Cuff", "1 3/4 in Cuff", "2 in Cuff"], item?.trouserCuff || "")}
          </select>
        </div>
      </div>
    </div>
  `;
}

function renderShirtMeasurements(item, idx) {
  const formatValue = (field) => formatSignedQuarter(item?.[field] || 0);
  return `
    <div class="measurement-controls">
      <div class="measurement-row">
        <label for="shirt-sleeve-${idx}">Sleeve Length</label>
        <div class="stepper" id="shirt-sleeve-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="shirt" data-index="${idx}" data-field="shirtSleeve" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("shirtSleeve")}" data-action="clear-value" data-type="shirt" data-index="${idx}" data-field="shirtSleeve" />
          <button type="button" class="stepper-btn" data-action="step" data-type="shirt" data-index="${idx}" data-field="shirtSleeve" data-dir="1">+</button>
        </div>
      </div>
      <div class="measurement-row">
        <label for="shirt-body-${idx}">Body Length</label>
        <div class="stepper" id="shirt-body-${idx}">
          <button type="button" class="stepper-btn" data-action="step" data-type="shirt" data-index="${idx}" data-field="shirtBody" data-dir="-1">-</button>
          <input class="stepper-value-input" type="text" inputmode="none" value="${formatValue("shirtBody")}" data-action="clear-value" data-type="shirt" data-index="${idx}" data-field="shirtBody" />
          <button type="button" class="stepper-btn" data-action="step" data-type="shirt" data-index="${idx}" data-field="shirtBody" data-dir="1">+</button>
        </div>
      </div>
    </div>
  `;
}

function formatMultiline(text) {
  const lines = String(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return "";
  }

  return lines.map((line) => `• ${escapeHtml(line)}`).join("<br>");
}

function getBalanceDueValue() {
  const selected = document.querySelector('input[name="balanceDue"]:checked');
  return selected ? selected.value : "";
}

function updateTailorOtherVisibility() {
  tailorOtherWrap.hidden = tailorInput.value !== "Other";
}

function getTailorDisplayName() {
  if (tailorInput.value !== "Other") {
    return tailorInput.value || "Luis";
  }

  return tailorOtherInput.value.trim() || "Other";
}

function formatDueDate(dateValue) {
  if (!dateValue) {
    return "Not set";
  }

  return new Date(`${dateValue}T00:00:00`).toLocaleDateString();
}

function isRushDueDate(dateValue, generatedAt) {
  if (!dateValue) {
    return false;
  }

  const due = new Date(`${dateValue}T23:59:59`);
  const diffMs = due.getTime() - generatedAt.getTime();
  const fourDaysMs = 4 * 24 * 60 * 60 * 1000;
  return diffMs >= 0 && diffMs < fourDaysMs;
}

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
  if (shouldPersist) {
    saveToStorage();
  }
}


let driveTokenClient = null;

function getDriveToken() {
  return localStorage.getItem(DRIVE_TOKEN_KEY) || "";
}

function getDriveTokenExpiry() {
  return Number(localStorage.getItem(DRIVE_TOKEN_EXP_KEY) || 0);
}

function setDriveToken(token, expiresInSeconds) {
  if (!token) return;
  const expiry = Date.now() + Number(expiresInSeconds || 0) * 1000;
  localStorage.setItem(DRIVE_TOKEN_KEY, token);
  localStorage.setItem(DRIVE_TOKEN_EXP_KEY, String(expiry));
  saveStatus.textContent = "Google Drive connected";
  updateDriveButtons();
}

function clearDriveToken() {
  localStorage.removeItem(DRIVE_TOKEN_KEY);
  localStorage.removeItem(DRIVE_TOKEN_EXP_KEY);
  updateDriveButtons();
}

function ensureDriveClient() {
  if (driveTokenClient) return driveTokenClient;
  if (!window.google?.accounts?.oauth2) {
    return null;
  }
  driveTokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: DRIVE_SCOPE,
    callback: () => {},
  });
  return driveTokenClient;
}

function updateDriveButtons() {
  const token = getDriveToken();
  const exp = getDriveTokenExpiry();
  const isValid = token && exp && Date.now() < exp - 30_000;
  driveAuthBtn.style.display = isValid ? "none" : "inline-flex";
}

function requestDriveToken(prompt) {
  return new Promise((resolve, reject) => {
    const client = ensureDriveClient();
    if (!client) {
      reject(new Error("Google Identity Services not loaded"));
      return;
    }
    client.callback = (response) => {
      if (response?.access_token) {
        setDriveToken(response.access_token, response.expires_in);
        resolve(response.access_token);
      } else {
        reject(new Error("No access token returned"));
      }
    };
    client.requestAccessToken({ prompt });
  });
}

async function getValidDriveToken() {
  const token = getDriveToken();
  const exp = getDriveTokenExpiry();
  if (token && exp && Date.now() < exp - 30_000) {
    return token;
  }
  clearDriveToken();
  return requestDriveToken("");
}

function buildDriveMultipart({ filename, content, mimeType }) {
  const boundary = `boundary_${Math.random().toString(36).slice(2)}`;
  const metadata = { name: filename };
  const body =
    `--${boundary}\r\n` +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    `${JSON.stringify(metadata)}\r\n` +
    `--${boundary}\r\n` +
    `Content-Type: ${mimeType}\r\n\r\n` +
    `${content}\r\n` +
    `--${boundary}--`;

  return { body, boundary };
}

function isIOSDevice() {
  const ua = navigator.userAgent || "";
  const isIOS = /iPad|iPhone|iPod/.test(ua);
  const isIpadOS = /Macintosh/.test(ua) && navigator.maxTouchPoints > 1;
  return isIOS || isIpadOS;
}

function getDueDateValue() {
  return dueDateHiddenInput ? dueDateHiddenInput.value : dueDateInput.value;
}

function setDueDateValue(value) {
  if (dueDateHiddenInput) {
    dueDateHiddenInput.value = value || "";
    updateDueDateDisplay();
    return;
  }
  dueDateInput.value = value || "";
}

function updateDueDateDisplay() {
  if (!dueDateDisplayInput) {
    return;
  }
  const value = dueDateHiddenInput ? dueDateHiddenInput.value : "";
  dueDateDisplayInput.value = value ? new Date(`${value}T00:00:00`).toLocaleDateString() : "";
}

function buildOptions(options, selectedValue) {
  const base = '<option value="">Select</option>';
  const values = options
    .map((value) => {
      const selected = selectedValue === value ? " selected" : "";
      const label = value === "custom" ? "Custom" : value;
      return `<option value="${value}"${selected}>${label}</option>`;
    })
    .join("");
  return `${base}${values}`;
}

function renderItemList(type) {
  const map = {
    jacket: { items: jackets, container: jacketItemsEl, label: "Jacket / Suit", sizes: JACKET_SIZES },
    trouser: { items: trousers, container: trouserItemsEl, label: "Trouser", sizes: TROUSER_SIZES },
    shirt: { items: shirts, container: shirtItemsEl, label: "Shirt", sizes: SHIRT_SIZES },
  };
  const config = map[type];
  if (!config) return;

  const { items, container, label, sizes } = config;

  container.innerHTML = items
    .map((item, idx) => {
      const canRemove = items.length > 1;
      const showTitle = items.length > 1;
      const itemTitle = showTitle && idx > 0 ? `${label} ${idx + 1}` : "";
      const measurementBlock =
        type === "jacket"
          ? renderJacketMeasurements(item, idx)
          : type === "trouser"
            ? renderTrouserMeasurements(item, idx)
            : type === "shirt"
              ? renderShirtMeasurements(item, idx)
              : "";
      return `
        <section class="repeat-item" data-type="${type}" data-index="${idx}">
          ${itemTitle ? `<p class="repeat-item-title">${itemTitle}</p>` : ""}
          <div class="garment-field size-row">
            <label for="${type}-size-${idx}">Size</label>
            <select id="${type}-size-${idx}" data-type="${type}" data-index="${idx}" data-field="size">
              ${buildOptions(sizes, item.size)}
            </select>
          </div>
          <div class="garment-field">
            <label for="${type}-description-${idx}">Description</label>
            <input
              id="${type}-description-${idx}"
              type="text"
              data-type="${type}"
              data-index="${idx}"
              data-field="description"
              value="${escapeAttr(item.description)}"
            />
          </div>
          ${measurementBlock}
          <label for="${type}-adjustments-${idx}">Additional Notes</label>
          <textarea
            id="${type}-adjustments-${idx}"
            rows="3"
            data-type="${type}"
            data-index="${idx}"
            data-field="adjustments"
          >${escapeHtml(item.adjustments)}</textarea>
          ${
            canRemove
              ? `<button type="button" class="remove-item-btn" data-action="remove" data-type="${type}" data-index="${idx}">Remove ${label}</button>`
              : ""
          }
        </section>
      `;
    })
    .join("");
}

function handleDynamicInput(event) {
  const target = event.target;
  if (!target || !target.dataset) {
    return;
  }

  const { type, field, index, action } = target.dataset;
  if (!type || !field || index === undefined) {
    return;
  }

  const idx = Number(index);
  if (Number.isNaN(idx)) {
    return;
  }

  const items = type === "jacket" ? jackets : type === "trouser" ? trousers : type === "shirt" ? shirts : null;
  if (!items || !items[idx]) {
    return;
  }

  if (action === "clear-value") {
    const rawValue = target.value.trim();
    if (!rawValue) {
      items[idx][field] = 0;
      renderItemList(type);
      onInputChange();
      return;
    }
    const formatted = formatSignedQuarter(items[idx][field] || 0);
    if (target.value !== formatted) {
      target.value = formatted;
    }
    return;
  }

  items[idx][field] = target.value;
  onInputChange();
}

function handleDynamicClick(event) {
  const target = event.target;
  if (!target || !target.dataset) {
    return;
  }

  const action = target.dataset.action;
  const type = target.dataset.type;
  const idx = Number(target.dataset.index);
  if (!action || !type || Number.isNaN(idx)) {
    return;
  }

  if (action === "step") {
    const field = target.dataset.field;
    const dir = Number(target.dataset.dir);
    if (!field || Number.isNaN(dir)) {
      return;
    }

    const items = type === "jacket" ? jackets : type === "trouser" ? trousers : type === "shirt" ? shirts : null;
    if (!items || !items[idx]) {
      return;
    }

    const current = Number(items[idx][field]) || 0;
    items[idx][field] = current + dir;
    renderItemList(type);
    onInputChange();
    return;
  }

  if (action !== "remove") {
    return;
  }

  if (type === "jacket") {
    jackets.splice(idx, 1);
    if (!jackets.length) jackets.push(createEmptyItem());
    renderItemList("jacket");
  } else if (type === "trouser") {
    trousers.splice(idx, 1);
    if (!trousers.length) trousers.push(createEmptyItem());
    renderItemList("trouser");
  } else if (type === "shirt") {
    shirts.splice(idx, 1);
    if (!shirts.length) shirts.push(createEmptyItem());
    renderItemList("shirt");
  }

  onInputChange();
}

function renderOutput() {
  const now = new Date();
  const customerName = customerNameInput.value.trim() || "Not provided";
  const tailor = getTailorDisplayName();
  const salesperson = salespersonInput.value || "Select";
  const dueDateValue = getDueDateValue();
  const dueDate = formatDueDate(dueDateValue);
  const rushFlag = isRushDueDate(dueDateValue, now);
  const balanceDue = getBalanceDueValue();

  const jacketFilled = jackets.filter(hasGarmentData);
  const trouserFilled = trousers.filter(hasGarmentData);
  const shirtFilled = shirts.filter(hasGarmentData);

  const garmentSections = [];

  if (jacketFilled.length) {
    garmentSections.push(
      ...jacketFilled.map((entry, idx) => {
        const measurements = [];
        if (Number(entry.halfBack)) {
          measurements.push(`<p><strong>1/2 Back:</strong> ${formatSignedQuarter(entry.halfBack)}"</p>`);
        }
        if (Number(entry.halfWaist)) {
          measurements.push(`<p><strong>1/2 Waist:</strong> ${formatSignedQuarter(entry.halfWaist)}"</p>`);
        }
        if (Number(entry.shortenBody)) {
          measurements.push(`<p><strong>Shorten Body:</strong> ${formatSignedQuarter(entry.shortenBody)}"</p>`);
        }
        if (Number(entry.sleeves)) {
          measurements.push(`<p><strong>Sleeves:</strong> ${formatSignedQuarter(entry.sleeves)}"</p>`);
        }
        if (Number(entry.tightenCollar)) {
          measurements.push(`<p><strong>Tighten Collar:</strong> ${formatSignedQuarter(entry.tightenCollar)}"</p>`);
        }
        if ((entry.buttons || "").trim()) {
          measurements.push(`<p><strong>Buttons:</strong> ${escapeHtml(entry.buttons)}</p>`);
        }

        const notes = formatMultiline(entry.adjustments);
        const outputPieces = [];
        const jacketLabel = idx === 0 ? "Jacket" : `Jacket ${idx + 1}`;
        const sizeDesc = [entry.size, entry.description].filter((value) => (value || "").trim());
        if (sizeDesc.length) {
          outputPieces.push(`<p><strong>${jacketLabel}:</strong> ${escapeHtml(sizeDesc.join(", "))}</p>`);
        }
        outputPieces.push(...measurements);
        if (notes) {
          outputPieces.push(
            `<p><strong>${idx === 0 ? "Jacket" : `Jacket ${idx + 1}`} Additional Notes:</strong><br>${notes}</p>`,
          );
        }

        if (!outputPieces.length) {
          return "";
        }
        return `
          <section class="garment-output-block">
            ${outputPieces.join("")}
          </section>
        `;
      }),
    );
  }

  if (trouserFilled.length) {
    garmentSections.push(
      ...trouserFilled.map((entry, idx) => {
        const notes = formatMultiline(entry.adjustments);
        const measurements = [];
        if (Number(entry.trouserWaist)) {
          measurements.push(`<p><strong>Waist:</strong> ${formatSignedQuarter(entry.trouserWaist)}"</p>`);
        }
        if (Number(entry.trouserSeat)) {
          measurements.push(`<p><strong>Seat:</strong> ${formatSignedQuarter(entry.trouserSeat)}"</p>`);
        }
        if (Number(entry.trouserThigh)) {
          measurements.push(`<p><strong>Thigh:</strong> ${formatSignedQuarter(entry.trouserThigh)}"</p>`);
        }
        if (Number(entry.trouserKnee)) {
          measurements.push(`<p><strong>Knee:</strong> ${formatSignedQuarter(entry.trouserKnee)}"</p>`);
        }
        if (Number(entry.trouserLegOpening)) {
          measurements.push(`<p><strong>Leg Opening:</strong> ${formatSignedQuarter(entry.trouserLegOpening)}"</p>`);
        }
        if (Number(entry.trouserInseam)) {
          measurements.push(`<p><strong>Inseam:</strong> ${formatSignedQuarter(entry.trouserInseam)}"</p>`);
        }
        if ((entry.trouserTotalLength || "").trim()) {
          measurements.push(`<p><strong>Total Length:</strong> ${escapeHtml(entry.trouserTotalLength)}"</p>`);
        }
        if ((entry.trouserCuff || "").trim()) {
          const cuffLabel = entry.trouserCuff.replace(/\bin\b/g, "\"");
          measurements.push(`<p><strong>Cuff Style:</strong> ${escapeHtml(cuffLabel)}</p>`);
        }

        const outputPieces = [];
        const trouserLabel = idx === 0 ? "Trouser" : `Trouser ${idx + 1}`;
        const sizeDesc = [entry.size, entry.description].filter((value) => (value || "").trim());
        if (sizeDesc.length) {
          outputPieces.push(`<p><strong>${trouserLabel}:</strong> ${escapeHtml(sizeDesc.join(", "))}</p>`);
        }
        outputPieces.push(...measurements);
        if (notes) {
          outputPieces.push(
            `<p><strong>${idx === 0 ? "Trouser" : `Trouser ${idx + 1}`} Additional Notes:</strong><br>${notes}</p>`,
          );
        }

        if (!outputPieces.length) {
          return "";
        }

        return `
          <section class="garment-output-block">
            ${outputPieces.join("")}
          </section>
        `;
      }),
    );
  }

  if (shirtFilled.length) {
    garmentSections.push(
      ...shirtFilled.map((entry, idx) => {
        const notes = formatMultiline(entry.adjustments);
        const measurements = [];
        if (Number(entry.shirtSleeve)) {
          measurements.push(`<p><strong>Sleeve Length:</strong> ${formatSignedQuarter(entry.shirtSleeve)}"</p>`);
        }
        if (Number(entry.shirtBody)) {
          measurements.push(`<p><strong>Body Length:</strong> ${formatSignedQuarter(entry.shirtBody)}"</p>`);
        }

        const outputPieces = [];
        const shirtLabel = idx === 0 ? "Shirt" : `Shirt ${idx + 1}`;
        const sizeDesc = [entry.size, entry.description].filter((value) => (value || "").trim());
        if (sizeDesc.length) {
          outputPieces.push(`<p><strong>${shirtLabel}:</strong> ${escapeHtml(sizeDesc.join(", "))}</p>`);
        }
        outputPieces.push(...measurements);
        if (notes) {
          outputPieces.push(
            `<p><strong>${idx === 0 ? "Shirt" : `Shirt ${idx + 1}`} Additional Notes:</strong><br>${notes}</p>`,
          );
        }

        if (!outputPieces.length) {
          return "";
        }

        return `
          <section class="garment-output-block">
            ${outputPieces.join("")}
          </section>
        `;
      }),
    );
  }

  printArea.innerHTML = `
    ${rushFlag ? '<p class="rush-flag">**RUSH**</p>' : ""}
    <p><strong>Customer Name:</strong> ${escapeHtml(customerName)}</p>
    <p><strong>Tailor:</strong> ${escapeHtml(tailor)}</p>
    <p><strong>Salesperson Name:</strong> ${escapeHtml(salesperson)}</p>
    <p><strong>Due Date:</strong> ${escapeHtml(dueDate)}</p>
    ${balanceDue ? `<p><strong>Balance Due?:</strong> ${escapeHtml(balanceDue)}</p>` : ""}
    ${garmentSections.join("")}
    <p class="meta">Generated: ${escapeHtml(now.toLocaleString())}</p>
  `;
}

function buildState(savedAt = new Date().toISOString()) {
  return {
    customerName: customerNameInput.value,
    tailor: tailorInput.value,
    tailorOther: tailorOtherInput.value,
    salesperson: salespersonInput.value,
    dueDate: getDueDateValue(),
    balanceDue: getBalanceDueValue(),
    jackets,
    trousers,
    shirts,
    activeTab: getActiveTab(),
    savedAt,
  };
}

function saveToStorage() {
  const state = buildState();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  const stamp = new Date(state.savedAt).toLocaleTimeString();
  saveStatus.textContent = `Autosaved at ${stamp}`;
}

function saveToLocalFile() {
  const state = buildState();
  const safeName = formatFileBaseName(state.customerName).slice(0, 40);
  const datePart = new Date(state.savedAt).toISOString().slice(0, 10);
  const filename = `${safeName}_${datePart}.doc`;
  renderOutput();

  const content = buildExportHtml();
  const blob = new Blob([content], { type: "application/msword;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  saveToStorage();
  saveStatus.textContent = `Saved file: ${filename}`;
}

function buildExportHtml() {
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Alterations Ticket</title>
    <style>
      @page { size: 4in 6in; margin: 0.1in; }
      body {
        margin: 0;
        padding: 0;
        background: white;
        font-family: "Avenir Next", "Trebuchet MS", sans-serif;
        color: #1e1d1a;
      }
      .output {
        width: 3.8in;
        min-height: 5.8in;
        box-sizing: border-box;
        padding: 0;
        margin: 0;
        font-size: 10pt;
        line-height: 1.25;
      }
      .output p { margin: 0.03in 0; }
      .output .doc-title { font-size: 12pt; margin: 0 0 0.06in; color: #3d352c; }
      .output .rush-flag { color: #b42318; font-weight: 800; margin: 0 0 0.08in; }
      .output .meta { margin: 0.08in 0 0; color: #6c645d; font-size: 6pt; }
      .output .garment-output-block {
        margin-top: 0.08in;
        padding-top: 0.06in;
        border-top: 1px solid #c8c0b4;
        break-inside: avoid;
      }
    </style>
  </head>
  <body>
    <article class="output">
      ${printArea.innerHTML}
    </article>
  </body>
</html>`;
}

function parseLegacyArray(primary, secondary) {
  const arr = [];
  if (hasGarmentData(primary)) arr.push(primary);
  if (hasGarmentData(secondary)) arr.push(secondary);
  if (!arr.length) arr.push(createEmptyItem());
  return arr;
}

function loadFromStorage() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    jackets = [createEmptyItem()];
    trousers = [createEmptyItem()];
    shirts = [createEmptyItem()];
    renderItemList("jacket");
    renderItemList("trouser");
    renderItemList("shirt");
    updateTailorOtherVisibility();
    renderOutput();
    return;
  }

  try {
    const parsed = JSON.parse(stored);
    customerNameInput.value = parsed.customerName || "";
    tailorInput.value = parsed.tailor || "Luis";
    tailorOtherInput.value = parsed.tailorOther || "";
    salespersonInput.value = parsed.salesperson || "";
    setDueDateValue(parsed.dueDate || "");

    if (Array.isArray(parsed.jackets)) {
      jackets = parsed.jackets.map((item) => ({
        size: item?.size || "",
        description: item?.description || "",
        adjustments: item?.adjustments || "",
        halfBack: Number(item?.halfBack) || 0,
        halfWaist: Number(item?.halfWaist) || 0,
        shortenBody: Number(item?.shortenBody) || 0,
        sleeves: Number(item?.sleeves) || 0,
        tightenCollar: Number(item?.tightenCollar) || 0,
        buttons: item?.buttons || "",
      }));
      if (!jackets.length) jackets = [createEmptyItem()];
    } else {
      jackets = parseLegacyArray(
        {
          size: parsed.jacketSize || "",
          description: parsed.jacketDescription || "",
          adjustments: parsed.jacketAdjustments || "",
          halfBack: Number(parsed.jacketHalfBack) || 0,
          halfWaist: Number(parsed.jacketHalfWaist) || 0,
          shortenBody: Number(parsed.jacketShortenBody) || 0,
          sleeves: Number(parsed.jacketSleeves) || 0,
          tightenCollar: Number(parsed.jacketTightenCollar) || 0,
          buttons: parsed.jacketButtons || "",
        },
        {
          size: parsed.jacketSize2 || "",
          description: parsed.jacketDescription2 || "",
          adjustments: parsed.jacketAdjustments2 || "",
          halfBack: Number(parsed.jacketHalfBack2) || 0,
          halfWaist: Number(parsed.jacketHalfWaist2) || 0,
          shortenBody: Number(parsed.jacketShortenBody2) || 0,
          sleeves: Number(parsed.jacketSleeves2) || 0,
          tightenCollar: Number(parsed.jacketTightenCollar2) || 0,
          buttons: parsed.jacketButtons2 || "",
        },
      );
    }

    if (Array.isArray(parsed.trousers)) {
      trousers = parsed.trousers.map((item) => ({
        size: item?.size || "",
        description: item?.description || "",
        adjustments: item?.adjustments || "",
        trouserWaist: Number(item?.trouserWaist) || 0,
        trouserSeat: Number(item?.trouserSeat) || 0,
        trouserThigh: Number(item?.trouserThigh) || 0,
        trouserKnee: Number(item?.trouserKnee) || 0,
        trouserLegOpening: Number(item?.trouserLegOpening) || 0,
        trouserInseam: Number(item?.trouserInseam) || 0,
        trouserTotalLength: item?.trouserTotalLength || "",
        trouserCuff: item?.trouserCuff || "",
      }));
      if (!trousers.length) trousers = [createEmptyItem()];
    } else {
      trousers = parseLegacyArray(
        {
          size: parsed.trouserSize || "",
          description: parsed.trouserDescription || "",
          adjustments: parsed.trouserAdjustments || "",
        },
        {
          size: parsed.trouserSize2 || "",
          description: parsed.trouserDescription2 || "",
          adjustments: parsed.trouserAdjustments2 || "",
        },
      );
    }

    if (Array.isArray(parsed.shirts)) {
      shirts = parsed.shirts.map((item) => ({
        size: item?.size || "",
        description: item?.description || "",
        adjustments: item?.adjustments || "",
        shirtSleeve: Number(item?.shirtSleeve) || 0,
        shirtBody: Number(item?.shirtBody) || 0,
      }));
      if (!shirts.length) shirts = [createEmptyItem()];
    } else {
      shirts = parseLegacyArray(
        {
          size: parsed.shirtSize || "",
          description: parsed.shirtDescription || "",
          adjustments: parsed.shirtAdjustments || "",
        },
        {
          size: parsed.shirtSize2 || "",
          description: parsed.shirtDescription2 || "",
          adjustments: parsed.shirtAdjustments2 || "",
        },
      );
    }

    renderItemList("jacket");
    renderItemList("trouser");
    renderItemList("shirt");

    setActiveTab(parsed.activeTab || "jacket", false);

    const balanceDue = parsed.balanceDue === "Yes" || parsed.balanceDue === "No" ? parsed.balanceDue : "";
    const selectedBalance = document.querySelector('input[name="balanceDue"]:checked');
    if (selectedBalance) {
      selectedBalance.checked = false;
    }
    if (balanceDue) {
      const target = document.querySelector(`input[name="balanceDue"][value="${balanceDue}"]`);
      if (target) {
        target.checked = true;
      }
    }

    if (parsed.savedAt) {
      saveStatus.textContent = `Last saved ${new Date(parsed.savedAt).toLocaleString()}`;
    }
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    jackets = [createEmptyItem()];
    trousers = [createEmptyItem()];
    shirts = [createEmptyItem()];
    renderItemList("jacket");
    renderItemList("trouser");
    renderItemList("shirt");
  }

  updateTailorOtherVisibility();
  renderOutput();
}

let saveTimer;
function onInputChange() {
  renderOutput();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveToStorage, 350);
}

function validateBeforePrint() {
  const missing = [];
  if (!customerNameInput.value.trim()) missing.push("Customer Name");
  if (!salespersonInput.value.trim()) missing.push("Salesperson Name");
  if (!getDueDateValue().trim()) missing.push("Due Date");
  return missing;
}

function clearAllFields() {
  clearTimeout(saveTimer);

  customerNameInput.value = "";
  tailorInput.value = "Luis";
  tailorOtherInput.value = "";
  salespersonInput.value = "";
  setDueDateValue("");

  const selectedBalance = document.querySelector('input[name="balanceDue"]:checked');
  if (selectedBalance) {
    selectedBalance.checked = false;
  }

  jackets = [createEmptyItem()];
  trousers = [createEmptyItem()];
  shirts = [createEmptyItem()];
  renderItemList("jacket");
  renderItemList("trouser");
  renderItemList("shirt");

  setActiveTab("jacket", false);
  updateTailorOtherVisibility();
  localStorage.removeItem(STORAGE_KEY);
  saveStatus.textContent = "Form cleared";
  renderOutput();
}

printBtn.addEventListener("click", () => {
  const missing = validateBeforePrint();
  if (missing.length) {
    alert(`Please complete these required fields before printing:\n- ${missing.join("\n- ")}`);
    return;
  }

  renderOutput();
  saveToStorage();
  window.print();
});

driveAuthBtn.addEventListener("click", () => {
  if (!GOOGLE_CLIENT_ID || GOOGLE_CLIENT_ID === "REPLACE_WITH_CLIENT_ID") {
    alert("Set your Google Client ID in app.js before connecting Drive.");
    return;
  }
  requestDriveToken("consent").catch(() => {
    alert("Google Drive connection failed. Please try again.");
  });
});

driveSaveBtn.addEventListener("click", async () => {
  const missing = validateBeforePrint();
  if (missing.length) {
    alert(`Please complete these required fields before saving:\n- ${missing.join("\n- ")}`);
    return;
  }

  if (!GOOGLE_CLIENT_ID || GOOGLE_CLIENT_ID === "REPLACE_WITH_CLIENT_ID") {
    alert("Set your Google Client ID in app.js before connecting Drive.");
    return;
  }

  const state = buildState();
  renderOutput();
  resetPrintScale();

  const safeName = formatFileBaseName(state.customerName).slice(0, 40);
  const datePart = new Date(state.savedAt).toISOString().slice(0, 10);
  const filename = `${safeName}_${datePart}.doc`;

  try {
    const token = await getValidDriveToken();
    const { body, boundary } = buildDriveMultipart({
      filename,
      content: buildExportHtml(),
      mimeType: "application/msword",
    });

    const response = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": `multipart/related; boundary=${boundary}`,
      },
      body,
    });

    if (!response.ok) {
      if (response.status === 401) {
        clearDriveToken();
        alert("Drive session expired. Click Save to Drive again.");
        return;
      }
      throw new Error("Upload failed");
    }

    saveStatus.textContent = `Saved to Drive: ${filename}`;
  } catch (err) {
    alert("Drive upload failed. Please try again.");
  }
});

saveBtn.addEventListener("click", () => {
  const missing = validateBeforePrint();
  if (missing.length) {
    alert(`Please complete these required fields before saving:\n- ${missing.join("\n- ")}`);
    return;
  }

  saveToLocalFile();
});

clearBtn.addEventListener("click", clearAllFields);

addJacketBtn.addEventListener("click", () => {
  jackets.push(createEmptyItem());
  renderItemList("jacket");
  onInputChange();
});

addTrouserBtn.addEventListener("click", () => {
  trousers.push(createEmptyItem());
  renderItemList("trouser");
  onInputChange();
});

addShirtBtn.addEventListener("click", () => {
  shirts.push(createEmptyItem());
  renderItemList("shirt");
  onInputChange();
});

jacketItemsEl.addEventListener("input", handleDynamicInput);
trouserItemsEl.addEventListener("input", handleDynamicInput);
jacketItemsEl.addEventListener("change", handleDynamicInput);
trouserItemsEl.addEventListener("change", handleDynamicInput);
shirtItemsEl.addEventListener("input", handleDynamicInput);
shirtItemsEl.addEventListener("change", handleDynamicInput);
jacketItemsEl.addEventListener("click", handleDynamicClick);
trouserItemsEl.addEventListener("click", handleDynamicClick);
shirtItemsEl.addEventListener("click", handleDynamicClick);

tailorInput.addEventListener("change", () => {
  updateTailorOtherVisibility();
  onInputChange();
});

garmentTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    setActiveTab(tab.dataset.tab || "jacket");
  });
});

[
  customerNameInput,
  tailorInput,
  tailorOtherInput,
  salespersonInput,
  ...document.querySelectorAll('input[name="balanceDue"]'),
].forEach((el) => el.addEventListener("input", onInputChange));

if (!dueDateHiddenInput) {
  dueDateInput.addEventListener("input", onInputChange);
}

function setupIOSDateFallback() {
  if (!isIOSDevice()) {
    return;
  }

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
  original.style.position = "absolute";
  original.style.opacity = "0";
  original.style.pointerEvents = "none";
  original.style.height = "0";
  original.style.width = "0";
  original.style.border = "0";
  original.style.padding = "0";
  original.style.margin = "0";
  original.setAttribute("aria-hidden", "true");

  original.parentNode.insertBefore(wrapper, original);
  wrapper.appendChild(display);
  wrapper.appendChild(original);

  dueDateHiddenInput = original;
  dueDateDisplayInput = display;

  const openPicker = () => {
    dueDateHiddenInput.focus();
    dueDateHiddenInput.click();
  };
  display.addEventListener("click", openPicker);
  display.addEventListener("focus", openPicker);
  dueDateHiddenInput.addEventListener("change", () => {
    updateDueDateDisplay();
    onInputChange();
  });

  updateDueDateDisplay();
}

setupIOSDateFallback();
loadFromStorage();
updateDriveButtons();
