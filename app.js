const STORAGE_KEY = "alterationsTicketStateV1";

const customerNameInput = document.getElementById("customerName");
const tailorInput = document.getElementById("tailor");
const tailorOtherInput = document.getElementById("tailorOther");
const tailorOtherWrap = document.getElementById("tailorOtherWrap");
const salespersonInput = document.getElementById("salesperson");
const dueDateInput = document.getElementById("dueDate");

const jacketItemsEl = document.getElementById("jacketItems");
const addJacketBtn = document.getElementById("addJacketBtn");

const trouserItemsEl = document.getElementById("trouserItems");
const addTrouserBtn = document.getElementById("addTrouserBtn");

const shirtItemsEl = document.getElementById("shirtItems");
const addShirtBtn = document.getElementById("addShirtBtn");

const printArea = document.getElementById("printArea");
const saveStatus = document.getElementById("saveStatus");
const saveBtn = document.getElementById("saveBtn");
const printBtn = document.getElementById("printBtn");
const clearBtn = document.getElementById("clearBtn");
const garmentTabs = Array.from(document.querySelectorAll(".garment-tab"));
const garmentPanels = Array.from(document.querySelectorAll(".garment-panel"));

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
  return { size: "", description: "", adjustments: "" };
}

function hasGarmentData(item) {
  return Boolean((item?.size || "").trim() || (item?.description || "").trim() || (item?.adjustments || "").trim());
}

function formatMultiline(text) {
  const lines = String(text)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return "None";
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
      const itemTitle = idx === 0 ? label : `${label} ${idx + 1}`;
      return `
        <section class="repeat-item" data-type="${type}" data-index="${idx}">
          <p class="repeat-item-title">${itemTitle}</p>
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
          <label for="${type}-adjustments-${idx}">Adjustments</label>
          <textarea
            id="${type}-adjustments-${idx}"
            rows="5"
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

  const { type, field, index } = target.dataset;
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

  items[idx][field] = target.value;
  onInputChange();
}

function handleDynamicClick(event) {
  const target = event.target;
  if (!target || !target.dataset || target.dataset.action !== "remove") {
    return;
  }

  const type = target.dataset.type;
  const idx = Number(target.dataset.index);
  if (!type || Number.isNaN(idx)) {
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
  const dueDate = formatDueDate(dueDateInput.value);
  const rushFlag = isRushDueDate(dueDateInput.value, now);
  const balanceDue = getBalanceDueValue();

  const jacketFilled = jackets.filter(hasGarmentData);
  const trouserFilled = trousers.filter(hasGarmentData);
  const shirtFilled = shirts.filter(hasGarmentData);

  const garmentSections = [];

  if (jacketFilled.length) {
    garmentSections.push(
      ...jacketFilled.map((entry, idx) => `
        <section class="garment-output-block">
          ${entry.size ? `<p><strong>${idx === 0 ? "Jacket" : `Jacket ${idx + 1}`} Size:</strong> ${escapeHtml(entry.size)}</p>` : ""}
          ${entry.description ? `<p><strong>Description:</strong> ${escapeHtml(entry.description)}</p>` : ""}
          <p><strong>${idx === 0 ? "Jacket" : `Jacket ${idx + 1}`} Adjustments:</strong><br>${formatMultiline(entry.adjustments)}</p>
        </section>
      `),
    );
  }

  if (trouserFilled.length) {
    garmentSections.push(
      ...trouserFilled.map((entry, idx) => `
        <section class="garment-output-block">
          ${entry.size ? `<p><strong>${idx === 0 ? "Trouser" : `Trouser ${idx + 1}`} Size:</strong> ${escapeHtml(entry.size)}</p>` : ""}
          ${entry.description ? `<p><strong>Description:</strong> ${escapeHtml(entry.description)}</p>` : ""}
          <p><strong>${idx === 0 ? "Trouser" : `Trouser ${idx + 1}`} Adjustments:</strong><br>${formatMultiline(entry.adjustments)}</p>
        </section>
      `),
    );
  }

  if (shirtFilled.length) {
    garmentSections.push(
      ...shirtFilled.map((entry, idx) => `
        <section class="garment-output-block">
          ${entry.size ? `<p><strong>${idx === 0 ? "Shirt" : `Shirt ${idx + 1}`} Size:</strong> ${escapeHtml(entry.size)}</p>` : ""}
          ${entry.description ? `<p><strong>Description:</strong> ${escapeHtml(entry.description)}</p>` : ""}
          <p><strong>${idx === 0 ? "Shirt" : `Shirt ${idx + 1}`} Adjustments:</strong><br>${formatMultiline(entry.adjustments)}</p>
        </section>
      `),
    );
  }

  printArea.innerHTML = `
    <h3 class="doc-title">Alterations Ticket</h3>
    ${rushFlag ? '<p class="rush-flag">**RUSH**</p>' : ""}
    <p><strong>Customer Name:</strong> ${escapeHtml(customerName)}</p>
    <p><strong>Tailor:</strong> ${escapeHtml(tailor)}</p>
    <p><strong>Salesperson Name:</strong> ${escapeHtml(salesperson)}</p>
    <p><strong>Due Date:</strong> ${escapeHtml(dueDate)}</p>
    ${balanceDue ? `<p><strong>Balance Due?:</strong> ${escapeHtml(balanceDue)}</p>` : ""}
    ${garmentSections.join("")}
  `;
}

function buildState(savedAt = new Date().toISOString()) {
  return {
    customerName: customerNameInput.value,
    tailor: tailorInput.value,
    tailorOther: tailorOtherInput.value,
    salesperson: salespersonInput.value,
    dueDate: dueDateInput.value,
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
  const safeName = (state.customerName || "ticket")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
  const datePart = new Date(state.savedAt).toISOString().slice(0, 10);
  const filename = `${safeName || "ticket"}-${datePart}.doc`;

  renderOutput();

  const content = `<!doctype html>
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
    dueDateInput.value = parsed.dueDate || "";

    if (Array.isArray(parsed.jackets)) {
      jackets = parsed.jackets.map((item) => ({
        size: item?.size || "",
        description: item?.description || "",
        adjustments: item?.adjustments || "",
      }));
      if (!jackets.length) jackets = [createEmptyItem()];
    } else {
      jackets = parseLegacyArray(
        {
          size: parsed.jacketSize || "",
          description: parsed.jacketDescription || "",
          adjustments: parsed.jacketAdjustments || "",
        },
        {
          size: parsed.jacketSize2 || "",
          description: parsed.jacketDescription2 || "",
          adjustments: parsed.jacketAdjustments2 || "",
        },
      );
    }

    if (Array.isArray(parsed.trousers)) {
      trousers = parsed.trousers.map((item) => ({
        size: item?.size || "",
        description: item?.description || "",
        adjustments: item?.adjustments || "",
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
  if (!dueDateInput.value.trim()) missing.push("Due Date");
  return missing;
}

function clearAllFields() {
  clearTimeout(saveTimer);

  customerNameInput.value = "";
  tailorInput.value = "Luis";
  tailorOtherInput.value = "";
  salespersonInput.value = "";
  dueDateInput.value = "";

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
  dueDateInput,
  ...document.querySelectorAll('input[name="balanceDue"]'),
].forEach((el) => el.addEventListener("input", onInputChange));
loadFromStorage();
