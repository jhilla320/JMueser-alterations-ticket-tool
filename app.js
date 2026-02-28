const STORAGE_KEY = "alterationsTicketStateV1";

const customerNameInput = document.getElementById("customerName");
const tailorInput = document.getElementById("tailor");
const salespersonInput = document.getElementById("salesperson");
const dueDateInput = document.getElementById("dueDate");
const jacketSizeInput = document.getElementById("jacketSize");
const trouserSizeInput = document.getElementById("trouserSize");
const shirtSizeInput = document.getElementById("shirtSize");
const notesInput = document.getElementById("notes");

const printArea = document.getElementById("printArea");
const saveStatus = document.getElementById("saveStatus");
const saveBtn = document.getElementById("saveBtn");
const printBtn = document.getElementById("printBtn");
const clearBtn = document.getElementById("clearBtn");
const garmentTabs = Array.from(document.querySelectorAll(".garment-tab"));
const garmentPanels = Array.from(document.querySelectorAll(".garment-panel"));

function escapeHtml(text) {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getBalanceDueValue() {
  const selected = document.querySelector('input[name="balanceDue"]:checked');
  return selected ? selected.value : "";
}

function formatMultiline(text) {
  const lines = text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return "None";
  }

  return lines.map((line) => `• ${escapeHtml(line)}`).join("<br>");
}

const measurementFields = [
  { key: "halfBack", label: "1/2 Back" },
  { key: "halfWaist", label: "1/2 Waist" },
  { key: "length", label: "Length" },
  { key: "leftSleeve", label: "Left Sleeve" },
  { key: "rightSleeve", label: "Right Sleeve" },
  { key: "bicep", label: "Bicep" },
  { key: "lowerCollar", label: "Lower Collar" },
  { key: "tightenCollar", label: "Tighten Collar" },
];

const trouserMeasurementFields = [
  { key: "trouserWaist", label: "Waist" },
  { key: "trouserSeat", label: "Seat" },
  { key: "trouserThighCrotch", label: "Thigh/Crotch" },
  { key: "trouserKnee", label: "Knee" },
  { key: "trouserCuffWidth", label: "Cuff Width" },
  { key: "trouserRise", label: "Rise" },
  { key: "trouserInseam", label: "Inseam" },
];

const shirtMeasurementFields = [
  { key: "shirtBodyLength", label: "Body Length" },
  { key: "shirtBody", label: "Body" },
  { key: "shirtLeftSleeve", label: "Left Sleeve" },
  { key: "shirtRightSleeve", label: "Right Sleeve" },
  { key: "shirtBicep", label: "Bicep" },
];

function getMeasurementInput(key) {
  return {
    plus: document.getElementById(`${key}Plus`),
    minus: document.getElementById(`${key}Minus`),
  };
}

function getMeasurementState() {
  return measurementFields.reduce((acc, field) => {
    const input = getMeasurementInput(field.key);
    acc[field.key] = {
      plus: input.plus.value.trim(),
      minus: input.minus.value.trim(),
    };
    return acc;
  }, {});
}

function formatMeasurementsForOutput() {
  const items = measurementFields
    .map((field) => {
      const input = getMeasurementInput(field.key);
      const plusValue = input.plus.value.trim();
      const minusValue = input.minus.value.trim();
      if (!plusValue && !minusValue) {
        return null;
      }

      const parts = [];
      if (plusValue) {
        parts.push(`+${escapeHtml(plusValue)}`);
      }
      if (minusValue) {
        parts.push(`-${escapeHtml(minusValue)}`);
      }

      return `• <strong>${field.label}:</strong> ${parts.join(" / ")}`;
    })
    .filter(Boolean);

  return items.length ? items.join("<br>") : "None";
}

function formatMeasurementStateForOutput(fields, state) {
  const withCm = (value) => {
    const normalized = (value || "").trim().replace(/\s*cm$/i, "");
    return `${escapeHtml(normalized)} cm`;
  };

  const items = fields
    .map((field) => {
      const values = state[field.key] || {};
      const plusValue = (values.plus || "").trim();
      const minusValue = (values.minus || "").trim();
      if (!plusValue && !minusValue) {
        return null;
      }

      const parts = [];
      if (plusValue) {
        parts.push(`+${withCm(plusValue)}`);
      }
      if (minusValue) {
        parts.push(`-${withCm(minusValue)}`);
      }

      return `• <strong>${field.label}:</strong> ${parts.join(" / ")}`;
    })
    .filter(Boolean);

  return items.length ? items.join("<br>") : "None";
}

function hasMeasurementState(fields, state) {
  return fields.some((field) => {
    const values = state[field.key] || {};
    return (values.plus || "").trim() || (values.minus || "").trim();
  });
}

function getTrouserMeasurementInput(key) {
  return {
    plus: document.getElementById(`${key}Plus`),
    minus: document.getElementById(`${key}Minus`),
  };
}

function getTrouserMeasurementState() {
  return trouserMeasurementFields.reduce((acc, field) => {
    const input = getTrouserMeasurementInput(field.key);
    acc[field.key] = {
      plus: input.plus.value.trim(),
      minus: input.minus.value.trim(),
    };
    return acc;
  }, {});
}

function formatTrouserMeasurementsForOutput() {
  const items = trouserMeasurementFields
    .map((field) => {
      const input = getTrouserMeasurementInput(field.key);
      const plusValue = input.plus.value.trim();
      const minusValue = input.minus.value.trim();
      if (!plusValue && !minusValue) {
        return null;
      }

      const parts = [];
      if (plusValue) {
        parts.push(`+${escapeHtml(plusValue)}`);
      }
      if (minusValue) {
        parts.push(`-${escapeHtml(minusValue)}`);
      }

      return `• <strong>${field.label}:</strong> ${parts.join(" / ")}`;
    })
    .filter(Boolean);

  return items.length ? items.join("<br>") : "None";
}

function getShirtMeasurementInput(key) {
  return {
    plus: document.getElementById(`${key}Plus`),
    minus: document.getElementById(`${key}Minus`),
  };
}

function getShirtMeasurementState() {
  return shirtMeasurementFields.reduce((acc, field) => {
    const input = getShirtMeasurementInput(field.key);
    acc[field.key] = {
      plus: input.plus.value.trim(),
      minus: input.minus.value.trim(),
    };
    return acc;
  }, {});
}

function formatShirtMeasurementsForOutput() {
  const items = shirtMeasurementFields
    .map((field) => {
      const input = getShirtMeasurementInput(field.key);
      const plusValue = input.plus.value.trim();
      const minusValue = input.minus.value.trim();
      if (!plusValue && !minusValue) {
        return null;
      }

      const parts = [];
      if (plusValue) {
        parts.push(`+${escapeHtml(plusValue)}`);
      }
      if (minusValue) {
        parts.push(`-${escapeHtml(minusValue)}`);
      }

      return `• <strong>${field.label}:</strong> ${parts.join(" / ")}`;
    })
    .filter(Boolean);

  return items.length ? items.join("<br>") : "None";
}

function formatDueDate(dateValue) {
  if (!dateValue) {
    return "Not set";
  }

  const date = new Date(`${dateValue}T00:00:00`);
  return date.toLocaleDateString();
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

function renderOutput() {
  const withCm = (value) => {
    const normalized = (value || "").trim().replace(/\s*cm$/i, "");
    return `${escapeHtml(normalized)} cm`;
  };

  const now = new Date();
  const customerName = customerNameInput.value.trim() || "Not provided";
  const tailor = tailorInput.value || "Luis";
  const salesperson = salespersonInput.value || "Select";
  const dueDate = formatDueDate(dueDateInput.value);
  const rushFlag = isRushDueDate(dueDateInput.value, now);
  const balanceDue = getBalanceDueValue();
  const jacketSize = jacketSizeInput.value.trim();
  const trouserSize = trouserSizeInput.value.trim();
  const shirtSize = shirtSizeInput.value.trim();
  const jacketMeasurements = getMeasurementState();
  const trouserMeasurements = getTrouserMeasurementState();
  const shirtMeasurements = getShirtMeasurementState();
  const jacketOutput = formatMeasurementStateForOutput(measurementFields, jacketMeasurements);
  const trouserOutput = formatMeasurementStateForOutput(trouserMeasurementFields, trouserMeasurements);
  const shirtOutput = formatMeasurementStateForOutput(shirtMeasurementFields, shirtMeasurements);
  const hasJacketChanges = Boolean(jacketSize || hasMeasurementState(measurementFields, jacketMeasurements));
  const hasTrouserChanges = Boolean(
    trouserSize || hasMeasurementState(trouserMeasurementFields, trouserMeasurements),
  );
  const hasShirtChanges = Boolean(shirtSize || hasMeasurementState(shirtMeasurementFields, shirtMeasurements));

  const garmentSections = [];
  if (hasJacketChanges) {
    garmentSections.push(`
      <section class="garment-output-block">
        ${jacketSize ? `<p><strong>Jacket Size:</strong> ${escapeHtml(jacketSize)}</p>` : ""}
        <p><strong>Jacket Measurements:</strong><br>${jacketOutput}</p>
      </section>
    `);
  }
  if (hasTrouserChanges) {
    garmentSections.push(`
      <section class="garment-output-block">
        ${trouserSize ? `<p><strong>Trouser Size:</strong> ${escapeHtml(trouserSize)}</p>` : ""}
        <p><strong>Trouser Measurements:</strong><br>${trouserOutput}</p>
      </section>
    `);
  }
  if (hasShirtChanges) {
    garmentSections.push(`
      <section class="garment-output-block">
        ${shirtSize ? `<p><strong>Shirt Size:</strong> ${escapeHtml(shirtSize)}</p>` : ""}
        <p><strong>Shirt Measurements:</strong><br>${shirtOutput}</p>
      </section>
    `);
  }

  printArea.innerHTML = `
    <h3 class="doc-title">${escapeHtml(`Alterations Ticket - ${customerName}`)}</h3>
    <p class="meta">Generated: ${now.toLocaleString()}</p>
    <p><strong>Customer Name:</strong> ${escapeHtml(customerName)}</p>
    <p><strong>Tailor:</strong> ${escapeHtml(tailor)}</p>
    <p><strong>Salesperson Name:</strong> ${escapeHtml(salesperson)}</p>
    <p><strong>Due Date:</strong> ${escapeHtml(dueDate)}</p>
    ${rushFlag ? "<p><strong>***RUSH***</strong></p>" : ""}
    ${balanceDue ? `<p><strong>Balance Due?:</strong> ${escapeHtml(balanceDue)}</p>` : ""}
    ${garmentSections.join("")}
    <section class="notes-output-block">
      <p><strong>Notes:</strong><br>${formatMultiline(notesInput.value)}</p>
    </section>
  `;
}

function buildState(savedAt = new Date().toISOString()) {
  return {
    customerName: customerNameInput.value,
    tailor: tailorInput.value,
    salesperson: salespersonInput.value,
    dueDate: dueDateInput.value,
    balanceDue: getBalanceDueValue(),
    jacketSize: jacketSizeInput.value,
    jacketMeasurements: getMeasurementState(),
    trouserSize: trouserSizeInput.value,
    trouserMeasurements: getTrouserMeasurementState(),
    shirtSize: shirtSizeInput.value,
    shirtMeasurements: getShirtMeasurementState(),
    activeTab: getActiveTab(),
    notes: notesInput.value,
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

  // Ensure saved output matches the latest rendered ticket.
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
      .output .meta { font-size: 8.5pt; margin: 0 0 0.08in; color: #6c645d; }
      .output .garment-output-block,
      .output .notes-output-block {
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

function loadFromStorage() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    renderOutput();
    return;
  }

  try {
    const parsed = JSON.parse(stored);
    customerNameInput.value = parsed.customerName || "";
    tailorInput.value = parsed.tailor || "Luis";
    salespersonInput.value = parsed.salesperson || "";
    dueDateInput.value = parsed.dueDate || "";
    jacketSizeInput.value = parsed.jacketSize || "";
    trouserSizeInput.value = parsed.trouserSize || "";
    shirtSizeInput.value = parsed.shirtSize || "";
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

    const jacketMeasurements = parsed.jacketMeasurements || parsed.measurements;
    if (jacketMeasurements && typeof jacketMeasurements === "object") {
      measurementFields.forEach((field) => {
        const input = getMeasurementInput(field.key);
        const value = jacketMeasurements[field.key] || {};

        // Backward compatibility with older sign/value storage.
        if ("sign" in value || "value" in value) {
          if (value.sign === "-") {
            input.minus.value = value.value || "";
            input.plus.value = "";
          } else {
            input.plus.value = value.value || "";
            input.minus.value = "";
          }
          return;
        }

        input.plus.value = value.plus || "";
        input.minus.value = value.minus || "";
      });
    }

    const trouserMeasurements = parsed.trouserMeasurements || {};
    if (parsed.trouserInseam) {
      trouserMeasurements.trouserInseam = {
        plus: parsed.trouserInseam,
        minus: "",
      };
    }
    if (trouserMeasurements && typeof trouserMeasurements === "object") {
      trouserMeasurementFields.forEach((field) => {
        const input = getTrouserMeasurementInput(field.key);
        const value = trouserMeasurements[field.key] || {};
        input.plus.value = value.plus || "";
        input.minus.value = value.minus || "";
      });
    }

    if (parsed.shirtMeasurements && typeof parsed.shirtMeasurements === "object") {
      shirtMeasurementFields.forEach((field) => {
        const input = getShirtMeasurementInput(field.key);
        const value = parsed.shirtMeasurements[field.key] || {};
        input.plus.value = value.plus || "";
        input.minus.value = value.minus || "";
      });
    }

    notesInput.value = parsed.notes || "";

    if (parsed.savedAt) {
      saveStatus.textContent = `Last saved ${new Date(parsed.savedAt).toLocaleString()}`;
    }
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }

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
  const customerName = customerNameInput.value.trim();
  const salesperson = salespersonInput.value.trim();
  const dueDate = dueDateInput.value.trim();

  if (!customerName) {
    missing.push("Customer Name");
  }
  if (!salesperson) {
    missing.push("Salesperson Name");
  }
  if (!dueDate) {
    missing.push("Due Date");
  }

  return missing;
}

function clearAllFields() {
  clearTimeout(saveTimer);

  customerNameInput.value = "";
  tailorInput.value = "Luis";
  salespersonInput.value = "";
  dueDateInput.value = "";

  const selectedBalance = document.querySelector('input[name="balanceDue"]:checked');
  if (selectedBalance) {
    selectedBalance.checked = false;
  }

  jacketSizeInput.value = "";
  trouserSizeInput.value = "";
  shirtSizeInput.value = "";
  notesInput.value = "";

  measurementFields.forEach((field) => {
    const input = getMeasurementInput(field.key);
    input.plus.value = "";
    input.minus.value = "";
  });

  trouserMeasurementFields.forEach((field) => {
    const input = getTrouserMeasurementInput(field.key);
    input.plus.value = "";
    input.minus.value = "";
  });

  shirtMeasurementFields.forEach((field) => {
    const input = getShirtMeasurementInput(field.key);
    input.plus.value = "";
    input.minus.value = "";
  });

  setActiveTab("jacket", false);
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

garmentTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    setActiveTab(tab.dataset.tab || "jacket");
  });
});

[
  customerNameInput,
  tailorInput,
  salespersonInput,
  dueDateInput,
  jacketSizeInput,
  trouserSizeInput,
  shirtSizeInput,
  notesInput,
  ...document.querySelectorAll('input[name="balanceDue"]'),
  ...measurementFields.flatMap((field) => {
    const input = getMeasurementInput(field.key);
    return [input.plus, input.minus];
  }),
  ...trouserMeasurementFields.flatMap((field) => {
    const input = getTrouserMeasurementInput(field.key);
    return [input.plus, input.minus];
  }),
  ...shirtMeasurementFields.flatMap((field) => {
    const input = getShirtMeasurementInput(field.key);
    return [input.plus, input.minus];
  }),
].forEach((el) => el.addEventListener("input", onInputChange));

loadFromStorage();
