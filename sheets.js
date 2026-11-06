import { DRIVE_FOLDER_ID, TICKET_LOG_FILE_NAME, LEDGER_SPREADSHEET_ID } from "./config.js";

const LEDGER_ID_CACHE_KEY = "ticketLedgerFileId";
const SHEET_TAB = "Tickets";
const HEADERS = [
  "Ticket ID",
  "Created At",
  "Client Name",
  "Tailor",
  "Salesperson",
  "Due Date",
  "Rush",
  "Balance",
  "Status",
  "Docx Filename",
  "Drive Link",
  "Data JSON",
  "Status Date",
  "Drive File ID",
  "Notes",
  "Ticket Number",
];

async function driveApi(token, path, options = {}) {
  const response = await fetch(`https://www.googleapis.com/drive/v3/${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
  });
  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please sign in again.");
    throw new Error(`Drive request failed (${response.status})`);
  }
  return response.json();
}

async function sheetsApi(token, path, options = {}) {
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please sign in again.");
    const detail = await response.json().catch(() => null);
    throw new Error(detail?.error?.message || "Google Sheets request failed");
  }
  return response.json();
}

async function findLedgerFile(token) {
  const params = new URLSearchParams({
    includeItemsFromAllDrives: "true",
    supportsAllDrives: "true",
    pageSize: "10",
    orderBy: "createdTime", // oldest first, so every user resolves duplicates to the same (original) file
    fields: "files(id,name)",
    q: `'${DRIVE_FOLDER_ID}' in parents and name = '${TICKET_LOG_FILE_NAME.replace(/'/g, "\\'")}' and trashed = false`,
  });
  const data = await driveApi(token, `files?${params.toString()}`);
  return data.files?.[0]?.id || null;
}

async function createLedgerFile(token) {
  const created = await sheetsApi(token, "", {
    method: "POST",
    body: JSON.stringify({
      properties: { title: TICKET_LOG_FILE_NAME },
      sheets: [{ properties: { title: SHEET_TAB } }],
    }),
  });
  const sheetId = created.spreadsheetId;

  // Move it from My Drive root into the shop's shared folder.
  try {
    await driveApi(token, `files/${sheetId}?addParents=${DRIVE_FOLDER_ID}&supportsAllDrives=true&fields=id,parents`, {
      method: "PATCH",
    });
  } catch (err) {
    // Don't leave an orphaned spreadsheet sitting in this account's personal
    // Drive if we can't actually put it where it belongs — clean it up and
    // surface a clear, actionable reason instead of Google's raw 404 text.
    try {
      await driveApi(token, `files/${sheetId}?supportsAllDrives=true`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ trashed: true }),
      });
    } catch {
      // If cleanup itself fails, the original error below is still the
      // important one to surface — don't let this mask it.
    }
    throw new Error(
      "Your Google account doesn't seem to have access to the shop's shared Drive folder. Ask whoever manages Drive sharing to add you to that folder (not just an individual file), then sign out and back in.",
    );
  }

  await sheetsApi(token, `${sheetId}/values/${encodeURIComponent(`${SHEET_TAB}!A1`)}?valueInputOption=RAW`, {
    method: "PUT",
    body: JSON.stringify({ values: [HEADERS] }),
  });

  return sheetId;
}

// The ticket log's spreadsheet ID if this browser already knows it (pinned in
// config, or remembered from an earlier load) — no network call needed.
export function getKnownLedgerId() {
  return LEDGER_SPREADSHEET_ID || localStorage.getItem(LEDGER_ID_CACHE_KEY) || "";
}

export async function getOrCreateLedger(token) {
  // A pinned ID wins over everything: no per-browser cache, no folder search.
  if (LEDGER_SPREADSHEET_ID) {
    localStorage.removeItem(LEDGER_ID_CACHE_KEY);
    return LEDGER_SPREADSHEET_ID;
  }

  const cached = localStorage.getItem(LEDGER_ID_CACHE_KEY);
  if (cached) {
    try {
      await sheetsApi(token, `${cached}?fields=spreadsheetId`);
      return cached;
    } catch {
      localStorage.removeItem(LEDGER_ID_CACHE_KEY);
    }
  }

  const found = await findLedgerFile(token);
  const ledgerId = found || (await createLedgerFile(token));
  localStorage.setItem(LEDGER_ID_CACHE_KEY, ledgerId);
  return ledgerId;
}

// Writes one row to the first free row of the tab, always starting in
// column A. This replaces Sheets' "append" command, which guesses where the
// data table starts and can drift to a different starting column if anyone
// has hand-edited the sheet (a stray value, a blank gap, an emptied header
// cell). Returns the 1-indexed row number that was written.
async function writeRowAtEnd(token, ledgerId, tab, lastCol, row) {
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const colA = await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${tab}!A:A`)}`);
    const nextRow = Math.max(colA.values?.length || 0, 1) + 1;

    // The tab's grid only has so many rows; add more when we run off the end.
    const info = await sheetsApi(token, `${ledgerId}?fields=sheets.properties`);
    const props = info.sheets?.find((sheet) => sheet.properties.title === tab)?.properties;
    const rowCount = props?.gridProperties?.rowCount ?? 0;
    if (props && nextRow > rowCount) {
      await sheetsApi(token, `${ledgerId}:batchUpdate`, {
        method: "POST",
        body: JSON.stringify({
          requests: [{ appendDimension: { sheetId: props.sheetId, dimension: "ROWS", length: Math.max(100, nextRow - rowCount) } }],
        }),
      });
    }

    const range = `${tab}!A${nextRow}:${lastCol}${nextRow}`;
    await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
      method: "PUT",
      body: JSON.stringify({ values: [row] }),
    });

    // Confirm our row is the one that landed (guards against two people
    // saving in the same instant and writing to the same row).
    const check = await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${tab}!A${nextRow}`)}`);
    if ((check.values?.[0]?.[0] ?? "") === String(row[0])) return nextRow;
  }
  throw new Error("Could not safely add the row to the sheet after several tries. Please try saving again.");
}

export async function appendTicketRow(token, ledgerId, ticket) {
  const row = [
    ticket.id,
    ticket.createdAt,
    ticket.customerName,
    ticket.tailor,
    ticket.salesperson,
    ticket.dueDate,
    ticket.rush ? "Yes" : "",
    ticket.balance,
    ticket.status,
    ticket.docxFilename,
    ticket.driveLink,
    JSON.stringify(ticket.formState || {}),
    ticket.createdAt,
    ticket.driveFileId || "",
    "",
    ticket.ticketNumber || "",
  ];
  const rowNumber = await writeRowAtEnd(token, ledgerId, SHEET_TAB, "P", row);
  return { rowNumber };
}

// Overwrites an existing ticket's row in place (used when re-saving a
// reopened ticket) instead of appending a new row. Ticket ID, Created At,
// and Status are left untouched — only the editable fields, the file
// reference, and Status Date (since the ticket was just altered) change.
export async function updateTicketRecord(token, ledgerId, rowNumber, ticket) {
  const now = new Date().toLocaleString();
  await sheetsApi(token, `${ledgerId}/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      valueInputOption: "USER_ENTERED",
      data: [
        {
          range: `${SHEET_TAB}!C${rowNumber}:H${rowNumber}`,
          values: [[ticket.customerName, ticket.tailor, ticket.salesperson, ticket.dueDate, ticket.rush ? "Yes" : "", ticket.balance]],
        },
        {
          range: `${SHEET_TAB}!J${rowNumber}:L${rowNumber}`,
          values: [[ticket.docxFilename, ticket.driveLink, JSON.stringify(ticket.formState || {})]],
        },
        { range: `${SHEET_TAB}!M${rowNumber}`, values: [[now]] },
        { range: `${SHEET_TAB}!N${rowNumber}`, values: [[ticket.driveFileId || ""]] },
      ],
    }),
  });
  return now;
}

export async function listTickets(token, ledgerId) {
  const data = await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${SHEET_TAB}!A2:P`)}`);
  const rows = data.values || [];
  const skippedRows = rows.filter((row) => row.some((cell) => String(cell).trim()) && !String(row[0] || "").trim()).length;
  const tickets = rows
    .map((row, index) => ({
      rowNumber: index + 2,
      id: row[0] || "",
      createdAt: row[1] || "",
      customerName: row[2] || "",
      tailor: row[3] || "",
      salesperson: row[4] || "",
      dueDate: row[5] || "",
      rush: row[6] === "Yes",
      balance: row[7] || "",
      status: row[8] || "Open",
      docxFilename: row[9] || "",
      driveLink: row[10] || "",
      formState: (() => {
        try {
          return JSON.parse(row[11] || "{}");
        } catch {
          return {};
        }
      })(),
      statusDate: row[12] || row[1] || "",
      driveFileId: row[13] || "",
      notes: row[14] || "",
      ticketNumber: row[15] || "",
    }))
    .filter((ticket) => ticket.id);
  tickets.skippedRows = skippedRows; // rows with data that don't start in column A, so they can't be read
  return tickets;
}

// Assigns each new ticket the next sequential number (existing max + 1).
// Not a strict atomic counter — in a small shop with a handful of staff,
// the odds of two saves landing in the exact same instant are low enough
// to accept, and the worst case is a rare duplicate number, not data loss.
export async function getNextTicketNumber(token, ledgerId) {
  const data = await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${SHEET_TAB}!P2:P`)}`);
  const rows = data.values || [];
  const highest = rows.reduce((max, row) => {
    const n = parseInt(row[0], 10);
    return Number.isFinite(n) && n > max ? n : max;
  }, 0);
  return highest + 1;
}

export async function updateTicketStatus(token, ledgerId, rowNumber, status) {
  const now = new Date().toLocaleString();
  await sheetsApi(token, `${ledgerId}/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      valueInputOption: "RAW",
      data: [
        { range: `${SHEET_TAB}!I${rowNumber}`, values: [[status]] },
        { range: `${SHEET_TAB}!M${rowNumber}`, values: [[now]] },
      ],
    }),
  });
  return now;
}

export async function updateTicketNotes(token, ledgerId, rowNumber, notes) {
  await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${SHEET_TAB}!O${rowNumber}`)}?valueInputOption=USER_ENTERED`, {
    method: "PUT",
    body: JSON.stringify({ values: [[notes]] }),
  });
}

async function getSheetGid(token, ledgerId, tabName = SHEET_TAB) {
  const data = await sheetsApi(token, `${ledgerId}?fields=sheets.properties`);
  const sheet = data.sheets?.find((s) => s.properties.title === tabName);
  return sheet?.properties.sheetId ?? null;
}

// rowNumber is the 1-indexed sheet row (matches listTickets' rowNumber).
// Deleting shifts every later row up by one, so callers should refresh
// their ticket list afterward rather than reusing cached row numbers.
export async function deleteTicketRow(token, ledgerId, rowNumber) {
  const sheetId = await getSheetGid(token, ledgerId, SHEET_TAB);
  if (sheetId === null) throw new Error("Could not find the Tickets sheet tab");
  await sheetsApi(token, `${ledgerId}:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          deleteDimension: {
            range: { sheetId, dimension: "ROWS", startIndex: rowNumber - 1, endIndex: rowNumber },
          },
        },
      ],
    }),
  });
}

/* ---------------------------------- studio -----------------------------------
   Garments that arrived for a client fitting but aren't (yet) an alteration
   ticket. Lives in its own tab of the same spreadsheet — deliberately not
   mixed into the Tickets tab, since the fields and lifecycle don't overlap
   (no tailor, no measurements, no saved document).
--------------------------------------------------------------------------- */
const STUDIO_SHEET_TAB = "Studio";
const STUDIO_HEADERS = [
  "Studio ID",
  "Created At",
  "Client Name",
  "Garment Description",
  "Arrival Date",
  "Salesperson",
  "Status",
  "Status Date",
  "Converted At",
];

async function ensureStudioTab(token, ledgerId) {
  const existingGid = await getSheetGid(token, ledgerId, STUDIO_SHEET_TAB);
  if (existingGid !== null) return;

  await sheetsApi(token, `${ledgerId}:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({ requests: [{ addSheet: { properties: { title: STUDIO_SHEET_TAB } } }] }),
  });

  await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${STUDIO_SHEET_TAB}!A1`)}?valueInputOption=RAW`, {
    method: "PUT",
    body: JSON.stringify({ values: [STUDIO_HEADERS] }),
  });
}

export async function getOrCreateStudioTab(token, ledgerId) {
  await ensureStudioTab(token, ledgerId);
  return ledgerId; // same spreadsheet — this just guarantees the tab exists
}

export async function appendStudioEntry(token, ledgerId, entry) {
  const row = [
    entry.id,
    entry.createdAt,
    entry.clientName,
    entry.garmentDescription,
    entry.arrivalDate,
    entry.salesperson,
    entry.status,
    entry.createdAt,
    "",
  ];
  await writeRowAtEnd(token, ledgerId, STUDIO_SHEET_TAB, "I", row);
}

export async function listStudioEntries(token, ledgerId) {
  const data = await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${STUDIO_SHEET_TAB}!A2:I`)}`);
  const rows = data.values || [];
  return rows
    .map((row, index) => ({
      rowNumber: index + 2,
      id: row[0] || "",
      createdAt: row[1] || "",
      clientName: row[2] || "",
      garmentDescription: row[3] || "",
      arrivalDate: row[4] || "",
      salesperson: row[5] || "",
      status: row[6] || "In Studio",
      statusDate: row[7] || row[1] || "",
      convertedAt: row[8] || "",
    }))
    .filter((entry) => entry.id);
}

export async function updateStudioStatus(token, ledgerId, rowNumber, status) {
  const now = new Date().toLocaleString();
  await sheetsApi(token, `${ledgerId}/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      valueInputOption: "RAW",
      data: [
        { range: `${STUDIO_SHEET_TAB}!G${rowNumber}`, values: [[status]] },
        { range: `${STUDIO_SHEET_TAB}!H${rowNumber}`, values: [[now]] },
      ],
    }),
  });
  return now;
}

export async function updateStudioEntry(token, ledgerId, rowNumber, entry) {
  const now = new Date().toLocaleString();
  await sheetsApi(token, `${ledgerId}/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      valueInputOption: "USER_ENTERED",
      data: [
        {
          range: `${STUDIO_SHEET_TAB}!C${rowNumber}:F${rowNumber}`,
          values: [[entry.clientName, entry.garmentDescription, entry.arrivalDate, entry.salesperson]],
        },
        { range: `${STUDIO_SHEET_TAB}!H${rowNumber}`, values: [[now]] },
      ],
    }),
  });
  return now;
}

// Marks a studio entry as converted (once its garment has become an actual
// alteration ticket) and advances it to Complete, since its job in the
// studio log is done — tracking continues on the ticket from here.
export async function markStudioConverted(token, ledgerId, rowNumber) {
  const now = new Date().toLocaleString();
  await sheetsApi(token, `${ledgerId}/values:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      valueInputOption: "RAW",
      data: [
        { range: `${STUDIO_SHEET_TAB}!G${rowNumber}`, values: [["Completed"]] },
        { range: `${STUDIO_SHEET_TAB}!H${rowNumber}`, values: [[now]] },
        { range: `${STUDIO_SHEET_TAB}!I${rowNumber}`, values: [[now]] },
      ],
    }),
  });
  return now;
}

export async function deleteStudioEntry(token, ledgerId, rowNumber) {
  const sheetId = await getSheetGid(token, ledgerId, STUDIO_SHEET_TAB);
  if (sheetId === null) throw new Error("Could not find the Studio sheet tab");
  await sheetsApi(token, `${ledgerId}:batchUpdate`, {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          deleteDimension: {
            range: { sheetId, dimension: "ROWS", startIndex: rowNumber - 1, endIndex: rowNumber },
          },
        },
      ],
    }),
  });
}
