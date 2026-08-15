import { DRIVE_FOLDER_ID, TICKET_LOG_FILE_NAME } from "./config.js";

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
    pageSize: "1",
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
  await driveApi(token, `files/${sheetId}?addParents=${DRIVE_FOLDER_ID}&supportsAllDrives=true&fields=id,parents`, {
    method: "PATCH",
  });

  await sheetsApi(token, `${sheetId}/values/${encodeURIComponent(`${SHEET_TAB}!A1`)}?valueInputOption=RAW`, {
    method: "PUT",
    body: JSON.stringify({ values: [HEADERS] }),
  });

  return sheetId;
}

export async function getOrCreateLedger(token) {
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
  ];
  await sheetsApi(
    token,
    `${ledgerId}/values/${encodeURIComponent(`${SHEET_TAB}!A:M`)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    { method: "POST", body: JSON.stringify({ values: [row] }) },
  );
}

export async function listTickets(token, ledgerId) {
  const data = await sheetsApi(token, `${ledgerId}/values/${encodeURIComponent(`${SHEET_TAB}!A2:M`)}`);
  const rows = data.values || [];
  return rows
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
    }))
    .filter((ticket) => ticket.id);
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
