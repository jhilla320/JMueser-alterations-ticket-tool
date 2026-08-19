// ---------------------------------------------------------------------------
// Shop configuration. Edit the values in this file to match your setup.
// Nothing in here is a secret — Google OAuth client IDs and Drive folder IDs
// are meant to be public; access is controlled by Google sign-in + Drive
// sharing permissions, not by hiding these values.
// ---------------------------------------------------------------------------

export const GOOGLE_CLIENT_ID = "617892178220-84fg83gdjhjssb3et6e5ufjnkb8cn1v2.apps.googleusercontent.com";

// The Drive folder tickets (and the ticket log spreadsheet) are saved into.
export const DRIVE_FOLDER_ID = "16r8SU6V02rAzgKVZvO6Ixc23ZT9eldsn";
export const DRIVE_ROOT_FOLDER_NAME = "Shared Drive";

// --- Sign-in restriction ---------------------------------------------------
// Only Google accounts matching ALLOWED_EMAIL_DOMAIN and/or ALLOWED_EMAILS
// will be let past the sign-in screen. Leave both empty to allow any Google
// account to sign in (not recommended for a shop tool).
//
// Examples:
//   ALLOWED_EMAIL_DOMAIN = "jmueser.com"   -> anyone @jmueser.com
//   ALLOWED_EMAILS = ["luis@gmail.com"]     -> specific personal accounts too
export const ALLOWED_EMAIL_DOMAIN = "jmueser.com";
export const ALLOWED_EMAILS = [];

// --- Ticket log (Google Sheet used as the shared ticket record) -----------
export const TICKET_LOG_FILE_NAME = "Alterations Ticket Log";
export const TICKET_STATUSES = [
  "In Progress",
  "Inbound",
  "ATTENTION",
  "Alts Finished",
  "Filed",
  "Client Contacted",
  "Completed",
];

export const STUDIO_STATUSES = ["In Studio", "Client Contact", "Fitting Scheduled", "Completed"];

// --- Form options ------------------------------------------------------
export const SALESPEOPLE = [
  "Bodhi", "Chase", "Chris", "Christian", "Colin", "Edris",
  "Frank", "Jake", "Jeff", "Jonas", "Mattia", "Nor", "Ryder",
];

// OAuth scopes: Drive (create/read files this app makes) + read-only
// metadata (for the folder picker) + email (to check ALLOWED_EMAIL_DOMAIN)
// + full Sheets access. That last one matters: drive.file only grants
// access to files *this app itself created* under the signed-in account —
// it does NOT extend to a shared spreadsheet someone else's session made,
// even if the current user has full Drive-level access to it. Without a
// separate spreadsheets scope, only whoever's account created the ticket
// log could actually read/write it — everyone else gets a 404 (Google
// deliberately returns "not found" rather than "forbidden" for
// unauthorized-scope file access, to avoid confirming a file exists).
export const DRIVE_SCOPE = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/drive.metadata.readonly",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/spreadsheets",
].join(" ");
