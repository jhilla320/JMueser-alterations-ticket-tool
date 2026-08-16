# J. Mueser Alterations Ticket Tool

A static web app for a bespoke tailoring shop to create alteration tickets, log them, and track garments awaiting a client fitting. No backend server — it's a plain static site (GitHub Pages) that talks directly to Google APIs from the browser.

## Stack & architecture

- Vanilla JS (ES modules, no bundler, no framework), plain HTML/CSS. Files load directly via `<script type="module">` — no build step.
- Auth: Google Identity Services, client-side OAuth. Sign-in is gated to an allowed email domain/list (see `config.js`).
- "Database": a Google Sheet, created/found automatically in a shared Drive folder. Two tabs:
  - **Tickets** — one row per alteration ticket.
  - **Studio** — a separate, simpler log for garments that arrived for a client fitting but aren't an alteration ticket (yet).
- Documents: each ticket's printed/saved output is a hand-built `.docx` (raw OOXML via JSZip — no server-side doc generation) stored in Google Drive.
- Deployment: push to `main`, GitHub Pages serves it. Currently **no custom domain configured** (DNS for `alterations.jmueser.com` was never finished — check current state before assuming it's live there vs. the default `github.io` URL).

## File map

- `index.html` — all markup, three views (Create Ticket / Ticket Log / Studio) toggled via JS, not routing.
- `app.js` — all UI logic and state. This is the biggest file by far.
- `styles.css` — design system: CSS custom properties for colors/fonts at the top, then component styles. Look there before adding new colors/fonts.
- `config.js` — shop-specific settings (OAuth client ID, Drive folder ID, allowed sign-in domain, status lists, salesperson list). Edit this, not the code, for shop-level config changes.
- `google-auth.js` — token handling + domain allow-list check.
- `drive.js` — folder picker (paginated), file upload/update/trash, folder creation.
- `sheets.js` — all Google Sheets reads/writes for both the Tickets and Studio tabs.
- `docx.js` — builds the `.docx` file from structured ticket data (not from HTML).

## Data model — the important part

A ticket's garments are **one flat array** (`garmentItems` in `app.js`), not separate jacket/trouser/shirt lists. Each item has a `garmentType` field ("jacket" | "trouser" | "shirt") and carries *all* possible measurement fields regardless of type — switching an item's type just changes which fields the UI shows/reads, it doesn't change the object shape.

**Backward compatibility matters here.** This app went through two earlier data shapes before landing on the current one:
1. Original: three separate `jackets` / `trousers` / `shirts` arrays.
2. Original: separate `size` and `description` fields per garment (now merged into one `sizeDescription` field).

Real tickets already saved in the production Sheet may be in either old shape. `normalizeGarmentItems()` and `migrateSizeDescription()` in `app.js` handle reading old data into the current shape on the fly. **Do not remove these** without confirming no old data depends on them — they're load-bearing for every ticket saved before this migration existed.

## Editing an existing ticket doesn't create a duplicate

Re-saving a ticket you opened via "Edit" updates the same Sheet row and the same Drive file in place (tracked via `editingTicket` state in `app.js`). This was a deliberate fix for an earlier bug where every re-save created a duplicate ticket and a duplicate file. If you touch the save flow, preserve this — check `editingTicket` before deciding whether to append a new row/file or update an existing one.

## Print mechanism (fragile — read before touching)

Printing (from anywhere — the live form, or a row in the Ticket Log) works by injecting the ticket's HTML into a single `#printArea` element that only physically exists inside the "Create Ticket" view. The print stylesheet (`@media print` in `styles.css`) force-shows that view and force-hides the others via ID selectors, **regardless of which view is currently active on screen**. If you add a new top-level view, you must add it to that hide-list in the print media query, or it'll bleed into printed output.

## Known quirks encountered building this (worth knowing before you "fix" something that looks like a bug)

- Status color-coding CSS is keyed on exact status string values (`data-status="..."`). If you rename or add a status in `config.js`, add a matching CSS rule or it'll render with no color.
- Mobile Safari has a real bug where an absolutely-positioned dropdown panel inside a `flex-direction: column` container with `align-items: stretch` gets miscounted into that container's height, even while hidden. Fixed once already (`align-items: flex-start` + explicit `flex-basis` scoping) — if a mysterious gap reappears on mobile, look there first.
- Sheets writes to specific columns use hardcoded column letters (e.g. `C${rowNumber}:H${rowNumber}`) in `sheets.js`. If a column is ever inserted/reordered, multiple spots need updating — there's no schema abstraction.
- OAuth scope is just `drive.file` + `drive.metadata.readonly` + `userinfo.email` — no dedicated Sheets scope, because Sheets API access to files the app itself created is covered by `drive.file`.

## Deployment

No CI/CD — changes are committed and pushed to `main` manually, and GitHub Pages redeploys automatically after a push. This project doesn't yet have git history from within Claude Code, so check `git log` / `git status` to get oriented before making changes.
