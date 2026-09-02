import { DRIVE_FOLDER_ID, DRIVE_ROOT_FOLDER_NAME } from "./config.js";

const DRIVE_FOLDER_MIME = "application/vnd.google-apps.folder";

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

export async function fetchDriveFolders(token, parentId, { query = "", onProgress, onBatch } = {}) {
  const baseQuery = `'${parentId}' in parents and mimeType = '${DRIVE_FOLDER_MIME}' and trashed = false`;
  const q = query ? `${baseQuery} and name contains '${query.replace(/[\\']/g, "\\$&")}'` : baseQuery;

  const files = [];
  let pageToken = "";

  do {
    const params = new URLSearchParams({
      includeItemsFromAllDrives: "true",
      supportsAllDrives: "true",
      orderBy: "name_natural",
      pageSize: "1000",
      fields: "nextPageToken, files(id,name)",
      q,
    });
    if (pageToken) params.set("pageToken", pageToken);

    const response = await fetch(`https://www.googleapis.com/drive/v3/files?${params.toString()}`, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      if (response.status === 401) throw new Error("Drive session expired. Please sign in again.");
      throw new Error("Could not load Google Drive folders");
    }

    const data = await response.json();
    files.push(...(data.files || []));
    pageToken = data.nextPageToken || "";
    onProgress?.(files.length);
    onBatch?.(files, Boolean(pageToken)); // (results so far, whether more are still loading)
  } while (pageToken);

  return files;
}

export async function createDriveFolder(token, parentId, name) {
  const response = await fetch("https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id,name", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ name, mimeType: DRIVE_FOLDER_MIME, parents: [parentId] }),
  });

  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please sign in again.");
    const detail = await response.json().catch(() => null);
    throw new Error(detail?.error?.message || "Could not create the folder");
  }

  return response.json();
}

const RENDER_CAP = 200;

// Caches a folder's full subfolder list for the rest of this browser
// session, so reopening the picker or navigating back to a folder you've
// already visited is instant instead of re-fetching from Drive.
const folderListSessionCache = new Map();

export function chooseDriveFolder(token, dom) {
  const { modal, path, search, list, upBtn, cancelBtn, chooseBtn, newFolderInput, newFolderBtn } = dom;

  return new Promise((resolve, reject) => {
    const stack = [{ id: DRIVE_FOLDER_ID, name: DRIVE_ROOT_FOLDER_NAME }];
    let settled = false;
    let handleFolderClick = null;
    let allFolders = [];
    let loadToken = 0;

    const cleanup = () => {
      cancelBtn.removeEventListener("click", handleCancel);
      chooseBtn.removeEventListener("click", handleChoose);
      upBtn.removeEventListener("click", handleBack);
      search.removeEventListener("input", handleSearchInput);
      newFolderBtn.removeEventListener("click", handleCreateFolder);
      if (handleFolderClick) list.removeEventListener("click", handleFolderClick);
      modal.hidden = true;
    };

    const finish = (value) => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(value);
    };

    const fail = (err) => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(err);
    };

    const currentFolder = () => stack[stack.length - 1];

    const renderFolderRows = (folders, totalCount, stillLoading) => {
      if (!folders.length) {
        list.innerHTML = stillLoading
          ? '<p class="drive-folder-empty">Loading folders…</p>'
          : '<p class="drive-folder-empty">No folders found here.</p>';
        return;
      }
      const rows = folders
        .slice(0, RENDER_CAP)
        .map(
          (item) =>
            `<button type="button" class="drive-folder-row" data-folder-id="${escapeHtml(item.id)}" data-folder-name="${escapeHtml(item.name)}">${escapeHtml(item.name)}</button>`,
        )
        .join("");
      const hint = stillLoading
        ? `<p class="drive-folder-hint">${totalCount.toLocaleString()} so far, still loading…</p>`
        : totalCount > RENDER_CAP
          ? `<p class="drive-folder-hint">Showing ${RENDER_CAP} of ${totalCount.toLocaleString()} — keep typing to narrow it down.</p>`
          : "";
      list.innerHTML = rows + hint;
    };

    const applyFilter = (stillLoading = false) => {
      const term = search.value.trim().toLowerCase();
      const matches = term ? allFolders.filter((item) => item.name.toLowerCase().includes(term)) : allFolders;
      renderFolderRows(matches, matches.length, stillLoading);
    };

    // Loads every subfolder of the current folder (paginating past Drive's
    // 1000-per-page limit as needed). Renders progressively after each page
    // rather than waiting for the whole (possibly thousands-long) list —
    // the picker becomes usable after ~1 page instead of the full load.
    const renderFolders = async () => {
      const folder = currentFolder();
      const myLoad = ++loadToken;
      path.textContent = stack.map((item) => item.name).join(" / ");
      search.value = "";
      upBtn.disabled = stack.length === 1;

      const cached = folderListSessionCache.get(folder.id);
      if (cached) {
        allFolders = cached;
        applyFilter();
        return;
      }

      list.innerHTML = '<p class="drive-folder-empty">Loading folders…</p>';
      try {
        const folders = await fetchDriveFolders(token, folder.id, {
          onBatch: (soFar, stillLoading) => {
            if (myLoad !== loadToken) return;
            allFolders = soFar;
            applyFilter(stillLoading); // respects a search term typed while this was loading
          },
        });
        if (myLoad !== loadToken) return; // superseded by a newer navigation
        allFolders = folders;
        folderListSessionCache.set(folder.id, folders);
        applyFilter();
      } catch (err) {
        fail(err);
      }
    };

    const handleCancel = () => finish(null);
    const handleChoose = () => finish(currentFolder());
    const handleSearchInput = () => applyFilter();

    const handleBack = () => {
      if (stack.length > 1) {
        stack.pop();
        renderFolders();
      }
    };

    const handleCreateFolder = async () => {
      const name = newFolderInput.value.trim();
      if (!name) {
        newFolderInput.focus();
        return;
      }
      newFolderBtn.disabled = true;
      try {
        const parentId = currentFolder().id;
        const created = await createDriveFolder(token, parentId, name);
        folderListSessionCache.delete(parentId);
        newFolderInput.value = "";
        stack.push({ id: created.id, name: created.name });
        await renderFolders();
      } catch (err) {
        fail(err);
      } finally {
        newFolderBtn.disabled = false;
      }
    };

    handleFolderClick = (event) => {
      const target = event.target.closest(".drive-folder-row");
      if (!target || settled) return;
      stack.push({ id: target.dataset.folderId, name: target.dataset.folderName });
      renderFolders();
    };

    list.addEventListener("click", handleFolderClick);
    search.addEventListener("input", handleSearchInput);
    cancelBtn.addEventListener("click", handleCancel);
    chooseBtn.addEventListener("click", handleChoose);
    upBtn.addEventListener("click", handleBack);
    newFolderBtn.addEventListener("click", handleCreateFolder);

    modal.hidden = false;
    renderFolders();
  });
}

function buildMultipartBlob({ filename, blob, mimeType, folderId = "" }) {
  const boundary = `boundary_${Math.random().toString(36).slice(2)}`;
  const metadata = folderId ? { name: filename, parents: [folderId] } : { name: filename };
  const body = new Blob(
    [
      `--${boundary}\r\n`,
      "Content-Type: application/json; charset=UTF-8\r\n\r\n",
      `${JSON.stringify(metadata)}\r\n`,
      `--${boundary}\r\n`,
      `Content-Type: ${mimeType}\r\n\r\n`,
      blob,
      `\r\n--${boundary}--`,
    ],
    { type: `multipart/related; boundary=${boundary}` },
  );
  return { body, boundary };
}

export async function uploadFileToDrive(token, { filename, blob, mimeType, folderId }) {
  const { body, boundary } = buildMultipartBlob({ filename, blob, mimeType, folderId });
  const response = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink",
    {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": `multipart/related; boundary=${boundary}` },
      body,
    },
  );

  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please save again.");
    throw new Error("Upload failed");
  }
  return response.json();
}

// Overwrites an existing file's content and name in place (used when
// re-saving a reopened ticket), rather than creating a new file.
export async function updateFileInDrive(token, fileId, { filename, blob, mimeType }) {
  const { body, boundary } = buildMultipartBlob({ filename, blob, mimeType });
  const response = await fetch(
    `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink`,
    {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": `multipart/related; boundary=${boundary}` },
      body,
    },
  );

  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please save again.");
    if (response.status === 404) throw new Error("The original file could not be found — it may have been moved or deleted.");
    throw new Error("Update failed");
  }
  return response.json();
}

// Moves a file to Drive's trash (recoverable), rather than permanently
// deleting it, so an accidental ticket deletion isn't unrecoverable.
export async function trashDriveFile(token, fileId) {
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?supportsAllDrives=true`, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ trashed: true }),
  });

  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please sign in again.");
    if (response.status === 404) return; // already gone — nothing to do
    throw new Error("Could not remove the file from Drive");
  }
}
