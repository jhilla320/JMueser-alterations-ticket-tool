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

export async function fetchDriveFolders(token, parentId) {
  const params = new URLSearchParams({
    includeItemsFromAllDrives: "true",
    supportsAllDrives: "true",
    orderBy: "name_natural",
    pageSize: "100",
    fields: "files(id,name)",
    q: `'${parentId}' in parents and mimeType = '${DRIVE_FOLDER_MIME}' and trashed = false`,
  });

  const response = await fetch(`https://www.googleapis.com/drive/v3/files?${params.toString()}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    if (response.status === 401) throw new Error("Drive session expired. Please sign in again.");
    throw new Error("Could not load Google Drive folders");
  }

  const data = await response.json();
  return data.files || [];
}

export function chooseDriveFolder(token, dom) {
  const { modal, path, search, list, upBtn, cancelBtn, chooseBtn } = dom;

  return new Promise((resolve, reject) => {
    const stack = [{ id: DRIVE_FOLDER_ID, name: DRIVE_ROOT_FOLDER_NAME }];
    let settled = false;
    let handleFolderClick = null;
    let currentFolders = [];

    const cleanup = () => {
      cancelBtn.removeEventListener("click", handleCancel);
      chooseBtn.removeEventListener("click", handleChoose);
      upBtn.removeEventListener("click", handleBack);
      search.removeEventListener("input", handleSearch);
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

    const renderFolderRows = (folders) => {
      const query = search.value.trim().toLowerCase();
      const visibleFolders = query ? folders.filter((item) => item.name.toLowerCase().includes(query)) : folders;

      if (!folders.length) {
        list.innerHTML = '<p class="drive-folder-empty">No folders inside this folder.</p>';
        return;
      }
      if (!visibleFolders.length) {
        list.innerHTML = '<p class="drive-folder-empty">No matching folders here.</p>';
        return;
      }
      list.innerHTML = visibleFolders
        .map(
          (item) =>
            `<button type="button" class="drive-folder-row" data-folder-id="${escapeHtml(item.id)}" data-folder-name="${escapeHtml(item.name)}">${escapeHtml(item.name)}</button>`,
        )
        .join("");
    };

    const renderFolders = async () => {
      const folder = currentFolder();
      path.textContent = stack.map((item) => item.name).join(" / ");
      search.value = "";
      list.innerHTML = '<p class="drive-folder-empty">Loading folders…</p>';
      upBtn.disabled = stack.length === 1;

      try {
        currentFolders = await fetchDriveFolders(token, folder.id);
        renderFolderRows(currentFolders);
      } catch (err) {
        fail(err);
      }
    };

    const handleCancel = () => finish(null);
    const handleChoose = () => finish(currentFolder());
    const handleSearch = () => renderFolderRows(currentFolders);
    const handleBack = () => {
      if (stack.length > 1) {
        stack.pop();
        renderFolders();
      }
    };

    handleFolderClick = (event) => {
      const target = event.target.closest(".drive-folder-row");
      if (!target || settled) return;
      stack.push({ id: target.dataset.folderId, name: target.dataset.folderName });
      renderFolders();
    };

    list.addEventListener("click", handleFolderClick);
    search.addEventListener("input", handleSearch);
    cancelBtn.addEventListener("click", handleCancel);
    chooseBtn.addEventListener("click", handleChoose);
    upBtn.addEventListener("click", handleBack);

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
