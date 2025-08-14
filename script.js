(function () {
  "use strict";

  // Elements
  const dateInput = document.getElementById("dateInput");
  const dayDisplay = document.getElementById("dayDisplay");
  const notesInput = document.getElementById("notesInput");
  const entryForm = document.getElementById("entryForm");
  const currentEditIdInput = document.getElementById("currentEditId");
  const saveBtn = document.getElementById("saveBtn");
  const cancelEditBtn = document.getElementById("cancelEditBtn");
  const clearFormBtn = document.getElementById("clearFormBtn");

  const searchInput = document.getElementById("searchInput");
  const fromDateInput = document.getElementById("fromDate");
  const toDateInput = document.getElementById("toDate");
  const clearFiltersBtn = document.getElementById("clearFiltersBtn");

  const exportBtn = document.getElementById("exportBtn");
  const importBtn = document.getElementById("importBtn");
  const importInput = document.getElementById("importInput");
  const printBtn = document.getElementById("printBtn");

  const signinBtn = document.getElementById("signinBtn");
  const signoutBtn = document.getElementById("signoutBtn");
  const syncBtn = document.getElementById("syncBtn");
  const cloudStatus = document.getElementById("cloudStatus");

  const pickLocalBtn = document.getElementById("pickLocalBtn");
  const saveLocalBtn = document.getElementById("saveLocalBtn");

  const entriesList = document.getElementById("entriesList");
  const emptyState = document.getElementById("emptyState");
  const resultsSummary = document.getElementById("resultsSummary");

  // Advanced filters
  const qrToday = document.getElementById("qrToday");
  const qr7 = document.getElementById("qr7");
  const qrThisMonth = document.getElementById("qrThisMonth");
  const qrLastMonth = document.getElementById("qrLastMonth");
  const qrAll = document.getElementById("qrAll");
  const weekdayFilters = document.getElementById("weekdayFilters");
  const sortSelect = document.getElementById("sortSelect");

  // Viewer modal elements
  const viewerModal = document.getElementById("viewerModal");
  const viewerBackdrop = document.getElementById("viewerBackdrop");
  const viewerCloseBtn = document.getElementById("viewerCloseBtn");
  const viewerTitle = document.getElementById("viewerTitle");
  const viewerMeta = document.getElementById("viewerMeta");
  const viewerNotes = document.getElementById("viewerNotes");
  const viewerCopyBtn = document.getElementById("viewerCopyBtn");
  const viewerPrintBtn = document.getElementById("viewerPrintBtn");

  const STORAGE_KEY = "dailyWorkEntries";

  // OneDrive / Graph setup (can be ignored if using local Excel only)
  const MSAL_CLIENT_ID = "YOUR_AZURE_APP_CLIENT_ID"; // optional
  const MSAL_AUTHORITY = "https://login.microsoftonline.com/common";
  const GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"];
  const FILE_NAME = "DailyWorkTracker.csv"; // OneDrive CSV name

  let msalInstance = null;
  let account = null;

  function setCloudStatus(text, type) {
    if (!cloudStatus) return;
    cloudStatus.textContent = text || "";
    cloudStatus.style.color = type === "error" ? "#ff6b6b" : "#a8b0c2";
  }

  function setAuthUi(isSignedIn) {
    if (!signinBtn || !signoutBtn || !syncBtn) return;
    signinBtn.hidden = !!isSignedIn;
    signoutBtn.hidden = !isSignedIn;
    syncBtn.hidden = !isSignedIn;
  }

  function initMsal() {
    if (!window.msal) return;
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: MSAL_CLIENT_ID,
        authority: MSAL_AUTHORITY,
        redirectUri: window.location.origin === "null" || window.location.origin.startsWith("file:")
          ? window.location.href
          : window.location.origin
      },
      cache: { cacheLocation: "localStorage" }
    });

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      account = accounts[0];
      setAuthUi(true);
      setCloudStatus(`Signed in as ${account.username}`);
    } else {
      setAuthUi(false);
    }

    msalInstance.handleRedirectPromise().catch(err => {
      console.error(err);
      setCloudStatus("Sign-in failed.", "error");
    });
  }

  async function signIn() {
    if (!msalInstance) return alert("MSAL not loaded");
    try {
      const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
      account = result.account;
      setAuthUi(true);
      setCloudStatus(`Signed in as ${account.username}`);
    } catch (e) {
      console.error(e);
      setCloudStatus("Sign-in canceled or failed.", "error");
    }
  }

  function signOut() {
    if (!msalInstance || !account) return;
    msalInstance.logoutPopup({ account }).finally(() => {
      account = null;
      setAuthUi(false);
      setCloudStatus("Signed out.");
    });
  }

  async function getToken() {
    if (!msalInstance || !account) throw new Error("Not signed in");
    try {
      const res = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
      return res.accessToken;
    } catch (e) {
      const res = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      return res.accessToken;
    }
  }

  // Local Excel (XLS - legacy) using File System Access + SheetJS
  const hasFS = "showSaveFilePicker" in window;
  let localExcelHandle = null;

  function setLocalUi() {
    if (!pickLocalBtn || !saveLocalBtn) return;
    if (!hasFS) {
      pickLocalBtn.hidden = true; // No FS Access API; use download fallback
      saveLocalBtn.textContent = "Download Excel (.xls)";
      return;
    }
  }

  function entriesToWorkbook(entries) {
    const headerOrder = ["dateISO", "day", "notes", "lastUpdated", "id"];
    const normalized = entries.map(e => ({
      dateISO: e.dateISO,
      day: e.day,
      notes: e.notes,
      lastUpdated: e.lastUpdated,
      id: e.id
    }));
    const ws = XLSX.utils.json_to_sheet(normalized, { header: headerOrder });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DailyWork");
    return wb;
  }

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }

  function workbookToBlob(wb, bookType) {
    try {
      const data = XLSX.write(wb, { bookType, type: "array" });
      return new Blob([data], { type: bookType === "xls" ? "application/vnd.ms-excel" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    } catch (e) {
      // Fallback to binary string if array not supported for this type
      const bin = XLSX.write(wb, { bookType, type: "binary" });
      const ab = s2ab(bin);
      return new Blob([ab], { type: bookType === "xls" ? "application/vnd.ms-excel" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }
  }

  async function ensureWritePermission(handle) {
    try {
      const q = await handle.queryPermission({ mode: "readwrite" });
      if (q === "granted") return true;
      const r = await handle.requestPermission({ mode: "readwrite" });
      return r === "granted";
    } catch (_) {
      return false;
    }
  }

  async function writeWorkbookToHandle(handle, entries) {
    const ok = await ensureWritePermission(handle);
    if (!ok) throw new Error("Permission to write the selected Excel file was not granted.");
    const wb = entriesToWorkbook(entries);
    const blob = workbookToBlob(wb, "xls");
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  // Persisting the chosen file handle in IndexedDB for auto-reuse
  const IDB_DB = "dwt-db";
  const IDB_STORE = "handles";
  function idbOpen() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(IDB_DB, 1);
      req.onupgradeneeded = () => {
        req.result.createObjectStore(IDB_STORE);
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function idbPut(key, value) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.objectStore(IDB_STORE).put(value, key);
    });
  }
  async function idbGet(key) {
    const db = await idbOpen();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readonly");
      const req = tx.objectStore(IDB_STORE).get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function restoreSavedExcelHandle() {
    try {
      const handle = await idbGet("excelHandle");
      if (!handle) return;
      const ok = await ensureWritePermission(handle);
      if (ok) {
        localExcelHandle = handle;
        setCloudStatus("Excel file reconnected. Auto-save enabled.");
      }
    } catch (_) {}
  }

  async function pickLocalWorkbook() {
    try {
      const suggestedName = "DailyWorkTracker.xls";
      localExcelHandle = await window.showSaveFilePicker({
        suggestedName,
        types: [{
          description: "Excel 97-2003 Workbook (*.xls)",
          accept: { "application/vnd.ms-excel": [".xls"] }
        }]
      });
      const ok = await ensureWritePermission(localExcelHandle);
      if (!ok) {
        setCloudStatus("Permission to write this file was denied.", "error");
        return;
      }
      try { await idbPut("excelHandle", localExcelHandle); } catch (_) {}
      setCloudStatus("Excel file selected. Changes will save to it.");
    } catch (e) {
      if (e && e.name === "AbortError") return; // user canceled
      console.error(e);
      setCloudStatus("Could not select file.", "error");
    }
  }

  async function saveLocalWorkbookNow() {
    const entries = loadEntries().sort((a, b) => a.dateISO.localeCompare(b.dateISO));
    if (hasFS) {
      try {
        if (!localExcelHandle) {
          await pickLocalWorkbook();
          if (!localExcelHandle) return;
        }
        await writeWorkbookToHandle(localExcelHandle, entries);
        setCloudStatus(`Saved to local Excel at ${new Date().toLocaleTimeString()}`);
      } catch (e) {
        console.error(e);
        const msg = e && e.message ? e.message : String(e);
        const hint = msg && msg.toLowerCase().includes("busy") || msg.toLowerCase().includes("locked")
          ? " (Close the Excel file if it is open, then try again.)"
          : "";
        setCloudStatus("Save to Excel failed: " + msg + hint, "error");
      }
    } else {
      // Fallback: download XLS file
      const wb = entriesToWorkbook(entries);
      const blob = workbookToBlob(wb, "xls");
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `DailyWorkTracker-${new Date().toISOString().slice(0,10)}.xls`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
      setCloudStatus("Excel (.xls) downloaded.");
    }
  }

  // Local storage utilities
  function getDayNameFromDateString(dateString) {
    if (!dateString) return "";
    const date = new Date(dateString + "T00:00:00");
    return date.toLocaleDateString(undefined, { weekday: "long" });
  }

  function formatDisplayDate(dateString) {
    if (!dateString) return "";
    const date = new Date(dateString + "T00:00:00");
    return date.toLocaleDateString(undefined, {
      year: "numeric",
      month: "short",
      day: "2-digit",
      weekday: "short"
    });
  }

  function loadEntries() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed;
    } catch (err) {
      console.error("Failed to load entries:", err);
      return [];
    }
  }

  function saveEntries(entries) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(entries));
    } catch (err) {
      console.error("Failed to save entries:", err);
      alert("Saving failed. Try exporting your data and clearing some space.");
    }
  }

  function upsertEntry(entry) {
    const entries = loadEntries();
    const index = entries.findIndex(e => e.id === entry.id);
    if (index >= 0) {
      entries[index] = entry;
    } else {
      entries.push(entry);
    }
    saveEntries(entries);
    return entries;
  }

  function deleteEntry(id) {
    const entries = loadEntries();
    const filtered = entries.filter(e => e.id !== id);
    saveEntries(filtered);
    return filtered;
  }

  function resetForm() {
    currentEditIdInput.value = "";
    saveBtn.textContent = "Save entry";
    cancelEditBtn.hidden = true;
    notesInput.value = "";
    updateDayDisplay();
  }

  function startEdit(entry) {
    currentEditIdInput.value = entry.id;
    dateInput.value = entry.dateISO;
    notesInput.value = entry.notes;
    updateDayDisplay();
    saveBtn.textContent = "Update entry";
    cancelEditBtn.hidden = false;
    notesInput.focus();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function updateDayDisplay() {
    dayDisplay.value = getDayNameFromDateString(dateInput.value);
  }

  // Viewer
  function openViewer(entry) {
    viewerTitle.textContent = `${formatDisplayDate(entry.dateISO)} (${entry.day})`;
    viewerMeta.textContent = `Last updated ${new Date(entry.lastUpdated).toLocaleString()}`;
    viewerNotes.textContent = entry.notes;
    viewerModal.hidden = false;
  }
  function closeViewer() { viewerModal.hidden = true; }
  if (viewerCloseBtn) viewerCloseBtn.addEventListener("click", closeViewer);
  if (viewerBackdrop) viewerBackdrop.addEventListener("click", closeViewer);
  window.addEventListener("keydown", (e) => { if (!viewerModal.hidden && e.key === "Escape") closeViewer(); });
  if (viewerCopyBtn) viewerCopyBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(`${viewerTitle.textContent}\n${viewerMeta.textContent}\n\n${viewerNotes.textContent}`);
      setCloudStatus("Copied entry to clipboard.");
    } catch (_) {
      setCloudStatus("Copy failed.", "error");
    }
  });
  if (viewerPrintBtn) viewerPrintBtn.addEventListener("click", () => window.print());

  // Rendering
  function renderEntries() {
    const query = (searchInput.value || "").trim();
    const queryLower = query.toLowerCase();
    const from = fromDateInput.value ? new Date(fromDateInput.value + "T00:00:00").getTime() : null;
    const to = toDateInput.value ? new Date(toDateInput.value + "T23:59:59").getTime() : null;

    // Weekday filter set
    const activeDays = new Set();
    if (weekdayFilters) {
      const toggles = weekdayFilters.querySelectorAll(".chip.toggle.active");
      toggles.forEach(t => activeDays.add(String(t.dataset.day || "")));
    }

    let entries = loadEntries();

    // Sort
    const sortMode = sortSelect ? sortSelect.value : "newest";
    entries = entries.sort((a, b) => {
      if (sortMode === "oldest") return a.dateISO.localeCompare(b.dateISO) || a.lastUpdated - b.lastUpdated;
      if (sortMode === "updated") return b.lastUpdated - a.lastUpdated;
      return b.dateISO.localeCompare(a.dateISO) || b.lastUpdated - a.lastUpdated; // newest
    });

    // Filter
    const filtered = entries.filter(e => {
      const time = new Date(e.dateISO + "T12:00:00").getTime();
      if (from !== null && time < from) return false;
      if (to !== null && time > to) return false;
      if (activeDays.size > 0) {
        const short = (e.day || "").slice(0,3);
        if (!activeDays.has(short)) return false;
      }
      if (!queryLower) return true;
      const hay = `${e.notes} ${e.dateISO} ${e.day}`.toLowerCase();
      return hay.includes(queryLower);
    });

    // Results summary
    if (resultsSummary) {
      const total = entries.length;
      const count = filtered.length;
      resultsSummary.textContent = count === total ? `${count} entries` : `${count} of ${total} entries`;
    }

    entriesList.innerHTML = "";

    if (filtered.length === 0) {
      emptyState.hidden = false;
      return;
    }

    emptyState.hidden = true;

    const highlight = (text) => {
      if (!query) return text;
      try {
        const safe = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const re = new RegExp(safe, "gi");
        return text.replace(re, (m) => `<mark>${m}</mark>`);
      } catch (_) {
        return text;
      }
    };

    for (const entry of filtered) {
      const card = document.createElement("article");
      card.className = "entry-card";

      const header = document.createElement("div");
      header.className = "entry-header";

      const left = document.createElement("div");
      left.className = "entry-left";

      const dayChip = document.createElement("span");
      const dayClass = (function(d){
        const k = d.toLowerCase();
        if (k.startsWith("mon")) return "w-monday";
        if (k.startsWith("tue")) return "w-tuesday";
        if (k.startsWith("wed")) return "w-wednesday";
        if (k.startsWith("thu")) return "w-thursday";
        if (k.startsWith("fri")) return "w-friday";
        if (k.startsWith("sat")) return "w-saturday";
        if (k.startsWith("sun")) return "w-sunday";
        return "";
      })(entry.day || "");
      dayChip.className = `day-chip ${dayClass}`.trim();
      dayChip.textContent = entry.day;

      const title = document.createElement("div");
      title.className = "entry-title";
      title.innerHTML = highlight(`${formatDisplayDate(entry.dateISO)} (${entry.day})`);

      left.appendChild(dayChip);
      left.appendChild(title);

      const meta = document.createElement("div");
      meta.className = "entry-meta";
      const updated = new Date(entry.lastUpdated);
      meta.textContent = `Last updated ${updated.toLocaleString()}`;

      header.appendChild(left);
      header.appendChild(meta);

      const notes = document.createElement("div");
      notes.className = "entry-notes";
      notes.innerHTML = highlight(entry.notes);

      const actions = document.createElement("div");
      actions.className = "entry-actions";

      const viewBtn = document.createElement("button");
      viewBtn.className = "btn";
      viewBtn.type = "button";
      viewBtn.textContent = "View";
      viewBtn.addEventListener("click", () => openViewer(entry));

      const editBtn = document.createElement("button");
      editBtn.className = "btn secondary";
      editBtn.type = "button";
      editBtn.textContent = "Edit";
      editBtn.addEventListener("click", () => startEdit(entry));

      const deleteBtn = document.createElement("button");
      deleteBtn.className = "btn ghost";
      deleteBtn.type = "button";
      deleteBtn.textContent = "Delete";
      deleteBtn.addEventListener("click", () => {
        const ok = confirm("Delete this entry? This cannot be undone.");
        if (!ok) return;
        deleteEntry(entry.id);
        renderEntries();
      });

      actions.appendChild(viewBtn);
      actions.appendChild(editBtn);
      actions.appendChild(deleteBtn);

      card.appendChild(header);
      card.appendChild(notes);
      card.appendChild(actions);

      entriesList.appendChild(card);
    }
  }

  // Quick ranges
  function setRange(fromISO, toISO) {
    fromDateInput.value = fromISO || "";
    toDateInput.value = toISO || "";
    renderEntries();
  }

  function startOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
  function endOfMonth(d) { return new Date(d.getFullYear(), d.getMonth()+1, 0); }

  if (qrToday) qrToday.addEventListener("click", () => {
    const t = new Date();
    const iso = t.toISOString().slice(0,10);
    setRange(iso, iso);
  });
  if (qr7) qr7.addEventListener("click", () => {
    const t = new Date();
    const to = t.toISOString().slice(0,10);
    const past = new Date(Date.now() - 6*24*60*60*1000).toISOString().slice(0,10);
    setRange(past, to);
  });
  if (qrThisMonth) qrThisMonth.addEventListener("click", () => {
    const now = new Date();
    const from = startOfMonth(now).toISOString().slice(0,10);
    const to = endOfMonth(now).toISOString().slice(0,10);
    setRange(from, to);
  });
  if (qrLastMonth) qrLastMonth.addEventListener("click", () => {
    const now = new Date();
    const prev = new Date(now.getFullYear(), now.getMonth()-1, 1);
    const from = startOfMonth(prev).toISOString().slice(0,10);
    const to = endOfMonth(prev).toISOString().slice(0,10);
    setRange(from, to);
  });
  if (qrAll) qrAll.addEventListener("click", () => setRange("", ""));

  // Weekday toggles
  if (weekdayFilters) {
    weekdayFilters.addEventListener("click", (e) => {
      const target = e.target;
      if (!(target instanceof HTMLElement)) return;
      if (!target.classList.contains("toggle")) return;
      target.classList.toggle("active");
      renderEntries();
    });
  }

  // Sort
  if (sortSelect) sortSelect.addEventListener("change", renderEntries);

  // Event bindings
  dateInput.addEventListener("change", updateDayDisplay);

  entryForm.addEventListener("submit", async function (e) {
    e.preventDefault();
    const dateISO = dateInput.value;
    const notes = (notesInput.value || "").trim();

    if (!dateISO) {
      alert("Please choose a date.");
      return;
    }
    if (!notes) {
      alert("Please enter your work details.");
      return;
    }

    const now = Date.now();
    const id = currentEditIdInput.value || `e_${now}`;

    const entry = {
      id,
      dateISO,
      day: getDayNameFromDateString(dateISO),
      notes,
      lastUpdated: now
    };

    upsertEntry(entry);
    renderEntries();
    resetForm();

    // Auto-save to local Excel if selected
    if (localExcelHandle) {
      try { await saveLocalWorkbookNow(); } catch (_) {}
    }
  });

  cancelEditBtn.addEventListener("click", function () {
    resetForm();
  });

  clearFormBtn.addEventListener("click", function () {
    notesInput.value = "";
    notesInput.focus();
  });

  searchInput.addEventListener("input", renderEntries);
  fromDateInput.addEventListener("change", renderEntries);
  toDateInput.addEventListener("change", renderEntries);
  clearFiltersBtn.addEventListener("click", function () {
    searchInput.value = "";
    fromDateInput.value = "";
    toDateInput.value = "";
    renderEntries();
  });

  exportBtn.addEventListener("click", function () {
    const entries = loadEntries().sort((a, b) => a.dateISO.localeCompare(b.dateISO) || a.lastUpdated - b.lastUpdated);
    const lines = [];
    lines.push(`Daily Work Tracker export - ${new Date().toLocaleString()}`);
    lines.push("");
    for (const e of entries) {
      lines.push(`Date: ${e.dateISO} (${e.day})`);
      lines.push("Notes:");
      String(e.notes || "").split(/\r?\n/).forEach(row => lines.push(row));
      lines.push(`Last updated: ${new Date(e.lastUpdated).toLocaleString()}`);
      lines.push("---");
      lines.push("");
    }
    const content = lines.join("\r\n");
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `daily-work-tracker-${new Date().toISOString().slice(0,10)}.txt`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  });

  importBtn.addEventListener("click", function () { importInput.click(); });
  importInput.addEventListener("change", function () {
    const file = importInput.files && importInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function () {
      try {
        const data = JSON.parse(String(reader.result || "[]"));
        if (!Array.isArray(data)) throw new Error("Invalid file format");
        const normalized = data.map(raw => ({
          id: String(raw.id || `e_${Date.now()}_${Math.random().toString(36).slice(2)}`),
          dateISO: String(raw.dateISO || "").slice(0, 10),
          day: String(raw.day || getDayNameFromDateString(String(raw.dateISO || ""))),
          notes: String(raw.notes || ""),
          lastUpdated: Number(raw.lastUpdated || Date.now())
        })).filter(x => x.dateISO && x.notes);

        saveEntries(normalized);
        renderEntries();
        alert("Import successful.");
      } catch (err) {
        console.error(err);
        alert("Import failed. Please select a valid JSON file exported from this app.");
      } finally {
        importInput.value = "";
      }
    };
    reader.readAsText(file);
  });

  printBtn.addEventListener("click", function () { window.print(); });

  // OneDrive events (optional)
  if (signinBtn) signinBtn.addEventListener("click", signIn);
  if (signoutBtn) signoutBtn.addEventListener("click", signOut);
  if (syncBtn) syncBtn.addEventListener("click", async () => {
    try {
      setCloudStatus("Syncing...");
      const token = await getToken();
      await (async function ensureFile(accessToken){
        const headers = { Authorization: `Bearer ${accessToken}` };
        const base = "https://graph.microsoft.com/v1.0/me/drive/root";
        const res = await fetch(`${base}/children/${encodeURIComponent(FILE_NAME)}` , { headers });
        if (res.ok) return;
        if (res.status !== 404) throw new Error(await res.text());
        const createRes = await fetch(`${base}/children/${encodeURIComponent(FILE_NAME)}/content`, {
          method: "PUT",
          headers: { ...headers, "Content-Type": "text/csv" },
          body: "dateISO,day,notes,lastUpdated,id\n"
        });
        if (!createRes.ok) throw new Error("Failed creating file");
      })(token);
      const entries = loadEntries().sort((a, b) => a.dateISO.localeCompare(b.dateISO));
      const header = ["dateISO","day","notes","lastUpdated","id"];
      const csv = [header.join(","), ...entries.map(e => header.map(k => {
        const s = String(e[k] ?? "");
        return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(","))].join("\n");
      const res2 = await fetch("https://graph.microsoft.com/v1.0/me/drive/root/children/" + encodeURIComponent(FILE_NAME) + "/content", {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "text/csv" },
        body: csv
      });
      if (!res2.ok) throw new Error("Upload failed");
      setCloudStatus(`Synced to OneDrive: ${FILE_NAME} at ${new Date().toLocaleTimeString()}`);
    } catch (e) {
      console.error(e);
      setCloudStatus("Sync failed.", "error");
    }
  });

  // Local Excel events
  if (pickLocalBtn) pickLocalBtn.addEventListener("click", pickLocalWorkbook);
  if (saveLocalBtn) saveLocalBtn.addEventListener("click", saveLocalWorkbookNow);

  // Init
  function initToday() {
    const todayISO = new Date().toISOString().slice(0, 10);
    dateInput.value = todayISO;
    updateDayDisplay();
  }

  initToday();
  renderEntries();
  initMsal();
  setLocalUi();
  if (navigator.storage && navigator.storage.persist) {
    try { navigator.storage.persist(); } catch (_) {}
  }
  restoreSavedExcelHandle();
})(); 