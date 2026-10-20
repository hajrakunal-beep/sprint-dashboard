"use strict";

(() => {
  const masterPicker = document.getElementById("master-excel-file");
  const masterStatus = document.getElementById("master-file-status");
  const adoInput = document.getElementById("ado-excel-file");
  const adoStatus = document.getElementById("ado-file-status");
  const mergeButton = document.getElementById("merge-data-button");
  const mergeStatus = document.getElementById("merge-data-status");
  const canSaveInPlace = typeof window.showOpenFilePicker === "function";
  const isAdoFile = file => Boolean(file && /\.(xlsx|xls|csv)$/i.test(file.name));
  let masterHandle = null;
  let adoSnapshot = null;
  let merging = false;
  let selecting = false;

  const isFileStateError = error => error.name === "NotReadableError"
    || /state cached in an interface object|state had changed since it was read from disk/i.test(error.message || "");

  async function readLatestMaster() {
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const file = await masterHandle.getFile();
        return { file, bytes: await file.arrayBuffer() };
      } catch (error) {
        if (!isFileStateError(error)) throw error;
        if (attempt === 1) throw new Error("The master file changed or could not be read. Save and close it in Excel, wait for any file sync to finish, then select the master again and retry.");
      }
    }
  }

  function updateReadyState() {
    mergeButton.disabled = merging || selecting || !masterHandle || !isAdoFile(adoInput.files?.[0]);
  }

  function showReadyStatus() {
    updateReadyState();
    mergeStatus.classList.remove("settings-file-error");
    mergeStatus.textContent = mergeButton.disabled
      ? "Select the master workbook and ADO dump to enable merging."
      : "Ready to merge and save directly to the selected master file.";
  }

  if (!canSaveInPlace) {
    masterPicker.disabled = true;
    masterStatus.textContent = "Direct file saving is unavailable. Open this dashboard in Chrome or Edge using HTTPS or a localhost server.";
    mergeStatus.textContent = "A browser with local file write access is required to update the master file.";
  }

  masterPicker.addEventListener("click", async () => {
    if (!canSaveInPlace || selecting || merging) return;
    selecting = true;
    masterPicker.disabled = true;
    updateReadyState();
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        excludeAcceptAllOption: true,
        types: [{ description: "Excel workbook", accept: {
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
        } }]
      });
      const file = await handle.getFile();
      if (!/\.xlsx$/i.test(file.name)) throw new Error("Select an .xlsx master to preserve formatting. Save .xls masters as .xlsx in Excel first.");
      masterHandle = handle;
      masterStatus.classList.remove("settings-file-error");
      masterStatus.textContent = `Selected: ${file.name}`;
    } catch (error) {
      if (error.name !== "AbortError") {
        masterStatus.classList.add("settings-file-error");
        masterStatus.textContent = `Could not select master file: ${error.message}`;
      }
    } finally {
      selecting = false;
      masterPicker.disabled = false;
      updateReadyState();
      if (!masterStatus.classList.contains("settings-file-error")) showReadyStatus();
    }
  });

  adoInput.addEventListener("change", () => {
    const file = adoInput.files?.[0];
    const invalid = Boolean(file && !isAdoFile(file));
    adoInput.setAttribute("aria-invalid", String(invalid));
    adoStatus.classList.toggle("settings-file-error", invalid);
    if (invalid) adoInput.value = "";
    // Keep bytes, rather than a File reference that can become stale on disk.
    // Capture failures as values so a read failure before Merge is handled too.
    adoSnapshot = file && !invalid
      ? { file, read: file.arrayBuffer().then(bytes => ({ bytes }), error => ({ error })) }
      : null;
    adoStatus.textContent = invalid ? "Please select an Excel or CSV file (.xlsx, .xls or .csv)."
      : file ? `Selected: ${file.name}` : "No file selected.";
    if (canSaveInPlace) showReadyStatus();
  });

  mergeButton.addEventListener("click", async () => {
    if (mergeButton.disabled || !masterHandle) return;
    const adoFile = adoInput.files?.[0];
    if (!isAdoFile(adoFile) || adoSnapshot?.file !== adoFile) return;
    merging = true;
    updateReadyState();
    masterPicker.disabled = true;
    adoInput.disabled = true;
    mergeButton.textContent = "Merging…";
    mergeStatus.classList.remove("settings-file-error");
    mergeStatus.textContent = "Merging data into the selected master file…";
    let writable;
    try {
      // Request permission while the Merge click still provides user activation.
      const permission = await masterHandle.requestPermission({ mode: "readwrite" });
      if (permission !== "granted") throw new Error("Write access was not granted. Allow access to save changes to the master file.");
      const adoRead = await adoSnapshot.read;
      if (adoRead.error) {
        const detail = isFileStateError(adoRead.error) ? "" : ` ${adoRead.error.message || ""}`;
        throw new Error(`Could not read the ADO dump. Save and close it in Excel, wait for any file sync to finish, then select the ADO file again.${detail}`);
      }
      // Refresh only reads; retrying a save could append the dump twice.
      const { file: masterFile, bytes: masterBytes } = await readLatestMaster();
      const result = mergeDailyDumpPreservingWorkbook(masterBytes, adoRead.bytes, adoFile.name);
      writable = await masterHandle.createWritable();
      await writable.write(result.bytes);
      await writable.close();
      writable = null;
      mergeStatus.textContent = `Saved ${result.appendedRows} new row${result.appendedRows === 1 ? "" : "s"} starting at row ${result.firstRow} in Master Daily Dump directly to ${masterFile.name}.`;
    } catch (error) {
      if (writable) {
        try { await writable.abort(); } catch (_) { /* The stream may already be closed. */ }
      }
      mergeStatus.classList.add("settings-file-error");
      const message = isFileStateError(error)
        ? "The file changed or is unavailable. Save and close both files in Excel, wait for any file sync to finish, select both files again, then retry."
        : error.message || "Unable to update the master file.";
      mergeStatus.textContent = `Merge failed: ${message}`;
    } finally {
      merging = false;
      masterPicker.disabled = false;
      adoInput.disabled = false;
      mergeButton.textContent = "Merge Data";
      updateReadyState();
    }
  });
})();
