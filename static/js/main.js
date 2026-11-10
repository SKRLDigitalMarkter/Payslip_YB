/**
 * YAYATH PAYROLL - CLIENT APPLICATION LOGIC
 */

document.addEventListener("DOMContentLoaded", () => {
  initDropzone();
  initFlashAlerts();
  initTableSearch();
  initMappingInteractions();
  initPreviewModal();
});

/* ==========================================================================
   DROPZONE & FILE UPLOAD
   ========================================================================== */
function initDropzone() {
  const dropzone = document.getElementById("payrollDropzone");
  const fileInput = document.getElementById("fileInput");
  const previewPill = document.getElementById("selectedFilePreview");
  const fileNameDisplay = document.getElementById("selectedFileName");
  const fileSizeDisplay = document.getElementById("selectedFileSize");
  const removeFileBtn = document.getElementById("removeFileBtn");
  const uploadForm = document.getElementById("uploadForm");
  const uploadSubmitBtn = document.getElementById("uploadSubmitBtn");

  if (!dropzone || !fileInput) return;

  // Prevent browser default open behavior on drop
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, () => dropzone.classList.add("dragover"));
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, () => dropzone.classList.remove("dragover"));
  });

  dropzone.addEventListener("drop", (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      fileInput.files = dt.files;
      handleFileSelected(dt.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  if (removeFileBtn) {
    removeFileBtn.addEventListener("click", (e) => {
      e.preventDefault();
      fileInput.value = "";
      if (previewPill) previewPill.style.display = "none";
      if (uploadSubmitBtn) uploadSubmitBtn.disabled = true;
    });
  }

  function handleFileSelected(file) {
    const allowed = [".xlsx", ".xls", ".csv"];
    const ext = "." + file.name.split(".").pop().toLowerCase();

    if (!allowed.includes(ext)) {
      alert("Please upload a valid Excel (.xlsx, .xls) or CSV (.csv) file.");
      fileInput.value = "";
      return;
    }

    if (fileNameDisplay) fileNameDisplay.textContent = file.name;
    if (fileSizeDisplay) {
      const sizeKb = (file.size / 1024).toFixed(1);
      fileSizeDisplay.textContent = `${sizeKb} KB`;
    }
    if (previewPill) previewPill.style.display = "flex";
    if (uploadSubmitBtn) {
      uploadSubmitBtn.disabled = false;
      uploadSubmitBtn.focus();
    }
  }

  if (uploadForm) {
    uploadForm.addEventListener("submit", () => {
      if (uploadSubmitBtn) {
        uploadSubmitBtn.innerHTML = `
          <div class="spinner" style="width:18px;height:18px;border-width:2px;display:inline-block;vertical-align:middle;margin-right:8px;"></div>
          Processing Spreadsheet...
        `;
        uploadSubmitBtn.disabled = true;
      }
    });
  }
}

/* ==========================================================================
   ALERT DISMISSAL
   ========================================================================== */
function initFlashAlerts() {
  document.querySelectorAll(".alert-close").forEach((btn) => {
    btn.addEventListener("click", () => {
      const alert = btn.closest(".alert");
      if (alert) {
        alert.style.opacity = "0";
        alert.style.transform = "translateY(-6px)";
        setTimeout(() => alert.remove(), 250);
      }
    });
  });
}

/* ==========================================================================
   LIVE TABLE SEARCH & FILTER
   ========================================================================== */
function initTableSearch() {
  const searchInput = document.getElementById("tableSearchInput");
  const tableRows = document.querySelectorAll(".preview-table tbody tr");
  const countBadge = document.getElementById("visibleRowCount");

  if (!searchInput || !tableRows.length) return;

  searchInput.addEventListener("input", (e) => {
    const term = e.target.value.toLowerCase().trim();
    let visibleCount = 0;

    tableRows.forEach((row) => {
      const text = row.textContent.toLowerCase();
      if (text.includes(term)) {
        row.style.display = "";
        visibleCount++;
      } else {
        row.style.display = "none";
      }
    });

    if (countBadge) {
      countBadge.textContent = visibleCount;
    }
  });
}

/* ==========================================================================
   DYNAMIC COLUMN MAPPING
   ========================================================================== */
function initMappingInteractions() {
  const mappingSelects = document.querySelectorAll(".mapping-select");
  if (!mappingSelects.length) return;

  mappingSelects.forEach((select) => {
    select.addEventListener("change", () => {
      const item = select.closest(".mapping-item");
      if (!item) return;
      const badge = item.querySelector(".status-badge");
      if (!badge) return;

      if (select.value) {
        badge.className = "status-badge badge-matched";
        badge.textContent = "Matched";
      } else {
        badge.className = "status-badge badge-unmapped";
        badge.textContent = "Unmapped";
      }
      updateMappingHiddenInputs();
    });
  });
}

function getCurrentMapping() {
  const mapping = {};
  document.querySelectorAll(".mapping-select").forEach((sel) => {
    const key = sel.dataset.fieldKey;
    if (key && sel.value) {
      mapping[key] = sel.value;
    }
  });
  return mapping;
}

function updateMappingHiddenInputs() {
  const current = getCurrentMapping();
  const mappingJson = JSON.stringify(current);
  document.querySelectorAll(".form-mapping-json").forEach((input) => {
    input.value = mappingJson;
  });
}

/* ==========================================================================
   IN-MODAL LIVE PDF PREVIEW
   ========================================================================== */
let activeBlobUrl = null;

function initPreviewModal() {
  const modal = document.getElementById("pdfPreviewModal");
  const modalCloseBtn = document.getElementById("modalCloseBtn");
  const previewIframe = document.getElementById("previewIframe");
  const modalLoader = document.getElementById("modalLoader");
  const modalTitle = document.getElementById("modalTitle");
  const modalDownloadBtn = document.getElementById("modalDownloadBtn");

  if (!modal) return;

  const closeModal = () => {
    modal.style.display = "none";
    if (previewIframe) previewIframe.src = "about:blank";
    if (activeBlobUrl) {
      URL.revokeObjectURL(activeBlobUrl);
      activeBlobUrl = null;
    }
  };

  if (modalCloseBtn) modalCloseBtn.addEventListener("click", closeModal);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeModal();
  });

  // Attach preview click event to all preview buttons in the employee table
  document.querySelectorAll(".btn-preview-slip").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const rowIndex = parseInt(btn.dataset.rowIndex, 10);
      if (typeof window.showEmployeePreview === "function") {
        window.showEmployeePreview(rowIndex);
        return;
      }
      const empName = btn.dataset.empName || "Employee";
      const empCode = btn.dataset.empCode || "Code";

      if (modalTitle) {
        modalTitle.textContent = `Payslip Preview — ${empName} (${empCode})`;
      }

      modal.style.display = "flex";
      if (modalLoader) modalLoader.style.display = "flex";
      if (previewIframe) previewIframe.style.opacity = "0";

      try {
        // Use the globally available payroll data from the inline <script> tag
        if (typeof _records === "undefined" || !_records) throw new Error("Employee records not found.");
        const record = _records[rowIndex];
        const currentMapping = (typeof getCurrentMapping === "function") ? getCurrentMapping() : _mapping;

        const month = (typeof _month !== "undefined") ? _month : "";
        const company = (typeof _company !== "undefined") ? _company : "YAYATH BUILDTECH";
        const address = (typeof _address !== "undefined") ? _address : "";

        const response = await fetch("/preview-pdf", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            record: record,
            mapping: currentMapping,
            month: month,
            company: company,
            address: address
          })
        });

        if (!response.ok) {
          throw new Error("Failed to render PDF preview.");
        }

        const blob = await response.blob();
        if (activeBlobUrl) URL.revokeObjectURL(activeBlobUrl);
        activeBlobUrl = URL.createObjectURL(blob);

        if (previewIframe) {
          previewIframe.src = activeBlobUrl;
          previewIframe.onload = () => {
            if (modalLoader) modalLoader.style.display = "none";
            previewIframe.style.opacity = "1";
          };
        }

        if (modalDownloadBtn) {
          modalDownloadBtn.onclick = () => {
            const a = document.createElement("a");
            a.href = activeBlobUrl;
            a.download = `Payslip_${empName.replace(/\s+/g, "_")}_${empCode}.pdf`;
            a.click();
          };
        }
      } catch (err) {
        console.error(err);
        if (modalLoader) {
          modalLoader.innerHTML = `
            <p style="color:#f43f5e;font-weight:600;">Error generating preview: ${err.message}</p>
            <button class="btn btn-secondary btn-sm" onclick="document.getElementById('pdfPreviewModal').style.display='none'">Close</button>
          `;
        }
      }
    });
  });
}
