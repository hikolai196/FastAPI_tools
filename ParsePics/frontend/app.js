(() => {
  const fileInput = document.getElementById("file-input");
  const dropzone = document.getElementById("dropzone");
  const fileList = document.getElementById("file-list");
  const processBtn = document.getElementById("process-btn");
  const clearBtn = document.getElementById("clear-btn");
  const statusEl = document.getElementById("status");
  const resultsEl = document.getElementById("results");
  const resultsBody = document.getElementById("results-body");

  /** @type {File[]} */
  let selected = [];

  const ALLOWED = new Set([".docx", ".pptx"]);

  function extOf(name) {
    const i = name.lastIndexOf(".");
    return i >= 0 ? name.slice(i).toLowerCase() : "";
  }

  function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function setStatus(text, kind = "") {
    statusEl.textContent = text;
    statusEl.className = `status${kind ? ` ${kind}` : ""}`;
  }

  function syncUi() {
    const hasFiles = selected.length > 0;
    processBtn.disabled = !hasFiles;
    clearBtn.disabled = !hasFiles;

    if (!hasFiles) {
      fileList.hidden = true;
      fileList.innerHTML = "";
      return;
    }

    fileList.hidden = false;
    fileList.innerHTML = selected
      .map(
        (f) =>
          `<li><span>${escapeHtml(f.name)}</span><span class="meta">${formatSize(f.size)}</span></li>`
      )
      .join("");
  }

  function escapeHtml(s) {
    return s
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  function addFiles(fileListLike) {
    const incoming = Array.from(fileListLike || []);
    const accepted = [];
    const rejected = [];

    for (const f of incoming) {
      if (ALLOWED.has(extOf(f.name))) accepted.push(f);
      else rejected.push(f.name);
    }

    // Deduplicate by name+size+lastModified
    const key = (f) => `${f.name}::${f.size}::${f.lastModified}`;
    const map = new Map(selected.map((f) => [key(f), f]));
    for (const f of accepted) map.set(key(f), f);
    selected = Array.from(map.values());

    if (rejected.length) {
      setStatus(`已略過不支援的檔案：${rejected.join(", ")}`, "error");
    } else if (accepted.length) {
      setStatus(`已選 ${selected.length} 個檔案`);
    }
    syncUi();
  }

  function clearAll() {
    selected = [];
    fileInput.value = "";
    resultsEl.hidden = true;
    resultsBody.innerHTML = "";
    setStatus("");
    syncUi();
  }

  function renderResults(payload) {
    resultsEl.hidden = false;
    const rows = payload.results || [];
    resultsBody.innerHTML = rows
      .map((r) => {
        if (r.status === "success") {
          return `
            <article class="result-card">
              <header>
                <span class="name">${escapeHtml(r.file_name)}</span>
                <span class="badge success">success</span>
              </header>
              <dl>
                <dt>圖片數</dt><dd>${r.image_count}</dd>
                <dt>輸出目錄</dt><dd>${escapeHtml(r.output_dir || "")}</dd>
                <dt>修改後文件</dt><dd>${escapeHtml(r.modified_file || "")}</dd>
                <dt>Excel</dt><dd>${escapeHtml(r.excel_file || "")}</dd>
                <dt>圖片資料夾</dt><dd>${escapeHtml(r.images_dir || "")}</dd>
              </dl>
            </article>`;
        }
        return `
          <article class="result-card">
            <header>
              <span class="name">${escapeHtml(r.file_name)}</span>
              <span class="badge error">error</span>
            </header>
            <p class="error-text">${escapeHtml(r.error || "處理失敗")}</p>
          </article>`;
      })
      .join("");
  }

  async function processFiles() {
    if (!selected.length) return;

    processBtn.disabled = true;
    clearBtn.disabled = true;
    setStatus("處理中…");

    const form = new FormData();
    for (const f of selected) form.append("files", f, f.name);

    try {
      const resp = await fetch("/process", { method: "POST", body: form });
      const data = await resp.json();

      if (!resp.ok && !data.results) {
        setStatus(data.detail || "請求失敗", "error");
        return;
      }

      renderResults(data);
      const ok = (data.results || []).filter((r) => r.status === "success").length;
      const total = (data.results || []).length;
      const kind = data.status === "success" ? "ok" : data.status === "partial" ? "" : "error";
      setStatus(`完成：${ok}/${total} 成功（${data.status}）`, kind);
    } catch (err) {
      setStatus(`網路或伺服器錯誤：${err.message || err}`, "error");
    } finally {
      syncUi();
    }
  }

  // Events
  fileInput.addEventListener("change", () => {
    addFiles(fileInput.files);
    fileInput.value = "";
  });

  ["dragenter", "dragover"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });

  dropzone.addEventListener("drop", (e) => {
    addFiles(e.dataTransfer?.files);
  });

  processBtn.addEventListener("click", processFiles);
  clearBtn.addEventListener("click", clearAll);

  syncUi();
})();
