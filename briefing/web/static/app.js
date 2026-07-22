(() => {
  const state = { jobId: null };

  const $ = (id) => document.getElementById(id);
  const drop = $("dropzone");
  const fileInput = $("file-input");
  const uploadMeta = $("upload-meta");
  const stepConfig = $("step-config");
  const stepGenerate = $("step-generate");
  const btnGenerate = $("btn-generate");
  const btnDownload = $("btn-download");
  const btnDownloadXlsx = $("btn-download-xlsx");
  const statusEl = $("status");
  const preview = $("preview");
  const previewList = $("preview-list");
  const previewIndustry = $("preview-industry");

  function setStatus(type, text) {
    statusEl.hidden = false;
    statusEl.className = `status ${type}`;
    statusEl.textContent = text;
  }

  function enableSteps() {
    stepConfig.classList.remove("is-disabled");
    stepGenerate.classList.remove("is-disabled");
    btnGenerate.disabled = false;
  }

  function fillMeta(data) {
    uploadMeta.hidden = false;
    uploadMeta.innerHTML = `
      <div class="meta-card"><span>文件</span><strong>${escapeHtml(data.filename)}</strong></div>
      <div class="meta-card"><span>数据行</span><strong>${data.rows}</strong></div>
      <div class="meta-card"><span>公司数</span><strong>${data.companies}</strong></div>
    `;
  }

  function fillForm(data) {
    $("period_label").value = data.period_label || "";
    $("period_type").value = data.period_type || "half_year";
    $("current").value = data.dates.current || "";
    $("previous_quarter").value = data.dates.previous_quarter || "";
    $("year_start").value = data.dates.year_start || "";
  }

  function validationSummary(data) {
    const v = data.draft_validation;
    if (!v) return "";
    const errs = (v.findings || []).filter((f) => f.severity === "error");
    const warns = (v.findings || []).filter((f) => f.severity === "warning");
    if (!errs.length && !warns.length) return "\n底稿校验：通过";
    const lines = [`\n底稿校验：error ${errs.length} · warning ${warns.length}`];
    [...errs, ...warns].slice(0, 8).forEach((f) => {
      const loc = f.sheet ? `[${f.sheet}] ` : "";
      lines.push(`- ${f.severity}: ${loc}${f.message}`);
    });
    if (errs.length + warns.length > 8) lines.push("- …（详见 validate-draft）");
    return lines.join("\n");
  }

  function escapeHtml(s) {
    return String(s)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  async function uploadFile(file) {
    if (!file) return;
    setStatus("info", "正在导入底稿…");
    btnDownload.hidden = true;
    btnDownloadXlsx.hidden = true;
    preview.hidden = true;

    const fd = new FormData();
    fd.append("file", file);
    try {
      const res = await fetch("/api/upload", { method: "POST", body: fd });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.detail || "上传失败");

      state.jobId = data.job_id;
      fillMeta(data);
      fillForm(data);
      enableSteps();
      const hasErr = data.draft_validation && data.draft_validation.error_count > 0;
      setStatus(
        hasErr ? "error" : "ok",
        `${data.message}\n识别日期：期末 ${data.dates.current} · 上季 ${data.dates.previous_quarter} · 年初 ${data.dates.year_start}\n品类：${data.categories.join("、")}${validationSummary(data)}`
      );
    } catch (err) {
      setStatus("error", String(err.message || err));
    }
  }

  drop.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => uploadFile(fileInput.files[0]));

  ["dragenter", "dragover"].forEach((ev) => {
    drop.addEventListener(ev, (e) => {
      e.preventDefault();
      drop.classList.add("is-drag");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    drop.addEventListener(ev, (e) => {
      e.preventDefault();
      drop.classList.remove("is-drag");
    });
  });
  drop.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    uploadFile(file);
  });

  btnGenerate.addEventListener("click", async () => {
    if (!state.jobId) return;
    const payload = {
      job_id: state.jobId,
      period_label: $("period_label").value.trim(),
      period_type: $("period_type").value,
      current: $("current").value.trim(),
      previous_quarter: $("previous_quarter").value.trim(),
      year_start: $("year_start").value.trim(),
      focus_company: $("focus_company").value.trim(),
      focus_company_short: $("focus_company_short").value.trim(),
    };

    if (!/^\d{8}$/.test(payload.current) || !/^\d{8}$/.test(payload.previous_quarter) || !/^\d{8}$/.test(payload.year_start)) {
      setStatus("error", "日期须为 8 位数字，如 20250630");
      return;
    }

    btnGenerate.disabled = true;
    setStatus("info", "正在生成 PPT + Excel 表图，大约需要几秒…");
    btnDownload.hidden = true;
    btnDownloadXlsx.hidden = true;
    preview.hidden = true;

    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(typeof data.detail === "string" ? data.detail : JSON.stringify(data.detail || "生成失败"));

      const src = data.table_images_from === "excel" ? "表图来自 Excel 条件格式导出" : "表图为程序绘制（Excel 静默导出未成功）";
      setStatus(
        "ok",
        `生成完成（${data.n_slides} 页）· ${data.generated_at}\nPPT：${data.pptx_name}\nExcel：${data.xlsx_name || "—"}\n${src}`
      );
      btnDownload.hidden = false;
      btnDownload.href = `/api/download/${state.jobId}?kind=pptx`;
      btnDownload.setAttribute("download", data.pptx_name);
      if (data.xlsx_name) {
        btnDownloadXlsx.hidden = false;
        btnDownloadXlsx.href = `/api/download/${state.jobId}?kind=xlsx`;
        btnDownloadXlsx.setAttribute("download", data.xlsx_name);
      }

      const pv = data.preview || {};
      previewIndustry.textContent = `行业合计 ${pv.industry || "—"} · 关注 ${pv.focus || ""}`;
      previewList.innerHTML = (pv.items || [])
        .map((it) => `<li><span>${escapeHtml(it.label)}</span><strong>${escapeHtml(it.text)}</strong></li>`)
        .join("");
      preview.hidden = false;
      refreshRecent();
    } catch (err) {
      setStatus("error", String(err.message || err));
    } finally {
      btnGenerate.disabled = false;
    }
  });

  async function refreshRecent() {
    try {
      const res = await fetch("/api/outputs");
      const data = await res.json();
      const list = $("recent-list");
      if (!data.items || !data.items.length) {
        list.innerHTML = `<li class="muted">暂无文件</li>`;
        return;
      }
      list.innerHTML = data.items
        .map((it) => {
          const links = [];
          if (it.pptx) {
            links.push(`<a href="/api/download-file/${encodeURIComponent(it.pptx)}">PPT</a>`);
          }
          if (it.xlsx) {
            links.push(`<a href="/api/download-file/${encodeURIComponent(it.xlsx)}">Excel</a>`);
          }
          return `
        <li>
          <div>
            <strong>${escapeHtml(it.key)}</strong>
            <div class="muted">${escapeHtml(it.mtime)} · ${it.size_mb} MB</div>
          </div>
          <div class="recent-links">${links.join(" · ")}</div>
        </li>`;
        })
        .join("");
    } catch {
      /* ignore */
    }
  }

  refreshRecent();
})();
