"use strict";

(() => {
  const $ = (id) => document.getElementById(id);
  const kinds = { observe: "观察记录", items: "型式评价项目一览表", result: "试验结果" };
  const statuses = { uploaded: "待分析", analyzing: "分析中", ready: "待核对", exporting: "导出中", completed: "已完成", failed: "失败" };
  const state = { job: null, draft: null, dirty: false, busy: false, timer: null, epoch: 0, page: 0, suggestedCount: 1, pollError: false, customNames: new Set(), labels: new Map() };
  const clone = (value) => JSON.parse(JSON.stringify(value));
  const active = () => ["analyzing", "exporting"].includes(state.job?.status);
  const editable = () => Boolean(state.draft && state.job?.total_pages > 0 && !active());
  const text = (value) => typeof value === "string" ? value : JSON.stringify(value) || "";

  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    Object.entries(attrs).forEach(([key, value]) => {
      if (key.startsWith("on")) node.addEventListener(key.slice(2), value);
      else if (["value", "checked", "disabled", "hidden"].includes(key)) node[key] = value;
      else if (value != null) node.setAttribute(key, String(value));
    });
    children.flat().forEach((child) => { if (child != null) node.append(child); });
    return node;
  }
  function button(label, action, style = "secondary") {
    return el("button", { type: "button", class: `button ${style}`, onclick: action }, label);
  }
  function field(label, input, className = "") {
    return el("label", { class: `field ${className}` }, el("span", { class: "field-label" }, label), input);
  }
  function feedback(message, error = false) {
    $("feedback").textContent = message;
    $("feedback").className = `feedback${error ? " error" : ""}`;
    $("feedback").setAttribute("role", error ? "alert" : "status");
  }
  function showErrors(errors = []) {
    $("validation-errors").hidden = !errors.length;
    $("validation-errors").replaceChildren(el("ul", {}, ...errors.map((error) => el("li", {}, error))));
  }
  function safeURL(value) {
    if (!value) return "";
    try {
      const url = new URL(value, location.origin);
      return url.origin === location.origin && ["http:", "https:"].includes(url.protocol) ? url.href : "";
    } catch { return ""; }
  }
  function endpoint(suffix = "") {
    return `/api/separations/${encodeURIComponent(state.job.id)}${suffix}`;
  }
  async function api(url, body) {
    let response;
    try {
      response = await fetch(url, {
        method: body === undefined ? "GET" : "POST", credentials: "same-origin",
        headers: body === undefined || body instanceof FormData ? {} : { "Content-Type": "application/json" },
        body: body === undefined ? undefined : body instanceof FormData ? body : JSON.stringify(body)
      });
    } catch { throw new Error("无法连接本机服务。未保存的编辑仍保留，请检查连接后重试。"); }
    let data;
    try { data = await response.json(); }
    catch { throw new Error(`服务响应不可读取（HTTP ${response.status}），未视为操作成功。`); }
    if (!response.ok) throw new Error(text(data.detail || data.error || data) || `请求失败（HTTP ${response.status}）`);
    return data;
  }
  function validJob(job) {
    if (!job || !job.id || !Object.hasOwn(statuses, job.status)) throw new Error("服务返回的任务信息不完整，未视为操作成功。");
    return job;
  }
  function customKey() { return `separation-custom-names:${state.job.id}`; }
  function readCustomNames() {
    try { state.customNames = new Set(JSON.parse(sessionStorage.getItem(customKey()) || "[]")); }
    catch { state.customNames = new Set(); }
  }
  function rememberCustomNames() {
    try { sessionStorage.setItem(customKey(), JSON.stringify([...state.customNames])); } catch { /* Storage may be disabled. */ }
  }
  function stopPolling() {
    clearTimeout(state.timer);
    state.timer = null;
    state.epoch += 1;
  }
  function adopt(job) {
    validJob(job);
    const previousJobId = state.job?.id;
    stopPolling();
    state.job = job;
    if (previousJobId !== job.id || !$("job-select").value) $("job-select").value = job.id;
    state.draft = {
      model_count: job.model_count ?? 1,
      root_directory_token: job.root_directory_token || "",
      parts: clone(job.parts || []).map((part) => ({ ...part, enabled: part.enabled !== false, directory_token: part.directory_token || "" }))
    };
    state.suggestedCount = Number(state.draft.model_count);
    state.dirty = false;
    state.pollError = false;
    state.labels.clear();
    (job.parts || []).forEach((part) => {
      const directory = job.plan?.outputs?.find((output) => output.id === part.id)?.directory;
      if (part.directory_token && directory) state.labels.set(part.directory_token, directory);
    });
    state.page = 0;
    readCustomNames();
    history.replaceState(null, "", `${location.pathname}?job=${encodeURIComponent(job.id)}`);
    $("confirm-plan").checked = false;
    showErrors();
    render();
    schedulePoll();
  }
  function schedulePoll() {
    if (!active() || state.timer || state.pollError) return;
    const epoch = state.epoch;
    state.timer = setTimeout(() => { state.timer = null; poll(epoch); }, 1800);
  }
  async function poll(epoch = state.epoch) {
    if (!active() || epoch !== state.epoch) return;
    try {
      const job = validJob(await api(endpoint()));
      if (epoch !== state.epoch) return;
      if (state.dirty) {
        state.job.status = job.status;
        state.job.stage = job.stage;
        state.job.error = job.error;
        renderStatus();
      } else {
        adopt(job);
      }
      state.pollError = false;
      $("poll-retry").hidden = true;
      if (job.status === "failed") feedback(text(job.error) || "任务失败，请查看错误后重试。", true);
      else if (job.status === "completed") feedback(job.exports?.length ? "服务端已完成导出，请核对下方实际文件记录。" : "服务报告完成，但未返回导出文件记录，请重新打开任务核对。", !job.exports?.length);
      else if (job.status === "ready") feedback("分析完成，请核对各文件的物理页码、名称与目录。");
      schedulePoll();
    } catch (error) {
      if (epoch !== state.epoch) return;
      state.pollError = true;
      $("poll-retry").hidden = false;
      feedback(error.message, true);
    }
  }
  async function operation(action) {
    if (state.busy) return;
    state.busy = true;
    controls();
    try { await action(); }
    catch (error) { feedback(error.message, true); }
    finally { state.busy = false; controls(); }
  }
  function mayDiscard() {
    return !state.dirty || window.confirm("当前拆分设置尚未保存。继续会丢弃这些改动，是否继续？");
  }
  function changed() {
    state.dirty = true;
    state.job.plan = null;
    $("confirm-plan").checked = false;
    showErrors();
    renderPlan();
    controls();
  }
  function controls() {
    const locked = state.busy || active();
    $("upload-button").disabled = state.busy;
    $("source-file").disabled = state.busy;
    $("open-job").disabled = state.busy;
    $("job-select").disabled = state.busy;
    $("refresh-history").disabled = state.busy;
    $("plan-fields").disabled = locked || !editable();
    $("plan-button").disabled = locked || !editable();
    $("analyze-button").disabled = locked;
    $("poll-retry").disabled = state.busy;
    const validPlan = completePlan(state.job?.plan) && !state.dirty && !active() && state.job.status !== "completed" && state.job.status !== "failed";
    $("confirm-plan").disabled = state.busy || !validPlan;
    $("export-button").disabled = state.busy || !validPlan || !$("confirm-plan").checked;
    $("dirty-status").textContent = state.dirty ? "有未保存修改 · 旧计划已失效，请重新保存并检查。" : "当前设置已保存；最终导出仍需核对计划并确认。";
  }
  function renderStatus() {
    const job = state.job;
    $("job-section").hidden = !job;
    if (!job) return;
    $("job-title").textContent = job.name || job.id;
    $("job-status").textContent = statuses[job.status];
    $("job-status").className = `pill${job.status === "failed" ? " error" : ""}`;
    $("job-stage").textContent = [text(job.stage), job.total_pages ? `共 ${job.total_pages} 个物理页` : "", active() ? "本机处理中，可稍后从历史任务恢复。" : ""].filter(Boolean).join(" · ");
    $("job-error").hidden = !job.error;
    $("job-error").textContent = text(job.error);
    $("job-warnings").replaceChildren(...(job.warnings || []).map((warning) => el("p", { class: "notice" }, text(warning))));
    $("analyze-button").hidden = !["uploaded", "failed"].includes(job.status);
    $("analyze-button").textContent = job.status === "failed" ? "重新分析（显式重试）" : "开始原生 Word 分析";
    $("source-pdf").hidden = !job.total_pages;
    $("source-pdf").href = endpoint("/source.pdf");
    $("poll-retry").hidden = !state.pollError;
    controls();
  }
  function updateModelCount(value) {
    const previous = state.suggestedCount;
    const next = Number(value);
    if (Number.isSafeInteger(next) && next > 0 && Number.isSafeInteger(previous) && previous > 0) {
      const suffix = new RegExp(`-${previous} models?\\.docx$`, "i");
      state.draft.parts.forEach((part) => {
        if (part.kind !== "items" && !state.customNames.has(part.id) && suffix.test(part.filename)) {
          part.filename = part.filename.replace(suffix, `-${next} ${next === 1 ? "model" : "models"}.docx`);
        }
      });
      state.suggestedCount = next;
    }
    state.draft.model_count = value;
    changed();
    renderParts();
  }
  function renderParts() {
    $("parts").replaceChildren(...Object.entries(kinds).map(([kind, label]) => {
      const parts = state.draft.parts.filter((part) => part.kind === kind);
      return el("section", { class: "part-group", "aria-label": label },
        el("h3", {}, `${label}（${parts.length} 份）`),
        ...(parts.length ? parts.map(renderPart) : [el("p", { class: "muted" }, "暂无文件，可通过下方按钮增加。")]));
    }));
  }
  function renderPart(part) {
    const enabled = el("input", { type: "checkbox", checked: part.enabled, onchange: () => {
      part.enabled = enabled.checked;
      row.classList.toggle("excluded", !part.enabled);
      changed();
    } });
    const category = el("select", { onchange: () => { part.kind = category.value; changed(); renderParts(); } },
      ...Object.entries(kinds).map(([kind, label]) => el("option", { value: kind }, label)));
    category.value = part.kind;
    const pageField = (key, label) => {
      const input = el("input", { type: "number", min: 1, max: state.job.total_pages, step: 1, required: "", value: part[key], oninput: () => { part[key] = input.value; changed(); } });
      return field(label, input);
    };
    const filename = el("input", { type: "text", value: part.filename, required: "", oninput: () => {
      part.filename = filename.value;
      state.customNames.add(part.id);
      rememberCustomNames();
      changed();
    } });
    const directory = part.directory_token
      ? state.labels.get(part.directory_token) || (state.job.plan?.outputs || []).find((output) => output.id === part.id)?.directory || "已选择专用目录（保存计划后核对完整路径）"
      : "使用根文件夹下的类别子文件夹";
    const clear = button("清除专用目录", () => { part.directory_token = ""; changed(); renderParts(); });
    clear.disabled = !part.directory_token;
    const row = el("article", { class: `part-row${part.enabled ? "" : " excluded"}` },
      el("div", { class: "row-heading" },
        el("label", { class: "check" }, enabled, el("strong", {}, part.project || kinds[part.kind])),
        button("删除此行", () => {
          if (!window.confirm(`删除“${part.filename || part.project || kinds[part.kind]}”这一行？不会删除已有导出文件。`)) return;
          state.draft.parts = state.draft.parts.filter((item) => item.id !== part.id);
          changed(); renderParts();
        }, "danger small")),
      el("div", { class: "part-fields" }, field("类别", category, "kind-field"), pageField("start_page", "起始物理页"), pageField("end_page", "结束物理页"), field("输出文件名（含 .docx）", filename, "filename-field")),
      el("div", { class: "part-directory" }, el("p", { class: "directory-path" }, directory),
        button("单份目录…", () => pickFolder(part)), clear,
        button("预览起始页", () => selectPhysicalPage(Number(part.start_page)), "quiet")));
    return row;
  }
  function renderPreview() {
    const pages = state.job.pages || [];
    $("preview-section").hidden = !pages.length;
    $("preview-page").replaceChildren(...pages.map((page, index) => el("option", { value: index }, `物理页 ${page.number}`)));
    state.page = Math.max(0, Math.min(state.page, pages.length - 1));
    showPage();
  }
  function showPage() {
    const pages = state.job.pages || [];
    const page = pages[state.page];
    if (!page) return;
    $("preview-page").value = String(state.page);
    $("previous-page").disabled = state.page === 0;
    $("next-page").disabled = state.page === pages.length - 1;
    $("page-caption").textContent = `物理第 ${page.number} 页 · 点击缩略图可在新标签页放大`;
    $("page-text").textContent = page.text || "服务未返回此页文字，请打开原生 PDF 核对。";
    const imageURL = safeURL(page.image_url);
    $("image-error").hidden = Boolean(imageURL);
    $("page-image-link").hidden = !imageURL;
    if (imageURL) {
      $("page-image-link").href = imageURL;
      $("page-image").alt = `源报告物理第 ${page.number} 页缩略图`;
      $("page-image").src = imageURL;
    } else $("page-image").removeAttribute("src");
  }
  function selectPhysicalPage(number) {
    const index = (state.job.pages || []).findIndex((page) => page.number === number);
    if (index < 0) { feedback("该物理页预览不可用，请检查页码或打开原生 PDF。", true); return; }
    state.page = index;
    showPage();
    $("preview-page").focus();
    $("preview-section").scrollIntoView({ behavior: "smooth", block: "start" });
  }
  function outputTable(outputs) {
    return el("table", {},
      el("caption", { class: "visually-hidden" }, "最终计划输出文件清单"),
      el("thead", {}, el("tr", {}, ...["类别", "原报告物理页码", "完整文件名", "最终保存路径"].map((title) => el("th", { scope: "col" }, title)))),
      el("tbody", {}, ...outputs.map((output) => el("tr", {},
        el("td", {}, kinds[output.kind] || output.kind),
        el("td", {}, `${output.start_page}–${output.end_page}`),
        el("td", {}, output.filename), el("td", {}, output.full_path || "服务未返回完整路径，不能导出")))));
  }
  function renderPlan() {
    const plan = !state.dirty && state.job?.plan;
    $("plan-section").hidden = !plan;
    if (!plan) return;
    $("plan-title").textContent = `3. 核对最终计划并确认导出 · 版本 ${plan.revision}`;
    $("plan-outputs").replaceChildren(outputTable(plan.outputs || []));
    $("plan-adjustments").replaceChildren(...(plan.adjustments?.length ? plan.adjustments.map((adjustment) => {
      const name = plan.outputs.find((output) => output.id === adjustment.id)?.filename || adjustment.id;
      return el("p", { class: "notice" }, `${name}：${adjustment.message || "建议调整项目边界"}（请求 ${adjustment.requested_start}–${adjustment.requested_end} 页 → 建议 ${adjustment.start_page}–${adjustment.end_page} 页）。待您确认后采用。`);
    }) : [el("p", { class: "notice" }, "未提出页码边界调整；仍请核对所有文件和保存路径。")]));
    if (!completePlan(plan)) $("plan-adjustments").append(el("p", { class: "notice error", role: "alert" }, "已保存计划不完整或与当前选中行不一致，已禁止导出。请重新保存并检查拆分计划。"));
  }
  function renderExports() {
    const exports = state.job.exports || [];
    $("exports-section").hidden = !exports.length && state.job.status !== "completed";
    if (!exports.length && state.job.status === "completed") {
      $("export-list").replaceChildren(el("p", { class: "notice error", role: "alert" }, "服务报告完成，但未返回任何导出文件。请重新打开此任务核对；不会生成虚假的下载链接。"));
      return;
    }
    $("export-list").replaceChildren(...exports.map((file) => {
      const url = safeURL(file.url);
      return el("div", { class: "export-entry" }, el("strong", {}, file.filename),
        el("p", { class: "muted" }, file.checks?.layout_policy
          ? `原报告 P${file.checks.source_start_page}–P${file.checks.source_end_page} → 实际输出 ${file.checks.output_pages} 页。${file.checks.review_note}`
          : "历史导出：按当时规则生成，未自动重新排版。"),
        el("p", { class: "directory-path" }, file.path || "服务未返回保存路径"),
        url ? (window.workbenchDownloads ? window.workbenchDownloads.link(url, "下载可编辑 Word", file.filename)
          : el("p", {class: "notice error", role: "alert"}, "下载组件未载入，请强制刷新后重试。"))
          : el("p", { class: "notice error" }, "未返回可用的本机下载链接。"));
    }));
  }
  function render() {
    renderStatus();
    if (!state.job) return;
    $("editor-section").hidden = !state.job.total_pages;
    $("model-count").value = state.draft.model_count;
    $("root-directory").textContent = state.labels.get(state.draft.root_directory_token) || state.job.root_directory_label || (state.draft.root_directory_token ? "已选择根目录（计划中核对路径）" : "尚未选择");
    renderParts(); renderPreview(); renderPlan(); renderExports(); controls();
  }
  function validateDraft() {
    const errors = [];
    const integer = (value) => String(value).trim() !== "" && Number.isSafeInteger(Number(value));
    if (!integer(state.draft.model_count) || Number(state.draft.model_count) < 1) errors.push("型号数量必须是大于零的整数。");
    const included = state.draft.parts.filter((part) => part.enabled);
    if (!included.length) errors.push("请至少勾选一份要导出的文件。");
    const paths = new Set();
    const ids = new Set();
    state.draft.parts.forEach((part, index) => {
      const name = `第 ${index + 1} 行（${part.filename || "未命名"}）`;
      if (!part.id || ids.has(part.id)) errors.push(`${name}的标识无效或重复。`);
      ids.add(part.id);
      if (!Object.hasOwn(kinds, part.kind)) errors.push(`${name}的类别无效。`);
      if (!integer(part.start_page) || !integer(part.end_page)) errors.push(`${name}的起止页码必须是整数。`);
      if (!part.enabled) return;
      const start = Number(part.start_page), end = Number(part.end_page);
      if (start < 1 || end < start || end > state.job.total_pages) errors.push(`${name}的范围必须在 1–${state.job.total_pages} 页内，结束页不能早于起始页。`);
      const filename = part.filename.trim();
      if (!filename || !/\.docx$/i.test(filename) || /[/\\:\u0000-\u001f]/.test(filename) || /^\.docx$/i.test(filename) || filename !== part.filename) errors.push(`${name}需要有效的 .docx 文件名，不能含路径分隔符、控制字符或首尾空格。`);
      if (!part.directory_token && !state.draft.root_directory_token) errors.push(`${name}尚未设置根目录或单份保存目录。`);
      const destination = part.directory_token ? `override:${part.directory_token}` : `root:${state.draft.root_directory_token}:${part.kind}`;
      const key = `${destination}/${filename.normalize("NFC").toLocaleLowerCase()}`;
      if (paths.has(key)) errors.push(`${name}与同目录另一份文件重名，请改名；服务端还会检查实际磁盘冲突。`);
      paths.add(key);
    });
    return errors;
  }
  function payload() {
    return {
      model_count: Number(state.draft.model_count), root_directory_token: state.draft.root_directory_token,
      parts: state.draft.parts.map((part) => ({
        id: part.id, kind: part.kind, enabled: part.enabled,
        start_page: Number(part.start_page), end_page: Number(part.end_page),
        filename: part.filename, directory_token: part.directory_token || ""
      }))
    };
  }
  function completePlan(plan) {
    if (!Number.isSafeInteger(plan?.revision) || !Array.isArray(plan.outputs) || !plan.outputs.length) return false;
    const included = state.draft.parts.filter((part) => part.enabled);
    const ids = new Set();
    return plan.outputs.length === included.length && plan.outputs.every((output) => {
      const part = included.find((item) => item.id === output.id);
      if (!part || ids.has(output.id)) return false;
      ids.add(output.id);
      return output.kind === part.kind && typeof output.filename === "string" && output.filename.length > 0 &&
        typeof output.directory === "string" && output.directory.length > 0 &&
        typeof output.full_path === "string" && output.full_path.length > 0 &&
        Number.isSafeInteger(output.start_page) && Number.isSafeInteger(output.end_page) &&
        output.start_page >= 1 && output.end_page >= output.start_page && output.end_page <= state.job.total_pages;
    });
  }
  async function loadHistory() {
    const jobs = await api("/api/separations");
    if (!Array.isArray(jobs)) throw new Error("本机任务列表格式无效，请重试。");
    const selected = state.job?.id || $("job-select").value;
    $("job-select").replaceChildren(el("option", { value: "" }, jobs.length ? "请选择本机任务" : "暂无拆分任务"),
      ...jobs.map((job) => el("option", { value: job.id }, `${job.name || job.id} · ${statuses[job.status] || job.status} · ${job.created_at || ""}`)));
    $("job-select").value = selected;
  }
  async function pickFolder(part = null) {
    await operation(async () => {
      const result = await api(endpoint("/folder"), { part_id: part?.id || null });
      if (result.cancelled) { feedback("已取消选择，原保存目录保持不变。"); return; }
      if (!result.token || !result.path) throw new Error("服务未返回有效的目录授权，保存目录未改变。");
      state.labels.set(result.token, result.path);
      if (part) part.directory_token = result.token;
      else state.draft.root_directory_token = result.token;
      changed();
      render();
      feedback("目录选择已记录，请保存并重新检查拆分计划。");
    });
  }
  async function savePlan() {
    const errors = validateDraft();
    showErrors(errors);
    if (errors.length) { feedback("请修正拆分设置中的错误。", true); $("validation-errors").scrollIntoView({ block: "center" }); return; }
    await operation(async () => {
      state.job.plan = null;
      $("confirm-plan").checked = false;
      renderPlan();
      try {
        const job = validJob(await api(endpoint("/plan"), payload()));
        if (!completePlan(job.plan)) throw new Error("服务未返回有效的完整拆分计划（需包含每份文件的名称、物理页码和完整路径），无法导出。");
        // Keep requested boundaries in the editor; adjustments are proposals until export confirmation.
        state.job = job;
        state.dirty = false;
        showErrors();
        render();
        feedback("拆分设置已保存并通过计划检查。请核对最终清单与调整建议，再明确确认导出。");
      } catch (error) {
        showErrors([error.message]);
        throw error;
      }
    });
  }
  async function exportPlan() {
    if ($("export-button").disabled || !completePlan(state.job.plan) || state.dirty || !$("confirm-plan").checked) return;
    const plan = state.job.plan;
    const summary = plan.outputs.map((output) => `${output.filename}｜物理页 ${output.start_page}–${output.end_page}｜${output.full_path || output.directory}`).join("\n");
    if (!window.confirm(`确认导出计划版本 ${plan.revision} 的 ${plan.outputs.length} 份文件？\n\n${summary}\n\n已核对并接受页面中列出的全部边界调整建议。不会覆盖已有文件。`)) return;
    await operation(async () => {
      $("confirm-plan").checked = false;
      try {
        const job = validJob(await api(endpoint("/export"), { plan_revision: plan.revision, confirm_adjustments: true }));
        adopt(job);
        if (job.status === "failed") feedback(text(job.error) || "导出失败，请查看错误。", true);
        else if (job.status === "completed") feedback(job.exports?.length ? "服务端返回导出已完成，请核对实际文件记录。" : "服务报告完成，但未返回导出文件记录，请重新打开任务核对。", !job.exports?.length);
        else feedback("导出请求已提交，正在等待本机处理结果。");
      } catch (error) {
        state.job.plan = null;
        renderPlan();
        showErrors([`${error.message} 请重新保存检查计划后再导出。`]);
        throw error;
      }
    });
  }
  $("upload-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const file = $("source-file").files[0];
    if (!file || !/\.docx$/i.test(file.name) || !file.size) { feedback("请选择非空的可编辑 .docx 源报告。", true); return; }
    if (!mayDiscard()) return;
    operation(async () => {
      const data = new FormData();
      data.append("file", file);
      adopt(await api("/api/separations/upload", data));
      $("source-file").value = "";
      if (state.job.status === "uploaded") adopt(await api(endpoint("/analyze"), {}));
      if (state.job.status === "failed") feedback(text(state.job.error) || "源报告分析失败，请查看任务错误。", true);
      else feedback("源报告已上传，正在读取分析状态；失败时可显式重试分析。");
      try { await loadHistory(); } catch (error) { feedback(`源报告已上传，但任务列表刷新失败：${error.message}`, true); }
    });
  });
  $("open-job").addEventListener("click", () => {
    const id = $("job-select").value;
    if (!id) { feedback("请先选择一个历史任务。", true); return; }
    if (!mayDiscard()) return;
    operation(async () => { adopt(await api(`/api/separations/${encodeURIComponent(id)}`)); feedback("已恢复本机任务；请核对已保存的设置与状态。"); });
  });
  $("refresh-history").addEventListener("click", () => operation(async () => { await loadHistory(); feedback("本机任务列表已刷新。"); }));
  $("analyze-button").addEventListener("click", () => {
    if (!mayDiscard()) return;
    operation(async () => { adopt(await api(endpoint("/analyze"), {})); feedback("已提交原生 Word 分析请求，请等待任务状态更新。"); });
  });
  $("poll-retry").addEventListener("click", () => { state.pollError = false; poll(); });
  $("root-picker").addEventListener("click", () => pickFolder());
  $("model-count").addEventListener("input", (event) => updateModelCount(event.target.value));
  $("add-part").addEventListener("click", () => {
    const kind = $("new-kind").value;
    const id = crypto.randomUUID();
    state.draft.parts.push({ id, kind, project: "", start_page: 1, end_page: 1, filename: "", enabled: true, directory_token: "" });
    state.customNames.add(id);
    changed(); renderParts();
    const inputs = $("parts").querySelectorAll(".filename-field input");
    inputs[inputs.length - 1]?.focus();
  });
  $("plan-form").addEventListener("submit", (event) => { event.preventDefault(); savePlan(); });
  $("confirm-plan").addEventListener("change", controls);
  $("export-button").addEventListener("click", exportPlan);
  $("preview-page").addEventListener("change", (event) => { state.page = Number(event.target.value); showPage(); });
  $("previous-page").addEventListener("click", () => { state.page -= 1; showPage(); });
  $("next-page").addEventListener("click", () => { state.page += 1; showPage(); });
  $("page-image").addEventListener("error", () => { $("image-error").hidden = false; $("page-image-link").hidden = true; });
  window.addEventListener("beforeunload", (event) => {
    if (!state.dirty) return;
    event.preventDefault();
    event.returnValue = "";
  });
  window.addEventListener("pagehide", stopPolling);
  window.addEventListener("pageshow", () => schedulePoll());
  operation(async () => {
    let listError = "";
    try { await loadHistory(); } catch (error) { listError = error.message; $("job-select").replaceChildren(el("option", { value: "" }, "任务列表读取失败，请刷新重试")); }
    const id = new URLSearchParams(location.search).get("job");
    if (id) adopt(await api(`/api/separations/${encodeURIComponent(id)}`));
    feedback(listError || (id ? "已恢复本机拆分任务。" : "请选择可编辑 DOCX，或打开本机历史任务。"), Boolean(listError));
  });
})();
