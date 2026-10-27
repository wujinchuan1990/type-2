"use strict";

// Independent copy of the splitter's upload, preview, planning and history UI.
(() => {
  const $ = (id) => document.getElementById(id);
  const API = "/api/separations-experimental";
  const kinds = {observe: "观察记录", items: "型式评价项目一览表", result: "试验结果"};
  const statuses = {uploaded: "待分析", analyzing: "分析中", ready: "待核对", exporting: "导出中", completed: "已完成", partial: "部分成功", failed: "失败"};
  const state = {job: null, draft: null, dirty: false, busy: false, timer: null, epoch: 0, page: 0,
    customNames: new Set(), labels: new Map(), invalidation: Promise.resolve(), pollError: false,
    expandedParts: new Set(), modelCounts: {result: new Set([1]), observe: new Set([1])},
    resultSampleCounts: new Map()};
  const clone = (value) => JSON.parse(JSON.stringify(value));
  const active = () => ["analyzing", "exporting"].includes(state.job?.status);
  const endpoint = (suffix = "") => `${API}/${encodeURIComponent(state.job.id)}${suffix}`;
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
  const button = (label, action, style = "secondary") => el("button", {type: "button", class: `button ${style}`, onclick: action}, label);
  const field = (label, input, extra = "") => el("label", {class: `field ${extra}`}, el("span", {class: "field-label"}, label), input);
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
    try { const url = new URL(value, location.origin); return url.origin === location.origin && ["http:", "https:"].includes(url.protocol) ? url.href : ""; }
    catch { return ""; }
  }
  async function api(url, body) {
    let response;
    try {
      response = await fetch(url, {method: body === undefined ? "GET" : "POST", credentials: "same-origin",
        headers: body === undefined || body instanceof FormData ? {} : {"Content-Type": "application/json"},
        body: body === undefined ? undefined : body instanceof FormData ? body : JSON.stringify(body)});
    } catch { throw new Error("无法连接本机服务；未保存改动仍保留，请检查连接。"); }
    let data;
    try { data = await response.json(); } catch { throw new Error(`响应不可读取（HTTP ${response.status}），未视为成功。`); }
    if (!response.ok) throw new Error(typeof data.detail === "string" ? data.detail : JSON.stringify(data.detail || data));
    return data;
  }
  function validJob(job) {
    if (!job?.id || !Object.hasOwn(statuses, job.status)) throw new Error("任务响应不完整，未视为成功。");
    return job;
  }
  function rememberNames() {
    try { sessionStorage.setItem(`experimental-split-names:${state.job.id}`, JSON.stringify([...state.customNames])); } catch {}
  }
  function adopt(job) {
    validJob(job);
    if (state.job?.id !== job.id || state.job?.status === "analyzing") state.expandedParts.clear();
    clearTimeout(state.timer); state.timer = null; state.epoch++;
    state.job = job;
    state.draft = {root_directory_token: job.root_directory_token || "",
      normalize_result_widths: job.normalize_result_widths === true,
      parts: clone(job.parts || []).map((p) => ({...p, enabled: p.enabled !== false, groups: p.groups || [], selected: p.selected || []}))};
    state.expandedParts = new Set([...state.expandedParts].filter((id) => state.draft.parts.some((p) => p.id === id)));
    for (const kind of ["result", "observe"]) {
      const counts = state.draft.parts.filter((p) => p.kind === kind && p.groups.length >= 1 && p.groups.length <= 6).map((p) => p.groups.length);
      state.modelCounts[kind] = new Set(counts.length ? counts : [1]);
    }
    state.resultSampleCounts.clear();
    state.dirty = false; state.pollError = false; state.page = 0; state.invalidation = Promise.resolve();
    try { state.customNames = new Set(JSON.parse(sessionStorage.getItem(`experimental-split-names:${job.id}`) || "[]")); } catch { state.customNames = new Set(); }
    state.draft.parts.filter((p) => p.manual_filename).forEach((p) => state.customNames.add(p.id));
    (job.plan?.outputs || []).forEach((o) => {
      const part = state.draft.parts.find((p) => p.id === o.id);
      if (part?.directory_token) state.labels.set(part.directory_token, o.base_directory || o.directory);
    });
    $("job-select").value = job.id;
    history.replaceState(null, "", `${location.pathname}?job=${encodeURIComponent(job.id)}`);
    let renamed = false;
    if (!active()) state.draft.parts.filter((p) => p.kind === "result" || p.kind === "observe").forEach((part) => {
      const previous = part.filename; suggest(part);
      renamed ||= part.filename !== previous;
    });
    $("confirm-plan").checked = false; showErrors(); render();
    if (renamed) changed();
    schedulePoll();
  }
  function schedulePoll() {
    if (!active() || state.timer || state.pollError) return;
    const epoch = state.epoch;
    state.timer = setTimeout(() => { state.timer = null; poll(epoch); }, 1800);
  }
  async function poll(epoch = state.epoch) {
    if (epoch !== state.epoch) return;
    try {
      const job = validJob(await api(endpoint()));
      if (epoch !== state.epoch) return;
      adopt(job);
      if (job.status === "failed") feedback(job.error || "任务失败，请核对后重试。", true);
      if (job.status === "completed") feedback(job.exports?.length ? "实际导出已保存，请核对记录。" : "完成响应没有实际导出记录，请核对。", !job.exports?.length);
      if (job.status === "partial") feedback(`处理结束：成功保存${job.exports?.length || 0}份，失败${job.export_failures?.length || 0}份。请查看导出记录中的失败清单。`, true);
    } catch (error) { state.pollError = true; $("poll-retry").hidden = false; feedback(error.message, true); }
  }
  async function operation(action) {
    if (state.busy) return;
    state.busy = true; controls();
    try { await action(); } catch (error) { feedback(error.message, true); }
    finally { state.busy = false; controls(); }
  }
  const mayDiscard = () => !state.dirty || window.confirm("当前方案尚未保存，继续将丢弃改动，是否继续？");
  function changed() {
    if (!state.dirty && state.job.plan) {
      const url = endpoint("/invalidate");
      state.invalidation = api(url, {}).catch((error) => { feedback(error.message, true); throw error; });
      // Keep the failure observable at save/export without an unhandled rejection.
      state.invalidation.catch(() => {});
    }
    state.dirty = true; state.job.plan = null; $("confirm-plan").checked = false;
    showErrors(); renderPlan(); controls();
  }
  function controls() {
    const locked = state.busy || active();
    ["upload-button", "source-file", "open-job", "job-select", "refresh-history", "poll-retry"].forEach((id) => { $(id).disabled = state.busy; });
    $("plan-fields").disabled = locked || !state.job?.total_pages;
    $("plan-button").disabled = locked || !state.job?.total_pages;
    $("analyze-button").disabled = locked;
    const valid = completePlan(state.job?.plan) && !state.dirty && state.job.status === "ready";
    $("confirm-plan").disabled = locked || !valid;
    $("export-button").disabled = locked || !valid || !$("confirm-plan").checked;
    $("dirty-status").textContent = state.dirty ? "有未保存修改；旧计划失效，请重新保存检查。" : "当前设置已保存；导出仍需逐项核对确认。";
  }
  function labels(part) {
    return part.groups.flatMap((g, i) => Array.from({length: Math.max(0, Math.min(99, Number(g.count) || 0))}, (_, n) => `${String.fromCharCode(65 + i)}${String(n + 1).padStart(3, "0")}`));
  }
  function partSources(part) {
    const sources = (state.job.source_units || []).filter((u) => u.kind === part.kind);
    const source = sources.find((u) => u.id === part.source_id
      && u.start_page <= part.start_page && part.end_page <= u.end_page);
    const units = source ? [source] : sources.filter((u) =>
      u.start_page <= part.end_page && u.end_page >= part.start_page);
    units.sort((a, b) => a.start_page - b.start_page);
    return units;
  }
  function modelFilename(counts) {
    return `${counts.length} ${counts.length === 1 ? "model" : "models"}-${counts.join("")}.docx`;
  }
  function suggest(part) {
    if (part.manual_filename || state.customNames.has(part.id)) return;
    if (part.kind === "items") { part.filename = "type evaluation items.docx"; return; }
    part.filename = modelFilename(part.groups.map((g) => g.count));
  }
  function resultFolder(part) {
    const sources = partSources(part);
    if (sources.length !== 1 || !/^[0-9]+$/.test(String(sources[0].sequence ?? "")) || !sources[0].project?.trim()) return "";
    return sources[0].folder_name || `${sources[0].sequence}.${(sources[0].display_project || sources[0].project).replace(/[/\\:*?"<>|\x00-\x1f]/g, "_")}`;
  }
  function reconfigure(part, previousLabels) {
    const next = labels(part), old = new Set(previousLabels), selected = new Set(part.selected);
    part.selected = next.filter((label) => !old.has(label) || selected.has(label));
    suggest(part); changed(); renderParts();
  }
  function resultSampleCounts(count) {
    if (!state.resultSampleCounts.has(count)) {
      const projects = new Map();
      for (const part of state.draft.parts.filter((p) => p.kind === "result")) {
        const sources = partSources(part);
        const key = JSON.stringify([sources.length === 1 ? sources[0].id : part.id, part.start_page, part.end_page]);
        if (!projects.has(key)) projects.set(key, []);
        projects.get(key).push(part);
      }
      const templates = [...projects.values()].map((parts) => parts.find((p) => p.groups.length === count)
        || parts.reduce((largest, p) => p.groups.length > largest.groups.length ? p : largest));
      state.resultSampleCounts.set(count, Array.from({length: count}, (_, i) => {
        const values = templates.map((part) => i < part.groups.length ? part.groups[i].count : 3);
        return !values.length ? 3 : values.every((value) => value === values[0]) ? values[0] : null;
      }));
    }
    return state.resultSampleCounts.get(count);
  }
  function resultSamplePlan(count) {
    const values = resultSampleCounts(count);
    const preview = el("p", {class: "muted", "aria-live": "polite"});
    const update = () => {
      preview.textContent = values.every((n) => Number.isInteger(n) && n >= 1 && n <= 99)
        && values.reduce((sum, n) => sum + n, 0) <= 120
        ? `默认文件名：${modelFilename(values)}；合计${values.reduce((sum, n) => sum + n, 0)}台。`
        : "请填写每个型号的台数（1～99整数，合计不超过120台）；不同项目原台数不同时不自动猜填。";
    };
    const inputs = values.map((value, index) => {
      const input = el("input", {type: "number", min: 1, max: 99, step: 1,
        value: value ?? "", placeholder: "请填写台数",
        "data-result-sample-count": count, "data-model-index": index, onchange: () => {
          if (state.busy || active()) { input.value = values[index] ?? ""; return; }
          values[index] = input.value === "" ? null : Number(input.value);
          update();
        }});
      return field(`${String.fromCharCode(65 + index)}组（第${index + 1}型号）台数`, input);
    });
    update();
    return el("fieldset", {class: "model-batch", "data-result-sample-plan": count},
      el("legend", {}, `${count}个型号方案：各型号台数`),
      el("div", {class: "grid three"}, ...inputs), preview);
  }
  function applyModelCounts(kind) {
    if (state.busy || active()) return;
    const counts = [...state.modelCounts[kind]].sort((a, b) => a - b);
    if (!counts.length || counts.some((n) => !Number.isInteger(n) || n < 1 || n > 6)) {
      feedback("请至少选择一个1～6的导出型号数量。", true); return;
    }
    if (kind === "result") {
      for (const count of counts) {
        const values = resultSampleCounts(count);
        if (values.length !== count || values.some((n) => !Number.isInteger(n) || n < 1 || n > 99)
            || values.reduce((sum, n) => sum + n, 0) > 120) {
          feedback(`${count}型号方案：请填写每个型号1～99的整数台数，合计不超过120台；未应用任何修改。`, true); return;
        }
      }
    }
    const projects = new Map();
    for (const part of state.draft.parts.filter((p) => p.kind === kind)) {
      const sources = partSources(part);
      if (sources.length !== 1) {
        feedback(`${part.filename}：请先核对范围，批量型号方案只支持单个明确的${kinds[kind]}来源。`, true); return;
      }
      const key = JSON.stringify([sources[0].id, part.start_page, part.end_page]);
      if (!projects.has(key)) projects.set(key, []);
      projects.get(key).push(part);
    }
    if (!projects.size) { feedback(`当前没有可配置的${kinds[kind]}。`, true); return; }
    const others = state.draft.parts.filter((p) => p.kind !== kind);
    const total = others.length + projects.size * counts.length;
    if (total > 100) {
      feedback(`应用后共有${total}张卡片，超过每任务100张的限制，请减少型号数量选项。`, true); return;
    }
    const variants = [];
    for (const parts of projects.values()) {
      const template = parts.reduce((largest, p) => p.groups.length > largest.groups.length ? p : largest);
      for (const count of counts) {
        const existing = parts.find((p) => p.groups.length === count);
        const part = clone(existing || template);
        if (!existing) {
          part.id = crypto.randomUUID(); part.manual_filename = false; part.filename = "";
          part.enabled = parts.some((p) => p.enabled);
        }
        if (!existing || kind === "result") {
          const old = new Set(labels(part)), selected = new Set(part.selected);
          part.groups = Array.from({length: count}, (_, i) => ({
            ...(part.groups[i] || {model: String.fromCharCode(65 + i).repeat(3), count: 3}),
            ...(kind === "result" ? {count: resultSampleCounts(count)[i]} : {})}));
          part.selected = labels(part).filter((label) => !old.has(label) || selected.has(label));
        }
        if (kind === "observe") part.selected = labels(part);
        const available = new Set(labels(part));
        if (part.groups.some((g) => !g.model.trim() || g.model !== g.model.trim()
              || g.model.length > 80 || /[\x00-\x1f]/.test(g.model)
              || !Number.isInteger(g.count) || g.count < 1 || g.count > 99)
            || new Set(part.groups.map((g) => g.model)).size !== part.groups.length
            || available.size > 120 || (kind === "result" && !part.selected.length)
            || new Set(part.selected).size !== part.selected.length
            || part.selected.some((label) => !available.has(label))) {
          feedback(`${template.filename}：${count}型号方案的型号、台数或参试编号无效，请先展开核对参数。`, true); return;
        }
        suggest(part);
        if (kind === "result" && !resultFolder(part)) {
          feedback(`${part.filename}：无法确定项目文件夹的序号或名称，请先核对源项目。`, true); return;
        }
        variants.push(part);
      }
    }
    const sampleSummary = kind === "result"
      ? `各方案台数：${counts.map((count) => `${count}型号=${resultSampleCounts(count).join("/")}台`).join("；")}。台数将应用到所有试验项目；保留型号名称和原参试取消勾选，新编号默认勾选，超出台数的旧编号移除。`
      : "新卡片沿用已有型号/台数，观察编号完整保留，新型号默认3台。";
    if (!window.confirm(`将为${projects.size}个${kinds[kind]}来源分别配置${counts.join("、")}个型号，共${variants.length}份方案。保留已有相同型号数的第一份卡片及其手动名称和目录，其他数量及重复卡片会从当前方案移除。${sampleSummary}原报告和历史文件不会删除。是否应用？`)) return;
    state.draft.parts = [...others, ...variants];
    const ids = new Set(state.draft.parts.map((p) => p.id));
    state.customNames = new Set([...state.customNames].filter((id) => ids.has(id)));
    state.expandedParts = new Set([...state.expandedParts].filter((id) => others.some((p) => p.id === id)));
    rememberNames(); changed(); renderParts();
    feedback(`已配置${projects.size}个${kinds[kind]}来源、${variants.length}份方案；请展开核对并重新保存检查计划。`);
  }
  function batchModelControls(parts, kind) {
    const hint = el("p", {class: "muted"});
    const apply = button(kind === "result" ? "应用到所有试验项目" : "应用到所有观察记录", () => applyModelCounts(kind));
    apply.setAttribute("data-apply-model-counts", kind);
    const selected = state.modelCounts[kind];
    const samplePlans = kind === "result" ? el("div", {"data-result-sample-plans": ""}) : null;
    const update = () => {
      hint.textContent = `已选：${[...selected].sort((a, b) => a - b).join("、") || "无"}个型号。多选会为每个来源分别生成独立方案；仅调整${kinds[kind]}，应用前确认替换范围。`;
      apply.disabled = !parts.length || !selected.size;
      if (samplePlans) samplePlans.replaceChildren(...[...selected].sort((a, b) => a - b)
        .filter((n) => Number.isInteger(n) && n >= 1 && n <= 6).map(resultSamplePlan));
    };
    const choices = Array.from({length: 6}, (_, i) => {
      const count = i + 1;
      const checkbox = el("input", {type: "checkbox", "data-model-count": count, "data-model-kind": kind,
        checked: selected.has(count), onchange: () => {
          if (state.busy || active()) { checkbox.checked = selected.has(count); return; }
          if (checkbox.checked) selected.add(count); else selected.delete(count);
          update();
        }});
      return el("label", {class: "check"}, checkbox, `${count}个型号`);
    });
    update();
    return el("fieldset", {class: "model-batch"}, el("legend", {}, "导出型号数量（可多选）"),
      el("div", {class: "checks"}, ...choices), hint,
      kind === "result" ? el("p", {class: "muted"},
        "分别设置每个方案中各型号的配置台数，再点击应用。台数设置不是实际参试勾选数量；应用前不修改卡片或已保存计划。") : null,
      samplePlans, apply);
  }
  function modelConfig(part) {
    if (part.kind === "items") return el("p", {class: "muted"}, "items保持原内容，不调整型号或结果。");
    const count = el("input", {type: "number", min: 1, max: 6, step: 1, value: part.groups.length, onchange: () => {
      const n = Number(count.value);
      if (!Number.isInteger(n) || n < 1 || n > 6) { feedback("型号数须为1～6。", true); return; }
      const old = labels(part);
      part.groups = Array.from({length: n}, (_, i) => part.groups[i] || {model: String.fromCharCode(65 + i).repeat(3), count: 3});
      reconfigure(part, old);
    }});
    return el("div", {class: "model-config"}, field("本卡片型号数量", count),
      ...part.groups.map((group, i) => {
        const model = el("input", {type: "text", value: group.model, maxlength: 80, onchange: () => { group.model = model.value; suggest(part); changed(); renderParts(); }});
        const number = el("input", {type: "number", min: 1, max: 99, step: 1, value: group.count, onchange: () => {
          const old = labels(part); group.count = Number(number.value); reconfigure(part, old);
        }});
        return el("div", {class: "model-row"}, field(`${String.fromCharCode(65 + i)}组型号`, model), field("本型号台数", number));
      }),
      el("p", {class: "muted"}, part.kind === "result" ? "实际参试编号（每个子条件均覆盖选中编号，全部实测格清空）" : "观察顶部完整编号列表，每组单独一行"),
      el("div", {class: "sample-choice"}, ...labels(part).map((label) => {
        if (part.kind !== "result") return el("span", {}, label);
        const checkbox = el("input", {type: "checkbox", checked: part.selected.includes(label), onchange: () => {
          const selected = new Set(part.selected);
          if (checkbox.checked) selected.add(label); else selected.delete(label);
          part.selected = labels(part).filter((x) => selected.has(x)); suggest(part); changed(); renderParts();
        }});
        return el("label", {class: "check"}, checkbox, label);
      })));
  }
  function renderPart(part) {
    const enabled = el("input", {type: "checkbox", checked: part.enabled, onchange: () => { part.enabled = enabled.checked; changed(); renderParts(); }});
    const category = el("select", {onchange: () => {
      part.kind = category.value; part.source_id = ""; part.groups = part.kind === "items" ? [] : [{model: "AAA", count: 3}];
      part.selected = labels(part); suggest(part); changed(); renderParts();
    }}, ...Object.entries(kinds).map(([kind, title]) => el("option", {value: kind}, title)));
    category.value = part.kind;
    const page = (key, title) => {
      const input = el("input", {type: "number", min: 1, max: state.job.total_pages, value: part[key], oninput: () => {
        part[key] = Number(input.value); suggest(part); filename.value = part.filename; changed();
      }});
      return field(title, input);
    };
    const filename = el("input", {type: "text", value: part.filename, oninput: () => {
      part.filename = filename.value; part.manual_filename = true; state.customNames.add(part.id); rememberNames(); changed();
    }});
    const base = part.directory_token ? state.labels.get(part.directory_token) || "已授权项目父目录"
      : `${state.labels.get(state.draft.root_directory_token) || state.job.root_directory_label || "根目录"}/type evaluation results`;
    const directory = part.kind === "result" ? `${base}/${resultFolder(part) || "待核对项目目录"}`
      : part.directory_token ? state.labels.get(part.directory_token) || "已授权单份目录（最终计划核对路径）" : "根目录下的类别子文件夹";
    const parameters = el("div", {id: `part-config-${part.id}`, "data-part-config": part.id,
      hidden: !state.expandedParts.has(part.id)},
      el("div", {class: "part-fields"}, field("类别", category, "kind-field"), page("start_page", "起始物理页"), page("end_page", "结束物理页"), field("输出文件名（可手动编辑）", filename, "filename-field")),
      modelConfig(part), el("div", {class: "part-directory"}, el("p", {class: "directory-path"}, directory),
        button(part.kind === "result" ? "项目父目录…" : "单份目录…", () => pickFolder(part)), button("使用根目录", () => { part.directory_token = ""; changed(); renderParts(); }),
        button("恢复建议名", () => { part.manual_filename = false; state.customNames.delete(part.id); rememberNames(); suggest(part); changed(); renderParts(); }),
        button("预览起始页", () => selectPage(part.start_page), "quiet")));
    const toggle = button(parameters.hidden ? "展开参数" : "收起参数", () => {
      parameters.hidden = !parameters.hidden;
      if (parameters.hidden) state.expandedParts.delete(part.id); else state.expandedParts.add(part.id);
      toggle.textContent = parameters.hidden ? "展开参数" : "收起参数";
      toggle.setAttribute("aria-expanded", String(!parameters.hidden));
    }, "secondary small");
    toggle.setAttribute("data-toggle-part", part.id);
    toggle.setAttribute("aria-expanded", String(!parameters.hidden));
    toggle.setAttribute("aria-controls", `part-config-${part.id}`);
    const title = partSources(part).map((u) => u.display_project || u.project).filter(Boolean).join("、");
    return el("article", {class: `part-row${part.enabled ? "" : " excluded"}`, "data-part-id": part.id},
      el("div", {class: "row-heading"}, el("label", {class: "check"}, enabled, el("strong", {}, title || part.project || kinds[part.kind])),
        part.kind !== "items" ? el("span", {class: "pill"}, `${part.groups.length}个型号`) : null,
        toggle,
        part.kind !== "items" ? button("复制方案", () => {
          const copy = clone(part); copy.id = crypto.randomUUID(); copy.filename = ""; copy.manual_filename = false; suggest(copy);
          state.draft.parts.splice(state.draft.parts.indexOf(part) + 1, 0, copy); changed(); renderParts();
        }) : null,
        button("删除卡片", () => { if (window.confirm("只删除本卡片，不删除历史文件？")) { state.draft.parts = state.draft.parts.filter((p) => p !== part); changed(); renderParts(); } }, "danger small")),
      parameters);
  }
  function renderParts() {
    $("parts").replaceChildren(...Object.entries(kinds).map(([kind, title]) => {
      const parts = state.draft.parts.filter((p) => p.kind === kind);
      const heading = el("h3", {}, `${title}（${parts.length}份）`);
      if (kind !== "result") return el("section", {class: "part-group"}, heading,
        kind === "observe" ? el("p", {class: "muted"},
          "技术要求按Word实际剩余空间跨页接排，不因整项要求放不下而强制整项换页。导出前自动复核内容完整性、续页表头和编号/叉号；可安全修正的版式重新排版核验，仍不通过则明确报错，不交付有疑问的文件。") : null,
        kind === "observe" ? batchModelControls(parts, kind) : null, ...parts.map(renderPart));
      const selectedCount = parts.filter((p) => p.enabled).length;
      const selectAll = el("input", {type: "checkbox", "data-select-all": "result",
        checked: parts.length > 0 && selectedCount === parts.length, disabled: !parts.length,
        onchange: () => {
          parts.forEach((p) => { p.enabled = selectAll.checked; });
          changed(); renderParts();
        }});
      selectAll.indeterminate = selectedCount > 0 && selectedCount < parts.length;
      const uniformWidths = el("input", {type: "checkbox", "data-normalize-result-widths": "",
        checked: state.draft.normalize_result_widths, onchange: () => {
          if (state.busy || active()) { renderParts(); return; }
          state.draft.normalize_result_widths = uniformWidths.checked;
          changed();
        }});
      return el("section", {class: "part-group"},
        el("div", {class: "row-heading"}, heading,
          el("label", {class: "check"}, selectAll, "全选试验结果")),
        el("p", {class: "muted"}, "默认文件名按型号顺序附加各组配置台数，例如1 model-3.docx、2 models-33.docx、5 models-33322.docx；不按参试勾选数量命名，手动名称保留。"),
        el("label", {class: "check"}, uniformWidths, "统一结果表列宽（总宽16.2cm）"),
        el("p", {class: "muted"}, "勾选后：序号1.3、项目名称2.5、技术要求6.2、样机编号3.7、试验结果2.5cm；项目名称内部子列等分，技术要求内部子列按源宽比例分配；指示装置不调整。不勾选则保留源文件宽度，仅影响试验结果。"),
        batchModelControls(parts, kind),
        ...parts.map(renderPart));
    }));
  }
  function selectPage(number) {
    const index = (state.job.pages || []).findIndex((p) => p.number === Number(number));
    if (index < 0) { feedback("源页预览不可用，请打开PDF核对。", true); return; }
    state.page = index; showPage(); $("preview-section").scrollIntoView({block: "start", behavior: "smooth"});
  }
  function showPage() {
    const pages = state.job?.pages || [], page = pages[state.page];
    if (!page) return;
    $("preview-page").value = String(state.page);
    $("previous-page").disabled = state.page === 0; $("next-page").disabled = state.page === pages.length - 1;
    $("page-caption").textContent = `原报告物理第${page.number}页 · 点击图片放大`;
    $("page-text").textContent = page.text || "未返回文字，请核对原生PDF。";
    const url = safeURL(page.image_url);
    $("page-image-link").hidden = !url; $("image-error").hidden = Boolean(url);
    if (url) { $("page-image-link").href = url; $("page-image").src = url; $("page-image").alt = `源报告物理第${page.number}页`; }
  }
  function completePlan(plan) {
    if (!Number.isSafeInteger(plan?.revision) || !Array.isArray(plan.outputs) || !plan.outputs.length || !state.draft) return false;
    if ((plan.normalize_result_widths ?? false) !== state.draft.normalize_result_widths) return false;
    const included = state.draft.parts.filter((p) => p.enabled), ids = new Set();
    return plan.outputs.length === included.length && plan.outputs.every((o) => {
      const p = included.find((p) => p.id === o.id);
      if (!p || ids.has(o.id)) return false;
      ids.add(o.id);
      return o.kind === p.kind && o.filename === p.filename && Boolean(o.directory && o.full_path)
        && Number.isSafeInteger(o.start_page) && Number.isSafeInteger(o.end_page) && o.start_page >= 1
        && o.end_page >= o.start_page && o.end_page <= state.job.total_pages
        && (p.kind === "items" || (p.groups.length >= 1 && p.groups.length <= 6
          && Array.isArray(o.models) && Array.isArray(o.selected) && Boolean(o.structure)));
    });
  }
  const summary = (o) => (o.models || []).map((g) => `${g.model}（${g.labels.length}台）：${g.labels.join("、")}`).join("；");
  function renderPlan() {
    const plan = !state.dirty && state.job?.plan;
    $("plan-section").hidden = !plan;
    if (!plan) return;
    $("plan-title").textContent = `3. 核对最终计划 · 版本${plan.revision}`;
    $("plan-adjustments").replaceChildren(...(plan.adjustments || []).map((a) => el("p", {class: "notice"}, `${a.id}：请求P${a.requested_start}～P${a.requested_end} → P${a.start_page}～P${a.end_page}；${a.message}`)));
    $("plan-outputs").replaceChildren(el("table", {}, el("thead", {}, el("tr", {}, ...["类别/源边界", "型号和编号", "参试子集/结构", "文件及目录", "格式调整"].map((t) => el("th", {}, t)))),
      el("tbody", {}, ...plan.outputs.map((o) => el("tr", {},
        el("td", {}, `${kinds[o.kind]} P${o.start_page}～P${o.end_page}`,
          el("p", {}, (o.source_projects || []).map((p) => `${p.display_name || p.name} [正文块${p.start_block}～${p.end_block_exclusive}，右边界不含]`).join("；"))),
        el("td", {}, summary(o) || "原样"),
        el("td", {}, `${(o.selected || []).join("、")} ${o.structure?.conditions ? `共${o.structure.conditions}个条件/${o.structure.rows}行，实测全空` : ""}`),
        el("td", {}, o.filename, el("p", {}, o.full_path)), el("td", {}, o.format_adjustments))))));
    if (!completePlan(plan)) $("plan-adjustments").append(el("p", {class: "notice error"}, "计划不完整，请重新保存，禁止导出。"));
  }
  function render() {
    const job = state.job; $("job-section").hidden = !job;
    if (!job) { controls(); return; }
    $("job-title").textContent = job.name || job.id; $("job-status").textContent = statuses[job.status];
    $("job-stage").textContent = `${job.stage || ""}${job.total_pages ? ` · 共${job.total_pages}个物理页` : ""}`;
    $("job-error").hidden = !job.error; $("job-error").textContent = job.error || "";
    $("job-warnings").replaceChildren(...(job.warnings || []).map((w) => el("p", {class: "notice"}, w)));
    $("analyze-button").hidden = !["uploaded", "failed"].includes(job.status);
    $("poll-retry").hidden = !state.pollError; $("source-pdf").hidden = !job.total_pages; $("source-pdf").href = endpoint("/source.pdf");
    $("editor-section").hidden = !job.total_pages;
    $("root-directory").textContent = state.labels.get(state.draft.root_directory_token) || job.root_directory_label || "尚未选择";
    renderParts(); renderPlan();
    $("preview-section").hidden = !job.pages?.length;
    $("preview-page").replaceChildren(...(job.pages || []).map((p, i) => el("option", {value: i}, `物理页${p.number}`))); showPage();
    $("exports-section").hidden = !job.exports?.length && !job.export_failures?.length && job.status !== "completed";
    $("export-list").replaceChildren(...(job.exports || []).map((file) => el("div", {class: "export-entry"}, el("strong", {}, file.filename),
      el("p", {}, file.path), el("p", {}, `输出${file.checks?.output_pages || "未知"}页；${file.checks?.review_note || ""}`),
      safeURL(file.url) ? (window.workbenchDownloads ? window.workbenchDownloads.link(file.url, "下载可编辑Word", file.filename)
        : el("p", {class: "notice error", role: "alert"}, "下载组件未载入，请强制刷新后重试。"))
        : el("p", {}, "没有有效下载链接"))));
    if (job.export_failures?.length) {
      $("export-list").prepend(el("p", {class: "notice error"},
        `成功保存${job.exports?.length || 0}份，失败${job.export_failures.length}份；以下失败文件未交付。`));
      $("export-list").append(el("h3", {}, "失败文件及原因"),
        ...job.export_failures.map((file) => el("div", {class: "export-entry"},
          el("strong", {}, file.filename), el("p", {}, file.path),
          el("p", {class: "notice error"}, `${file.phase || "处理"}失败：${file.error}`))));
    }
    if (job.status === "completed" && !job.exports?.length) $("export-list").append(el("p", {class: "notice error"}, "没有实际导出记录，不能视为交付完成。"));
    controls();
  }
  function payload() {
    return {root_directory_token: state.draft.root_directory_token,
      normalize_result_widths: state.draft.normalize_result_widths,
      parts: state.draft.parts.map((p) => ({id: p.id, source_id: p.source_id || "", kind: p.kind, enabled: p.enabled,
        start_page: Number(p.start_page), end_page: Number(p.end_page), filename: p.filename, manual_filename: Boolean(p.manual_filename),
        directory_token: p.directory_token || "", groups: p.groups.map((g) => ({model: g.model, count: Number(g.count)})), selected: p.selected}))};
  }
  function validate() {
    const errors = [];
    if (state.draft.parts.length > 100) errors.push("每任务最多100张卡片，请减少方案。");
    if (!state.draft.parts.some((p) => p.enabled)) errors.push("至少勾选一个启用方案。");
    for (const p of state.draft.parts) {
      if (!p.enabled) continue;
      if (!Number.isInteger(p.start_page) || !Number.isInteger(p.end_page) || p.start_page < 1 || p.end_page < p.start_page || p.end_page > state.job.total_pages) errors.push(`${p.filename}：页码无效。`);
      if (!p.filename || !/\.docx$/i.test(p.filename) || /[/\\:*?"<>|\x00-\x1f]/.test(p.filename) || p.filename !== p.filename.trim() || p.filename.includes("待核对序号")) errors.push("请填写有效且已核对序号的DOCX文件名。");
      if (!p.directory_token && !state.draft.root_directory_token) errors.push("请选择根目录或单份目录。");
      if (p.kind === "result" && !resultFolder(p)) errors.push(`${p.filename}：请核对单个源项目的序号、名称和范围，才能生成项目文件夹。`);
      if (p.kind !== "items" && (!p.groups.length || p.groups.length > 6 || p.groups.some((g) => !g.model.trim() || !Number.isInteger(g.count) || g.count < 1 || g.count > 99) || labels(p).length > 120)) errors.push("型号/台数无效（最多6个型号，每组1～99，总计最多120台）。");
      if (p.kind === "result" && !p.selected.length) errors.push("结果方案至少勾选一台参试样机。");
    }
    return errors;
  }
  async function loadHistory() {
    const jobs = await api(API);
    if (!Array.isArray(jobs)) throw new Error("任务列表响应无效。");
    $("job-select").replaceChildren(el("option", {value: ""}, "请选择试验任务"), ...jobs.map((j) => el("option", {value: j.id}, `${j.name} · ${statuses[j.status]} · ${j.created_at}`)));
    $("job-select").value = state.job?.id || "";
  }
  async function pickFolder(part = null) {
    await operation(async () => {
      const result = await api(endpoint("/folder"), {part_id: part?.id || null});
      if (result.cancelled) { feedback("已取消，原目录保持不变。"); return; }
      if (!result.token || !result.path) throw new Error("目录授权响应无效。");
      state.labels.set(result.token, result.path);
      if (part) part.directory_token = result.token; else state.draft.root_directory_token = result.token;
      changed(); render();
    });
  }
  $("plan-form").addEventListener("submit", (event) => {
    event.preventDefault(); const errors = validate(); showErrors(errors);
    if (errors.length) return;
    operation(async () => {
      state.job.plan = null; $("confirm-plan").checked = false; renderPlan();
      // A failed invalidation can be recovered by saving: /plan invalidates the
      // previous server plan before any structural validation.
      await state.invalidation.catch(() => {});
      try {
        const job = validJob(await api(endpoint("/plan"), payload()));
        if (!completePlan(job.plan)) throw new Error("服务器未返回完整计划，禁止导出。");
        state.job = job; state.dirty = false; state.invalidation = Promise.resolve(); render();
        feedback("计划已保存，请核对所有型号、参试编号、边界和格式调整。");
      } catch (error) { state.dirty = true; controls(); showErrors([error.message]); throw error; }
    });
  });
  $("export-button").addEventListener("click", () => {
    if ($("export-button").disabled || !completePlan(state.job.plan) || state.dirty) return;
    const plan = state.job.plan;
    if (!window.confirm(`确认计划版本${plan.revision}？\n\n${plan.outputs.map((o) => `${o.filename}｜P${o.start_page}～P${o.end_page}｜${summary(o)}｜参试${o.selected.join("、")}｜${o.full_path}`).join("\n")}\n\n实测格全部清空，接受全部边界及V1格式调整；报告端兼容尚待适配。`)) return;
    operation(async () => { await state.invalidation; adopt(await api(endpoint("/export"), {plan_revision: plan.revision, confirm_adjustments: true, configuration: payload()})); });
  });
  $("upload-form").addEventListener("submit", (event) => {
    event.preventDefault(); const file = $("source-file").files[0];
    if (!file?.size || !/\.docx$/i.test(file.name)) { feedback("请选择非空DOCX。", true); return; }
    if (!mayDiscard()) return;
    operation(async () => { const data = new FormData(); data.append("file", file); adopt(await api(`${API}/upload`, data)); $("source-file").value = ""; await loadHistory(); });
  });
  $("open-job").addEventListener("click", () => { const id = $("job-select").value; if (id && mayDiscard()) operation(async () => { adopt(await api(`${API}/${encodeURIComponent(id)}`)); }); });
  $("refresh-history").addEventListener("click", () => operation(loadHistory));
  $("analyze-button").addEventListener("click", () => { if (mayDiscard()) operation(async () => adopt(await api(endpoint("/analyze"), {}))); });
  $("poll-retry").addEventListener("click", () => { state.pollError = false; poll(); });
  $("root-picker").addEventListener("click", () => pickFolder());
  $("confirm-plan").addEventListener("change", controls);
  $("add-part").addEventListener("click", () => {
    const kind = $("new-kind").value, part = {id: crypto.randomUUID(), kind, enabled: true, start_page: 1, end_page: 1,
      directory_token: "", groups: kind === "items" ? [] : [{model: "AAA", count: 3}], selected: []};
    part.selected = labels(part); suggest(part); state.draft.parts.push(part); changed(); renderParts();
  });
  $("previous-page").addEventListener("click", () => { state.page--; showPage(); });
  $("next-page").addEventListener("click", () => { state.page++; showPage(); });
  $("preview-page").addEventListener("change", () => { state.page = Number($("preview-page").value); showPage(); });
  $("page-image").addEventListener("error", () => { $("image-error").hidden = false; });
  window.addEventListener("beforeunload", (event) => { if (state.dirty) { event.preventDefault(); event.returnValue = ""; } });
  operation(async () => { await loadHistory(); const id = new URLSearchParams(location.search).get("job"); if (id) adopt(await api(`${API}/${encodeURIComponent(id)}`)); else feedback("请选择源报告或恢复独立试验任务。"); });
})();
