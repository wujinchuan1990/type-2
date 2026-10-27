"use strict";

(() => {
  const $ = (id) => document.getElementById(id);
  const clone = (value) => JSON.parse(JSON.stringify(value));
  const environmentFields = {
    environment_temperature: "型式评价环境温度",
    environment_humidity: "型式评价环境相对湿度"
  };
  const automaticOverviewFields = new Set(["test_location", "evaluation_period"]);
  const displayOptions = ["机械", "电动机械", "电子"];
  const measurementColumns = [
    ["measurement_parameter", "测量参数"], ["parameter_unit", "测量参数单位"],
    ["measurement_range", "测量区间"], ["display_digits", "显示位数"],
    ["metrological_performance", "计量性能指标"]
  ];
  const steps = [
    ["资料与任务", "建立本次报告，导入原始记录与当前大纲。所有操作仅连接本机服务。"],
    ["基本信息总览", "填写新值则替换，未填写则保留模板原文；沿用内容仍需核对，不算本次已确认数据。"],
    ["素材与章节", "选择一览表和观察模块，按文件夹展开并选择结果素材方案。"],
    ["仪器设备库", "维护可复用设备库，明确添加本次快照；库内修改不会自动影响本次报告。"],
    ["项目与人员", "补齐每个样机的真实试验数据，指定设备、评价人员和复核人员。"],
    ["附件2照片", "上传整体外形、内部结构、显示部分与零部件照片；零部件名称取自当前大纲。图片居中、等比例插入报告。"],
    ["预览与生成", "生成可编辑 Word，下载历史版本并查看相邻版本改动。"]
  ];
  const state = {
    catalog: { assets: [], overview: [], hidden_overview_fields: [], warnings: [], photo_slots: [] }, equipment: [], uploads: [], photoFiles: [], principles: [],
    sampleSchemes: [], sampleSchemeId: "", sampleSchemeName: "", sampleSchemeIssues: [],
    materialSchemes: [], materialSchemeId: "", materialSchemeName: "", materialSchemeType: "",
    task: null, step: 0, dirty: false, busy: false, ready: false, epoch: 0,
    exports: [], validation: null, loadErrors: [], orderVerified: false, nativeWord: false, paginationStrategy: "standard",
    loadingComparisons: new Set(), materialFolders: new Set(), materialQuery: "",
    equipmentQuery: "", equipmentPage: 1, equipmentPageSize: 10, equipmentSelectedOnly: false,
    equipmentEditorDirty: false, resultsEditorDirty: false,
    clockDrafts: new Map(), autoEndTimes: new Map(), autoStartTimes: new Map(),
    environmentSignature: null, environmentTimer: null,
    evaluationSignature: null, evaluationTimer: null,
    sealingSignature: null, sealingTimer: null, sealingPending: false, sealingError: ""
  };
  state.generationJobs = [];
  let draftTimer = null, jobTimer = null, draftKey = null;
  const recoveryPrefix = "report-recovery-v1:";

  function persistDraft() {
    if (!state.task || (!state.dirty && !state.resultsEditorDirty && !state.equipmentEditorDirty && !state.clockDrafts.size)) return;
    try {
      if (!draftKey) draftKey = recoveryPrefix + crypto.randomUUID();
      localStorage.setItem(draftKey, JSON.stringify({
        format: 1, task: state.task, step: state.step, updated: new Date().toISOString(),
        clockDrafts: [...state.clockDrafts], resultDraft: state.recoveryResultDraft || null,
        equipmentDraft: state.recoveryEquipmentDraft || null
      }));
      if ($("recovery-state")) $("recovery-state").textContent = "未保存草稿已暂存本机浏览器；仍需保存任务。";
    } catch (error) {
      if ($("recovery-state")) $("recovery-state").textContent = `暂存失败：${error.message}。请立即保存任务。`;
      feedback("未保存草稿暂存失败，请立即保存任务：" + error.message, true);
    }
  }
  function scheduleDraft() {
    clearTimeout(draftTimer);
    draftTimer = setTimeout(persistDraft, 800);
  }
  function clearDraft() {
    clearTimeout(draftTimer);
    if (draftKey) {
      try { localStorage.removeItem(draftKey); }
      catch (error) { feedback("已保存任务，但清理浏览器暂存失败：" + error.message, true); return; }
    }
    draftKey = null;
    if ($("recovery-state")) $("recovery-state").textContent = "";
  }
  function recoveryEntry(key) {
    let raw = "";
    try {
      raw = localStorage.getItem(key);
      const data = JSON.parse(raw);
      if (data?.format !== 1 || typeof data.task?.name !== "string" ||
          !Array.isArray(data.task.projects) || !Number.isInteger(data.step) ||
          data.step < 0 || data.step >= steps.length || !Number.isFinite(Date.parse(data.updated)))
        throw new Error("暂存格式、步骤或时间无效");
      if (data.clockDrafts != null && (!Array.isArray(data.clockDrafts) ||
          data.clockDrafts.some(pair => !Array.isArray(pair) || pair.length !== 2 ||
            typeof pair[0] !== "string" || !pair[1] || typeof pair[1] !== "object" ||
            !["base", "date", "hour", "minute"].every(field => typeof pair[1][field] === "string") ||
            (pair[1].seconds != null && typeof pair[1].seconds !== "string"))))
        throw new Error("暂存的未完成时间格式无效");
      return {key, raw, data};
    } catch (error) { return {key, raw, error: error.message}; }
  }
  function openRecovery() {
    const dialog = el("dialog", {}, el("h2", {}, "恢复未保存草稿"));
    const content = el("div");
    dialog.append(content, button("关闭", () => { dialog.close(); dialog.remove(); }));
    try {
      const entries = Object.keys(localStorage).filter(key => key.startsWith(recoveryPrefix))
        .map(recoveryEntry)
        .sort((a, b) => (b.data?.updated || "").localeCompare(a.data?.updated || ""));
      if (!entries.length) content.append(notice("没有未保存草稿。"));
      entries.forEach(({key, raw, data, error}) => content.append(el("section", {class: "card"},
        el("strong", {}, data ? data.task.name : "损坏的暂存"),
        data ? el("p", {}, `暂存时间：${new Date(data.updated).toLocaleString()}`) :
          notice(`无法恢复：${error}。可下载原始内容核对或删除；其他暂存不受影响。`, "error"),
        data ? button("恢复到页面（不覆盖已保存任务）", async () => {
          try {
          const checked = await api("/api/tasks/recovery-check", {
            task: data.task, result: data.resultDraft || null, equipment: data.equipmentDraft || null
          });
          if (data.task.id) {
              const saved = await api(`/api/tasks/${encodeURIComponent(data.task.id)}`);
              if (saved.revision !== data.task.revision) {
                content.prepend(notice("已保存任务已有新版本，不能直接套用旧草稿。请下载暂存内容核对，未覆盖任何任务。", "error"));
                return;
              }
          }
          if (!discardAllowed()) return;
          persistDraft();
          state.task = normalizeTask(checked.task); state.step = data.step; state.dirty = true;
          state.clockDrafts = new Map(data.clockDrafts || []); state.epoch += 1;
          state.exports = []; state.validation = null; state.orderVerified = false; draftKey = null;
          dialog.close(); dialog.remove(); render();
          if (checked.result) {
            const project = state.task.projects.find(p => p.id === checked.result.id);
            openResultsEditor(project, document.activeElement, checked.result);
          }
          if (checked.equipment)
            openEquipmentEditor(checked.equipment, render, true);
          scheduleDraft();
          feedback("已恢复草稿。尚未保存或覆盖任务；生成前请重新核对一览表顺序。");
          } catch (error) { content.prepend(notice("恢复失败：" + error.message, "error")); feedback("恢复失败：" + error.message, true); }
        }) : null,
        button("下载暂存内容", () => {
          const url = URL.createObjectURL(new Blob([raw], {type:"application/json"}));
          const link = el("a", {href:url, download:"未保存草稿.json"});
          document.body.append(link); link.click(); link.remove(); URL.revokeObjectURL(url);
        }),
        button("删除这份暂存", () => {
          if (!window.confirm("删除这份浏览器暂存？不会删除已保存任务。")) return;
          try { localStorage.removeItem(key); dialog.close(); dialog.remove(); openRecovery(); }
          catch (error) { content.prepend(notice(error.message, "error")); }
        }, "danger small")
      )));
    } catch (error) { content.append(notice("无法读取浏览器暂存：" + error.message, "error")); }
    document.body.append(dialog); dialog.showModal();
  }
  async function openBackups() {
    const dialog = el("dialog", {}, el("h2", {}, "本机备份与恢复校验"));
    const content = el("div", {}, notice("备份包含业务资料，请下载后存放到可靠位置。恢复校验只使用独立目录，不覆盖当前系统。备份期间暂缓其他操作，Word任务须先完成。"));
    const status = el("p", {role:"status"});
    let working = false;
    const setWorking = value => { working = value; create.disabled = value; upload.disabled = value; };
    const create = button("生成备份包", async () => {
      if (working) return;
      setWorking(true);
      try {
        status.textContent = "正在生成一致性备份，请稍候…";
        const item = await api("/api/backups", {});
        content.append(downloadLink(item.url, `下载备份包（${(item.bytes / 1048576).toFixed(1)}MB）`));
        status.textContent = "备份包已生成，请下载并妥善保存。";
      } catch (error) { status.textContent = "备份失败：" + error.message; }
      finally { setWorking(false); }
    });
    const upload = el("input", {type:"file", accept:".zip",
      onchange: async event => {
        if (working || !event.target.files[0]) return;
        setWorking(true);
        const data = new FormData(); data.append("file", event.target.files[0]);
        try { status.textContent = "正在校验备份…"; const result = await api("/api/backups/verify", data);
          status.textContent = `${result.message} 共${result.files}个文件、${result.tasks}个任务。校验编号：${result.verification_id}`;
        } catch (error) { status.textContent = "校验失败：" + error.message; }
        finally { setWorking(false); event.target.value = ""; }
      }});
    content.append(status);
    dialog.append(content, create, el("label", {}, "上传备份包进行独立恢复校验", upload),
      button("关闭", () => { dialog.close(); dialog.remove(); }));
    document.body.append(dialog); dialog.showModal();
    try {
      const items = await api("/api/backups");
      items.slice(0, 5).forEach(item => content.append(el("p", {},
        downloadLink(item.url, "下载备份 · " + new Date(item.created).toLocaleString()))));
    } catch (error) { content.append(notice(error.message, "error")); }
  }
  function generationStatus() {
    const root = el("div", {id: "generation-status", "aria-live": "polite"});
    if (state.generationJobError) {
      root.append(notice("生成状态读取失败：" + state.generationJobError, "error"),
        button("重试读取", pollGenerationJobs));
    }
    if (state.task?.id) root.append(state.generationJobs
      .filter(job => job.task_id === state.task.id && job.status !== "completed")
      .map(job => notice(`版本 ${job.task_revision} · ${job.stage}${job.error ? "：" + job.error : ""}`,
        ["failed", "interrupted"].includes(job.status) ? "error" : "")));
    return root;
  }
  function paintGenerationJobs() {
    const root = document.getElementById("generation-status");
    if (root) root.replaceChildren(...generationStatus().childNodes);
  }
  async function pollGenerationJobs() {
    clearTimeout(jobTimer);
    try {
      state.generationJobs = await api("/api/generation-jobs");
      state.generationJobError = "";
      paintGenerationJobs();
      let added = false;
      for (const job of state.generationJobs) {
        if (job.export_id && job.task_id === state.task?.id && !state.exports.some(output => output.id === job.export_id)) {
          const output = await api(`/api/exports/${job.export_id}`);
          if (job.task_id !== state.task?.id) continue;
          state.exports.unshift({...output, epoch: !state.dirty && output.revision === state.task.revision ? state.epoch : -1});
          added = true;
        }
      }
      if (added && state.step === 6 && !state.busy) $("panel").replaceChildren(renderPreview());
      jobTimer = setTimeout(pollGenerationJobs, state.generationJobs.some(j => ["queued", "running"].includes(j.status)) ? 1500 : 15000);
    } catch (error) {
      state.generationJobError = error.message;
      paintGenerationJobs();
    }
  }
  let uid = 0;

  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    Object.entries(attrs).forEach(([key, value]) => {
      if (key.startsWith("on")) node.addEventListener(key.slice(2), value);
      else if (key === "class") node.className = value;
      else if (["checked", "disabled", "open", "multiple", "required", "value"].includes(key)) node[key] = value;
      else if (value !== false && value != null) node.setAttribute(key, String(value));
    });
    children.flat(Infinity).forEach((child) => {
      if (child !== null && child !== undefined && child !== false)
        node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    });
    return node;
  }
  function button(text, action, type = "secondary", disabled = false) {
    return el("button", { type: "button", class: `button ${type}`, onclick: action, disabled }, text);
  }
  function notice(text, type = "") {
    const content = Array.isArray(text) ? el("ul", {}, text.map((t) => el("li", {}, asText(t)))) : asText(text);
    return el("div", { class: `notice ${type}` }, content);
  }
  function asText(value) {
    return typeof value === "string" ? value : JSON.stringify(value, null, 2) ?? "";
  }
  function reportValue(value) {
    return value?.trim() ? value : "未填写（输出保留原文）";
  }
  function feedback(text, error = false) {
    $("feedback").className = `feedback${error ? " error" : ""}${state.busy ? " busy" : ""}`;
    $("feedback").setAttribute("role", error ? "alert" : "status");
    $("feedback").textContent = text;
  }
  async function api(url, body, method) {
    const task = body?.task || (body?.projects ? body : null);
    if (task && url !== "/api/tasks/recovery-check") {
      for (const project of task.projects) {
        for (const key of ["start", "end"]) {
          const draft = state.clockDrafts.get(`${project.id}:${key}`);
          if (draft && draft.base === project[key])
            throw new Error(`${project.name}的${key === "start" ? "开始" : "结束"}时间尚未选完整，请补齐日期、时和分，或清空时分后仅保留日期。`);
        }
      }
    }
    let response;
    try {
      response = await fetch(url, {
        method: method || (body === undefined ? "GET" : "POST"),
        credentials: "same-origin",
        headers: body === undefined || body instanceof FormData ? {} : { "Content-Type": "application/json" },
        body: body === undefined ? undefined : body instanceof FormData ? body : JSON.stringify(body)
      });
    } catch {
      throw new Error("无法连接本机服务。请检查服务是否运行；当前未保存内容仍保留在页面中。");
    }
    let data;
    try { data = await response.json(); }
    catch { throw new Error(`服务返回了无法读取的响应（HTTP ${response.status}）。当前编辑内容未清除。`); }
    if (!response.ok) {
      const detail = Array.isArray(data.detail)
        ? data.detail.map((item) => typeof item === "string" ? item : `${(item.loc || []).join(".")} ${item.msg || asText(item)}`).join("；")
        : asText(data.detail || data);
      throw new Error(`请求失败（HTTP ${response.status}）：${detail}`);
    }
    return data;
  }
  async function run(message, action, success) {
    if (state.busy) return;
    if (state.photoUpload || state.photoUploadFailure?.task === state.task) {
      feedback("附件2照片上传尚未完成。可继续浏览和编辑；请到附件2等待、重试或结束本次上传后，再保存、生成或执行其他后台修改。", true);
      return false;
    }
    if (state.materialSelection) {
      feedback("结果素材尚在处理、未选入任务。请等待完成或取消本次选择后，再保存、生成或执行其他后台修改。", true);
      return false;
    }
    state.busy = true;
    $("app").inert = true;
    $("task-dialog").inert = true;
    $("app").setAttribute("aria-busy", "true");
    feedback(message);
    let failure = null;
    try { await action(); }
    catch (error) { failure = error; }
    finally {
      state.busy = false;
      $("app").inert = false;
      $("task-dialog").inert = false;
      $("app").removeAttribute("aria-busy");
      if (failure) {
        feedback(failure.message || String(failure), true);
        if ($("task-dialog").open) {
          const errorNotice = notice(failure.message || String(failure), "error");
          errorNotice.setAttribute("role", "alert");
          $("task-list").prepend(errorNotice);
        }
      }
      else feedback(typeof success === "function" ? success() : success || "操作完成。");
    }
    return !failure;
  }
  const resultMaterialCache = new Map();
  const resultMaterialRequests = new Map();
  const resultMaterialErrors = new Map();
  let resultMaterialVersion = "";
  function ensureResultMaterialVersion() {
    const version = state.catalog.result_material_version || "";
    if (version !== resultMaterialVersion) {
      resultMaterialVersion = version;
      resultMaterialCache.clear(); resultMaterialRequests.clear(); resultMaterialErrors.clear();
      state.materialPreload = null;
    }
    return version;
  }
  function fetchResultMaterialBatch(ids, version) {
    const pending = (async () => {
      const data = await api("/api/assets/projects", {version, asset_ids: ids});
      if (ensureResultMaterialVersion() !== version || data.version !== version)
        throw new Error("结果素材版本已变化，请刷新目录后重试。");
      if (!Array.isArray(data.projects) || !Array.isArray(data.errors))
        throw new Error("后台返回的结果素材批次格式无效。");
      const returned = [...data.projects.map(project => project.asset_id), ...data.errors.map(error => error.asset_id)];
      if (returned.length !== ids.length || new Set(returned).size !== ids.length ||
          returned.some(id => !ids.includes(id)) || data.projects.some(project => !Array.isArray(project.rows)) ||
          data.errors.some(error => typeof error.detail !== "string" || !error.detail))
        throw new Error("后台返回的结果素材与请求不一致，未使用该批次。");
      const projects = new Map(data.projects.map(project => [project.asset_id, project]));
      data.projects.forEach(project => {
        resultMaterialErrors.delete(project.asset_id);
        resultMaterialCache.set(project.asset_id, {project, size: JSON.stringify(project).length});
      });
      data.errors.forEach(error => resultMaterialErrors.set(error.asset_id, error.detail));
      while (resultMaterialCache.size > 128 ||
          [...resultMaterialCache.values()].reduce((sum, entry) => sum + entry.size, 0) > 8 * 1024 * 1024)
        resultMaterialCache.delete(resultMaterialCache.keys().next().value);
      return projects;
    })();
    ids.forEach(id => resultMaterialRequests.set(id, pending));
    const cleanup = () => ids.forEach(id => {
      if (resultMaterialRequests.get(id) === pending) resultMaterialRequests.delete(id);
    });
    pending.then(cleanup, cleanup);
    return pending;
  }
  async function loadResultMaterial(identity) {
    const version = ensureResultMaterialVersion();
    if (!version) return api(`/api/assets/${encodeURIComponent(identity)}/project`);
    if (resultMaterialCache.has(identity)) return clone(resultMaterialCache.get(identity).project);
    const projects = await (resultMaterialRequests.get(identity) ||
      fetchResultMaterialBatch([identity], version));
    if (ensureResultMaterialVersion() !== version)
      throw new Error("结果素材版本已变化，请刷新目录后重试。");
    if (!projects.has(identity))
      throw new Error(resultMaterialErrors.get(identity) || "结果素材未返回，请重试。");
    return clone(projects.get(identity));
  }
  function paintMaterialLoading(target = document.querySelector("[data-material-loading]")) {
    if (!target) return;
    const version = ensureResultMaterialVersion();
    if (!version) { target.replaceChildren(); return; }
    const assets = state.catalog.assets.filter(asset =>
      asset.kind === "result" && asset.report_type === state.task.report_type);
    const cached = assets.filter(asset => resultMaterialCache.has(asset.id)).length;
    const operation = state.materialSelection;
    target.replaceChildren(
      notice(`当前类别公共素材缓存：${cached}/${assets.length}。仅缓存原始定义，不包含本次实测值；已有素材可直接复用。`),
      ...(state.materialPreload?.error ? [notice(`素材预加载未全部完成：${state.materialPreload.error}`, "warning"),
        button("重试预加载", () => { state.materialPreload = null; preloadResultMaterials(); }, "secondary small")] : []),
      ...(operation ? [notice(`正在${operation.phase === "processing" ? "匹配并校验" : "读取"}「${operation.name}」，尚未选入任务。可浏览或编辑其他内容；若任务改变，本次返回结果不会覆盖新编辑。`, "warning"),
        button("取消本次素材选择", () => cancelResultSelection("已取消素材选择，当前任务保持不变。"), "secondary small")] : []));
  }
  async function preloadResultMaterials() {
    const version = ensureResultMaterialVersion();
    if (!version || !state.task) return;
    const family = state.task.report_type;
    const key = `${version}:${family}`;
    if (state.materialPreload?.key === key) return;
    const progress = {key, error: ""};
    state.materialPreload = progress;
    const assets = state.catalog.assets.filter(asset => asset.kind === "result" && asset.report_type === family);
    const preferred = (state.materialSchemes || []).find(scheme => scheme.report_type === family)?.asset_ids ||
      (family === "水表" ? (state.sampleSchemes || [])[0]?.projects.map(project => project.asset_id) : []) || [];
    const ids = [...new Set([...preferred, ...assets.map(asset => asset.id)])]
      .filter(id => assets.some(asset => asset.id === id));
    const limit = Math.min(24, Math.max(1, state.catalog.result_material_batch_limit || 24));
    try {
      for (let offset = 0; offset < ids.length; offset += limit) {
        if (state.materialPreload !== progress || state.task.report_type !== family) return;
        const batch = ids.slice(offset, offset + limit).filter(id =>
          !resultMaterialCache.has(id) && !resultMaterialRequests.has(id));
        if (batch.length) await fetchResultMaterialBatch(batch, version);
        if (state.materialPreload !== progress) return;
        paintMaterialLoading();
      }
      const failures = ids.filter(id => resultMaterialErrors.has(id));
      if (failures.length) progress.error = failures.map(id =>
        `${assets.find(asset => asset.id === id).name}：${resultMaterialErrors.get(id)}`).join("；");
    } catch (error) {
      if (state.materialPreload !== progress) return;
      progress.error = error.message || String(error);
    } finally {
      if (state.materialPreload === progress) paintMaterialLoading();
    }
  }
  function cancelResultSelection(message = "", repaint = true) {
    if (!state.materialSelection) return;
    state.materialSelection = null;
    if (repaint) render();
    if (message) feedback(message);
  }
  function materialTaskSignature() {
    const task = {...state.task};
    // These summaries can finish in parallel; the authoritative task request recomputes them.
    for (const key of ["report_environment", "evaluation_summary", "sealing_model", "sealing_model_source",
      "sealing_issues", "supplementary_notes", "supplementary_note_ids", "other_information_order",
      "supplementary_source", "supplementary_issues", "other_sealing_number"]) delete task[key];
    task.overview = (task.overview || []).filter(field =>
      !["environment_temperature", "environment_humidity", "test_location", "evaluation_period"].includes(field.key) &&
      !(field.key === "sealing" && task.sealing_mode === "auto"));
    return JSON.stringify([task, state.epoch, [...state.clockDrafts],
      state.resultsEditorDirty, state.equipmentEditorDirty,
      state.recoveryResultDraft, state.recoveryEquipmentDraft]);
  }
  async function selectResultMaterial(asset) {
    if (state.photoUpload || state.photoUploadFailure?.task === state.task) {
      feedback("附件2照片上传尚未完成，请等待、重试或结束本次上传后再选择结果素材。", true); return;
    }
    if (state.busy || state.materialSelection) return;
    if (state.resultsEditorDirty || state.equipmentEditorDirty) {
      feedback("请先应用或取消尚未完成的编辑，再选择结果素材。", true); return;
    }
    const originalTask = state.task;
    const operation = {asset_id: asset.id, name: asset.name, phase: "loading"};
    const version = ensureResultMaterialVersion();
    const select = async () => {
      const loaded = normalizeProject(await loadResultMaterial(asset.id));
      if (version && state.materialSelection !== operation) return;
      if (version && (state.task !== originalTask || state.task.report_type !== asset.report_type))
        throw new Error("当前任务或报告对象已变化，本次素材未选入，请在当前任务重新选择。");
      loaded.asset_id = asset.id;
      loaded.evaluators = []; loaded.reviewers = []; loaded.confirmed = false;
      inheritOrdinaryStaff(loaded, state.task.projects[0]);
      if (!loaded.id) loaded.id = `project-${Date.now()}-${++uid}`;
      const signature = materialTaskSignature();
      const candidate = clone(state.task);
      candidate.projects.push(loaded);
      operation.phase = "processing";
      if (version) paintMaterialLoading();
      const result = candidate.record_ids.length ? await api("/api/tasks/parse", candidate)
        : version ? await api("/api/tasks/order", candidate) : await orderedTask(candidate);
      if (version && state.materialSelection !== operation) return;
      if (version && (ensureResultMaterialVersion() !== version ||
          state.task !== originalTask || materialTaskSignature() !== signature))
        throw new Error("等待期间任务或编辑内容已改变，本次素材未选入。新编辑已保留，请重新选择该素材。");
      state.task = normalizeTask(result);
      state.orderVerified = Boolean(state.task.items_asset_id);
      dirty();
    };
    if (!version) {
      return run("正在载入项目并按所选一览表自动排列…", async () => { await select(); render(); },
        "结果素材已选入，请核对样机行、技术要求和人员。");
    }
    state.materialSelection = operation;
    render();
    feedback("结果素材正在后台读取和校验，尚未选入任务；可继续浏览和编辑已加载内容。");
    try {
      await select();
      if (state.materialSelection === operation)
        feedback("结果素材已选入并完成后台校验；请核对并保存当前任务。");
    } catch (error) {
      if (state.materialSelection === operation) feedback(error.message || String(error), true);
    } finally {
      if (state.materialSelection === operation) {
        state.materialSelection = null;
        render();
      }
    }
  }
  function newTask() {
    return {
      id: "", name: "型式评价报告", report_type: "水表", overview: clone(state.catalog.overview).map((field) => ({
        ...field, value: "", confirmed: false, manual_override: false, candidates: []
      })), items_asset_id: "", observe_asset_id: "", outline_id: "", record_ids: [],
      is_new_type: null, measurement_parameters: null, include_volume_display: false,
      volume_display_label: "体积量显示", display_types: null, classification_code: "",
      equipment: [], projects: [], observations: [], items: [], revision: 0,
      include_atmospheric_pressure: false,
      samples: [], sample_schema: {}, sample_table_excluded_columns: [], sample_issues: [], samples_confirmed: false, legacy_sample_fields: [],
      water_sample_source_schema: {}, water_sample_next_numbers: {},
      sample_candidates: {}, sample_mapping_issues: [], photos: {}, photo_edits: {}, photo_order: [], photo_layout: {}, gas_photo_layout_version: 0, photo_components: [], photo_issues: [],
      construction_usage: "", construction_custom: "", construction_instrument: "", construction_issues: [],
      test_basis_issues: [],
      sealing_mode: "auto", sealing_model: "", sealing_model_source: "", sealing_issues: [], other_notes: [],
      supplementary_notes: [], supplementary_source: "", supplementary_issues: [], other_sealing_number: 1,
      supplementary_note_ids: [], custom_supplementary_notes: [], other_information_order: [],
      record_receipt_location: "张衡路1500号", record_test_conclusion: "",
      report_environment: { values: {}, issues: [] },
      sample_delivery_date: "", evaluation_summary: { values: {}, issues: [] },
      observation_summary: { values: {}, sources: {}, issues: [] }
    };
  }
  function normalizeTask(task) {
    const base = newTask();
    const stored = task.overview || [];
    const storedByKey = new Map(stored.map((field) => [field.key, field]));
    const catalogKeys = new Set(base.overview.map((field) => field.key));
    const overview = [
      ...base.overview.map((field) => ({ ...field, ...storedByKey.get(field.key), label: field.label })),
      ...stored.filter((field) => !catalogKeys.has(field.key) &&
        !state.catalog.hidden_overview_fields.includes(field.key))
    ];
    const normalized = { ...base, ...task,
      report_type: task.report_type === "流量计" ? "流量计（液体）" : (task.report_type ?? "水表"),
      record_receipt_location: "张衡路1500号", record_test_conclusion: "",
      sealing_mode: task.sealing_mode || (storedByKey.get("sealing")?.value?.trim() ? "manual" : "auto"),
      other_notes: task.other_notes || [],
      supplementary_notes: task.supplementary_notes || [], supplementary_source: task.supplementary_source || "",
      supplementary_issues: task.supplementary_issues || [], other_sealing_number: task.other_sealing_number || 1,
      supplementary_note_ids: task.supplementary_note_ids || [],
      custom_supplementary_notes: task.custom_supplementary_notes || [],
      other_information_order: task.other_information_order || [],
      sample_table_excluded_columns: task.sample_table_excluded_columns ?? [],
      samples: (task.samples || []).map((sample) => ({ sealed: false, ...sample })),
      overview: overview.map((f) => ({
      old: "", value: "", confirmed: false, manual_override: false, candidates: [], ...f
    })), projects: (task.projects || []).map(normalizeProject) };
    clearInactiveDisplayPhotos(normalized);
    return normalized;
  }
  function normalizeProject(project) {
    return {
      id: "", asset_id: "", name: "", variant: "", rows: [], start: "", end: "",
      water_result_labels: [],
      water_output_labels: null,
      temperature: "", humidity: "", conclusion: "", abnormal: "", evaluators: [],
      atmospheric_pressure: "", pressure_sources: [], pressure_issues: [],
      pressure_manual_override: false, pressure_auto_value: "",
      reviewers: [], equipment_ids: [], confirmed: false, notes: [], source: "",
      record_summary: { values: {}, records: [], issues: [] }, ...project, conclusion: "", abnormal: ""
    };
  }
  async function orderedTask(candidate) {
    if (!candidate.items_asset_id) {
      state.orderVerified = false;
      return normalizeTask(candidate);
    }
    try {
      const ordered = normalizeTask(await api("/api/tasks/order", candidate));
      state.orderVerified = true;
      return ordered;
    } catch (error) {
      state.orderVerified = false;
      render();
      throw new Error(`${error.message} 排序未完成，本次变更未应用；当前任务内容仍保留，请解决匹配问题后重新核对。`);
    }
  }
  function orderNotice() {
    if (!state.task.items_asset_id) return notice("请先选择项目一览表，再选择结果素材。当前项目顺序未经一览表确认，不能视为已排序。", "warning");
    return el("div", {},
      notice(state.orderVerified
        ? "已按当前所选项目一览表的项目顺序自动排列。取消后重新选择仍归入一览表对应位置；同项目多个方案保持相对顺序，不按文件名或勾选顺序决定项目位置。"
        : "项目顺序尚未核对成功。存在未匹配或歧义时不会猜测排序，请核对一览表、项目语义及条款。", state.orderVerified ? "" : "warning"),
      button("按一览表重新核对顺序", () => run("正在按一览表项目语义及条款核对顺序…", async () => {
        state.task = await orderedTask(state.task);
        dirty(); render();
      }, "项目顺序已按当前一览表重新核对。预览与导出采用此顺序。"), "secondary small"));
  }
  async function acceptParsedTask(parsed) {
    state.task = normalizeTask(parsed);
    state.orderVerified = Boolean(state.task.items_asset_id);
    state.clockDrafts.clear(); state.autoEndTimes.clear(); state.autoStartTimes.clear();
    state.task.overview.forEach((item) => { item.confirmed = false; });
    state.task.projects.forEach((project) => { project.confirmed = false; });
    dirty();
    render();
  }
  async function refreshOutlineFields(outlineId) {
    const candidate = { ...state.task, outline_id: outlineId, photos: { ...state.task.photos },
      photo_edits: { ...state.task.photo_edits }, photo_layout: { ...state.task.photo_layout } };
    if (outlineId !== state.task.outline_id && (Object.keys(candidate.photos).some((slot) => slot.startsWith("component:"))
        || Object.keys(candidate.photo_layout).some((key) => /^(photo|heading):component:/.test(key)))) {
      if (!window.confirm("更换大纲需要清除本次零部件照片关联及对应排版设置，避免照片对应到错误名称。样机、铭牌、封样、内部结构和显示照片保持不变；原照片及历史报告不会删除。是否继续？")) {
        render();
        throw new Error("已取消更换大纲，原大纲及照片关联保持不变。");
      }
      Object.keys(candidate.photos).filter((slot) => slot.startsWith("component:")).forEach((slot) => {
        delete candidate.photos[slot]; delete candidate.photo_edits[slot];
      });
      candidate.photo_order = (candidate.photo_order || []).filter((slot) => !slot.startsWith("component:"));
    }
    if (outlineId !== state.task.outline_id) {
      candidate.sample_table_excluded_columns = [];
      Object.keys(candidate.photo_layout).filter((key) => /^(photo|heading):component:/.test(key))
        .forEach((key) => { delete candidate.photo_layout[key]; });
    }
    let parsed;
    try { parsed = await api("/api/tasks/outline-fields", candidate); }
    catch (error) { render(); throw error; }
    state.task = normalizeTask(parsed);
    dirty();
    render();
  }
  async function refreshRecordFields(recordIds) {
    const candidate = { ...state.task, record_ids: recordIds };
    let parsed;
    try { parsed = await api("/api/tasks/record-fields", candidate); }
    catch (error) { render(); throw error; }
    state.task = normalizeTask(parsed);
    dirty();
    render();
  }
  function dirty(project) {
    state.photoPreviewPaint?.();
    state.dirty = true;
    state.epoch += 1;
    state.validation = null;
    if (project) project.confirmed = false;
    refreshReportEnvironment();
    refreshEvaluationSummary();
    refreshOtherInformation();
    updateSummary();
    scheduleDraft();
  }
  function sealingSignature() {
    const task = state.task;
    return JSON.stringify([task.id, task.revision, task.outline_id, task.sealing_mode,
      task.custom_supplementary_notes, task.other_information_order,
      task.overview.find((f) => f.key === "instrument_name")?.value || "",
      task.sealing_mode === "manual" ? task.overview.find((f) => f.key === "sealing")?.value || "" : "",
      task.samples.map((s) => [s.label, s.number, s.sealed, s.disassembly])]);
  }
  function refreshOtherInformation() {
    let signature = sealingSignature();
    if (state.sealingSignature === signature) return;
    state.sealingSignature = signature;
    const previousError = state.sealingError;
    state.sealingPending = true; state.sealingError = "";
    clearTimeout(state.sealingTimer);
    paintOtherInformation();
    const task = clone(state.task);
    state.sealingTimer = setTimeout(async () => {
      try {
        const result = await api("/api/other-information", task);
        if (sealingSignature() !== signature) return;
        if (!Array.isArray(result.supplementary_notes) || !Array.isArray(result.supplementary_issues) ||
            !Array.isArray(result.supplementary_note_ids) || !Array.isArray(result.other_information_order))
          throw new Error("后台尚未启用大纲补充说明，请重启服务后刷新页面。");
        const current = state.task.overview.find((f) => f.key === "sealing");
        const incoming = result.overview.find((f) => f.key === "sealing");
        if (current && incoming) {
          if (current.value !== incoming.value || JSON.stringify(current.candidates) !== JSON.stringify(incoming.candidates) ||
              JSON.stringify(state.task.supplementary_notes) !== JSON.stringify(result.supplementary_notes))
            current.confirmed = false;
          current.value = incoming.value; current.candidates = incoming.candidates;
        }
        state.task.sealing_model = result.sealing_model;
        state.task.sealing_model_source = result.sealing_model_source;
        state.task.sealing_issues = result.sealing_issues;
        state.task.supplementary_notes = result.supplementary_notes;
        state.task.supplementary_note_ids = result.supplementary_note_ids;
        state.task.other_information_order = result.other_information_order;
        state.task.supplementary_source = result.supplementary_source;
        state.task.supplementary_issues = result.supplementary_issues;
        state.task.other_sealing_number = result.other_sealing_number;
        signature = sealingSignature();
        state.sealingSignature = signature;
        if (previousError && $("feedback").textContent === previousError)
          feedback("大纲补充说明和封样信息已重新读取；当前其他编辑已保留。");
      } catch (error) {
        if (sealingSignature() !== signature) return;
        state.sealingError = `封样信息更新失败：${error.message} 请重试，保存及生成时仍由服务端重新校验。`;
        feedback(state.sealingError, true);
      }
      if (sealingSignature() !== signature) return;
      state.sealingPending = false;
      paintOtherInformation(); updateSummary();
      if (state.step === 6) render();
    }, 200);
  }
  function otherInformationText() {
    return otherInformationClauses().map((clause, index) => numberedOtherText(
      clause.text.trim() ? (clause.kind === "custom" ? otherClauseBody(clause.text).trim() || "待补录" : clause.text) : clause.kind === "sealing"
        ? "封样信息未填写（输出保留模板封样文字）" : "待补录", index + 1)).join("\n");
  }
  function numberedOtherText(text, number) {
    const match = text.match(/^\s*(?:([0-9０-９]+)[.．、)）](?![0-9０-９])|[（(]([0-9０-９]+)[）)])\s*/);
    const digits = match && /^[０-９]+$/.test(match[1] || match[2])
      ? String(number).replace(/[0-9]/g, digit => String.fromCharCode(digit.charCodeAt(0) + 0xfee0)) : String(number);
    return match ? match[0].replace(/[0-9０-９]+/, digits) + text.slice(match[0].length) : `${number}.` + text;
  }
  function otherClauseBody(text) {
    return text.replace(/^\s*(?:[0-9０-９]+[.．、)）](?![0-9０-９])|[（(][0-9０-９]+[）)])\s*/, "");
  }
  function otherInformationClauses() {
    const task = state.task;
    const entries = task.supplementary_notes.map((text, index) => ({
      id: task.supplementary_note_ids?.[index] || `outline:${task.outline_id || "current"}:${index+1}`,
      text, kind: "outline"
    }));
    entries.push(...(task.custom_supplementary_notes || []).map(note => ({...note, kind:"custom"})));
    entries.push({id:"sealing", kind:"sealing", text:task.overview.find(f=>f.key==="sealing")?.value || ""});
    const order = (task.other_information_order || []).filter(id => entries.some(entry=>entry.id===id));
    entries.forEach(entry => {if (!order.includes(entry.id)) order.push(entry.id);});
    return order.map(id=>entries.find(entry=>entry.id===id));
  }
  function otherClauseEditor() {
    const clauses = otherInformationClauses();
    return clauses.map((clause, index) => {
      const move = delta => {
        const order = clauses.map(entry=>entry.id);
        [order[index],order[index+delta]] = [order[index+delta],order[index]];
        state.task.other_information_order = order;
        dirty(); render();
      };
      const content = clause.kind === "custom"
        ? field(`第${index+1}条 · 新增补充说明`, clause.text, value => {
          state.task.custom_supplementary_notes.find(note=>note.id===clause.id).text = value;
          dirty(); paintOtherInformation();
        }, {multiline:true, attrs:{"data-custom-note":clause.id}, help:"只填写正文，条款号自动生成。"})
        : el("p", {class:"preformatted"}, numberedOtherText(clause.text || "封样信息未填写",index+1));
      return el("div", {class:"other-clause", "data-other-clause":clause.id},
        el("strong", {}, `${index+1}. ${clause.kind==="outline" ? "大纲原文" : clause.kind==="sealing" ? "封样信息（下方编辑）" : "新增条款"}`),
        el("div", {class:"toolbar"},
          button("上移", ()=>move(-1), "secondary small", index===0),
          button("下移", ()=>move(1), "secondary small", index===clauses.length-1),
          clause.kind==="custom" ? button("删除此条", ()=>{
            if (!window.confirm("删除本次新增的补充说明？不修改大纲原文或历史报告。")) return;
            state.task.custom_supplementary_notes = state.task.custom_supplementary_notes.filter(note=>note.id!==clause.id);
            state.task.other_information_order = clauses.filter(entry=>entry.id!==clause.id).map(entry=>entry.id);
            dirty(); render();
          }, "secondary small") : null), content);
    });
  }
  function paintOtherInformation() {
    const item = state.task.overview.find((f) => f.key === "sealing");
    if (!item) return;
    document.querySelectorAll("[data-sealing-value]").forEach((input) => {
      if (input.value !== item.value) input.value = item.value;
    });
    document.querySelectorAll("[data-other-preview]").forEach((box) => { box.textContent = otherInformationText(); });
    document.querySelectorAll("[data-other-clauses]").forEach(box => {
      if (document.activeElement?.closest("[data-other-clauses]") !== box) box.replaceChildren(...otherClauseEditor());
    });
    document.querySelectorAll("[data-supplementary-notes]").forEach((box) => {
      box.replaceChildren(
        ...(state.task.supplementary_notes.length ? state.task.supplementary_notes.map((note) =>
          el("p", { class: "preformatted" }, note)) : [el("p", { class: "muted" }, "未提取到带编号的补充条款，封样信息作为第一条。")]),
        el("small", {}, state.task.supplementary_source));
    });
    document.querySelectorAll("[data-sealing-messages]").forEach((box) => {
      box.replaceChildren(
        ...(state.sealingPending ? [notice("正在读取大纲补充条款并更新封样信息…")] : []),
        ...(state.sealingError ? [notice(state.sealingError, "error")] : []),
        ...(state.task.sealing_issues.length ? [notice(state.task.sealing_issues, "warning")] : []),
        ...(state.task.supplementary_issues.length ? [notice(state.task.supplementary_issues, "warning")] : []),
        el("p", { class: "muted" }, state.task.sealing_model
          ? `型号：${state.task.sealing_model}；来源：${state.task.sealing_model_source}`
          : "尚未识别到本次大纲封面型号。"),
        ...(state.task.sealing_mode === "manual" ? [notice("封样信息保留人工/旧任务文字，仅调整条款号，不随样机选择自动覆盖。可切回自动生成。", "warning")] : []));
    });
  }
  function environmentSignature() {
    return JSON.stringify([state.task.id, state.task.projects.map((p) => [p.id, p.name, p.temperature, p.humidity])]);
  }
  function evaluationSignature() {
    return JSON.stringify([state.task.id, state.task.sample_delivery_date,
      state.task.equipment.map((e) => [e.id, e.storage_location]),
      state.task.projects.map((p) => [p.id, p.name, p.end])]);
  }
  function setEvaluationSummary(summary) {
    state.task.evaluation_summary = summary;
    state.task.overview.forEach((item) => {
      if (automaticOverviewFields.has(item.key)) {
        item.value = summary.values[item.key] || "";
        item.confirmed = false;
        item.candidates = [];
      }
    });
    document.querySelectorAll("[data-evaluation-value]").forEach((box) => {
      box.textContent = summary.values[box.dataset.evaluationValue] || "待补齐来源后自动填写";
    });
    document.querySelectorAll("[data-evaluation-issues]").forEach((box) => {
      box.replaceChildren(...summary.issues.map((issue) => notice(issue, "warning")));
    });
  }
  function refreshEvaluationSummary() {
    const signature = evaluationSignature();
    if (state.evaluationSignature === signature) return;
    state.evaluationSignature = signature;
    clearTimeout(state.evaluationTimer);
    setEvaluationSummary({ values: {}, issues: ["地点与型式评价时间正在自动汇总…"] });
    const task = clone(state.task);
    state.evaluationTimer = setTimeout(async () => {
      try {
        const summary = await api("/api/evaluation-summary", task);
        if (evaluationSignature() !== signature) return;
        setEvaluationSummary(summary);
      } catch (error) {
        if (evaluationSignature() !== signature) return;
        setEvaluationSummary({ values: {}, issues: [`自动汇总失败：${error.message}；请重试。`] });
        feedback(`地点与型式评价时间汇总失败：${error.message}`, true);
      }
      if (state.step === 6) render();
    }, 200);
  }
  function setReportEnvironment(summary) {
    state.task.report_environment = summary;
    state.task.overview.forEach((field) => {
      if (field.key in environmentFields) {
        field.value = summary.values[field.key] || "";
        field.confirmed = false;
        field.candidates = [];
      }
    });
    document.querySelectorAll("[data-report-environment]").forEach(paintReportEnvironment);
  }
  function refreshReportEnvironment() {
    const signature = environmentSignature();
    if (state.environmentSignature === signature) return;
    state.environmentSignature = signature;
    clearTimeout(state.environmentTimer);
    setReportEnvironment({ values: {}, issues: ["总环境温湿度正在汇总，请稍候。"] });
    const projects = clone(state.task.projects);
    state.environmentTimer = setTimeout(async () => {
      try {
        const summary = await api("/api/environment-summary", { projects });
        if (environmentSignature() === signature) {
          setReportEnvironment(summary);
          if (state.step === 6) render();
        }
      } catch (error) {
        if (environmentSignature() !== signature) return;
        setReportEnvironment({ values: {}, issues: [`总环境温湿度汇总失败：${error.message}；请重试。`] });
        feedback(`总环境温湿度汇总失败：${error.message}`, true);
        if (state.step === 6) render();
      }
    }, 200);
  }
  function paintReportEnvironment(box) {
    const summary = state.task.report_environment;
    box.replaceChildren(
      el("dl", { class: "data-list" }, Object.entries(environmentFields).flatMap(([key, label]) => [
        el("dt", {}, label), el("dd", {}, summary.values[key] || "尚未形成完整范围")
      ])),
      el("p", { class: "muted" }, `汇总当前 ${state.task.projects.length} 个已选项目的已填写值；取全部最小值～最大值，不取平均值。修改或增删项目后自动更新，报告同步使用此范围，无须在总览重复填写。`),
      el("p", { class: "muted" }, "项目中未写单位的温度按℃、湿度按%RH理解。任一项目缺项或格式无效时，该项总范围暂不写入；草稿沿用模板原文并保留缺项提醒，请勿视作本次范围。"),
      ...(summary.issues.length ? [notice(summary.issues, "warning")] : []),
      button("重新汇总", () => { state.environmentSignature = null; refreshReportEnvironment(); }, "secondary small")
    );
  }
  function renderReportEnvironment() {
    const box = el("div", { "data-report-environment": "" });
    paintReportEnvironment(box);
    return card("本次报告总环境温湿度（自动汇总）", box);
  }
  function updateSummary() {
    const task = state.task;
    if (!task) return;
    const fields = applicableOverviewFields(task).filter((f) => !(f.key in environmentFields) && !automaticOverviewFields.has(f.key));
    $("save-state").textContent = `${task.id ? `已保存版本 ${task.revision}` : "新任务"} · ${state.dirty ? "有未保存修改" : task.id ? "当前内容已保存" : "尚未保存"}`;
    const metrics = [
      [task.record_ids.length, "份原始记录"],
      [`${fields.filter((f) => f.value.trim()).length}/${fields.length}`, "项总览已填写"],
      [task.projects.length, "个已选结果项目"],
      [task.equipment.length, "台本次设备"]
    ];
    $("summary").replaceChildren(...metrics.map(([count, label]) => el("div", { class: "summary-item" }, el("strong", {}, count), el("span", {}, label))));
  }
  function render() {
    if (!state.task) return;
    refreshReportEnvironment();
    refreshEvaluationSummary();
    refreshOtherInformation();
    $("navigation").replaceChildren(...steps.map(([name], index) => el("button", {
      type: "button", class: "nav-step", "aria-current": state.step === index ? "step" : null,
      onclick: () => navigate(index)
    }, el("span", { class: "nav-number", "aria-hidden": "true" }, String(index + 1).padStart(2, "0")), el("span", { class: "nav-name" }, name))));
    $("step-title").textContent = steps[state.step][0];
    $("step-description").textContent = steps[state.step][1];
    $("step-indicator").textContent = `${String(state.step + 1).padStart(2, "0")} / ${String(steps.length).padStart(2, "0")}`;
    $("previous-step").disabled = state.step === 0;
    $("next-step").disabled = state.step === steps.length - 1;
    $("summary").hidden = state.step === steps.length - 1;
    const warnings = catalogWarnings();
    $("catalog-warnings").replaceChildren(...(state.step !== steps.length - 1 && warnings.length
      ? [notice(warnings, "warning")] : []));
    if (state.loadErrors.length) $("catalog-warnings").append(notice(state.loadErrors, "error"), button("重试载入资料（保留当前编辑）", initialize, "secondary"));
    const panels = [renderImports, renderOverview, renderMaterials, renderEquipment, renderProjects, renderPhotos, renderPreview];
    $("panel").replaceChildren(panels[state.step]());
    paintOtherInformation();
    updateSummary();
    if (state.ready) preloadResultMaterials();
  }
  function navigate(index) {
    if (state.busy || !state.task) return;
    if (index === 6 && state.task.projects.some((project) => project.name === "示值误差曲线")) {
      run("正在从本次原始记录刷新曲线自动结果…", async () => {
        await refreshCurvePreview();
        state.step = index;
        render();
        $("main").focus();
        window.scrollTo({ top: 0, behavior: "instant" });
      }, "曲线自动结果已按当前记录刷新；人工结果与主动清空保持不变。");
      return;
    }
    state.step = index;
    render();
    $("main").focus();
    window.scrollTo({ top: 0, behavior: "instant" });
  }
  async function refreshCurvePreview() {
    if (!state.task.projects.some((project) => project.name === "示值误差曲线")) return;
    const previous = JSON.stringify(state.task);
    const refreshed = normalizeTask(await api("/api/tasks/record-fields", state.task));
    state.task = refreshed;
    if (JSON.stringify(refreshed) !== previous) {
      state.epoch += 1;
      state.validation = null;
      state.dirty = true;
    }
  }
  function field(label, value, change, options = {}) {
    const id = `field-${++uid}`;
    const input = el(options.multiline ? "textarea" : options.choices ? "select" : "input", {
      id, value: value ?? "", type: options.type || "text", required: !!options.required,
      placeholder: options.placeholder || "", ...(options.hideLabel ? { "aria-label": label } : {}),
      ...(options.attrs || {})
    });
    if (options.choices) {
      input.append(...options.choices.map((choice) => {
        const [v, text] = Array.isArray(choice) ? choice : [choice, choice];
        return el("option", { value: v }, text);
      }));
      input.value = value || "";
    }
    input.addEventListener(options.choices ? "change" : "input", () => change(input.value));
    return el("div", { class: "field" },
      options.hideLabel ? null : el("label", { for: id, class: "field-label" },
        label, options.required ? el("span", { class: "required" }, "必填") : null),
      input, options.help ? el("small", {}, options.help) : null);
  }
  function checkbox(label, checked, change, help) {
    const input = el("input", { type: "checkbox", checked });
    input.addEventListener("change", () => change(input.checked));
    return el("label", { class: "check" }, input, el("span", {}, label, help ? el("small", {}, help) : null));
  }
  function card(title, ...children) { return el("section", { class: "card" }, el("h2", {}, title), ...children); }
  function empty(title, text) { return el("div", { class: "empty" }, el("strong", {}, title), text); }
  function uploadName(id) { return state.uploads.find((u) => u.id === id)?.name || `资料 ${id}（详情未载入）`; }
  function selectedUpload(id) { return state.uploads.find((u) => u.id === id); }

  function renderReportType() {
    const locked = !!(state.task.items_asset_id || state.task.observe_asset_id || state.task.projects.length);
    return el("fieldset", { "data-report-type": "" },
      el("legend", {}, "本次型评报告对象类型"),
      el("div", { class: "toolbar" }, ["水表", "流量计（气体）", "流量计（液体）", "燃气表"].map((type) =>
        el("label", { class: "check" },
          el("input", { type: "radio", name: "report-type", value: type,
            disabled: locked,
            checked: (state.task.report_type ?? "水表") === type,
            onchange: (event) => {
              if (!event.target.checked || state.task.report_type === type) return;
              if (locked) {
                feedback("任务已选择素材或结果项目，不能切换报告对象类型；请新建任务。", true);
                return;
              }
              if (!["燃气表", "流量计（气体）"].includes(type) && Object.keys(state.task.photos).some((slot) => slot.startsWith("sealed_overall:"))) {
                feedback("仅燃气表和流量计（气体）支持多张封样照片；请先移除追加的封样照片或新建任务，再切换报告对象。", true);
                render(); return;
              }
              const mismatchedRecords = state.task.record_ids.filter(id => selectedUpload(id)?.report_type !== type);
              const mismatchedOutline = state.task.outline_id && selectedUpload(state.task.outline_id)?.report_type !== type;
              if (mismatchedOutline && (Object.keys(state.task.photos).some(slot => slot.startsWith("component:"))
                  || Object.keys(state.task.photo_layout || {}).some(key => /^(photo|heading):component:/.test(key)))) {
                feedback("当前大纲有关联零部件照片或排版设置，请先取消大纲关联并确认照片处理，再切换报告对象。", true);
                render(); return;
              }
              if ((mismatchedRecords.length || mismatchedOutline) &&
                  !window.confirm(`切换为${type}将解除${mismatchedRecords.length}份不匹配原始记录${mismatchedOutline ? "及当前大纲" : ""}的本次关联。上传原件、其他任务和已填写内容保留，请重新核对。是否继续？`)) {
                render(); return;
              }
              state.task.record_ids = state.task.record_ids.filter(id => !mismatchedRecords.includes(id));
              if (mismatchedOutline) {
                state.task.outline_id = "";
                state.task.sample_table_excluded_columns = [];
              }
              state.task.report_type = type;
              clearInactiveDisplayPhotos(state.task);
              state.orderVerified = false;
              state.materialFolders.clear();
              state.materialQuery = "";
              dirty();
              render();
            } }), el("span", {}, type)))),
      notice("每份任务选择一种类型，默认水表。流量计分为气体、液体两类，资料和素材独立；原流量计任务及素材归入流量计（液体）。燃气表和流量计（气体）均支持各自格式的Excel原始记录及DOCX；Excel结果保留两位小数、负零归零，燃气表误差曲线按本次数据重绘。水表及流量计（液体）仅支持DOCX记录；四类不混用资料或素材。"),
      locked ? notice("已选择素材或结果项目，报告对象类型已锁定。需要更换类型请新建任务。", "warning")
        : el("p", { class: "muted" }, "未选择素材时可切换类型，保留已上传资料、基本信息、样机、设备及照片，不自动修改器具名称或主模板。"));
  }
  function renderImports() {
    const task = state.task;
    const root = el("div");
    root.append(card("01  本次报告",
      field("报告任务名称", task.name, (value) => { task.name = value; dirty(); }, { required: true }),
      renderReportType(),
      notice("原始记录可上传多份；大纲只使用本次明确选择的一份。重新解析保留总览已填值和人工修改的实测结果，更新自动识别值并汇总项目时间、环境温湿度；自动匹配不等于人工确认。"),
      button("解析本次资料 → 总览与项目", () => run("正在解析本次资料，保留人工修改并匹配已选项目…", async () => {
        const parsed = await api("/api/tasks/parse?refresh_records=true", task);
        await acceptParsedTask(parsed);
      }, "解析完成。请逐项核对总览候选、项目记录、来源与缺项；当前内容待保存。"), "primary",
      !task.record_ids.length && !task.outline_id)
    ));
    root.append(el("div", { class: "grid" }, uploadCard("record", "02  原始记录", true), uploadCard("outline", "03  当前大纲", false)));
    const selected = selectedUpload(task.outline_id);
    const current = selected?.report_type === (task.report_type ?? "水表") ? selected : null;
    root.append(card("大纲原文提取范围",
      notice("第4项完整表格 → 报告「（七）关键零部件和材料」；第5项覆盖型号引导语或“技术参数如下：”之后、至第6项“本次型式评价项目”之前的全部内容 → 下方「建议批准的型号范围」及报告「七、型式评价结论及建议」。以下是提取核对文本，原生表格与格式保留在生成的 Word 中，合并、跨页与版式仍需核对。", "warning"),
      current ? el("div", {},
        el("h3", {}, current.name),
        el("h4", {}, "关键零部件和材料"),
        el("div", { class: "preformatted" }, asText(current.outline?.components) || "未提取到范围，请核对大纲。"),
        el("h4", {}, "建议批准的型号范围"),
        el("div", { class: "preformatted" }, asText(current.outline?.coverage) || "未提取到范围，请核对大纲。"))
        : empty("尚未选择当前大纲", "上传或从本机已上传资料中选择，不会复用其他任务的大纲。")
    ));
    return root;
  }
  function uploadCard(kind, title, multiple) {
    const supportsExcel = kind === "record" && ["燃气表", "流量计（气体）"].includes(state.task.report_type);
    const formats = supportsExcel ? "DOCX / XLSX" : "DOCX";
    const input = el("input", { type: "file", accept: supportsExcel ? ".docx,.xlsx" : ".docx", multiple, "aria-label": `${title}，仅 ${formats}` });
    const box = el("div", { class: "upload-box" }, el("strong", {}, multiple ? `添加原始记录 ${formats}` : "上传并设为当前大纲"), el("p", { class: "muted" }, multiple ? "支持一次选择多份；请确认属于同一报告范围。Excel“打印页”按报告项目及编号顺序追加，共用页只追加一次；未选项目或该编号无本次实测结果的页不追加。生成结果不再单独输出Excel，上传原件保留。" : "大纲补充件即可，无需再上传基础大纲。"), input,
      button("上传至本机", () => {
        const files = Array.from(input.files);
        if (!files.length) { feedback(`请先选择 ${formats} 文件。`, true); return; }
        if (files.some((file) => !file.name.toLowerCase().endsWith(".docx") &&
          !(supportsExcel && file.name.toLowerCase().endsWith(".xlsx")))) { feedback(`当前仅支持 ${formats}；Excel识别仅适用于燃气表和流量计（气体），两者格式独立。`, true); return; }
        run("正在上传与提取资料，请稍候…", async () => {
          for (const file of files) {
            feedback(`正在处理 ${file.name}…`);
            const form = new FormData();
            form.append("file", file);
            form.append("kind", kind);
            form.append("report_type", state.task.report_type ?? "水表");
            const uploaded = await api("/api/uploads", form);
            state.uploads = [...state.uploads.filter((u) => u.id !== uploaded.id), uploaded];
            if (kind === "record") {
              await refreshRecordFields([...new Set([...state.task.record_ids, uploaded.id])]);
            } else await refreshOutlineFields(uploaded.id);
            dirty();
            render();
          }
        }, kind === "outline"
          ? "大纲已关联，已读取计量器具名称、制造单位、申请书编号及生产依据的可识别内容，请到总览核对候选和缺项提示；已有人工值保留。"
          : "原始记录已关联，报告编号已读取到总览，请核对；其他项目数据请点击「解析本次资料」。");
      }, "secondary")
    );
    const reportType = state.task.report_type ?? "水表";
    const known = state.uploads.filter((u) => u.kind === kind && u.report_type === reportType);
    const unclassified = state.uploads.filter(u => u.kind === kind && !u.report_type);
    const linkedIds = kind === "record" ? state.task.record_ids : state.task.outline_id ? [state.task.outline_id] : [];
    const list = el("div");
    list.append(el("p", { class: "muted" }, `仅显示${reportType}资料（${known.length}份），新上传资料归入当前报告对象。`));
    if (unclassified.length) list.append(notice(`有${unclassified.length}份旧资料无法可靠分类，未混入当前列表。请在正确报告对象下重新上传；原件未删除。`, "warning"));
    const mismatched = linkedIds.filter(id => selectedUpload(id) && !known.some(u => u.id === id));
    if (mismatched.length) list.append(notice("当前任务仍关联了其他类型或未分类资料，请解除关联后重新选择。已填写内容不会自动清除，请核对。", "warning"),
      button("解除不匹配资料关联", () => run("正在解除不匹配资料关联…", async () => {
        if (!window.confirm("仅解除不匹配资料的本次关联，上传原件和其他任务保留。是否继续？")) return;
        if (kind === "record") await refreshRecordFields(linkedIds.filter(id => !mismatched.includes(id)));
        else await refreshOutlineFields("");
      }), "secondary"));
    if (!known.length && !linkedIds.length) list.append(el("p", { class: "muted" }, "暂无已上传资料。"));
    [...known, ...linkedIds.filter((id) => !selectedUpload(id)).map((id) => ({ id, name: uploadName(id) }))].forEach((upload) => {
      const checked = linkedIds.includes(upload.id);
      list.append(el("div", { class: "file-row" },
        checkbox(upload.name, checked, (enabled) => {
          if (kind === "outline") {
            run("正在读取当前大纲基本信息…", async () => {
              await refreshOutlineFields(enabled ? upload.id : "");
            }, "当前大纲候选已更新，请核对总览；已有人工值保留。");
            return;
          }
          run("正在读取原始记录报告编号…", async () => {
            await refreshRecordFields(enabled ? [...state.task.record_ids, upload.id] : state.task.record_ids.filter((id) => id !== upload.id));
          }, "报告编号候选已更新，请核对总览；已有人工值保留。");
        }, kind === "record" ? `${upload.records?.length ?? "未知"} 条提取记录` : checked ? "本次使用的大纲" : "选择后替换当前大纲"),
        upload.kind ? downloadLink(`/api/uploads/${encodeURIComponent(upload.id)}/file`, "下载原件") : null,
        button("删除", () => {
          if (checked) {
            feedback("这份资料仍关联当前页面的任务，请先取消勾选；若任务已保存，请解除关联后保存任务，再删除。", true);
            return;
          }
          if (!window.confirm(`永久删除上传资料「${upload.name}」？\n只删除系统中的上传副本及提取记录，不会删除电脑中的原文件，也不会删除已生成的报告。\n若仍被已保存任务使用，系统会阻止删除并列出关联任务。`)) return;
          run("正在检查关联任务并删除上传副本…", async () => {
            await api(`/api/uploads/${encodeURIComponent(upload.id)}`, undefined, "DELETE");
            state.uploads = state.uploads.filter((item) => item.id !== upload.id);
            state.epoch += 1; state.validation = null;
            render();
          }, "上传副本及提取记录已永久删除；原文件和已生成报告未改动。");
        }, "secondary small")));
      if (checked && upload.warnings?.length) list.append(kind === "record"
        ? el("details", {}, el("summary", {}, `查看识别说明与提醒（${upload.warnings.length} 条）`),
          notice(upload.warnings, "warning"))
        : notice(upload.warnings, "warning"));
      if (checked && kind === "record" && upload.records?.length) list.append(recordDetails(upload));
    });
    return card(title, box, el("h3", {}, "本机已上传资料"), el("p", { class: "muted" }, "取消勾选只解除本次关联。删除会永久移除系统上传副本及提取记录，不动原文件；已保存任务仍在使用时须先解除关联并保存。更换资料后请重新核对已有人工值。"), list);
  }
  function recordDetails(upload) {
    const details = el("details", {}, el("summary", {}, `查看提取记录与来源（${upload.records.length} 条）`));
    let populated = false;
    details.addEventListener("toggle", () => {
      if (!details.open || populated) return;
      populated = true;
      details.append(notice(upload.parser?.startsWith("gas-flowmeter-excel-")
        ? "气体流量计Excel独立识别：初始平均示值误差、重复性及试验后复测按打印页流量点、真实样机编号匹配，结果保留两位小数、负零归零，原始精度和单元格来源保留。耐压最大工作压力按对应样机的唯一记录更新，源模板不改写；缺少的贮存、振动或电磁兼容试验记录不代填符合要求。"
        : ["流量计（气体）", "流量计（液体）"].includes(upload.report_type)
        ? "流量计独立识别：按样机编号、项目及实际流量读取平均示值误差和重复性。流量点不按行序；先确定目标流量，所有流量点均按与目标值的相对差异≤5%匹配。qt须来自大纲或原始记录的明确参数，不默认取0.2qmax。缺项、多个候选或重复批次逐行提示，不代填合格结论。"
        : upload.parser?.startsWith("gas-excel-")
        ? "燃气表Excel独立识别：按本台样机、项目、条件及流量点填写，报告数值保留两位小数、负零归零。原始精度及单元格来源保留；空白、除零错误及未开展项目的模板残留不作实测。误差曲线按本次初始平均误差重绘，缺少日期或定性检查时需补录。"
        : "以下是原始提取明细。先核对一次样机对应关系，再按项目、样机及子试验读取实测结果。静磁场/安装条件取三次误差中绝对值最大的一次并保留符号；安装前后三次对应上游/下游，水温前后两次对应低温/最高温度。缺失、重复批次或无法对应的结果逐行提示，不代填合格结论。", "warning"));
      upload.records.forEach((record, index) => {
        const metadata = [
          ["型号 / 样机编号", `${record.model || "未提取型号"} / ${record.sample_id || "未提取编号"}`],
          ["试验开始 / 结束", `${record.start || "未提取"} — ${record.end || "未提取"}`],
          ["环境温度 / 湿度", `${record.temperature || "未提取"} / ${record.humidity || "未提取"}`],
          ["来源", record.source || "来源未注明"]
        ];
        const table = el("table", {},
          el("thead", {}, el("tr", {}, ["原始测量项 / 流量点 / 次数", "原始值", "来源 / 位置"].map((text) => el("th", { scope: "col" }, text)))),
          el("tbody", {}, (record.rows || []).map((row) => el("tr", {},
            el("td", {}, row.label || "未注明测量项"), el("td", {}, row.raw_value != null
              ? `${row.raw_value}${row.unit || ""} → ${row.value}` : asText(row.value) || "未提取"),
            el("td", {}, row.source || record.source || "来源未注明")))));
        details.append(el("section", {},
          el("h4", {}, `${index + 1}. ${record.project || "未匹配项目"}`),
          el("dl", { class: "data-list" }, metadata.flatMap(([label, value]) => [el("dt", {}, label), el("dd", {}, value)])),
          record.rows?.length ? el("div", { class: "table-wrap" }, table) : notice("没有提取到测量行，需人工核对原始文件。", "warning")));
      });
    });
    return details;
  }
  function renderApplicationType(editable = true) {
    const defaults = state.catalog.basic_options;
    if (!defaults) return notice("未取得报告模板选项，请保存任务并重启服务、刷新素材目录。", "error");
    const selected = state.task.is_new_type ?? defaults.is_new_type;
    return card("申请类型", editable
      ? checkbox("新型（不勾选即为改进型）", selected, (value) => {
        state.task.is_new_type = value; dirty();
      }, "仅切换原报告的新型/改进型方框，保留标签、空格及原标志格式。")
      : el("p", {}, selected ? "新型 ☒　改进型 □" : "新型 □　改进型 ☒"));
  }
  function clearInactiveDisplayPhotos(task) {
    if (task.report_type === "流量计（气体）") {
      task.gas_photo_layout_version = 2;
    } else if (task.report_type === "燃气表") {
      if (!task.gas_photo_layout_version) {
        task.photo_layout = {};
      }
      if ((task.gas_photo_layout_version || 0) < 2) {
        task.photo_edits = Object.fromEntries(Object.entries(task.photo_edits || {})
          .map(([slot, edit]) => [slot, { ...edit, width_cm: null }]));
        task.gas_photo_layout_version = 2;
      }
    }
    const chosen = task.display_types ?? state.catalog.basic_options?.display_types ?? [];
    const inactive = chosen.includes("机械") && chosen.includes("电子")
      ? ["display"] : ["display_mechanical", "display_electronic"];
    task.photos = { ...task.photos };
    task.photo_edits = { ...task.photo_edits };
    Object.keys(task.photos).filter((slot) => inactive.includes(slot.split(":")[0])).forEach((slot) => {
      delete task.photos[slot];
      delete task.photo_edits[slot];
    });
    task.photo_order = (task.photo_order || []).filter((slot) => !inactive.includes(slot.split(":")[0]));
    task.photo_layout = { ...task.photo_layout };
    Object.keys(task.photo_layout).filter((key) => key.startsWith("photo:") && inactive.includes(key.split(":")[1]))
      .forEach((key) => { delete task.photo_layout[key]; });
  }
  function renderSampleTablePreview() {
    const task = state.task, schema = task.sample_schema || {};
    const editable = state.catalog.sample_table_column_selection === true;
    const root = card("主报告（三）样机表预览",
      notice("按当前大纲每个型号规格数据行预览，源参数只读。表头右侧勾选仅控制主报告（三）的输出列，不改变大纲、样机对应或下方测量参数。序号、准确度及编号固定输出。"));
    root.setAttribute("data-sample-report-preview", "");
    if (!editable) root.append(notice("后台尚未启用样机表列选择，请重启本机服务并刷新素材目录后再编辑输出列。当前仅提供只读预览，不能确认旧后台会按排除列设置输出。", "warning"));
    if (!schema.headers?.length || !schema.groups?.length) {
      root.append(notice("请先选择含「提供试验样机」原生表格的当前大纲。", "warning"));
      return root;
    }
    const fixed = (index) => index === schema.quantity_column ||
      String(schema.headers[index]).replace(/\s/g, "").includes("准确度");
    const table = el("table", { class: "sample-report-table" });
    const redraw = () => {
      const excluded = new Set(task.sample_table_excluded_columns || []);
      const cellAttrs = (index) => ({ "data-sample-source-column": String(index),
        class: !fixed(index) && excluded.has(index) ? "sample-column-excluded" : "" });
      table.replaceChildren(
        el("thead", {}, el("tr", {}, el("th", {}, "序号"),
          schema.headers.map((header, index) => {
            if (fixed(index)) return el("th", cellAttrs(index),
              index === schema.quantity_column ? "编号" : header);
            const included = !excluded.has(index);
            return el("th", cellAttrs(index),
              editable ? el("label", { class: "sample-column-choice" }, el("span", {}, header),
                el("input", { type: "checkbox", checked: included,
                  "data-sample-output-column": String(index), "aria-label": `输出${header}列`,
                  onchange: (event) => {
                    const next = new Set(task.sample_table_excluded_columns || []);
                    if (event.target.checked) next.delete(index); else next.add(index);
                    task.sample_table_excluded_columns = [...next].filter((column) => !fixed(column)).sort((a, b) => a - b);
                    dirty();
                    const left = wrap.scrollLeft;
                    redraw();
                    wrap.scrollLeft = left;
                    table.querySelector(`[data-sample-output-column="${index}"]`)?.focus({ preventScroll: true });
                  } })) : header,
              !included ? el("small", { class: "sample-column-status" }, "不输出") : null);
          }))),
        el("tbody", {}, schema.groups.map((group, rowIndex) =>
          el("tr", {}, el("td", {}, rowIndex + 1),
            schema.headers.map((_, index) => el("td", cellAttrs(index),
              index === schema.quantity_column
                ? group.labels.map((label) => {
                  const number = task.samples.find((sample) => sample.label === label)?.number;
                  return el("div", { "data-sample-preview-label": label },
                    number?.trim() ? number : `${label}（编号待填写）`);
                })
                : el("span", { class: "preformatted" }, group.values[index] ?? "")))))));
    };
    const wrap = el("div", { class: "table-wrap" }, table);
    root.append(wrap);
    redraw();
    return root;
  }
  function renderMeasurementOptions(editable = true) {
    const defaults = state.catalog.basic_options;
    if (!defaults) return notice("未取得报告模板测量参数表及显示选项，不能编辑；请刷新素材目录。", "error");
    const root = el("div", { "data-basic-options": "" });
    const emptyRow = () => ({ ...Object.fromEntries(measurementColumns.map(([key]) => [key, ""])), volume_display: "" });
    const rows = () => state.task.measurement_parameters ??
      defaults.measurement_parameters.map(emptyRow);
    const ensureRows = () => {
      if (state.task.measurement_parameters === null || state.task.measurement_parameters === undefined)
        state.task.measurement_parameters = rows();
      return state.task.measurement_parameters;
    };
    const redraw = () => {
      const parameters = card("计量器具的测量参数",
        notice("按原报告表头逐行填写，可增加或删除参数行。原有行未填写的格保留模板原文；新增行空格保持空白。输入m3时，Word自动将3设为上标。体积量显示可选标况、工况或手填，列名可修改；勾选才输出，Word在原表总宽内自动分配列宽。"));
      const included = state.task.include_volume_display ?? false;
      const volumeLabel = state.task.volume_display_label ?? "体积量显示";
      const columns = [...measurementColumns, ...((editable || included) ? [["volume_display", volumeLabel]] : [])];
      const volumeHeader = editable ? el("div", {},
        el("div", { class: "sample-column-choice" },
          field("体积量显示列名", volumeLabel, (value) => {
            state.task.volume_display_label = value; dirty();
          }, { hideLabel: true, attrs: { "data-volume-display-label": "" } }),
          el("input", { type: "checkbox", checked: included, "data-volume-display-output": "",
            "aria-label": "输出体积量显示列", onchange: (event) => {
              state.task.include_volume_display = event.target.checked; dirty(); redraw();
              root.querySelector("[data-volume-display-output]")?.focus({ preventScroll: true });
            } })),
        el("small", { class: "sample-column-status" }, included ? "输出" : "不输出")) : volumeLabel;
      parameters.append(el("div", { class: "table-wrap" }, el("table", { class: "measurement-table" },
        el("thead", {}, el("tr", {}, ["序号", ...measurementColumns.map(([, label]) => label),
          ...((editable || included) ? [volumeHeader] : []), ...(editable ? ["操作"] : [])]
          .map((label) => el("th", {}, label)))),
        el("tbody", {}, rows().map((row, index) => el("tr", {}, el("td", {}, index + 1),
          columns.map(([key, label]) => el("td", {}, editable && key === "volume_display"
            ? el("div", {}, field(`第${index + 1}行${label}`, row[key] ?? "", (value) => {
                ensureRows()[index][key] = value; dirty();
              }, { hideLabel: true, placeholder: "选择或手填", attrs: {
                list: `volume-display-options-${index}`, "data-volume-display-row": String(index) } }),
              el("datalist", { id: `volume-display-options-${index}` },
                ["标况", "工况"].map((value) => el("option", { value }, value))))
            : editable
            ? field(`第${index + 1}行${label}`, row[key], (value) => {
              ensureRows()[index][key] = value; dirty();
            }, { multiline: true, hideLabel: true, attrs: { rows: 1 },
              placeholder: defaults.measurement_parameters[index]?.[key] || "请填写" })
            : (row[key]?.trim() ? row[key] : defaults.measurement_parameters[index]?.[key]
              ? `${defaults.measurement_parameters[index][key]}（沿用模板）` : "未填写"))),
          editable ? el("td", {}, button("删除此行", () => {
            if (!window.confirm(`删除第${index + 1}条测量参数？保存任务后生效。`)) return;
            ensureRows().splice(index, 1); dirty(); redraw();
          }, "danger small", rows().length <= 1)) : null))))));
      if (editable) parameters.append(button("增加测量参数", () => {
        ensureRows().push(emptyRow()); dirty(); redraw();
      }, "secondary small"));
      const chosen = state.task.display_types ?? defaults.display_types;
      const display = card("显示型式", editable
        ? el("div", { class: "toolbar" }, displayOptions.map((label) =>
          checkbox(label, chosen.includes(label), (checked) => {
            const next = new Set(state.task.display_types ?? defaults.display_types);
            if (checked) next.add(label); else next.delete(label);
            state.task.display_types = displayOptions.filter((value) => next.has(value));
            clearInactiveDisplayPhotos(state.task);
            dirty(); redraw();
          })))
        : el("p", {}, displayOptions.map((label) => `${label}${chosen.includes(label) ? "☒" : "□"}`).join("　")),
      notice("可多选；同时勾选机械和电子时，附件2依次上传“机械部分照片”和“电子部分照片”，每张下方输出居中题注。其他组合使用单张显示照片。单张与双张模式互相切换时，自动清除原模式照片及其旋转、裁剪、宽度设置，需重新上传；不删除原图或历史报告。"));
      root.replaceChildren(parameters, display);
    };
    redraw();
    return root;
  }
  function automaticClassification(task) {
    const name = (task.overview.find((item) => item.key === "instrument_name")?.value || "").replace(/\s+/g, "");
    const matches = (state.catalog.classification_options || []).filter((item) =>
      [item.name, ...item.aliases].some((alias) => item.exact ? name === alias : name.includes(alias)));
    return matches.length === 1 ? matches[0] : null;
  }

  function renderClassification() {
    const options = state.catalog.classification_options || [];
    const automatic = automaticClassification(state.task);
    const selected = options.find((item) => item.code === state.task.classification_code);
    const current = selected || automatic;
    return card("分类编码",
      field("分类编码选项", state.task.classification_code, (value) => {
        state.task.classification_code = value; dirty(); render();
      }, { choices: [
        ["", automatic ? `自动匹配：${automatic.name} — ${automatic.code}` : "自动匹配：未找到唯一结果，请选择"],
        ...options.map((item) => [item.code, `${item.name} — ${item.code}`])
      ], attrs: { disabled: !options.length } }),
      !options.length ? notice("当前后台尚未支持分类编码选项，请保存任务并在服务更新后刷新页面。", "warning")
        : !current ? notice("分类编码待选择；草稿显示“待补录”，不沿用模板旧编码。", "warning")
        : el("p", {}, `封面和正文统一输出：${current.code}（${selected ? "人工选择，保持至您改选或恢复自动匹配" : "按当前器具名称自动匹配"}）`));
  }

  function renderOverview() {
    const root = el("div", { class: "overview-layout" }, notice("按报告章节顺序核对填写，同一信息只维护一次。填写新值则替换，未填写（包括仅有空格）则保留原文；沿用内容不代表已核实。来源和说明默认折叠，人工填写及原有输出规则不变。"));
    const definitions = [
      ["cover", "封面信息", "封面 · 器具名称、分类编码、报告编号", ["instrument_name", "classification_code", "report_number"]],
      ["application", "申请与委托信息", "正文一 · 申请与委托的基本情况", ["manufacturer", "agent", "commissioned_date", "application_number"]],
      ["type", "关于型式的基本信息", "正文二 · 原理、构造、生产依据、样机、测量参数及显示型式", ["principle", "construction", "production_basis", "supply_voltage"]],
      ["basis", "型式评价的依据", "正文三 · 本次型式评价使用的技术依据，与生产依据分开", ["test_basis"]],
      ["other", "其他说明", "正文八 · 补充说明及封样信息，保留本次条款排序", ["sealing"]],
      ["period", "型式评价时间", "正文九 · 送样日期及自动汇总的型式评价时间", ["evaluation_period"]],
      ["observation", "观察与证书信息", "附件1观察项目 · 仅显示当前对象适用项，证书编号与第五部分一览表共用", ["meter_manufacture_month", "battery_replacement_deadline"]],
      ["record", "配套原始记录说明", "配套文件 · 自动取值、默认项及留空规则", []]
    ];
    const groups = {};
    const groupKeys = new Map();
    const fieldOrder = [];
    const navigation = el("nav", { class: "overview-navigation", "aria-label": "总览分组导航" });
    root.append(navigation);
    definitions.forEach(([id, title, subtitle, keys], index) => {
      const content = el("div", { class: "overview-group-content" });
      const section = el("section", { class: "overview-group", "data-overview-group": id, tabindex: "-1" },
        el("div", { class: "overview-group-heading" },
          el("span", { class: "overview-group-number", "aria-hidden": "true" }, String(index + 1).padStart(2, "0")),
          el("div", {}, el("h2", {}, title), el("p", { class: "muted" }, subtitle))), content);
      groups[id] = content;
      keys.forEach(key => { groupKeys.set(key, id); fieldOrder.push(key); });
      navigation.append(button(title, () => {
        section.scrollIntoView({ block: "start" });
        section.focus({ preventScroll: true });
      }, "secondary small"));
      root.append(section);
    });
    const classification = el("div", { "data-classification-picker": "" }, renderClassification());
    if (!state.task.overview.length) root.append(empty("未获取总览字段", "请检查素材目录与服务状态；未取得字段时不能完成正式输出。"));
    const fields = applicableOverviewFields(state.task);
    const positions = new Map(fieldOrder.map((key, index) => [key, index]));
    const ordered = [...fields].sort((left, right) =>
      (positions.get(left.key) ?? fieldOrder.length) - (positions.get(right.key) ?? fieldOrder.length));
    let samplesPlaced = false;
    let additional = null;
    ordered.forEach((item) => {
      if (item.key in environmentFields || item.key === "test_location") return;
      let host = groups[groupKeys.get(item.key)];
      if (!host) {
        if (!additional) {
          additional = card("其他基本字段");
          root.append(additional);
        }
        host = additional;
      }
      if (automaticOverviewFields.has(item.key)) {
        host.append(card(item.label,
          item.key === "evaluation_period" ? field("送样日期", state.task.sample_delivery_date, (value) => {
            state.task.sample_delivery_date = value; dirty();
          }, { type: "date", required: true, help: "型式评价开始日期采用此日期；结束日期取全部已选项目中最晚的截止日期，不需重复确认。" }) : null,
          el("div", { "data-evaluation-value": item.key },
            state.task.evaluation_summary.values[item.key] || "待补齐来源后自动填写"),
          el("small", {}, item.key === "test_location"
            ? "按本次所选设备快照的存放地点自动去重汇总，与原始记录检测地点表使用同一清单，无需再次确认。"
            : "自动汇总，不使用模板历史时间，也不受项目显示顺序影响。")));
        return;
      }
      if (item.key === "sealing") {
        host.append(renderOtherInformation(item));
        return;
      }
      const warnings = el("div");
      const detailLabel = item.key === "principle" ? "来源、原理选项与说明"
        : item.key === "construction" ? "来源、用途设置与说明" : "来源与说明";
      const summary = el("summary");
      const update = () => {
        if (item.key === "instrument_name") classification.replaceChildren(renderClassification());
        const messages = [];
        if (item.key === "production_basis") {
          messages.push(...(state.task.production_basis_issues || []));
          if (item.value.trim() && item.candidates.length &&
              !item.candidates.some((candidate) => candidate.value === item.value))
            messages.push("当前生产依据与本次大纲3.2产品标准不同，已保留您的文字；请核对并按需采用大纲候选。");
        }
        if (item.key === "test_basis") {
          messages.push(...(state.task.test_basis_issues || []));
          if (!item.value.trim()) messages.push("试验依据待补录；空白输出保留模板原文，不代表本次识别成功。");
          if (item.value.trim() && item.candidates.length &&
              !item.candidates.some((candidate) => candidate.value === item.value))
            messages.push("当前试验依据与本次大纲第2项技术依据不同，已保留您的文字，请核对或采用大纲候选。");
          if (!state.catalog.test_basis_from_outline) messages.push("当前后台尚未支持试验依据同步，请保存任务并在服务更新后刷新页面。");
        }
        if (item.key === "report_number" && state.task.record_ids.length && !item.candidates.length)
          messages.push("所选原始记录中未识别到报告编号，请核对记录右上方并手动填写。");
        if (item.key === "report_number" && item.value.trim() && item.candidates.length &&
            !item.candidates.some((candidate) => candidate.value === item.value))
          messages.push("当前填写的报告编号与所选原始记录不同，已保留您的文字，请核对后决定是否采用候选。");
        if (new Set(item.candidates.map((c) => c.value)).size > 1)
          messages.push("不同来源存在候选差异，请人工选择或修正，不会静默覆盖。");
        warnings.replaceChildren(...messages.map((message) => notice(message, "warning")));
        summary.replaceChildren(`${detailLabel}（${item.candidates.length}条候选）`,
          ...(messages.length ? [el("span", { class: "pill warning" }, `${messages.length}项待核对提醒`)] : []));
      };
      const valueField = field(`${item.label}：拟替换新值`, item.value, (value) => {
        item.value = value; item.manual_override = true; item.confirmed = false; dirty(); update();
      }, { multiline: true, required: true, attrs: {
        rows: /[\r\n]/.test(item.value) || item.value.length > 100 ? 3 : 1,
        disabled: item.key === "test_basis" && !state.catalog.test_basis_from_outline
      },
        placeholder: "未填写则保留原文", help: item.key === "meter_manufacture_month"
          ? "手工填写水表制造年月，例如2026.09；同步替换型评报告和配套原始记录中observe的标志备注，不从大纲推测。未填写时保留所选素材原文。"
          : item.key === "battery_replacement_deadline"
            ? "手工填写更换电池的最后期限，例如2033.09；同步替换两份输出中observe的标志和7.7.4备注，不根据制造年月自动计算。未填写时保留所选素材原文。"
          : ["instrument_name", "manufacturer", "application_number"].includes(item.key)
        ? "读取当前上传大纲的对应字段；制造单位对应大纲的制造厂。已有人工值保留，无需额外勾选确认。"
        : item.key === "report_number"
          ? "读取当前原始记录的报告编号栏或样品信息首行右侧证书编号栏；不取样品、设备或申请书编号。候选不一致时请核对，人工值保留。"
          : item.key === "production_basis"
            ? "读取本次大纲3.2产品标准括号内原文。每条标准另起一行；报告中每条独立成段，首行缩进两个字符。已有人工文字保留供核对。"
          : item.key === "test_basis"
            ? "读取本次大纲第2项技术依据。每条标准一行：第三部分逐条成段，第六部分用“、”连接全部依据。人工修改保存保留；可采用大纲候选恢复原文。"
          : "多型号、多样机请保留对应关系；原理与日期不推测填写。" });
      if (item.key === "construction" && state.task.construction_usage)
        valueField.querySelector("textarea").readOnly = true;
      const help = valueField.querySelector("small");
      let principlePicker = null;
      if (item.key === "principle") {
        const preview = el("div", { class: "preformatted muted" });
        let selected = "";
        const apply = button("采用所选原理", () => {
          const choice = state.principles.find((entry) => entry.id === selected);
          if (!choice) { feedback("请先选择一种原理。", true); return; }
          if (item.value.trim() && item.value !== choice.text &&
              !window.confirm("采用所选原理会替换当前已填写的原理文字，是否继续？")) return;
          item.value = choice.text; item.manual_override = true; item.confirmed = false; dirty(); render();
          feedback("所选原理已带入本次任务，可继续修改并保存任务。后台原理选项不会随本次修改改变。");
        }, "secondary small", true);
        principlePicker = el("div", { class: "principle-picker" },
          field("选择已保存的原理", "", (identity) => {
            selected = identity;
            const choice = state.principles.find((entry) => entry.id === identity);
            preview.textContent = choice?.text || "";
            apply.disabled = !choice;
          }, { choices: [["", "请选择原理（不自动带入）"], ...state.principles.map((choice) => [choice.id, choice.name])] }),
          preview, apply,
          el("p", { class: "muted" }, "每次主动选择后点击采用；仍可手动填写或修改，保存的是本次文字，不联动修改后台选项。"));
      }
      const candidates = el("div");
      item.candidates.forEach((candidate) => candidates.append(el("div", { class: "candidate" },
        el("span", { class: item.key === "production_basis" ? "production-basis-text" : "" },
          candidate.value, el("small", {}, `来源：${candidate.source || "来源未注明"}`)),
        button("采用此候选", () => {
          item.value = candidate.value; item.manual_override = item.key !== "test_basis"; item.confirmed = false; dirty(); render();
        }, "secondary small", item.key === "test_basis" && !state.catalog.test_basis_from_outline))));
      if (item.key === "construction") {
        candidates.className = "construction-candidates";
        candidates.hidden = !!state.task.construction_usage;
      }
      const details = el("details", { class: "overview-details" }, summary,
        el("div", { class: "overview-detail-content" },
          el("p", { class: "overview-original" }, el("strong", {}, "模板原文："), item.old || "空白"),
          help, principlePicker,
          item.key === "construction" ? constructionPicker(item, valueField, update, candidates) : null,
          warnings, candidates));
      update();
      host.append(el("section", { class: "card overview-field", "data-overview-key": item.key },
        el("div", { class: "overview-main" },
          el("div", { class: "overview-heading" }, el("h2", {}, item.label)),
          el("div", { class: "overview-old", title: item.old || "空白" }, item.old || "空白"),
          el("span", { class: "overview-arrow" }, "替换为"),
          valueField),
        details));
      if (item.key === "instrument_name") host.append(classification);
      if (item.key === "application_number") host.append(renderApplicationType());
      if (item.key === "production_basis") {
        host.append(el("h3", { class: "overview-subheading" }, "样机信息"),
          renderSamples(true), renderSampleTablePreview(), renderMeasurementOptions());
        samplesPlaced = true;
      }
    });
    if (!samplesPlaced) groups.type.append(renderSamples(true));
    groups.period.append(el("div", { "data-evaluation-issues": "" },
      state.task.evaluation_summary.issues.map((issue) => notice(issue, "warning"))),
      button("重新汇总型评时间", () => {
        state.evaluationSignature = null; refreshEvaluationSummary();
      }, "secondary small"));
    groups.observation.append(renderObservationCertificates());
    if ((state.task.report_type ?? "水表") === "水表") groups.observation.append(renderObservationSummary());
    groups.record.append(el("details", { class: "card overview-record-details" },
      el("summary", {}, "配套原始记录封面 · 查看自动取值与留空规则"),
      notice("生成型评报告时同时输出另一份配套原始记录。封面读取本次大纲编号和封面型号、总览器具名称/制造厂、已选项目的最早/最晚试验日期及汇总温湿度。样机编号填写“见后页”；封面评价人员、复核人员均留空，不影响各试验项目的人员选择。受样日期取最早试验日期，月日补齐两位。填写内容在固定横线上居中，过长时自动缩小字号保持单行；试验日期和环境温湿度横线与上方对齐。业务流转单编号没有本次来源时保留空白。"),
      el("p", {}, "受样地点：张衡路1500号（默认填写，无需确认）"),
      el("p", {}, "原始记录封面试验结论：留空，不自动填写“通过”。"),
      el("p", { class: "muted" }, state.task.report_type === "燃气表"
        ? "误差曲线和耐久性项目保留本次曲线图，位于对应数据表下方、合格判定要求上方，不放在第一页封面。实测结果表、日期、人员等其他填写项仍留空，供后续填写和核对。"
        : "项目页仅保留技术要求、样机编号和固定标签；日期、实测结果、合格判定要求、结论、异常、设备引用、环境条件及人员全部留空，供后续填写。")));
    return root;
  }
  function applicableOverviewFields(task) {
    return (task.report_type ?? "水表") !== "水表"
      ? task.overview.filter((item) => !["meter_manufacture_month", "battery_replacement_deadline"].includes(item.key))
      : task.overview;
  }
  function catalogWarnings() {
    const messages = state.catalog.observation_copy_warnings;
    return state.catalog.warnings.map((message) => message === messages?.["水表"]
      ? messages[state.task.report_type ?? "水表"] : message);
  }
  function renderObservationSummary() {
    const summary = state.task.observation_summary || { values: {}, sources: {}, issues: [] };
    return el("details", { class: "card overview-details" },
      el("summary", {}, "观察备注 · 本次大纲参数与来源",
        summary.issues.length ? el("span", { class: "pill warning" }, `${summary.issues.length}项待核对提醒`) : null),
      el("p", { class: "muted" }, "观察首部使用总览制造单位、器具名称及样机清单的型号规格与真实编号。指定备注中Q3、Q3/Q1取第5项实际送检样机表，不混入覆盖范围型号。制造年月、电池期限和安装方式可在上方补录；未填保留对应原文并提示。其他要求、备注和已有勾选保留，两份Word输出同步生效。"),
      summary.issues.length ? notice(summary.issues, "warning") : null,
      el("dl", { class: "data-list" }, Object.entries(summary.values).map(([key, value]) => [
        el("dt", {}, state.catalog.observation_labels?.[key] || key),
        el("dd", {}, value, el("small", {}, ` · ${summary.sources[key] || "来源待核对"}`))
      ])));
  }
  function renderOtherInformation(item) {
    const task = state.task;
    const mode = field("封样信息填写方式", task.sealing_mode, (value) => {
      if (value === "auto" && item.value.trim() &&
          !window.confirm("切换为自动生成将用当前大纲型号、总览器具名称及封样选择替换封样文字；其他补充条款和排序不变。是否继续？")) {
        mode.querySelector("select").value = task.sealing_mode;
        return;
      }
      task.sealing_mode = value; item.confirmed = false;
      dirty(); render();
    }, { choices: [["auto", "根据大纲和封样样机自动生成"], ["manual", "保留或手工修改封样信息"]] });
    const first = field("封样信息正文（按上方条款顺序编号）", item.value, (value) => {
      if (task.sealing_mode === "auto") {
        task.sealing_mode = "manual";
        mode.querySelector("select").value = "manual";
      }
      item.value = value; item.confirmed = false; dirty(); paintOtherInformation();
    }, { multiline: true, attrs: { "data-sealing-value": "" },
      help: "自动生成后可直接修改；首次输入即转为人工模式并保留本次文字，重新生成Word时采用修改值。切回自动模式才会按当前大纲型号、器具名称及封样样机重新生成。" });
    const legacy = task.other_notes.filter((text) => text.trim());
    return el("section", { class: "card" },
      el("div", { class: "card-heading" }, el("h2", {}, "其他")),
      el("h3", {}, "大纲第7节 · 补充说明"),
      el("p", { class: "muted" }, "大纲只复制带条款号的原文，正文只读。可新增补充说明，所有条款（包括封样信息）均可上移、下移，预览与Word按当前顺序连续编号。"),
      el("small", {}, task.supplementary_source),
      el("div", { "data-other-clauses": "" }, ...otherClauseEditor()),
      button("新增补充说明", () => {
        const order = otherInformationClauses().map(clause=>clause.id);
        const id = `custom:${crypto.randomUUID()}`;
        task.custom_supplementary_notes.push({id,text:""});
        order.splice(order.indexOf("sealing"),0,id);
        task.other_information_order = order;
        dirty(); render();
      }, "secondary small"),
      legacy.length ? el("details", {}, el("summary", {}, "旧任务手工说明（仅保留备查，不参与本次输出）"),
        ...legacy.map((text) => el("p", { class: "preformatted" }, text))) : null,
      el("details", {}, el("summary", {}, "查看模板原有封样文字"), el("p", { class: "preformatted" }, item.old || "空白")),
      mode, el("div", { "data-sealing-messages": "" }), first,
      button("重新读取大纲补充说明及封样信息", () => { state.sealingSignature = null; refreshOtherInformation(); }, "secondary small"),
      el("h3", {}, "本次其他信息预览"), el("div", { class: "preformatted", "data-other-preview": "" }));
  }
  function constructionPicker(item, valueField, update, candidates) {
    const task = state.task;
    const messages = el("div");
    const custom = field("其他计量用途", task.construction_custom, (value) => {
      task.construction_custom = value;
      refreshSentence();
    }, { help: "只填写用途词，例如“蒸汽”；系统自动组合为“适用于蒸汽计量”。" });
    custom.hidden = task.construction_usage !== "其他";
    const usage = field("选择计量用途", task.construction_usage, (value) => {
      if (!task.construction_usage && value && item.value.trim() &&
          !window.confirm("启用自动生成会替换当前“构造和应用场合”文字，名称和零部件将随当前大纲更新。是否继续？")) {
        usage.querySelector("select").value = task.construction_usage;
        return;
      }
      task.construction_usage = value;
      item.manual_override = !value;
      custom.hidden = value !== "其他";
      candidates.hidden = !!value;
      valueField.querySelector("textarea").readOnly = !!value;
      refreshSentence();
    }, { choices: [["", "手动填写 / 尚未选择"], "用水", "液体", "气体", "燃气", ["其他", "其他（自行填入）"]] });
    function refreshSentence() {
      task.construction_issues = [];
      if (task.construction_usage) {
        const names = task.photo_components.map((part) => part.name);
        const target = task.construction_usage === "其他" ? task.construction_custom.trim() : task.construction_usage;
        if (!task.construction_instrument) task.construction_issues.push("当前大纲的计量器具名称缺失或不唯一，请上传并选择有效大纲。");
        if (!names.length) task.construction_issues.push("当前大纲的具体零部件名称尚未识别。");
        if (!target) task.construction_issues.push("请填写“其他”的具体计量用途。");
        const parts = names.length > 1 ? `${names.slice(0, -1).join("、")}和${names.at(-1)}` : names[0];
        item.value = task.construction_issues.length ? "" : `${task.construction_instrument}主要由${parts}组成，适用于${target}计量。`;
      }
      item.confirmed = false;
      valueField.querySelector("textarea").value = item.value;
      messages.replaceChildren(...(task.construction_issues.length ? [notice(task.construction_issues, "warning")] : []));
      dirty(); update();
    }
    if (task.construction_issues.length) messages.append(notice(task.construction_issues, "warning"));
    return el("div", { class: "construction-picker" }, usage, custom,
      el("p", { class: "muted" }, "选择用途后，自动使用当前大纲的计量器具名称及附件2的具体零部件名称；前面用“、”分隔，最后两项用“和”连接。自动模式下换纲会更新句子，请重新核对。"),
      messages);
  }
  function updateSampleNumber(sample, value, { task = state.task, manual = true, source = "" } = {}) {
    const previous = sample.number;
    sample.number = value.trim();
    if (manual) {
      sample.record_number ||= sample.auto_number || "";
      sample.number_origin = "manual"; sample.auto_number = ""; sample.number_source = source;
    }
    task.samples_confirmed = false;
    const other = task.overview.find((f) => f.key === "sealing");
    if (other) other.confirmed = false;
    task.projects.forEach((project) => {
      project.rows.filter((row) => row.sample_label === sample.label).forEach((row) => {
        if (previous !== sample.number) {
          if (!row.manual_override && !row.source?.includes("实测结果人工修改")
              && (!row.value?.trim() || (row.auto_value && row.value === row.auto_value))) {
            row.value = ""; row.source = ""; row.auto_value = ""; row.matched_sample_id = "";
            row.match_issue = "样机对应已改变，请重新解析本次资料";
          } else {
            row.matched_sample_id ||= previous;
            row.auto_value = "";
            row.match_issue = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存";
          }
        }
        row.sample_id = sample.number;
        project.confirmed = false;
      });
    });
    if (task === state.task) {
      const confirmation = document.querySelector("input[data-samples-confirm]");
      if (confirmation) confirmation.checked = false;
      document.querySelectorAll("[data-sample-preview-label]").forEach((node) => {
        if (node.dataset.samplePreviewLabel === sample.label)
          node.textContent = sample.number || `${sample.label}（编号待填写）`;
      });
      dirty();
    }
  }
  function reorderSamples(sourceLabel, targetLabel) {
    const task = state.task, groups = task.sample_schema.groups;
    const group = groups.find((item) => item.labels.includes(sourceLabel));
    if (!group || !group.labels.includes(targetLabel)) {
      feedback("只能在同一大纲型号规格组内调整编号，不能跨组拖动；原对应关系未改变。", true);
      return false;
    }
    const from = group.labels.indexOf(sourceLabel), to = group.labels.indexOf(targetLabel);
    if (from === to) return false;
    const slots = group.labels.map((label) => task.samples.find((sample) => sample.label === label));
    if (slots.some((sample) => !sample)) {
      feedback("样机清单与大纲分组不一致，请重新选择大纲后核对。", true);
      return false;
    }
    const values = slots.map((sample) => ({
      number: sample.number, disassembly: sample.disassembly, sealed: sample.sealed,
      number_source: sample.number_source || "", record_number: sample.record_number || sample.auto_number || "",
      added_manually: !!sample.added_manually
    }));
    const [moved] = values.splice(from, 1);
    values.splice(to, 0, moved);
    slots.forEach((sample, index) => {
      if (index < Math.min(from, to) || index > Math.max(from, to)) return;
      sample.disassembly = values[index].disassembly; sample.sealed = values[index].sealed;
      sample.record_number = values[index].record_number; sample.added_manually = values[index].added_manually;
      updateSampleNumber(sample, values[index].number, { source: values[index].number_source });
    });
    return true;
  }
  function focusSampleHandle(label) {
    [...document.querySelectorAll("[data-drag-sample]")].find((node) => node.dataset.dragSample === label)
      ?.focus({ preventScroll: true });
  }
  async function reparseSampleMapping(label) {
    const task = state.task, epoch = state.epoch, candidate = clone(task);
    const top = window.scrollY, left = document.querySelector("[data-sample-table]")?.scrollLeft || 0;
    const ok = await run("样机顺序已调整，正在按新编号重新解析…", async () => {
      const parsed = await api("/api/tasks/parse", candidate);
      if (state.task !== task || state.epoch !== epoch)
        throw new Error("解析期间已有新修改，解析结果未应用。请点击“解析本次资料”后再核对结果。");
      state.task = normalizeTask(parsed);
      state.task.samples_confirmed = false;
      state.task.projects.forEach((project) => { project.confirmed = false; });
      dirty(); render();
    }, "编号及拆样、封样选择已同步移动，报告代号位置不变。结果已按新编号重新解析，请核对并保存；旧人工结果的警告仍需处理。");
    if (!ok) feedback(`顺序修改已保留，旧自动结果已清除。请点击“解析本次资料”后再核对新编号结果。${$("feedback").textContent}`, true);
    const table = document.querySelector("[data-sample-table]");
    if (table) table.scrollLeft = left;
    window.scrollTo({ top });
    focusSampleHandle(label);
  }
  function renderSamples(editable) {
    const task = state.task, schema = task.sample_schema;
    const dynamic = task.report_type === "水表" && !!task.water_sample_source_schema?.groups;
    const root = card("样机清单 · 大纲第5项",
      notice(`${dynamic ? "水表按本次原始记录识别几台就显示几台，不按大纲数量预留空行；可在现有型号规格组内添加样机并手写编号。" : "选择大纲和原始记录后，按规格及记录出现顺序自动填入真实编号；未识别项保留空白。"}型号、规格及其他参数直接复制当前大纲。人工修改及主动清空不会被自动覆盖。只能在同一规格组内拖动编号（或使用上移、下移 / Alt+方向键），拆样和封样选择随编号一起移动，A001等报告代号的位置固定。手工改号或添加后请“解析本次资料”再核对项目结果。`));
    if (!schema.groups?.length) {
      root.append(notice("请先选择含「提供试验样机」原生表格的当前大纲。", "warning"));
      return root;
    }
    root.append(el("p", { class: "preformatted" }, schema.note || "大纲未提供拆样说明，请人工核对。"));
    if (task.sample_issues.length) root.append(notice(task.sample_issues, "error"));
    if (task.sample_mapping_issues.length) root.append(notice(task.sample_mapping_issues, "warning"));
    if (editable) root.append(button("匹配原始记录样机编号", () => run("正在读取本次样机候选…", async () => {
      const parsed = await api("/api/tasks/parse", state.task);
      openSampleMatching(normalizeTask(parsed));
    }, "请在匹配窗口核对后应用；尚未更改任务。"), "secondary"),
    notice("编号已在下表直接显示；匹配窗口仅用于人工纠正。请核对型号及证书范围，未出现在记录中的样机不编造编号。"));
    let draggedLabel = "";
    const move = (from, to) => {
      if (state.busy || !reorderSamples(from, to)) return;
      const position = root.querySelector(".table-wrap")?.scrollLeft || 0;
      const top = window.scrollY;
      render();
      const table = document.querySelector("[data-sample-table]");
      if (table) table.scrollLeft = position;
      window.scrollTo({ top });
      focusSampleHandle(to);
      reparseSampleMapping(to);
    };
    const header = ["序号", ...schema.headers.map((h, i) => i === schema.quantity_column ? "编号" : h),
      "报告代号", ...(editable ? ["拆样样机", "封样样机", ...(dynamic ? ["样机管理"] : [])] : [])];
    const rows = schema.groups.map((group, groupIndex) => {
      const samples = group.labels.map((label) => task.samples.find((s) => s.label === label));
      return el("tr", {}, el("td", {}, groupIndex + 1), group.values.map((value, index) => {
        if (index !== schema.quantity_column) return el("td", {}, value);
        return el("td", { class: "sample-number-cell" }, !samples.length ? "未识别样机，可手工添加" : null, samples.map((sample, index) => {
          if (!sample) return el("div", { class: "sample-slot" }, notice("样机位置缺失，请重新核对大纲。", "error"));
          if (!editable) return el("div", { class: "sample-slot" }, sample.number || "未识别，请填写");
          const origin = el("small", { class: "sample-number-origin", title: sample.number_source || "" },
            sample.number_origin === "auto" ? "自动识别 · 请核对" : sample.number_origin === "manual" ? "人工指定" : "尚未识别");
          const handle = el("button", {
            type: "button", class: "button secondary small sample-drag-handle", draggable: true,
            "data-drag-sample": sample.label, "aria-label": `调整 ${sample.label} 编号顺序`,
            title: "拖动编号；也可按 Alt+上 / 下方向键",
            ondragstart: (event) => {
              draggedLabel = sample.label;
              event.dataTransfer.effectAllowed = "move";
              event.dataTransfer.setData("text/plain", sample.label);
            },
            ondragend: () => {
              draggedLabel = "";
              root.querySelectorAll(".sample-drop-target").forEach((node) => node.classList.remove("sample-drop-target"));
            },
            onkeydown: (event) => {
              if (!event.altKey || !["ArrowUp", "ArrowDown"].includes(event.key)) return;
              event.preventDefault();
              const target = samples[index + (event.key === "ArrowUp" ? -1 : 1)];
              if (target) move(sample.label, target.label);
            }
          }, "↕");
          return el("div", { class: "sample-slot sample-edit-slot",
            ondragover: (event) => {
              if (!draggedLabel) return;
              event.preventDefault(); event.currentTarget.classList.add("sample-drop-target");
            },
            ondragleave: (event) => event.currentTarget.classList.remove("sample-drop-target"),
            ondrop: (event) => {
              event.preventDefault(); event.currentTarget.classList.remove("sample-drop-target");
              if (draggedLabel) move(draggedLabel, sample.label);
              draggedLabel = "";
            }
          }, handle, el("div", { class: "sample-number-input" }, el("input", {
            value: sample.number, placeholder: "未识别，请填写", "aria-label": `${sample.label} 样机编号`,
            oninput: (event) => {
              updateSampleNumber(sample, event.target.value);
              origin.textContent = "人工指定"; origin.title = "";
            }
          }), origin), el("div", { class: "sample-order-buttons" },
            el("button", { type: "button", class: "button secondary small", disabled: index === 0,
              "aria-label": `${sample.label} 编号上移`, onclick: () => move(sample.label, samples[index - 1]?.label) }, "↑"),
            el("button", { type: "button", class: "button secondary small", disabled: index === samples.length - 1,
              "aria-label": `${sample.label} 编号下移`, onclick: () => move(sample.label, samples[index + 1]?.label) }, "↓")));
        }));
      }), el("td", {}, group.labels.map((label) => el("div", { class: "sample-slot" }, label))),
      editable ? el("td", {}, samples.filter(Boolean).map((sample) => el("div", { class: "sample-slot" }, el("input", {
        type: "radio", name: "disassembly", checked: sample.disassembly,
        "aria-label": `${sample.label} 用于拆样`, onchange: () => {
          task.samples.forEach((s) => { s.disassembly = s === sample; });
          const other = task.overview.find((f) => f.key === "sealing");
          if (other) other.confirmed = false;
          task.samples_confirmed = false; dirty(); render();
        } })))) : null,
      editable ? el("td", {}, samples.filter(Boolean).map((sample) => el("div", { class: "sample-slot" }, el("input", {
        type: "checkbox", checked: sample.sealed, "aria-label": `${sample.label} 用于封样`,
        onchange: (event) => {
          sample.sealed = event.target.checked;
          const other = task.overview.find((f) => f.key === "sealing");
          if (other) other.confirmed = false;
          task.samples_confirmed = false; dirty(); render();
        }
      })))) : null,
      editable && dynamic ? el("td", {}, samples.map((sample) =>
        el("div", { class: "sample-slot" }, button(`移除 ${sample.label}`, () => {
          if (state.busy) return;
          const protectedRows = task.projects.flatMap((project) => project.rows)
            .filter((row) => row.sample_label === sample.label &&
              (row.manual_override || row.source?.includes("实测结果人工修改") ||
               (row.value?.trim() && row.value !== row.auto_value)));
          if (protectedRows.length) {
            feedback("该样机有已手填或主动清空的实测结果，不能移除；请先核对相关项目，人工内容未删除。", true);
            return;
          }
          if (task.projects.some((project) => !(project.variant === "1 model" ||
              (sampleSelectable(project) && project.water_output_labels != null)) &&
              project.rows.some((row) => row.sample_label === sample.label))) {
            feedback("该样机仍被固定行数的结果方案引用，请先改选水表1 model或取消相关项目，再移除样机。", true);
            return;
          }
          if (!window.confirm(`移除 ${sample.label}（${sample.number || "编号待填"}）及其自动结果？若原始记录仍识别到该样机，重新解析时会重新添加。`)) return;
          group.labels = group.labels.filter((label) => label !== sample.label);
          task.samples = task.samples.filter((item) => item !== sample);
          task.projects.forEach((project) => {
            project.rows = project.rows.filter((row) => row.sample_label !== sample.label);
            project.water_result_labels = (project.water_result_labels || []).filter((label) => label !== sample.label);
            if (project.water_output_labels != null)
              project.water_output_labels = project.water_output_labels.filter(label => label !== sample.label);
            project.confirmed = false;
          });
          task.samples_confirmed = false; dirty(); render();
        }, "secondary small"))),
        button("添加样机", () => {
          if (state.busy) return;
          const original = task.water_sample_source_schema.groups[groupIndex];
          const prefix = original.labels[0].replace(/\d{3}$/, "");
          const minimum = Math.max(0, ...group.labels.map((label) => Number(label.slice(prefix.length)))) + 1;
          const number = Math.max(task.water_sample_next_numbers[prefix] || 1, minimum);
          if (number > 999) {
            feedback("该组样机报告代号已用尽，不能重用旧样机代号。", true); return;
          }
          const label = `${prefix}${String(number).padStart(3, "0")}`;
          group.labels.push(label);
          const added = { label, number: "", number_origin: "manual", auto_number: "", record_number: "",
            number_source: "", added_manually: true, disassembly: false, sealed: false };
          const index = task.samples.findIndex((item) => schema.groups.slice(groupIndex + 1)
            .some((next) => next.labels.includes(item.label)));
          task.samples.splice(index < 0 ? task.samples.length : index, 0, added);
          task.water_sample_next_numbers[prefix] = number + 1;
          task.samples_confirmed = false;
          task.projects.forEach((project) => { project.confirmed = false; });
          dirty(); render();
          feedback(`已添加 ${label}，请手写编号，再解析本次资料；没有原始记录的结果将提示待补录。`);
          document.querySelector(`input[aria-label="${label} 样机编号"]`)?.focus({ preventScroll: true });
        }, "secondary small")) : null);
    });
    root.append(el("div", { class: "table-wrap", "data-sample-table": "" }, el("table", { "aria-label": "样机型号规格与编号" },
      el("thead", {}, el("tr", {}, header.map((h) => el("th", { scope: "col" }, h)))),
      el("tbody", {}, rows))));
    if (!editable) {
      const disassembly = task.samples.find((s) => s.disassembly);
      root.append(el("p", {}, `拆样样机：${disassembly ? disassembly.number || `${disassembly.label}（编号待填）` : "未选择（输出保留模板说明）"}`));
      root.append(el("p", {}, `封样样机：${task.samples.filter((s) => s.sealed).map((s) => s.number || `${s.label}（编号待填）`).join("、") || "未选择"}`));
    }
    if (editable) root.append(
      button("未填编号使用对应代号", () => {
        task.samples.filter((s) => !s.number.trim()).forEach((s) => updateSampleNumber(s, s.label));
        render();
      }, "secondary small"),
      button("清空样机对应关系", () => {
        if (!window.confirm("清空所有样机编号、拆样及封样选择和项目中的对应编号？旧自动结果清除，人工结果保留并提示核对。主动清空的编号不会被自动补填。用于更换不同样机表的大纲。")) return;
        task.samples.forEach((s) => { updateSampleNumber(s, ""); s.disassembly = false; s.sealed = false; });
        task.projects.forEach((p) => { p.rows.forEach((r) => { r.sample_id = ""; }); p.confirmed = false; });
        dirty(); render();
      }, "secondary small"));
    if (task.legacy_sample_fields.length) root.append(
      notice(["旧任务样机字段保留如下，请对照上方清单核对，不会猜测或覆盖对应关系：",
        ...task.legacy_sample_fields.map((f) => `${f.label}：${f.value}`)], "warning"),
      editable ? button("处理旧字段，使用上方清单", () => {
        if (!window.confirm("确认旧字段已核对并迁入上方清单？")) return;
        task.legacy_sample_fields = []; dirty(); render();
      }, "secondary small") : null);
    return root;
  }
  function openSampleMatching(parsed) {
    const draft = clone(parsed), errorBox = el("div", { role: "alert" });
    const dialog = el("dialog", { class: "results-editor", "aria-label": "核对样机对应关系" });
    const edited = new Set();
    const rows = [];
    draft.sample_schema.groups.forEach((group) => group.labels.forEach((label) => {
      const sample = draft.samples.find((s) => s.label === label);
      const choices = draft.sample_candidates[label] || [];
      const evidence = el("small");
      const showEvidence = () => {
        evidence.textContent = choices.filter((r) => r.sample_id === sample.number)
          .map((r) => `证书：${r.certificate || "未注明"}；${r.source}`).join("；");
      };
      const select = el("select", { class: "cell-input", "aria-label": `${label} 对应原始编号`,
        onchange: (event) => {
          sample.number = event.target.value;
          sample.number_origin = "manual"; sample.auto_number = ""; sample.number_source = "";
          edited.add(label); showEvidence(); state.resultsEditorDirty = true;
        } },
        el("option", { value: "" }, "未匹配／保留空白"),
        choices.map((r) => el("option", { value: r.sample_id }, `${r.sample_id} · ${r.model}`)));
      if (sample.number && !choices.some((r) => r.sample_id === sample.number))
        select.append(el("option", { value: sample.number }, `${sample.number}（原有手填值，不在当前候选中）`));
      select.value = sample.number;
      showEvidence();
      rows.push(el("tr", {}, el("td", {}, label),
        el("td", {}, group.values.filter((_, i) => i !== draft.sample_schema.quantity_column).join(" / ")),
        el("td", {}, select, evidence)));
    }));
    const close = () => {
      if (state.busy) return;
      if (state.resultsEditorDirty && !window.confirm("放弃尚未应用的样机匹配修改？")) return;
      state.resultsEditorDirty = false; dialog.close(); dialog.remove();
    };
    const fields = el("fieldset", {}, el("div", { class: "table-wrap" }, el("table", {},
      el("thead", {}, el("tr", {}, ["报告代号", "大纲型号规格", "原始记录编号（按出现顺序）"].map((h) => el("th", {}, h)))),
      el("tbody", {}, rows))), el("div", { class: "toolbar results-editor-actions" },
      el("button", { type: "submit", class: "button primary" }, "应用对应关系并识别、保存"),
      button("取消", close)));
    const form = el("form", {}, fields);
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (state.busy) return;
      fields.disabled = true; errorBox.replaceChildren();
      const ok = await run("正在按所选编号识别并保存…", async () => {
        const numbers = draft.samples.map((s) => s.number).filter(Boolean);
        if (new Set(numbers).size !== numbers.length) throw new Error("不同报告代号不能选择同一个真实编号。");
        const candidate = clone(state.task);
        candidate.samples.forEach((sample) => {
          const incoming = draft.samples.find((s) => s.label === sample.label);
          if (!incoming) return;
          if (sample.number !== incoming.number || edited.has(sample.label))
            updateSampleNumber(sample, incoming.number, { task: candidate, manual: edited.has(sample.label) });
          Object.assign(sample, incoming);
        });
        candidate.samples_confirmed = false;
        const matched = await api("/api/tasks/parse", candidate);
        const saved = await api("/api/tasks", matched);
        state.task = normalizeTask(saved);
        state.epoch += 1; state.validation = null; state.dirty = false;
        render();
      }, "样机对应及实测结果已保存。未匹配的项目或结果请查看逐行提示。");
      fields.disabled = false;
      if (ok) { state.resultsEditorDirty = false; dialog.close(); dialog.remove(); }
      else errorBox.replaceChildren(notice($("feedback").textContent, "error"));
    });
    dialog.append(el("h2", {}, "核对一次样机对应关系"),
      notice("此窗口仅用于人工纠正下表对应关系，不自动补填空白或跨规格分配。未改动的自动编号保留来源；手动选择或清空的编号按人工指定保存。应用后重新识别全部已选项目并保存当前任务。"));
    if (draft.sample_mapping_issues.length) dialog.append(notice(draft.sample_mapping_issues, "warning"));
    dialog.append(errorBox, form);
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); });
    document.body.append(dialog); dialog.showModal();
  }
  function renderObservationCertificates() {
    const panel = card("观察与一览表备注 · 防爆与防护证书",
      el("p", { class: "muted" }, "适用于所有对象，同一证书编号同步填写观察项目及第五部分项目一览表的对应备注，不改变评价标记。防爆编号从本次大纲的补充说明/其他说明查询，防护编号手工填写。封印和防护项目不修改；空值保留各素材原文并提示，不代表已核实证书。"));
    panel.setAttribute("data-observation-certificates", "");
    const certificates = state.task.observation_certificates;
    if (!certificates) {
      panel.append(notice("选择观察素材或项目一览表后载入证书项目；旧任务请保存或重新打开以查询。"));
      return panel;
    }
    if (certificates.issues?.length) panel.append(notice(certificates.issues, "warning"));
    let count = 0;
    [["explosion", "防爆证书编号"], ["protection", "防护证书编号"]].forEach(([key, label]) => {
      const item = certificates[key];
      if (!item?.targets?.length) return;
      count++;
      const box = card(label,
        el("p", { class: "muted" }, `对应项目：${item.targets.join("、")}`),
        field(label, item.value, (value) => {
          item.value = value; item.manual_override = true; item.issues = [];
          dirty();
        }, { placeholder: "仅填写证书编号；未填写保留素材原文", attrs: { "data-certificate-field": key } }));
      if (item.issues?.length) box.append(notice(item.issues, "warning"));
      const sources = el("details", { class: "overview-details" },
        el("summary", {}, "原备注、来源与查询结果",
          item.source_issues?.length ? el("span", { class: "pill warning" }, `${item.source_issues.length}项来源提醒`) : null),
        el("p", { class: "muted" }, `素材原备注：${item.originals.join("；") || "空白"}`));
      if (key === "explosion") {
        box.append(el("p", { class: "muted" }, item.manual_override
          ? "当前采用人工填写值；重新解析、保存和重开不会覆盖，主动清空也会保留。"
          : "当前采用大纲自动查询；来源缺失或有歧义时不自动填写。"));
        if (item.source_issues?.length) sources.append(notice(item.source_issues, "warning"));
        (item.candidates || []).forEach((candidate) => sources.append(card("大纲查询结果",
          el("p", {}, candidate.text),
          el("small", {}, candidate.source),
          button(`采用编号 ${candidate.value}`, () => {
            item.value = candidate.value; item.manual_override = true; item.issues = [];
            dirty(); render();
          }, "small"))));
        box.append(button("重新查询并恢复自动填写", () => {
          if (item.manual_override && !window.confirm("恢复自动查询将替换当前人工填写（包括主动清空）的防爆证书编号，是否继续？")) return;
          run("正在查询本次大纲防爆信息…", async () => {
            const candidate = clone(state.task);
            candidate.observation_certificates.explosion.manual_override = false;
            state.task = normalizeTask(await api("/api/tasks/order", candidate));
            dirty(); render();
          }, "已重新查询本次大纲，请核对证书信息并保存任务。");
        }, "small"));
      }
      box.append(sources);
      panel.append(box);
    });
    if (!count) panel.append(el("p", { class: "muted" }, "当前所选观察素材和项目一览表没有需填写证书的防爆/防护项目。"));
    return panel;
  }
  function projectVariantLabel(project) {
    const asset = state.catalog.assets.find(item => item.kind === "result" && item.id === project.asset_id);
    return asset?.variant || project.variant || "未标注方案";
  }
  function projectMaterialTitle(project) {
    const asset = state.catalog.assets.find(item => item.kind === "result" && item.id === project.asset_id);
    const candidates = [asset?.name, asset?.variant, ...[...(asset?.folders || [])].reverse(), project.variant, project.name];
    const numbered = candidates.find(value => /^\s*\d+\s*[.．、]\s*\S/.test(value || ""));
    return (numbered || asset?.project || project.name)
      .replace(/\.docx$/i, "")
      .replace(/\s*[-–—]\s*\d+\s*models?\s*$/i, "")
      .replace(/^\s*(\d+)\s*[.．、]\s*/, "$1.").trim();
  }
  function outputRows(project) {
    return project.rows.filter(row => project.water_output_labels == null ||
      project.water_output_labels.includes(row.sample_label));
  }
  function sampleSelectable(project) {
    return ["1 model", "B级", "O级", "E1级", "E2级"].includes(project.variant);
  }
  function effectiveSampleLabels(project) {
    if (project.water_output_labels != null) return [...project.water_output_labels];
    const eligible = new Set(state.task.samples.filter(sample => !sample.disassembly).map(sample => sample.label));
    return [...new Set(project.rows.map(row => row.sample_label))].filter(label => eligible.has(label));
  }
  function projectSampleChoices(project) {
    if ((state.task.report_type ?? "水表") !== "水表" || !sampleSelectable(project)) return null;
    if (!state.catalog.water_sample_selection)
      return notice("后台尚未启用样机勾选，请先保存当前编辑并重新连接新版服务。", "warning");
    const selected = effectiveSampleLabels(project);
    const samples = state.task.samples;
    const choices = el("div", {class: "project-sample-choices", "aria-label": `${project.name} 输出样机`});
    if (!samples.length) choices.append(el("small", {}, "尚无实际样机，请先在总览识别或添加。"));
    const labels = [...samples.map(sample => sample.label),
      ...selected.filter(label => !samples.some(sample => sample.label === label))];
    labels.forEach(label => {
      const sample = samples.find(item => item.label === label);
      const key = `${project.asset_id}|${label}`;
      const input = el("input", {type: "checkbox", checked: selected.includes(label),
        disabled: !!sample?.disassembly, "data-project-sample": key,
        "aria-label": `${projectMaterialTitle(project)} ${label} 输出样机`});
      input.addEventListener("change", () => {
        if (state.busy) { input.checked = selected.includes(label); return; }
        const next = input.checked ? [...selected, label] : selected.filter(value => value !== label);
        const top = window.scrollY, left = window.scrollX;
        run("正在更新本项目参试样机…", async () => {
          const candidate = clone(state.task);
          candidate.projects.find(item => item.id === project.id).water_output_labels = next;
          const updated = await api("/api/tasks/order", candidate);
          if (JSON.stringify(updated.projects.find(item => item.id === project.id)?.water_output_labels) !== JSON.stringify(next))
            throw new Error("后台未保留样机勾选，本次变更未应用。请重新连接新版服务；当前任务内容仍保留。");
          state.task = normalizeTask(updated);
          dirty(); render();
          [...document.querySelectorAll("[data-project-sample]")].find(item =>
            item.dataset.projectSample === key)?.focus({preventScroll: true});
          window.scrollTo({top, left, behavior: "instant"});
        }, "样机勾选已更新，取消勾选的人工结果仍保留；请保存任务。").then(ok => {
          if (!ok) render();
        });
      });
      choices.append(el("label", {class: "check", title: sample
        ? `真实编号：${sample.number || "待填写"}${sample.disassembly ? "；拆样不参加实测" : ""}`
        : "已不在本次样机清单，请取消勾选"}, input, label,
        sample?.disassembly ? el("small", {}, "拆样") : !sample ? el("small", {class: "unconfirmed"}, "已缺失") : null));
    });
    return choices;
  }
  function sampleSchemePayload(update) {
    const projects = state.task.projects;
    if ((state.task.report_type ?? "水表") !== "水表" || !projects.length ||
        projects.some(project => !sampleSelectable(project)))
      throw new Error("整套样机方案目前支持水表单型号及B/O/E1/E2等级素材，不支持2 models。");
    const saved = state.sampleSchemes.find(item => item.id === state.sampleSchemeId);
    if (update && !saved) throw new Error("请先选择要更新的已保存方案。");
    return {id: update ? saved.id : "", revision: update ? saved.revision : 0,
      name: state.sampleSchemeName.trim(), report_type: "水表",
      projects: projects.map(project => ({asset_id: project.asset_id, labels: effectiveSampleLabels(project)}))};
  }
  function renderSampleSchemes() {
    if ((state.task.report_type ?? "水表") !== "水表") return null;
    if (!state.catalog.water_sample_selection)
      return notice("当前后台尚未启用样机方案库。请先保存编辑，待新版服务启用后重新打开任务。", "warning");
    const save = update => {
      let completed = false;
      return run("正在保存整套样机方案…", async () => {
      const payload = sampleSchemePayload(update);
      if (!payload.name) throw new Error("请填写样机方案名称。");
      if (update && !window.confirm("用当前所有项目和样机勾选更新所选方案？只修改方案库，不覆盖其他任务。")) return;
      const saved = await api("/api/sample-schemes", payload);
      state.sampleSchemes = [saved, ...state.sampleSchemes.filter(item => item.id !== saved.id)];
      state.sampleSchemeId = saved.id; state.sampleSchemeName = saved.name;
      completed = true; render();
      }, () => completed ? "方案已保存到库；任务本身仍需单独保存。" : "已取消更新，方案库未改变。");
    };
    return card("可复用样机方案",
      el("p", {class: "muted"}, "整套保存结果素材项目及A001、B001等代号勾选，不保存真实编号、实测值、日期或人员。应用时选入方案项目，保留已有其他项目及人工数据；不同任务缺少的代号明确提示，不自动补造。"),
      el("p", {class: "muted"}, "修改已有方案：先选择并应用，调整项目或样机勾选，再点击“更新所选方案”。不改变项目一览表及观察素材。"),
      state.sampleSchemeIssues.length ? notice(["上次应用方案的核对提醒：", ...state.sampleSchemeIssues], "warning") : null,
      el("div", {class: "grid"},
        field("已保存方案", state.sampleSchemeId, identity => {
          state.sampleSchemeId = identity;
          state.sampleSchemeName = state.sampleSchemes.find(item => item.id === identity)?.name || "";
          render();
        }, {choices: [["", "请选择整套方案"], ...state.sampleSchemes.map(item =>
          [item.id, `${item.name} · ${item.projects.length}项目 · 版本${item.revision}`])]}),
        field("方案名称", state.sampleSchemeName, value => {state.sampleSchemeName = value;},
          {placeholder: "例如：水表常规六台方案"})),
      el("div", {class: "toolbar"},
        button("一键应用方案", () => {
          let applied = false;
          run("正在选入方案项目并匹配本次样机…", async () => {
            const result = await api(`/api/sample-schemes/${encodeURIComponent(state.sampleSchemeId)}/apply`, clone(state.task));
            if (result.issues.length && !window.confirm(result.issues.join("\n") +
                "\n只应用本次存在的参试代号，其余项目保留技术要求但不输出样机数据，是否继续？")) return;
            state.task = normalizeTask(result.task); state.orderVerified = true;
            state.sampleSchemeIssues = result.issues;
            dirty(); render(); applied = true;
          }, () => applied ? "方案已应用到当前页面，未覆盖已保存任务；请核对并保存。" : "已取消应用，当前任务未改变。");
        }, "primary small", !state.sampleSchemeId),
        button("保存为新方案", () => save(false), "secondary small"),
        button("更新所选方案", () => save(true), "secondary small", !state.sampleSchemeId),
        button("刷新方案库", () => run("正在读取方案库…", async () => {
          state.sampleSchemes = await api("/api/sample-schemes"); render();
        }, "方案库已刷新；本次项目勾选及任务未改变。"), "secondary small")));
  }
  function materialSchemePayload(update) {
    const reportType = state.task.report_type ?? "水表";
    const saved = state.materialSchemes.find(item =>
      item.id === state.materialSchemeId && item.report_type === reportType);
    if (update && !saved) throw new Error("请先选择要更新的素材方案。");
    if (!state.materialSchemeName.trim()) throw new Error("请填写素材方案名称。");
    if (!state.task.projects.length) throw new Error("请先勾选需要保存的结果素材。");
    return {id: update ? saved.id : "", revision: update ? saved.revision : 0,
      name: state.materialSchemeName.trim(), report_type: reportType,
      asset_ids: state.task.projects.map(project => project.asset_id)};
  }
  async function applyMaterialScheme(scheme) {
    const current = state.task;
    if (!scheme || scheme.report_type !== (current.report_type ?? "水表"))
      throw new Error("请先选择当前报告对象的素材方案。");
    if (!current.items_asset_id) throw new Error("请先选择项目一览表，再应用素材方案。");
    const selected = new Set(scheme.asset_ids);
    const removed = current.projects.filter(project => !selected.has(project.asset_id));
    if (removed.length && !window.confirm(`应用「${scheme.name}」将替换当前勾选，并移除以下${removed.length}个项目及其本次编辑数据：\n` +
        removed.map(project => `${projectMaterialTitle(project)} · ${projectVariantLabel(project)}`).join("\n") +
        "\n仍被选中的项目保留本次填写；已保存任务暂不覆盖。是否继续？")) return false;
    const updated = normalizeTask(await api(
      `/api/material-schemes/${encodeURIComponent(scheme.id)}/apply?revision=${scheme.revision}`, clone(current)));
    const returned = new Set(updated.projects.map(project => project.asset_id));
    if (updated.projects.length !== selected.size || returned.size !== selected.size ||
        updated.projects.some(project => !selected.has(project.asset_id)))
      throw new Error("后台返回的素材勾选与方案不一致，本次方案未应用，当前任务仍保留。");
    const existing = new Set(current.projects.map(project => project.asset_id));
    updated.projects.filter(project => !existing.has(project.asset_id))
      .forEach(project => inheritOrdinaryStaff(project, current.projects[0]));
    state.task = updated;
    removed.forEach(project => {
      state.clockDrafts.delete(`${project.id}:start`); state.clockDrafts.delete(`${project.id}:end`);
      state.autoEndTimes.delete(project.id); state.autoStartTimes.delete(project.id);
    });
    state.orderVerified = true;
    dirty(); render();
    return true;
  }
  function renderMaterialSchemes() {
    const reportType = state.task.report_type ?? "水表";
    if (state.materialSchemeType !== reportType) {
      state.materialSchemeType = reportType;
      state.materialSchemeId = ""; state.materialSchemeName = "";
    }
    const schemes = state.materialSchemes.filter(item => item.report_type === reportType);
    const saved = schemes.find(item => item.id === state.materialSchemeId);
    const save = update => {
      let completed = false;
      return run("正在保存结果素材勾选方案…", async () => {
        const payload = materialSchemePayload(update);
        if (update && !window.confirm("用当前结果素材勾选更新所选方案？其他任务和原素材不会改变。")) return;
        const result = await api("/api/material-schemes", payload);
        state.materialSchemes = [result, ...state.materialSchemes.filter(item => item.id !== result.id)];
        state.materialSchemeId = result.id; state.materialSchemeName = result.name;
        completed = true; render();
      }, () => completed ? "素材方案已命名并保存到本机，关闭页面或重启服务后仍可复用。"
        : "已取消更新，方案库未改变。");
    };
    return card("我的素材方案",
      el("p", {class: "muted"}, `保存当前${reportType}结果素材勾选，命名后可在以后任务中选择并一键应用。这里只保存素材组合，不保存样机勾选、实测值、日期或人员，也不改变一览表和观察素材。`),
      el("p", {class: "muted"}, "应用会替换当前勾选；移除项目及其编辑数据前需确认。仍选中的项目保留本次填写，新增项目按本次资料匹配；应用后须核对并保存任务。"),
      el("div", {class: "grid"},
        field("已保存素材方案", saved?.id || "", identity => {
          state.materialSchemeId = identity;
          state.materialSchemeName = schemes.find(item => item.id === identity)?.name || "";
          render();
        }, {choices: [["", `请选择${reportType}素材方案`], ...schemes.map(item =>
          [item.id, `${item.name} · ${item.asset_ids.length}个素材 · 版本${item.revision}`])]}),
        field("素材方案名称", state.materialSchemeName, value => {state.materialSchemeName = value;},
          {placeholder: "例如：超声波燃气表三型号常规项目"})),
      el("div", {class: "toolbar"},
        button("应用所选素材方案", () => {
          let applied = false;
          return run("正在应用素材方案并核对一览表…", async () => {
            applied = await applyMaterialScheme(saved);
          }, () => applied ? "素材方案已应用，勾选已替换；请核对并保存当前任务。"
            : "已取消应用，当前任务和勾选均未改变。");
        }, "primary small", !saved || !state.task.items_asset_id),
        button("保存当前勾选为新方案", () => save(false), "secondary small", !state.task.projects.length),
        button("更新所选素材方案", () => save(true), "secondary small", !saved || !state.task.projects.length),
        button("刷新素材方案库", () => run("正在读取素材方案库…", async () => {
          state.materialSchemes = await api("/api/material-schemes");
          if (!state.materialSchemes.some(item => item.id === state.materialSchemeId && item.report_type === reportType)) {
            state.materialSchemeId = ""; state.materialSchemeName = "";
          }
          render();
        }, "素材方案库已刷新，当前任务和勾选未改变。"), "secondary small")));
  }
  function renderMaterials() {
    const task = state.task;
    const root = el("div");
    const reportType = task.report_type ?? "水表";
    const folder = state.catalog.report_type_folders?.[reportType] ?? reportType;
    const typedAssets = state.catalog.assets.filter((asset) => (asset.report_type ?? "水表") === reportType);
    const loading = el("div", {"data-material-loading": ""});
    root.append(loading);
    paintMaterialLoading(loading);
    root.append(notice(`本次对象：${reportType}。观察、一览表和结果素材分别读取对应的「${folder}」文件夹；新增Word后点击“刷新素材目录”。`));
    const selectors = el("div", { class: "grid" });
    [["items", "项目一览表", "items_asset_id", "items"], ["observe", "基本信息与观察记录", "observe_asset_id", "observations"]].forEach(([kind, title, key, rows]) => {
      const assets = typedAssets.filter((a) => a.kind === kind);
      selectors.append(card(title,
        !assets.length ? notice(`${reportType}的${title}目录暂无Word素材，请先在对应的「${folder}」文件夹添加文件。`, "warning") : null,
        field("素材方案", task[key], (id) => {
          run("正在选择章节素材…", async () => {
            const candidate = clone(task);
            candidate[key] = id; candidate[rows] = [];
            state.task = kind === "items" ? await orderedTask(candidate) : normalizeTask(await api("/api/tasks/order", candidate));
            dirty(); render();
          }, () => kind === "items" && !id ? "已取消一览表选择。请重新选择一览表，当前项目顺序不视为已核对。"
            : reportType !== "水表" ? "章节素材已更新；导出时替换观察型号和样机编号，同步填写观察及一览表证书编号。"
              + (reportType === "燃气表" ? "流量范围要求按总览各型号参数替换，分界流量=0.1×最大流量；其余要求与结论保留。" : "技术要求与结论原样合成。")
            : "章节素材已更新。观察模块更新指定水表备注，制造年月和电池期限请到总览填写；观察与一览表同步填写证书编号，其余一览表内容保留。").then((ok) => { if (!ok) render(); });
        }, { required: true, choices: [["", "请选择，不设默认方案"], ...assets.map((a) => [a.id, a.name])] }),
        el("small", {}, kind === "observe"
          ? reportType !== "水表"
            ? "两份Word中的观察型号和样机编号按基本信息总览替换，排除拆样样机；同步填写防爆、防护证书编号，不套用水表参数。"
              + (reportType === "燃气表" ? "流量范围按各型号的总览最小/最大流量填写，分界流量=0.1×最大流量；其余技术要求与已有结论保留。" : "技术要求与已有结论原样合成。")
              + "素材旧分页续段安全接排，统一英文Arial、中文宋体。"
            : "观察备注中的指定水表参数取本次大纲，制造年月和电池更换期限在总览手工填写；型评报告和配套原始记录同步替换。其他文字、基本信息、要求及已有勾选保留，统一英文Arial、中文宋体。"
          : "项目一览表防爆、防护性能备注使用总览中与观察项目共用的证书编号，其余文字、表格和已有勾选原样保留，不自动同步结论。统一英文Arial、中文宋体，标题另起一页。")));
    });
    root.append(selectors);
    root.append(card("观察与证书信息",
      el("p", { class: "muted" }, "防爆、防护证书统一在基本信息总览维护，观察项目与一览表共用同一编号，这里不重复填写。"),
      button("前往总览 · 观察与证书信息", () =>
        locateIssue({ step: 1, overview_group: "observation" }), "secondary small")));
    const selected = el("div");
    task.projects.forEach((p) => selected.append(el("div", { class: "file-row selected-result-row", "data-selected-result": p.asset_id },
      el("span", { class: "selected-result-title", title: projectVariantLabel(p) }, projectMaterialTitle(p)),
      projectSampleChoices(p),
      button("移除", () => removeProject(p), "danger small"))));
    root.append(card("本次结果素材顺序",
      orderNotice(),
      reportType === "水表" ? el("p", {class: "muted"}, "单型号及B/O/E1/E2等级项目只输出勾选样机的数据；拆样不能勾选。未勾选任何样机时保留标题和技术要求，不输出编号及实测数据。") : null,
      task.projects.length ? selected : empty("尚未选择结果素材", "在下方素材库逐个勾选需要的方案。")));
    root.append(renderSampleSchemes());
    const list = el("div");
    const search = field("查找结果素材", state.materialQuery, (text) => {
      state.materialQuery = text; paint();
    }, { placeholder: "输入文件夹、项目名称或素材方案" });
    search.classList.add("search");
    function paint() {
      const query = state.materialQuery.trim().toLowerCase();
      const assets = typedAssets.filter((a) => a.kind === "result" &&
        `${a.folders.join(" ")} ${a.project} ${a.name}`.toLowerCase().includes(query));
      function assetCard(asset) {
        const project = task.projects.find((p) => p.asset_id === asset.id);
        const pending = state.materialSelection;
        const input = el("input", { type: "checkbox", id: `asset-${++uid}`,
          checked: !!project || pending?.asset_id === asset.id,
          disabled: (!task.items_asset_id && !project) || Boolean(pending && pending.asset_id !== asset.id) });
        input.addEventListener("change", () => {
          input.checked = !!project;
          if (state.materialSelection?.asset_id === asset.id) {
            cancelResultSelection("已取消素材选择，当前任务保持不变。"); return;
          }
          if (project) { removeProject(project); return; }
          if (!task.items_asset_id) { feedback("请先选择项目一览表，再选择结果素材。", true); return; }
          return selectResultMaterial(asset);
        });
        return el("div", { class: `check-card${project ? " selected" : ""}`, "data-asset-id": asset.id }, input,
          el("label", { for: input.id }, el("strong", {}, asset.name)),
          project ? el("span", { class: "pill" }, `第 ${task.projects.indexOf(project) + 1} 项`) : null);
      }
      const tree = { folders: new Map(), assets: [] };
      assets.forEach((asset) => {
        let node = tree;
        asset.folders.forEach((name) => {
          if (!node.folders.has(name)) node.folders.set(name, { folders: new Map(), assets: [] });
          node = node.folders.get(name);
        });
        node.assets.push(asset);
      });
      function folderContent(node, parents = []) {
        return [
          ...[...node.folders].sort(([a], [b]) => a.localeCompare(b, "zh-CN", { numeric: true })).map(([name, child]) => {
            const folders = [...parents, name];
            const path = folders.join("/");
            const contained = assets.filter((asset) => folders.every((part, i) => asset.folders[i] === part));
            const selectedCount = contained.filter((asset) => task.projects.some((p) => p.asset_id === asset.id)).length;
            const details = el("details", {
              class: "material-folder", open: !!query || state.materialFolders.has(path), "data-folder": path
            }, el("summary", {}, el("strong", {}, name),
              el("span", { class: "pill" }, `${contained.length} 个方案${selectedCount ? ` · 已选 ${selectedCount}` : ""}`)),
            el("div", { class: "material-folder-content" }, folderContent(child, folders)));
            details.addEventListener("toggle", () => {
              if (!details.isConnected || query) return;
              if (details.open) state.materialFolders.add(path);
              else state.materialFolders.delete(path);
            });
            return details;
          }),
          ...[...node.assets]
            .sort((a, b) => a.name.localeCompare(b.name, "zh-CN", { numeric: true }))
            .map(assetCard)
        ];
      }
      list.replaceChildren(...folderContent(tree));
      if (!assets.length) list.append(empty("没有匹配的素材",
        typedAssets.some((asset) => asset.kind === "result") ? "请更换搜索词或核对本机素材目录。"
          : `${reportType}结果素材目录「${folder}」暂无Word文件，请添加后刷新。`));
    }
    paint();
    root.append(card("结果素材库 · 按文件夹选择",
      el("p", { class: "muted" }, "文件夹和方案按名称中的序号自然排列；点击文件夹展开方案。在本机素材目录添加文件夹或 Word 后，点击刷新素材目录即可发现。搜索时自动展开匹配项。"),
      button("刷新素材目录（保留当前编辑）", () => run("正在重新读取本机素材目录…", async () => {
        state.catalog = await api("/api/catalog");
        state.orderVerified = false;
        render();
      }, "素材目录已刷新，当前选择和填写内容保留。生成前请按一览表重新核对顺序。"), "secondary small"),
      renderMaterialSchemes(),
      search, list));
    return root;
  }
  function removeProject(project) {
    if (!window.confirm(`移除「${project.name}」及其当前编辑数据？重新选择时按一览表归位，不恢复已移除的编辑数据。`)) { render(); return; }
    run("正在移除项目并重新核对一览表顺序…", async () => {
      const candidate = clone(state.task);
      candidate.projects = candidate.projects.filter((p) => p.id !== project.id);
      state.task = await orderedTask(candidate);
      state.clockDrafts.delete(`${project.id}:start`); state.clockDrafts.delete(`${project.id}:end`);
      state.autoEndTimes.delete(project.id); state.autoStartTimes.clear();
      dirty(); render();
    }, "项目已移除；剩余项目按当前一览表核对。未选一览表时请先选择一览表。");
  }
  function staffGroups(project) {
    const definitions = [
      ["液体组", ["吴锦川", "姚新红"]],
      ["气体组", ["陈超", "沈琼", "朱碧玉", "刘夷平"]],
      ["环境试验组", ["张宇", "黄薛青"]],
      ["电磁兼容组", ["刘雨宸", "秦夏臻", "杨先望"]]
    ];
    const required = new Set();
    if ((state.task.report_type ?? "水表") === "水表") required.add("液体组");
    if (state.task.report_type === "燃气表") required.add("气体组");
    state.task.equipment.filter((item) => project.equipment_ids.includes(item.id))
      .forEach((item) => { if (item.trial_group) required.add(item.trial_group); });
    return definitions.filter(([label]) => required.has(label));
  }
  function replaceStaffGroup(project, key, names, name) {
    const next = project[key].filter((person) => !names.includes(person));
    if (name) next.push(name);
    if (next.length === project[key].length && next.every((person) => project[key].includes(person))) return false;
    project[key] = next;
    project.confirmed = false;
    return true;
  }
  function inheritOrdinaryStaff(project, first) {
    if (!first) return;
    const sources = new Map(staffGroups(first));
    staffGroups(project).forEach(([label, target]) => {
      const source = sources.get(label);
      if (!source) return;
      ["evaluators", "reviewers"].forEach((key) => {
        const name = first[key].find((person) => source.includes(person));
        if (name) replaceStaffGroup(project, key, target, name);
      });
    });
  }
  function trimProjectStaff(project) {
    const allowed = new Set(staffGroups(project).flatMap(([, names]) => names));
    ["evaluators", "reviewers"].forEach((key) => {
      project[key] = project[key].filter((name) => allowed.has(name));
    });
  }
  function changeProjectStaff(project, key, label, names, name) {
    if (!replaceStaffGroup(project, key, names, name)) return;
    const affected = [project];
    state.task.projects.forEach((other) => {
      if (other === project) return;
      const group = staffGroups(other).find(([groupLabel]) => groupLabel === label);
      if (group && replaceStaffGroup(other, key, group[1], name)) affected.push(other);
    });
    feedback(`已将${label}${key === "evaluators" ? "评价" : "复核"}人员${name ? `“${name}”` : "清空选择"}同步到本次全部使用该组的项目；其他小组及另一角色不变。${affected.some((item) => item.evaluators.some((person) => item.reviewers.includes(person))) ? "部分项目评价与复核重名，请处理人员待核对提示。" : ""}`);
    affected.forEach((item) => {
      dirty(item);
      const pane = [...document.querySelectorAll("[data-project-row]")].find((row) => row.dataset.projectRow === item.id);
      if (!pane) return;
      syncStaffChoices(pane, item);
      pane.querySelector("[data-staff-warning]").textContent = staffIssues(item).length ? "人员待核对" : "";
    });
  }
  function staffEditor(project, key, title) {
    const groups = staffGroups(project);
    const group = el("fieldset", {}, el("legend", {}, `${title} · ${groups.length} 名（必选）`));
    groups.forEach(([label, names]) => group.append(field(label, project[key].find((n) => names.includes(n)) || "", (name) => {
      const opposite = key === "evaluators" ? "reviewers" : "evaluators";
      if (name && project[opposite].includes(name)) {
        const select = group.querySelector(`select[data-staff-role="${key}"] option[value="${name}"]`)?.parentElement;
        if (select) select.value = project[key].find((n) => names.includes(n)) || "";
        feedback("同一项目的评价与复核名单不得相交。请先取消另一角色对该人员的选择。", true);
        return;
      }
      changeProjectStaff(project, key, label, names, name);
    }, { required: true, choices: [["", label], ...names], attrs: {
      "data-staff-role": key, "data-staff-group": label, "aria-label": `${project.name} ${title} ${label}`
    } })));
    return group;
  }
  function syncStaffChoices(pane, project) {
    pane.querySelectorAll("select[data-staff-role]").forEach((select) => {
      const names = staffGroups(project).find(([label]) => label === select.dataset.staffGroup)?.[1] || [];
      select.value = project[select.dataset.staffRole].find((name) => names.includes(name)) || "";
      const opposite = select.dataset.staffRole === "evaluators" ? "reviewers" : "evaluators";
      const label = opposite === "evaluators" ? "评价" : "复核";
      Array.from(select.options).forEach((option) => {
        if (!option.value) return;
        option.disabled = project[opposite].includes(option.value);
        option.textContent = `${option.value}${option.disabled ? `（已选为${label}人员）` : ""}`;
      });
    });
  }
  function staffIssues(project) {
    const issues = [];
    const groups = staffGroups(project);
    [["evaluators", "评价"], ["reviewers", "复核"]].forEach(([key, label]) => {
      if (project[key].length !== groups.length || groups.some(([, names]) => project[key].filter((n) => names.includes(n)).length !== 1))
        issues.push(`${label}人员需${groups.map(([name]) => `${name} 1 名`).join(" + ")}。`);
    });
    if (project.evaluators.some((name) => project.reviewers.includes(name))) issues.push("评价与复核人员不得相同或交叉，请重新选择。");
    return issues;
  }
  function projectClock(project, key, change) {
    const label = `${project.name} ${key === "start" ? "开始" : "结束"}`;
    const identity = `${project.id}:${key}`;
    const raw = project[key] || "";
    const match = raw.trim().match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T]*(\d{1,2})[:时](\d{2})(?:[:分](\d{2})秒?|分)?)?$/);
    const parts = match ? { date: `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`,
      hour: match[4]?.padStart(2, "0") || "", minute: match[5] || "", seconds: match[6] || "" }
      : { date: "", hour: "", minute: "", seconds: "" };
    const pending = state.clockDrafts.get(identity);
    if (pending && pending.base !== raw) state.clockDrafts.delete(identity);
    if (pending?.base === raw) Object.assign(parts, pending);
    const date = el("input", { type: "date", value: parts.date, "aria-label": `${label}日期` });
    const extended = staffGroups(project).some(([label]) => ["环境试验组", "电磁兼容组"].includes(label));
    const hour = el("select", { "aria-label": `${label}时` }, el("option", { value: "" }, "时"),
      Array.from({ length: extended ? 24 : 10 }, (_, i) => {
        const v = String(i + (extended ? 0 : 8)).padStart(2, "0"); return el("option", { value: v }, v);
      }));
    const minute = el("select", { "aria-label": `${label}分` }, el("option", { value: "" }, "分"),
      ["00", "15", "30", "45"].map((v) => el("option", { value: v }, v)),
      el("option", { value: "expand" }, "扩展（00～59）"));
    for (const [control, value] of [[hour, parts.hour], [minute, parts.minute]]) {
      if (value && ![...control.options].some((o) => o.value === value))
        control.append(el("option", { value, disabled: true }, `${value}（原值）`));
      control.value = value;
    }
    let selectedMinute = parts.minute;
    const hint = el("small", { class: "clock-hint", role: "status" },
      pending?.base === raw ? "请补齐日期、时、分" :
      raw && (!match || !date.value) ? `保留原值：${raw}` :
      parts.seconds ? `原值含 ${parts.seconds} 秒；改时分后采用新选择` : "");
    function update(event) {
      if (event.target === minute && minute.value === "expand") {
        minute.replaceChildren(el("option", { value: "" }, "分"),
          ...Array.from({ length: 60 }, (_, i) => {
            const value = String(i).padStart(2, "0");
            return el("option", { value }, value);
          }));
        minute.value = selectedMinute;
        minute.focus();
        return;
      }
      if (key === "start") state.autoStartTimes.delete(project.id);
      if (event.target === date && !date.value) { hour.value = ""; minute.value = ""; parts.seconds = ""; }
      selectedMinute = minute.value;
      if (event.target !== date) parts.seconds = "";
      const incomplete = (hour.value || minute.value) && (!date.value || !hour.value || !minute.value);
      if (incomplete) {
        state.clockDrafts.set(identity, { base: project[key], date: date.value, hour: hour.value,
          minute: minute.value, seconds: parts.seconds });
        if (key === "end") {
          state.autoEndTimes.delete(project.id);
          state.autoStartTimes.delete(project.id);
        }
        dirty(project);
        hint.textContent = "请补齐日期、时、分（尚未写入时间）";
        return;
      }
      state.clockDrafts.delete(identity);
      hint.textContent = "";
      const value = date.value + (hour.value && minute.value ? ` ${hour.value}:${minute.value}${parts.seconds ? `:${parts.seconds}` : ""}` : "");
      change(value);
    }
    [date, hour, minute].forEach((control) => control.addEventListener("change", update));
    return el("div", { class: "project-clock" }, el("div", { class: "clock-controls" }, date, hour, minute), hint);
  }
  function captureProjectView(trigger = document.activeElement) {
    const table = document.querySelector(".project-table-wrap");
    const rowId = trigger?.closest("[data-project-row]")?.dataset.projectRow;
    const action = trigger?.dataset.projectAction, results = trigger?.dataset.resultsProject;
    const top = table?.scrollTop || 0, left = table?.scrollLeft || 0;
    const pageTop = window.scrollY, pageLeft = window.scrollX;
    return () => {
      const current = document.querySelector(".project-table-wrap");
      const row = [...document.querySelectorAll("[data-project-row]")].find((item) => item.dataset.projectRow === rowId);
      const replacement = [...(row?.querySelectorAll("button") || [])].find((item) =>
        action ? item.dataset.projectAction === action : results && item.dataset.resultsProject === results);
      (trigger?.isConnected ? trigger : replacement || current)?.focus({ preventScroll: true });
      if (current) { current.scrollTop = top; current.scrollLeft = left; }
      window.scrollTo({ top: pageTop, left: pageLeft, behavior: "instant" });
    };
  }
  function openProjectDetails(project, trigger = document.activeElement) {
    if (state.busy) return;
    const restoreView = captureProjectView(trigger);
    const dialog = el("dialog", { class: "results-editor", "aria-label": `${project.name} 项目详情` });
    const close = (restore = true) => {
      if (state.busy) return;
      dialog.close(); dialog.remove();
      if (restore) { render(); restoreView(); }
    };
    dialog.append(el("h2", {}, `${project.name} · 来源与设备`),
      notice("这里的修改立即进入当前任务，关闭后请保存任务。项目结论、异常情况固定沿用素材原文；不代表本次已核实或签发。"),
      recordConditions(project));
    if (project.notes.length) dialog.append(notice(project.notes, "warning"));
    const issues = [...staffIssues(project), ...new Set(outputRows(project).map((r) => r.match_issue).filter(Boolean))];
    if (issues.length) dialog.append(notice(issues, "warning"));
    dialog.append(field("项目记录来源", project.source, (value) => { project.source = value; dirty(project); },
      { multiline: true, help: "保留文件、表格、行或记录单位置；手工补录请注明。" }),
      el("h3", {}, "本项目使用的设备"));
    if (!state.task.equipment.length) dialog.append(notice("尚无本次设备，请先到仪器设备库添加快照。", "warning"),
      button("管理设备", () => { if (!state.busy) { close(false); navigate(3); } }, "secondary"));
    state.task.equipment.forEach((equipment, i) => dialog.append(checkbox(
      el("span", {}, `${i + 1}. `, equipmentText(equipment, "name"), " · ", equipmentText(equipment, "number")),
      project.equipment_ids.includes(equipment.id), (checked) => {
        project.equipment_ids = checked ? [...project.equipment_ids, equipment.id]
          : project.equipment_ids.filter((id) => id !== equipment.id);
        trimProjectStaff(project);
        dirty(project);
      }, el("span", {}, equipmentText(equipment, "model"),
        ` / ${equipment.trial_group || "未分组"} / 证书有效期 ${equipment.valid_until || "未填"}`))));
    dialog.append(el("div", { class: "toolbar results-editor-actions" }, button("关闭详情", close, "primary")));
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); });
    document.body.append(dialog); dialog.showModal();
  }
  function environmentInputValue(value, key) {
    const unit = key === "temperature" ? /(?:℃|°\s*C)/gi : /[％%]\s*(?:RH)?/gi;
    return value.replace(unit, "").trim();
  }
  function renderProjects() {
    const task = state.task;
    const root = el("div", {}, notice("按项目归集本次所选原始记录单：开始取最早、结束取最晚，环境温湿度取全部记录的最小～最大范围。重新解析更新自动识别值，但不覆盖人工修改的实测结果。缺失或无法识别的记录逐行提示；输出时间、温湿度遵循结果素材的原有格式。"));
    root.append(el("label", { class: "check" }, el("input", {
      type: "checkbox", checked: task.include_atmospheric_pressure, "aria-label": "显示并输出大气压力",
      onchange: (event) => { task.include_atmospheric_pressure = event.target.checked; dirty(); render(); }
    }), "大气压力（显示并输出）"));
    if (!task.projects.length) {
      root.append(empty("还没有已选项目", "先到「素材与章节」选择本次结果素材。"), button("选择结果素材", () => navigate(2), "primary"), renderReportEnvironment());
      return root;
    }
    root.append(el("div", {class: "toolbar"}, button("解析本次资料到已选项目", () => run("正在匹配已选项目与原始记录…", async () => {
      await acceptParsedTask(await api("/api/tasks/parse?refresh_records=true", task));
    }, "匹配完成，来源已保留。请查看总览、项目来源、备注与缺项，无需额外勾选确认。"), "secondary"),
      button("查看全部空白实测结果", (event) => openBlankResultsReview(event.currentTarget), "secondary")));
    root.append(el("details", {}, el("summary", {}, "项目顺序与填写说明"), orderNotice(),
      notice("使用环境试验组或电磁兼容组设备的项目，小时00～23、分钟按5分钟选择；其他项目小时08～17、分钟00/15/30/45。原记录非选项时间及秒数保持原值，不取整；改时分才采用新选择。开始时间自动补齐空结束时间，单独编辑结束时间后不再跟随。人员按报告对象基础组及项目实际使用设备的试验小组确定，每个不同小组各选1名评价人员和1名复核人员；任一项目修改某组人员会同步全部使用该组的项目。评价与复核不能相同。")));
    if (task.projects.some((p) => p.rows.some((r) => !r.sample_id)))
      root.append(notice("请先到基本信息总览核对一次真实样机编号，对应关系用于所有项目。", "warning"),
        button("匹配样机编号", () => navigate(1), "secondary"));
    const missing = task.projects.filter((p) => p.notes.some((n) => n.startsWith("缺失原始记录：")));
    if (missing.length) root.append(notice("本次原始记录缺少：" + missing.map((p) => p.name).join("、"), "warning"));
    const body = el("tbody");
    const startPainters = new Map();
    const table = el("table", { class: "project-table" + (task.include_atmospheric_pressure ? " with-pressure" : ""), "aria-label": "试验项目信息" },
      el("thead", {}, el("tr", {}, ["项目 / 方案", "开始日期 / 时 / 分", "结束日期 / 时 / 分", "环境温度", "环境湿度",
        ...(task.include_atmospheric_pressure ? ["大气压力（kPa）"] : []),
        "设备", "评价人员", "复核人员", "实测结果", "详情"].map((text) => el("th", { scope: "col" }, text)))), body);
    task.projects.forEach((project, index) => {
      function change(key, value) { project[key] = value; dirty(project); }
      const startCell = el("td"), endCell = el("td");
      function endChanged(value) {
        const next = task.projects[index + 1];
        const followNext = next && (
          state.autoStartTimes.get(next.id) === next.start
          || (!next.start.trim() && !next.end.trim())
        );
        state.autoStartTimes.delete(project.id);
        state.autoEndTimes.delete(project.id);
        change("end", value);
        if (followNext) {
          next.start = value;
          state.autoStartTimes.set(next.id, value);
          dirty(next);
          startPainters.get(next.id)?.();
        }
      }
      function paintEnd() { endCell.replaceChildren(projectClock(project, "end", endChanged)); }
      paintEnd();
      function paintStart() { startCell.replaceChildren(projectClock(project, "start", (value) => {
        state.autoStartTimes.delete(project.id);
        const follow = !state.clockDrafts.has(`${project.id}:end`) &&
          (!project.end.trim() || state.autoEndTimes.get(project.id) === project.end);
        change("start", value);
        if (follow) {
          change("end", value); state.autoEndTimes.set(project.id, value); paintEnd();
        }
      })); }
      startPainters.set(project.id, paintStart);
      paintStart();
      const numbers = task.equipment.flatMap((e, i) => project.equipment_ids.includes(e.id) ? [i + 1] : []);
      const content = el("tr", { class: "project-content", "data-project-row": project.id },
        el("td", {}, el("strong", {}, `${index + 1}. ${project.name}`), el("small", {}, projectVariantLabel(project))),
        startCell, endCell,
        el("td", {}, field(`${project.name} 环境温度`, environmentInputValue(project.temperature, "temperature"), (v) => change("temperature", v), { placeholder: "值/范围" })),
        el("td", {}, field(`${project.name} 环境湿度`, environmentInputValue(project.humidity, "humidity"), (v) => change("humidity", v), { placeholder: "值/范围" })),
        ...(task.include_atmospheric_pressure ? [el("td", {},
          field(`${project.name} 大气压力`, (project.atmospheric_pressure || "").replace(/\s*kPa\s*$/i, ""), value => {
            project.pressure_manual_override = true;
            change("atmospheric_pressure", value);
          }, {placeholder:"值/范围", attrs:{"data-pressure-input":project.id}}))] : []),
        el("td", {}, el("button", { type: "button", class: "button secondary small", "data-project-action": "equipment",
          onclick: (event) => openProjectDetails(project, event.currentTarget) }, numbers.length ? `序号 ${numbers.join("、")}` : "选择设备")),
        el("td", {}, staffEditor(project, "evaluators", "评价人员")),
        el("td", {}, staffEditor(project, "reviewers", "复核人员")),
        el("td", {}, el("button", { type: "button", class: "button secondary small", "data-results-project": project.id,
          "aria-label": `${project.name} 预览／编辑实测结果`, onclick: (event) => openResultsEditor(project, event.currentTarget) }, "实测结果"),
          el("button", { type: "button", class: "button secondary small", "data-fill-empty-results": project.id,
            disabled: !outputRows(project).some((row) => !row.value?.trim()),
            onclick: (event) => fillEmptyProjectResults(project, event.currentTarget) }, "空白填符合要求"),
          el("small", {}, `${outputRows(project).filter((r) => r.value?.trim()).length}/${outputRows(project).length} 行已填`)),
        el("td", {}, el("button", { type: "button", class: "button secondary small", "data-project-action": "details",
          onclick: (event) => openProjectDetails(project, event.currentTarget) }, "详情"),
          el("small", { "data-staff-warning": "" }, staffIssues(project).length ? "人员待核对" : ""),
          el("small", { class: "project-issue" }, outputRows(project).some((r) => r.match_issue) ? "有识别提醒" : "")));
      syncStaffChoices(content, project);
      body.append(content);
    });
    root.append(el("div", { class: "table-wrap project-table-wrap", tabindex: 0, "aria-label": "试验项目表，可横向滚动" }, table));
    root.append(renderReportEnvironment());
    return root;
  }
  function recordConditions(project) {
    const summary = project.record_summary;
    if (!summary?.records?.length) return notice("尚无本项目的记录归集明细。请关联原始记录并点击「解析本次资料到已选项目」。", "warning");
    return el("div", {}, summary.issues.length ? notice(summary.issues, "warning") : null,
      el("details", {}, el("summary", {}, `查看本项目对应记录（${summary.records.length} 张记录单）及汇总来源`),
        el("p", { class: "muted" }, "记录单序号按每张样品信息表定位，不是 Word 自动重排后的物理页码。仅归集本次勾选原始记录中项目标识匹配的记录单。"),
        el("div", { class: "table-wrap result-table" }, el("table", { "aria-label": `${project.name} 原始记录归集` },
          el("thead", {}, el("tr", {}, ["来源 / 记录单", "型号", "样机编号", "开始", "结束", "环境温度", "环境湿度"].map((title) => el("th", { scope: "col" }, title)))),
          el("tbody", {}, summary.records.map((record) => el("tr", {},
            ["source", "model", "sample_id", "start", "end", "temperature", "humidity"].map((key) => el("td", {}, record[key] || "未填写")))))))));
  }
  const rowFields = [["sample_label", "型号 / 子试验 / 样机代号"], ["sample_id", "真实样机编号"], ["value", "实测结果"], ["requirement", "技术要求"], ["source", "来源 / 位置"], ["match_issue", "识别提醒"]];
  function openBlankResultsReview(trigger = document.activeElement) {
    if (state.busy || state.materialSelection || state.photoUpload ||
        state.photoUploadFailure?.task === state.task || state.resultsEditorDirty || state.equipmentEditorDirty) {
      feedback("请先完成或取消正在处理的素材、照片或未应用编辑，再核对空白实测结果。", true); return;
    }
    if (document.querySelector(".results-editor[open]")) return;
    const task = state.task;
    const entries = task.projects.flatMap(project => outputRows(project)
      .filter(row => !row.value?.trim()).map(row => ({project, row, signature: JSON.stringify(row)})));
    if (!entries.length) {
      feedback("所有已选参试样机的实测结果均非空白，未作修改。"); return;
    }
    const restoreView = captureProjectView(trigger);
    const dialog = el("dialog", {class: "results-editor blank-results-review",
      "aria-label": "全部空白实测结果核对"});
    const selected = new Set();
    const status = el("p", {role: "status"});
    const errors = el("div", {role: "alert"});
    const close = () => { dialog.close(); dialog.remove(); restoreView(); };
    const update = () => {
      status.textContent = `共 ${new Set(entries.map(entry => entry.project)).size} 个项目、${entries.length} 行空白实测结果；已逐行确认 ${selected.size} 行。未勾选的行保持空白。`;
      apply.disabled = !selected.size;
    };
    const apply = button("确认勾选行，填入符合要求", () => {
      errors.replaceChildren();
      if (!selected.size) { errors.append(notice("请先逐行核实并勾选需要填写的结果。", "error")); return; }
      if (state.busy || state.materialSelection || state.photoUpload ||
          state.photoUploadFailure?.task === state.task || state.task !== task ||
          [...selected].some(({project, row, signature}) => !task.projects.includes(project) ||
            !outputRows(project).includes(row) || JSON.stringify(row) !== signature || row.value?.trim())) {
        errors.append(notice("任务、参试样机或待填写行已变化，未作任何修改。请关闭清单后重新核对。", "error")); return;
      }
      if (!window.confirm(`确认已核实所勾选的 ${selected.size} 行有实际依据，且可用定性结论“符合要求”表述？\n需要数值的实测项不能用该文字代替。仅填写勾选的空白行，识别提醒仍保留，不自动确认项目或判定合格。`)) return;
      const changed = new Set();
      for (const {project, row} of selected) {
        row.value = "符合要求";
        row.manual_override = true;
        row.auto_value = "";
        if (!row.match_issue) row.matched_sample_id = row.sample_id || "";
        row.source = [row.source, "人工逐行核实补填：符合要求", "实测结果人工修改"].filter(Boolean).join("；");
        changed.add(project);
      }
      changed.forEach(project => dirty(project));
      dialog.close(); dialog.remove();
      render(); restoreView();
      feedback(`已按你的逐行确认填写 ${changed.size} 个项目、${selected.size} 行空白实测结果。未勾选行及已有结果不变，识别提醒保留；尚未保存任务，请核对后保存。`);
    }, "primary", true);
    const body = el("tbody", {}, entries.map(entry => {
      const {project, row} = entry;
      const input = el("input", {type: "checkbox", checked: false,
        "aria-label": `已核实 ${project.name} ${row.sample_label || "未标注样机代号"} ${row.sample_id || "未匹配编号"} 可填写定性结论`,
        onchange: event => { if (event.target.checked) selected.add(entry); else selected.delete(entry); update(); }});
      return el("tr", {}, el("td", {}, input),
        el("td", {}, project.name, el("small", {}, projectVariantLabel(project))),
        el("td", {}, row.sample_label || "未标注"), el("td", {}, row.sample_id || "未匹配"),
        el("td", {title: row.working_pressure?.source || ""}, row.resolved_requirement || row.requirement || "未提供，请核对原始记录"),
        el("td", {}, row.source || "未识别到来源"), el("td", {}, row.match_issue || "无"));
    }));
    dialog.append(el("h2", {}, "全部空白实测结果核对"),
      notice("空白或未识别不代表符合要求。请核对真实记录，只勾选有实际依据、可填写定性结论的行；数值型实测项请补录真实数值。所有行默认不勾选，不提供一键全选。这里只列出已选参试样机的空白结果，不修改未参试行或已有结果。", "warning"),
      status, errors, el("div", {class: "table-wrap result-table"}, el("table", {"aria-label": "空白实测结果清单"},
        el("thead", {}, el("tr", {}, ["已核实", "项目 / 方案", "样机代号 / 子试验", "真实样机编号", "技术要求", "原来源", "识别提醒"]
          .map(label => el("th", {scope: "col"}, label)))), body)),
      el("div", {class: "toolbar"}, apply, button("关闭，不填写", close)));
    dialog.addEventListener("cancel", event => { event.preventDefault(); close(); });
    update();
    document.body.append(dialog); dialog.showModal();
  }
  function fillEmptyProjectResults(project, trigger) {
    if (state.busy || document.querySelector(".results-editor[open]")) return;
    const draft = clone(project);
    const rows = outputRows(draft).filter((row) => !row.value?.trim());
    if (!rows.length) {
      feedback("当前项目没有空白试验结果，未作修改。");
      return;
    }
    if (!window.confirm(`请确认已人工核实「${project.name}」以下 ${rows.length} 行试验结果符合要求。\n`
      + "此操作仅将当前项目已勾选参试样机的空白结果填为“符合要求”，不覆盖已有结果，不代表系统自动判定合格；需要数值的项目不能用此文字代替实测数据。\n"
      + "确认后打开预览，点击“保存实测结果”才会写入任务。")) return;
    for (const row of rows) {
      row.value = "符合要求";
      row.source = [row.source, "人工批量补填：符合要求"].filter(Boolean).join("；");
    }
    openResultsEditor(project, trigger, draft);
  }
  function resultTable(project, editable, changed) {
    const table = el("table", { "aria-label": `${project.name} 实测结果` }, el("thead", {}, el("tr", {}, rowFields.map(([, label]) => el("th", { scope: "col" }, label)))));
    table.append(el("tbody", {}, project.rows.map((row, index) =>
      project.water_output_labels != null && !project.water_output_labels.includes(row.sample_label) ? null : el("tr", {},
      rowFields.map(([key, label]) => {
        if (!editable || key !== "value") return el("td", {title: key === "requirement" ? row.working_pressure?.source || "" : ""},
          key === "requirement" ? row.resolved_requirement || row.requirement || "未填写"
            : ["sample_id", "value"].includes(key) ? reportValue(row[key]) : row[key] || (key === "match_issue" ? "—" : "未填写"));
        return el("td", {}, el("textarea", {
          class: "cell-input", rows: 2, value: row[key] || "", "aria-label": `${project.name} 第 ${index + 1} 行 ${label}`,
          placeholder: "未填写时，报告保留素材原文",
          oninput: (event) => { row[key] = event.target.value; changed(); }
        }));
      })))));
    if (project.water_output_labels?.length === 0)
      return notice("本项目未勾选参试样机。Word保留标题和技术要求，不输出编号及实测数据；原人工结果未删除，重新勾选后可继续编辑。");
    return outputRows(project).length ? el("div", {},
      project.water_result_labels?.length ? notice("本次明细按大纲参试清单及原始记录识别生成；单型号或等级素材只提供试验条件和版式，不限制型号或样机数量。缺少记录仍保留待补录行，拆样样机不列入实测结果。") : null,
      el("div", { class: "table-wrap result-table" }, table)) : notice("没有可映射的样机结果行，请核对参试清单、原始记录及所选结果素材。", "warning");
  }
  function resultsCandidate(project, draft) {
    const candidate = clone(state.task);
    const savedProject = candidate.projects.find((item) => item.id === project.id);
    if (!savedProject) throw new Error("当前项目已变更，无法保存或预览。请关闭窗口后重新核对项目。");
    savedProject.rows = clone(draft.rows);
    savedProject.rows.forEach((row, index) => {
      if (row.value === project.rows[index].value && !row.match_issue?.includes("属于旧编号")) return;
      savedProject.confirmed = false;
      row.auto_value = ""; row.matched_sample_id = row.sample_id; row.match_issue = "";
      row.manual_override = true;
      if (!row.source.includes("实测结果人工修改"))
        row.source = [row.source, "实测结果人工修改"].filter(Boolean).join("；");
    });
    return candidate;
  }
  function openResultsEditor(project, trigger = document.activeElement, recovered = null) {
    if (state.busy || document.querySelector(".results-editor[open]")) return;
    const draft = clone(recovered || project);
    const restoreView = captureProjectView(trigger);
    const titleId = `results-editor-${++uid}`;
    const dialog = el("dialog", { class: "results-editor", "aria-labelledby": titleId });
    const errorBox = el("div", { role: "alert" });
    const status = el("p", { role: "status", class: "muted" }, "预览中，尚未修改。");
    const hasCurves = state.task.report_type === "燃气表" &&
      ["误差曲线", "示值误差曲线", "耐久性"].includes(project.name.replace(/-\d+\s*models?$/i, ""));
    const curveContent = el("div", { class: "gas-curve-preview-content", "aria-live": "polite" });
    const curvePanel = hasCurves ? el("section", { class: "gas-curve-preview" },
      el("h3", {}, "曲线图预览"),
      el("p", { class: "muted" }, "与报告使用同一绘图规则，按样机分别显示。初始图的误差限和人工平均误差取所选“准确度等级和最大允许误差”项目；耐久性图取本项目。当前预览不会保存任务，修改后仍须点击“保存实测结果”。"),
      button("刷新曲线预览", () => refreshCurves(), "small"), curveContent) : null;
    let curveTimer = null, curveVersion = 0, closed = false;
    async function refreshCurves() {
      if (!hasCurves || closed) return;
      clearTimeout(curveTimer);
      const version = ++curveVersion;
      curveContent.replaceChildren(notice("正在根据本次数据绘制曲线…"));
      try {
        const result = await api(`/api/gas-curve-previews/${encodeURIComponent(project.id)}`, resultsCandidate(project, draft));
        if (closed || version !== curveVersion) return;
        curveContent.replaceChildren(...result.charts.map((chart) => {
          if (chart.warning) return notice(chart.warning, "warning");
          const image = el("img", { src: chart.image_url, width: 1480, height: 670,
            alt: `${result.kind}图 · 样机${chart.sample_id} · ${chart.model}`,
            onerror: () => image.replaceWith(notice(`样机${chart.sample_id}的曲线图片加载失败，请刷新预览。`, "error")) });
          return el("figure", {}, image,
            el("figcaption", { class: "muted" }, `样机：${chart.sample_id}；型号：${chart.model}；误差限来源：${chart.limit_source}`));
        }));
      } catch (error) {
        if (!closed && version === curveVersion)
          curveContent.replaceChildren(notice(`曲线预览失败：${error.message} 当前编辑未丢失，可刷新重试。`, "error"));
      }
    }
    const finishClose = () => {
      closed = true; curveVersion++; clearTimeout(curveTimer);
      state.resultsEditorDirty = false;
      state.recoveryResultDraft = null;
      if (state.dirty) scheduleDraft(); else clearDraft();
      dialog.close(); dialog.remove();
      restoreView();
    };
    const close = () => {
      if (state.busy) return;
      if (state.resultsEditorDirty && !window.confirm("放弃尚未保存的实测结果修改？这些修改不会写入报告。")) return;
      finishClose();
    };
    const changed = () => {
      state.resultsEditorDirty = draft.rows.some((row, index) => row.value !== project.rows[index].value);
      status.textContent = state.resultsEditorDirty ? "有未保存修改；点击「保存实测结果」后才会用于报告。" : "与打开时一致，尚未修改。";
      state.recoveryResultDraft = state.resultsEditorDirty ? draft : null;
      scheduleDraft();
      if (hasCurves) {
        curveVersion++; clearTimeout(curveTimer);
        curveContent.replaceChildren(notice("实测结果已修改，曲线预览正在更新…"));
        curveTimer = setTimeout(refreshCurves, 450);
      }
    };
    const fields = el("fieldset", {}, resultTable(draft, true, changed), curvePanel,
      el("div", { class: "toolbar results-editor-actions" },
        el("button", { type: "submit", class: "button primary" }, "保存实测结果"),
        button("关闭", close)));
    const form = el("form", {}, fields);
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (state.busy) return;
      errorBox.replaceChildren(); fields.disabled = true;
      const ok = await run("正在保存实测结果到本机任务…", async () => {
        const candidate = resultsCandidate(project, draft);
        const saved = await api("/api/tasks", candidate);
        state.task = normalizeTask(saved);
        state.epoch += 1; state.validation = null; state.dirty = false;
        render();
      }, "实测结果及当前任务已保存，后续报告使用保存后的数据，无需额外勾选确认。");
      fields.disabled = false;
      if (ok) finishClose();
      else errorBox.replaceChildren(notice(`${$("feedback").textContent} 窗口中的修改仍保留，尚未写入报告。`, "error"));
    });
    dialog.append(el("h2", { id: titleId }, `实测结果预览与编辑 · ${project.name} · ${projectVariantLabel(project)}`),
      notice("只修改实测结果，不改技术要求、样机代号和素材行结构。保存会同时保存主页面当前任务内容；若提示结果属于旧编号，请先逐项核对，保存表示已核对结果属于当前编号。空白仍按既定规则保留素材原文。"),
      status, errorBox, form);
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); });
    document.body.append(dialog); dialog.showModal();
    dialog.querySelector("textarea")?.focus();
    if (recovered) changed();
    if (hasCurves) refreshCurves();
  }
  const defaultEquipmentStorageLocation = "张衡路1500号";
  const equipmentTrialGroups = ["液体组", "气体组", "环境试验组", "电磁兼容组"];
  function equipmentStorageLocation(equipment) {
    return equipment.storage_location ?? defaultEquipmentStorageLocation;
  }
  function equipmentDescription(equipment) {
    return el("dl", { class: "data-list" },
      ...[["型号", equipmentText(equipment, "model")], ["测量性能", equipmentText(equipment, "performance")],
        ["设备编号", equipmentText(equipment, "number")], ["证书有效期", equipment.valid_until],
        ["存放地点", equipmentStorageLocation(equipment)], ["试验小组", equipment.trial_group]]
        .flatMap(([key, value]) => [el("dt", {}, key), el("dd", {}, value || "未填写")]));
  }
  function equipmentRuns(item, key) {
    return item.formatting?.[key] || [{ text: item[key] || "", italic: false, script: "normal" }];
  }
  function equipmentContentMatches(left, right) {
    return left.valid_until === right.valid_until && left.trial_group === right.trial_group &&
      equipmentStorageLocation(left) === equipmentStorageLocation(right) &&
      ["name", "model", "performance", "number"].every((key) =>
        left[key] === right[key] && JSON.stringify(equipmentRuns(left, key).map((run) =>
          [run.text, !!run.italic, run.script || "normal"])) === JSON.stringify(equipmentRuns(right, key).map((run) =>
          [run.text, !!run.italic, run.script || "normal"])));
  }
  function equipmentText(item, key) {
    const runs = equipmentRuns(item, key);
    const formatted = (parts) => parts.map((run) => {
      const content = run.italic ? el("i", {}, run.text) : run.text;
      return run.script === "sup" ? el("sup", {}, content) : run.script === "sub" ? el("sub", {}, content) : content;
    });
    if (["model", "number", "valid_until"].includes(key)) return el("span", { class: "equipment-text" }, formatted(runs));
    const value = runs.map(run => run.text).join("");
    const pattern = /(^|[^A-Za-z0-9_.+\-−])([-−](?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?)/g;
    const parts = [];
    let start = 0, match;
    while ((match = pattern.exec(value))) {
      const from = match.index + match[1].length, end = from + match[2].length;
      parts.push(...formatted(sliceEquipmentRuns(runs, start, from)));
      parts.push(el("span", { class: "negative-number" }, formatted(sliceEquipmentRuns(runs, from, end))));
      start = end;
    }
    parts.push(...formatted(sliceEquipmentRuns(runs, start, value.length)));
    return el("span", { class: "equipment-text" }, parts);
  }
  function sliceEquipmentRuns(runs, start, end) {
    let offset = 0;
    return runs.flatMap((run) => {
      const from = Math.max(0, start - offset), to = Math.min(run.text.length, end - offset);
      offset += run.text.length;
      return to > from ? [{ ...run, text: run.text.slice(from, to) }] : [];
    });
  }
  function mergeEquipmentRuns(runs) {
    const merged = [];
    runs.filter((run) => run.text).forEach((run) => {
      const normalized = { text: run.text, italic: !!run.italic, script: run.script || "normal" };
      const previous = merged[merged.length - 1];
      if (previous && previous.italic === normalized.italic && previous.script === normalized.script) previous.text += normalized.text;
      else merged.push(normalized);
    });
    return merged;
  }
  function equipmentRichField(label, item, key, changed) {
    const id = `equipment-field-${++uid}`;
    const input = el("textarea", { id, value: item[key] || "", required: true,
      rows: key === "performance" ? 4 : 1, class: "equipment-rich-input" });
    const preview = el("div", { class: "equipment-rich-preview", "aria-label": `${label}格式预览` });
    const help = el("small", { id: `${id}-help`, role: "status" }, "选中文字，再点击上标、下标或斜体；下方实时预览。粘贴按纯文本处理。");
    input.setAttribute("aria-describedby", `${id}-help`);
    let before = null;
    const snapshot = () => ({ value: item[key] || "", runs: clone(equipmentRuns(item, key)),
      start: input.selectionStart, end: input.selectionEnd });
    const history = [snapshot()];
    let cursor = 0;
    const paint = () => preview.replaceChildren(equipmentText(item, key));
    const commit = (runs) => {
      item.formatting ||= {};
      item.formatting[key] = mergeEquipmentRuns(runs);
      item[key] = item.formatting[key].map((r) => r.text).join("");
      history.splice(cursor + 1);
      history.push(snapshot()); cursor += 1;
      changed(); paint();
    };
    const undo = (delta) => {
      const next = cursor + delta;
      if (next < 0 || next >= history.length) return;
      cursor = next;
      const saved = history[cursor];
      item[key] = saved.value; item.formatting ||= {}; item.formatting[key] = clone(saved.runs);
      input.value = saved.value; input.focus(); input.setSelectionRange(saved.start, saved.end);
      before = null; changed(); paint();
    };
    input.addEventListener("beforeinput", (event) => {
      before = { start: input.selectionStart, end: input.selectionEnd, type: event.inputType };
      history[cursor].start = input.selectionStart; history[cursor].end = input.selectionEnd;
      if (["historyUndo", "historyRedo"].includes(event.inputType) && event.cancelable) {
        event.preventDefault(); undo(event.inputType === "historyUndo" ? -1 : 1);
      }
    });
    input.addEventListener("input", () => {
      const old = item[key] || "", value = input.value, runs = equipmentRuns(item, key);
      let start = 0, oldEnd = old.length, newEnd = value.length;
      let limit = before?.start ?? Math.min(old.length, value.length);
      if (before?.type === "deleteContentBackward" && before.start === before.end)
        limit = Math.max(0, before.start - (old.length - value.length));
      while (start < limit && old[start] === value[start]) start += 1;
      const endLimit = before?.end ?? start;
      while (oldEnd > Math.max(start, endLimit) && newEnd > start && old[oldEnd - 1] === value[newEnd - 1]) {
        oldEnd -= 1; newEnd -= 1;
      }
      const inherited = sliceEquipmentRuns(runs, start, start + 1)[0] || sliceEquipmentRuns(runs, start - 1, start)[0] || {};
      commit([...sliceEquipmentRuns(runs, 0, start),
        { ...inherited, text: value.slice(start, newEnd) }, ...sliceEquipmentRuns(runs, oldEnd, old.length)]);
      before = null;
    });
    input.addEventListener("keydown", (event) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
      const keyName = event.key.toLowerCase();
      if (keyName === "z" || keyName === "y") {
        event.preventDefault(); undo(keyName === "y" || event.shiftKey ? 1 : -1);
      }
    });
    const apply = (style) => {
      const start = input.selectionStart, end = input.selectionEnd, runs = equipmentRuns(item, key);
      if (start === end) { help.textContent = "请先在输入框中选中需要设置格式的文字。"; input.focus(); return; }
      history[cursor].start = start; history[cursor].end = end;
      const selected = sliceEquipmentRuns(runs, start, end);
      const italic = !selected.every((r) => r.italic);
      const script = selected.every((r) => r.script === style) ? "normal" : style;
      const formatted = selected.map((r) => style === "italic" ? { ...r, italic } : style === "clear"
        ? { ...r, italic: false, script: "normal" } : { ...r, script });
      commit([...sliceEquipmentRuns(runs, 0, start), ...formatted, ...sliceEquipmentRuns(runs, end, item[key].length)]);
      input.focus(); input.setSelectionRange(start, end);
      help.textContent = "已更新所选文字格式；可再次点击取消，或使用清除格式、撤销。";
    };
    const toolbar = el("div", { class: "toolbar equipment-format-toolbar", role: "group", "aria-label": `${label}格式` },
      [["sup", "上标"], ["sub", "下标"], ["italic", "斜体"], ["clear", "清除格式"]].map(([style, title]) =>
        el("button", { type: "button", class: "button secondary small",
          onmousedown: (event) => event.preventDefault(), onclick: () => apply(style) }, title)),
      button("撤销", () => undo(-1), "quiet small"), button("重做", () => undo(1), "quiet small"));
    paint();
    return el("div", { class: "field equipment-rich-field" },
      el("label", { for: id, class: "field-label" }, label, el("span", { class: "required" }, "必填")),
      toolbar, input, help, preview);
  }
  function openEquipmentEditor(editing, savedCallback, recovered = false) {
    if (document.querySelector(".equipment-editor[open]")) return;
    const equipment = clone(editing || { id: "", name: "", model: "", performance: "", number: "", valid_until: "" });
    equipment.storage_location = equipmentStorageLocation(equipment);
    const trigger = document.activeElement;
    const titleId = `equipment-editor-${++uid}`;
    const dialog = el("dialog", { class: "equipment-editor", "aria-labelledby": titleId });
    const changed = () => {
      state.equipmentEditorDirty = true; state.recoveryEquipmentDraft = equipment; scheduleDraft();
    };
    const errorBox = el("div", { role: "alert" });
    const finishClose = () => {
      state.equipmentEditorDirty = false;
      state.recoveryEquipmentDraft = null;
      if (state.dirty) scheduleDraft(); else clearDraft();
      dialog.close(); dialog.remove();
      const target = trigger?.isConnected ? trigger : $("panel").querySelector("button");
      target?.focus({ preventScroll: true });
    };
    const close = () => {
      if (state.busy) return;
      if (state.equipmentEditorDirty && !window.confirm("放弃尚未保存的设备修改？")) return;
      finishClose();
    };
    const fields = el("fieldset", {},
      el("div", { class: "grid" },
        equipmentRichField("设备名称", equipment, "name", changed),
        equipmentRichField("型号", equipment, "model", changed),
        equipmentRichField("设备编号（不是证书编号）", equipment, "number", changed),
        field("证书有效期截止日", equipment.valid_until, (v) => { equipment.valid_until = v; changed(); }, { type: "date", required: true }),
        field("试验小组", equipment.trial_group || "", (v) => { equipment.trial_group = v; changed(); },
          { required: true, choices: [["", "请选择试验小组"], ...equipmentTrialGroups] }),
        field("存放地点", equipment.storage_location, (v) => { equipment.storage_location = v; changed(); }, { required: true })),
      equipmentRichField("测量区间 / 测量不确定度 / 准确度等级 / 最大允许误差", equipment, "performance", changed),
      el("div", { class: "toolbar" },
        el("button", { class: "button primary", type: "submit" }, editing ? "保存设备库修改" : "保存到设备库"),
        button("取消编辑", close)));
    const form = el("form", {}, fields);
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (state.busy || !form.reportValidity()) return;
      errorBox.replaceChildren();
      fields.disabled = true;
      const ok = await run("正在保存设备及文字格式到本机设备库…", async () => {
        const saved = await api("/api/equipment", equipment);
        const index = state.equipment.findIndex((e) => e.id === saved.id);
        if (index < 0) state.equipment.push(saved); else state.equipment[index] = saved;
        savedCallback();
      }, "设备库及格式已保存。本次报告快照未自动改变；请勾选设备或明确刷新本次快照。");
      fields.disabled = false;
      if (ok) finishClose();
      else errorBox.replaceChildren(notice($("feedback").textContent, "error"));
    });
    dialog.append(el("h2", { id: titleId }, editing ? `修改设备库 · ${editing.name}` : "新增仪器设备"),
      notice("格式支持名称、型号、编号及测量性能，保存后用于预览和 Word。设备库不保留历史版本，以最后保存的内容为准；修改设备库不自动改动已选快照。"), errorBox, form);
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); });
    document.body.append(dialog); dialog.showModal();
    dialog.querySelector("textarea").focus();
    if (recovered) {
      changed();
      errorBox.append(notice("已恢复设备编辑草稿，尚未写入设备库；请与当前库内内容核对后再保存。", "warning"));
    }
  }
  function removeEquipmentSnapshot(snapshot) {
    if (!window.confirm(`从本次报告移除「${snapshot.name}」，并取消项目对它的引用？设备库记录不删除。`)) return false;
    state.task.equipment = state.task.equipment.filter((e) => e.id !== snapshot.id);
    state.task.projects.forEach((p) => {
      if (p.equipment_ids.includes(snapshot.id)) {
        p.equipment_ids = p.equipment_ids.filter((id) => id !== snapshot.id);
        trimProjectStaff(p); p.confirmed = false;
      }
    });
    dirty();
    return true;
  }
  function renderEquipment() {
    const root = el("div");
    root.append(el("p", { class: "muted" }, "设备库保存时统一全角括号；测量性能中的 U、k 为斜体，Urel 的 rel 为正体下标，不改变 kPa、kHz 等单位或型号字母。已有任务设备快照不自动覆盖，新输出统一采用这些格式。"));
    const headers = ["选用", "设备名称", "型号", "测量性能", "设备编号", "证书有效期", "存放地点", "试验小组", "操作"];
    const body = el("tbody");
    const table = el("table", { class: "equipment-table equipment-library-table", "aria-label": "本机设备库" },
      el("thead", {}, el("tr", {}, headers.map((title) => el("th", { scope: "col" }, title)))), body);
    const scroll = el("div", { class: "table-wrap equipment-table-wrap", tabindex: "0", "aria-label": "设备库表格滚动区域" }, table);
    const summary = el("p", { class: "muted", role: "status" });
    const pages = el("div", { class: "toolbar equipment-pages", "aria-label": "设备库分页" });
    const snapshots = el("div");
    let draggedId = "";
    function reorderEquipment(identity, targetIndex) {
      const sourceIndex = state.task.equipment.findIndex((item) => item.id === identity);
      if (sourceIndex < 0 || targetIndex < 0 || targetIndex >= state.task.equipment.length || sourceIndex === targetIndex) return;
      const previous = new Map(state.task.equipment.map((item, index) => [item.id, index]));
      const [item] = state.task.equipment.splice(sourceIndex, 1);
      state.task.equipment.splice(targetIndex, 0, item);
      const moved = new Set(state.task.equipment.filter((entry, index) => previous.get(entry.id) !== index).map((entry) => entry.id));
      state.task.projects.filter((project) => project.equipment_ids.some((id) => moved.has(id)))
        .forEach((project) => { project.confirmed = false; });
      const position = snapshots.querySelector(".table-wrap")?.scrollTop || 0;
      dirty(); paint();
      snapshots.querySelector(".table-wrap").scrollTop = position;
      [...snapshots.querySelectorAll("[data-drag-equipment]")].find((handle) => handle.dataset.dragEquipment === identity)?.focus({ preventScroll: true });
      feedback(`「${item.name}」已移至第 ${targetIndex + 1} 位。两处设备表及项目设备序号随之更新，请保存任务；相关项目待重新核对。`);
    }
    function paint() {
      const query = state.equipmentQuery.trim().toLocaleLowerCase();
      const selected = new Map(state.task.equipment.map((s, i) => [s.id, { item: s, index: i }]));
      const filtered = state.equipment.filter((item) =>
        (!state.equipmentSelectedOnly || selected.has(item.id)) &&
        [item.name, item.model, item.performance, item.number, item.valid_until,
          equipmentStorageLocation(item), item.trial_group].some((v) => (v || "").toLocaleLowerCase().includes(query)));
      const totalPages = Math.max(1, Math.ceil(filtered.length / state.equipmentPageSize));
      state.equipmentPage = Math.min(state.equipmentPage, totalPages);
      const start = (state.equipmentPage - 1) * state.equipmentPageSize;
      const position = scroll.scrollTop;
      body.replaceChildren(...filtered.slice(start, start + state.equipmentPageSize).map((item) => {
        const snapshot = selected.get(item.id);
        const check = el("input", { type: "checkbox", checked: !!snapshot,
          "aria-label": `选用 ${item.name} ${item.number}`, "data-equipment-id": item.id,
          onchange: () => {
            if (check.checked) { state.task.equipment.push(clone(item)); dirty(); }
            else if (!removeEquipmentSnapshot(snapshot.item)) { check.checked = true; return; }
            paint();
            [...body.querySelectorAll("input[data-equipment-id]")].find((input) => input.dataset.equipmentId === item.id)?.focus({ preventScroll: true });
            feedback(check.checked ? "已勾选为本次报告快照；请逐项目选择实际使用设备。" : "已移除本次快照及项目引用，设备库记录仍保留。");
          } });
        const refresh = snapshot ? button("刷新快照", () => {
          if (!window.confirm(`将本次「${item.name}」快照更新为设备库当前内容？旧校验及导出状态将失效。`)) return;
          state.task.equipment[snapshot.index] = clone(item);
          state.task.projects.filter((p) => p.equipment_ids.includes(item.id)).forEach((p) => {
            trimProjectStaff(p); p.confirmed = false;
          });
          dirty(); paint(); feedback("已明确刷新本次快照，请重新核对关联项目。");
        }, "secondary small") : null;
        return el("tr", { class: snapshot ? "equipment-selected" : "" },
          el("td", {}, check, snapshot ? el("small", {}, `本次 ${snapshot.index + 1}`) : null),
          el("td", {}, equipmentText(item, "name")), el("td", {}, equipmentText(item, "model")),
          el("td", {}, equipmentText(item, "performance")), el("td", {}, equipmentText(item, "number")),
          el("td", {}, item.valid_until), el("td", {}, equipmentStorageLocation(item) || "未填写"),
          el("td", {}, item.trial_group || "未分组"),
          el("td", {}, el("div", { class: "toolbar" }, button("修改", () => openEquipmentEditor(item, paint), "secondary small"), refresh),
            snapshot && !equipmentContentMatches(snapshot.item, item) ? el("small", {}, "设备库内容与本次快照不同，尚未同步") : null));
      }));
      if (!filtered.length) body.append(el("tr", {}, el("td", { colspan: headers.length },
        state.equipment.length ? "没有匹配设备，请调整搜索条件。" : "设备库暂无记录，请点击新增设备。")));
      scroll.scrollTop = position;
      summary.textContent = `设备库共 ${state.equipment.length} 台 · 匹配 ${filtered.length} 台 · 本次已选 ${state.task.equipment.length} 台（跨页保留）`;
      pages.replaceChildren(
        button("上一页", () => { state.equipmentPage -= 1; scroll.scrollTop = 0; paint(); }, "secondary small", state.equipmentPage === 1),
        el("span", {}, `第 ${state.equipmentPage} / ${totalPages} 页`),
        button("下一页", () => { state.equipmentPage += 1; scroll.scrollTop = 0; paint(); }, "secondary small", state.equipmentPage === totalPages));
      const snapshotTable = el("table", { class: "equipment-table", "aria-label": "本次设备快照" },
        el("thead", {}, el("tr", {}, ["序号", ...headers.slice(1, -1), "操作", "排序"].map((title) => el("th", { scope: "col" }, title)))),
        el("tbody", {}, state.task.equipment.map((item, index) => el("tr", {
          "data-snapshot-id": item.id,
          ondragover: (event) => {
            if (!draggedId || draggedId === item.id) return;
            event.preventDefault(); event.dataTransfer.dropEffect = "move";
            event.currentTarget.classList.add("equipment-drop-target");
          },
          ondragleave: (event) => {
            if (!event.currentTarget.contains(event.relatedTarget)) event.currentTarget.classList.remove("equipment-drop-target");
          },
          ondrop: (event) => {
            event.preventDefault();
            event.currentTarget.classList.remove("equipment-drop-target");
            if (!draggedId) return;
            const identity = draggedId; draggedId = "";
            reorderEquipment(identity, index);
          }
        },
          el("td", {}, index + 1),
          ...["name", "model", "performance", "number"].map((key) => el("td", {}, equipmentText(item, key))),
          el("td", {}, item.valid_until), el("td", {}, equipmentStorageLocation(item) || "未填写"),
          el("td", {}, item.trial_group || "未分组"),
          el("td", {}, button("移除", () => { if (removeEquipmentSnapshot(item)) paint(); }, "danger small")),
          el("td", { class: "equipment-order-cell" },
            el("button", { type: "button", draggable: "true", class: "button secondary small equipment-drag-handle",
              "data-drag-equipment": item.id, "aria-label": `拖动排序 ${item.name} ${item.number}`,
              title: "拖到目标行调整顺序；也可用 Alt＋↑／↓",
              ondragstart: (event) => {
                draggedId = item.id; event.dataTransfer.effectAllowed = "move";
                event.dataTransfer.setData("text/plain", item.id);
              },
              ondragend: () => {
                draggedId = "";
                snapshots.querySelectorAll(".equipment-drop-target").forEach((row) => row.classList.remove("equipment-drop-target"));
              },
              onkeydown: (event) => {
                if (event.altKey && ["ArrowUp", "ArrowDown"].includes(event.key)) {
                  event.preventDefault(); reorderEquipment(item.id, index + (event.key === "ArrowUp" ? -1 : 1));
                }
              } }, "↕ 拖动"),
            el("div", { class: "toolbar" },
              button("上移", () => reorderEquipment(item.id, index - 1), "quiet small", index === 0),
              button("下移", () => reorderEquipment(item.id, index + 1), "quiet small", index === state.task.equipment.length - 1)))))));
      snapshots.replaceChildren(el("div", { class: "table-wrap equipment-table-wrap" }, snapshotTable));
      snapshotSummary.textContent = `本次报告设备快照（${state.task.equipment.length}台）· 展开核对顺序与内容`;
    }
    const search = field("搜索设备（名称 / 型号 / 性能 / 编号 / 存放地点 / 试验小组）", state.equipmentQuery, (value) => {
      state.equipmentQuery = value; state.equipmentPage = 1; scroll.scrollTop = 0; paint();
    }, { type: "search" });
    const snapshotSummary = el("summary");
    const selectedDetails = el("details", { class: "equipment-snapshots" }, snapshotSummary,
      notice("拖动最右侧「↕ 拖动」到目标行可调整设备顺序，也可点击上移／下移。当前顺序用于两处设备表和项目设备序号，调整后请保存任务。库内修改仍须明确刷新快照；取消勾选会清除项目引用。"),
      snapshots);
    root.append(card("本机持久设备库",
      el("div", { class: "toolbar equipment-library-tools" }, button("＋ 新增设备", () => openEquipmentEditor(null, paint), "primary"),
        checkbox("仅看本次已选", state.equipmentSelectedOnly, (checked) => {
          state.equipmentSelectedOnly = checked; state.equipmentPage = 1; paint();
        }),
        field("每页条数", String(state.equipmentPageSize), (value) => {
          state.equipmentPageSize = Number(value); state.equipmentPage = 1; scroll.scrollTop = 0; paint();
        }, { choices: ["10", "20", "50"] })),
      search, summary, scroll, pages,
      el("p", { class: "muted" }, "勾选即可加入本次报告；搜索、翻页不丢失选择。证书有效性须结合实际试验日期核对。")),
      selectedDetails);
    paint();
    return root;
  }
  const photoImageCache = new Map();
  const photoImageRequests = new Map();
  async function loadCloudPhoto(image, identity, edit = {}, thumbnail = false) {
    let url = photoImageUrl(identity, edit);
    if (thumbnail) url += `${url.includes("?") ? "&" : "?"}preview=thumbnail`;
    if (photoImageCache.has(url)) {
      const source = photoImageCache.get(url);
      photoImageCache.delete(url); photoImageCache.set(url, source);
      image.src = source;
      return;
    }
    let pending = photoImageRequests.get(url);
    if (!pending) {
      pending = (async () => {
        const data = await api(`${url}${url.includes("?") ? "&" : "?"}encoding=base64`);
        if (!["image/png", "image/jpeg"].includes(data.media_type) || typeof data.data_base64 !== "string" ||
            !data.data_base64.length || data.data_base64.length % 4 ||
            !/^[A-Za-z0-9+/]+={0,2}$/.test(data.data_base64))
          throw new Error("后台返回的照片数据无效，请重试。");
        const source = `data:${data.media_type};base64,${data.data_base64}`;
        photoImageCache.set(url, source);
        while (photoImageCache.size > 8 ||
            [...photoImageCache.values()].reduce((sum, value) => sum + value.length, 0) > 16 * 1024 * 1024)
          photoImageCache.delete(photoImageCache.keys().next().value);
        return source;
      })();
      photoImageRequests.set(url, pending);
    }
    try {
      image.src = await pending;
    } finally {
      if (photoImageRequests.get(url) === pending) photoImageRequests.delete(url);
    }
  }
  function photoImageUrl(identity, edit = {}) {
    const query = new URLSearchParams();
    if (edit.rotation) query.set("rotation", edit.rotation);
    if (edit.crop) query.set("crop", edit.crop.join(","));
    const suffix = query.toString();
    return `/api/photos/${encodeURIComponent(identity)}/image${suffix ? "?" + suffix : ""}`;
  }
  function rotatePhotoCrop(crop, width, height, clockwise) {
    if (!crop) return null;
    const [left, top, right, bottom] = crop;
    return clockwise ? [height - bottom, left, height - top, right]
      : [top, width - right, bottom, width - left];
  }
  function openPhotoEditor(slot, title, defaultWidth) {
    if (!state.catalog.photo_editing || !Number.isFinite(state.catalog.photo_max_width_cm)) {
      feedback("当前后台尚不支持照片编辑，请先保存任务，重启服务后刷新页面。", true); return;
    }
    const identity = state.task.photos[slot];
    const gas = ["燃气表", "流量计（气体）"].includes(state.task.report_type);
    if (gas) defaultWidth = slot.startsWith("component:") ? 8 : slot === "internal" ? 10 : 12;
    const original = state.task.photo_edits?.[slot] || {};
    let rotation = original.rotation || 0, crop = original.crop ? [...original.crop] : null;
    let loaded = false, pixelWidth = 0, pixelHeight = 0, dragging = null;
    const previousFocus = document.activeElement;
    const pageTop = window.scrollY, pageLeft = window.scrollX;
    const dialog = el("dialog", { class: "photo-editor", "aria-label": `编辑${title}` });
    const status = el("p", { role: "status", class: "muted" }, "正在载入原图…");
    const canvas = el("canvas", { "aria-label": "旋转后的完整原图" });
    const selection = el("div", { class: "photo-crop-selection" });
    for (const corner of ["nw", "ne", "sw", "se"])
      selection.append(el("span", { class: `photo-crop-handle ${corner}`, "data-corner": corner }));
    const stage = el("div", { class: "photo-crop-stage" }, canvas, selection);
    const image = new Image();
    const widthInput = el("input", { type: "number", min: "0.1", max: String(state.catalog.photo_max_width_cm),
      step: "any", value: original.width_cm ?? defaultWidth, "aria-label": "报告输出宽度（cm）" });
    const cropInputs = ["左", "上", "右", "下"].map((label) =>
      el("input", { type: "number", min: "0", step: "1", "aria-label": `裁剪${label}边界（像素）` }));
    const sizeText = el("p", { class: "muted" });
    const controls = el("fieldset", { disabled: true });
    function close() {
      dialog.close(); dialog.remove();
      if (previousFocus?.isConnected) previousFocus.focus({ preventScroll: true });
      window.scrollTo({ top: pageTop, left: pageLeft });
    }
    function validBox(box) {
      return box.every(Number.isInteger) && box[0] >= 0 && box[1] >= 0 &&
        box[2] > box[0] && box[3] > box[1] && box[2] <= pixelWidth && box[3] <= pixelHeight;
    }
    function paintSelection() {
      const box = crop || [0, 0, pixelWidth, pixelHeight];
      selection.style.left = `${100 * box[0] / pixelWidth}%`;
      selection.style.top = `${100 * box[1] / pixelHeight}%`;
      selection.style.width = `${100 * (box[2] - box[0]) / pixelWidth}%`;
      selection.style.height = `${100 * (box[3] - box[1]) / pixelHeight}%`;
      cropInputs.forEach((input, index) => { input.value = box[index]; });
      paintSize();
    }
    function paintSize() {
      const box = crop || [0, 0, pixelWidth, pixelHeight];
      const width = Number(widthInput.value);
      sizeText.textContent = `${box[2] - box[0]} × ${box[3] - box[1]} 像素 · 输出 ${
        Number.isFinite(width) && width > 0 ? `${width} × ${(width * (box[3] - box[1]) / (box[2] - box[0])).toFixed(2)}cm` : "宽度待填写"
      }（高度自动等比例计算）`;
    }
    function draw() {
      pixelWidth = rotation % 180 ? image.naturalHeight : image.naturalWidth;
      pixelHeight = rotation % 180 ? image.naturalWidth : image.naturalHeight;
      const scale = Math.min(1, 900 / pixelWidth, 500 / pixelHeight);
      canvas.width = Math.max(1, Math.round(pixelWidth * scale));
      canvas.height = Math.max(1, Math.round(pixelHeight * scale));
      stage.style.width = `${canvas.width}px`;
      const context = canvas.getContext("2d");
      context.translate(canvas.width / 2, canvas.height / 2);
      context.scale(canvas.width / pixelWidth, canvas.height / pixelHeight);
      context.rotate(rotation * Math.PI / 180);
      context.drawImage(image, -image.naturalWidth / 2, -image.naturalHeight / 2);
      paintSelection();
    }
    function rotate(clockwise) {
      crop = rotatePhotoCrop(crop, pixelWidth, pixelHeight, clockwise);
      rotation = (rotation + (clockwise ? 90 : 270)) % 360;
      draw();
    }
    function point(event) {
      const rect = stage.getBoundingClientRect();
      return [
        Math.max(0, Math.min(pixelWidth, Math.round((event.clientX - rect.left) / rect.width * pixelWidth))),
        Math.max(0, Math.min(pixelHeight, Math.round((event.clientY - rect.top) / rect.height * pixelHeight))),
      ];
    }
    stage.addEventListener("pointerdown", (event) => {
      if (!loaded || event.button !== 0) return;
      event.preventDefault();
      const mode = event.target.dataset.corner || (event.target === selection ? "move" : "new");
      dragging = { pointer: event.pointerId, start: point(event), box: crop || [0, 0, pixelWidth, pixelHeight], mode };
      stage.setPointerCapture(event.pointerId);
    });
    stage.addEventListener("pointermove", (event) => {
      if (!dragging || dragging.pointer !== event.pointerId) return;
      const [x, y] = point(event), { start, box, mode } = dragging;
      let next;
      if (mode === "move") {
        const dx = Math.max(-box[0], Math.min(pixelWidth - box[2], x - start[0]));
        const dy = Math.max(-box[1], Math.min(pixelHeight - box[3], y - start[1]));
        next = [box[0] + dx, box[1] + dy, box[2] + dx, box[3] + dy];
      } else if (mode === "new") {
        next = [Math.min(start[0], x), Math.min(start[1], y), Math.max(start[0], x), Math.max(start[1], y)];
      } else {
        next = [...box];
        next[mode.includes("w") ? 0 : 2] = x;
        next[mode.includes("n") ? 1 : 3] = y;
      }
      if (validBox(next)) { crop = next; paintSelection(); }
    });
    for (const name of ["pointerup", "pointercancel", "lostpointercapture"])
      stage.addEventListener(name, () => { dragging = null; });
    cropInputs.forEach((input) => input.addEventListener("input", () => {
      const box = cropInputs.map((control) => control.value === "" ? NaN : Number(control.value));
      if (!validBox(box)) { status.textContent = "裁剪边界须为照片内的整数像素，右大于左、下大于上。"; return; }
      crop = box; status.textContent = "裁剪已调整，尚未应用。"; paintSelection();
    }));
    widthInput.addEventListener("input", () => { if (loaded) paintSize(); });
    const apply = button("应用到本次任务", () => {
      const box = cropInputs.map((control) => control.value === "" ? NaN : Number(control.value));
      if (!loaded || !validBox(box) || !widthInput.value || !widthInput.checkValidity()) {
        status.textContent = `请填写有效的裁剪区域和输出宽度（0.1～${state.catalog.photo_max_width_cm.toFixed(2)}cm）。`; return;
      }
      if (state.task.photos[slot] !== identity) {
        status.textContent = "照片关联已改变，请取消并重新打开编辑。"; return;
      }
      state.task.photo_edits ||= {};
      state.task.photo_edits[slot] = { rotation, crop, width_cm: Number(widthInput.value) };
      dirty(); close(); render();
      window.scrollTo({ top: pageTop, left: pageLeft });
      feedback("照片调整已应用于本次任务，请保存任务。原图及其他任务不变。");
    });
    controls.append(el("div", { class: "toolbar" },
      button("左旋转90°", () => rotate(false), "secondary"),
      button("右旋转90°", () => rotate(true), "secondary"),
      button("取消裁剪", () => { crop = null; draw(); }, "secondary"),
      button("恢复原图及默认宽度", () => { rotation = 0; crop = null; widthInput.value = defaultWidth; draw(); }, "secondary")),
      el("div", { class: "photo-crop-fields" }, cropInputs.map((input, index) =>
        el("label", {}, `${["左", "上", "右", "下"][index]}边界（像素）`, input))),
      el("label", {}, "报告输出宽度（cm）", widthInput), sizeText, apply);
    dialog.append(el("h2", {}, `编辑照片 · ${title}`),
      el("p", { class: "muted" }, "拖动四角裁剪，框内拖动移动，框外拖动重新选区；也可输入像素边界。仅保留框内区域，输出不拉伸。"),
      status, stage, controls, button("取消", close, "secondary"));
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); });
    image.addEventListener("load", () => {
      loaded = true; controls.disabled = false;
      draw(); status.textContent = "原图已载入。修改后点击应用，再保存任务。";
    }, { once: true });
    image.addEventListener("error", () => {
      status.textContent = "原图加载失败，请检查网络连接并重试；本次编辑未应用。";
    }, { once: true });
    document.body.append(dialog); dialog.showModal();
    if (document.querySelector('meta[name="workbench-photo-transport"]')?.content === "json")
      loadCloudPhoto(image, identity).catch(error => {
        status.textContent = `原图加载失败：${error.message}；本次编辑未应用。`;
      });
    else image.src = photoImageUrl(identity);
  }
  function openPhotoViewer(title, source, trigger = document.activeElement, cloudPhoto = null) {
    const pageTop = window.scrollY, pageLeft = window.scrollX;
    let loaded = false, scale = 1, fitted = true, drag = null;
    const dialog = el("dialog", { class: "photo-viewer", "aria-label": `查看照片 · ${title}` });
    const status = el("span", { role: "status", class: "muted" }, "正在载入照片…");
    const image = el("img", { class: "photo-viewer-image", alt: title, draggable: "false" });
    const surface = el("div", { class: "photo-viewer-surface" }, image);
    const viewport = el("div", { class: "photo-viewer-viewport", tabindex: "0",
      "aria-label": "照片查看区，可拖动或使用滚动条查看放大后的细节" }, surface);
    const fitScale = () => Math.min(1, Math.max(1, viewport.clientWidth - 32) / image.naturalWidth,
      Math.max(1, viewport.clientHeight - 32) / image.naturalHeight);
    function zoom(next, x = viewport.clientWidth / 2, y = viewport.clientHeight / 2, reset = false) {
      if (!loaded) return;
      const oldX = (viewport.scrollLeft + x - (surface.clientWidth - image.naturalWidth * scale) / 2) / scale;
      const oldY = (viewport.scrollTop + y - (surface.clientHeight - image.naturalHeight * scale) / 2) / scale;
      scale = Math.max(Math.min(0.1, fitScale()), Math.min(8, next));
      const width = image.naturalWidth * scale, height = image.naturalHeight * scale;
      const surfaceWidth = Math.max(viewport.clientWidth, width + 32);
      const surfaceHeight = Math.max(viewport.clientHeight, height + 32);
      image.style.width = `${width}px`; image.style.height = `${height}px`;
      surface.style.width = `${surfaceWidth}px`; surface.style.height = `${surfaceHeight}px`;
      viewport.scrollLeft = reset ? 0 : (surfaceWidth - width) / 2 + oldX * scale - x;
      viewport.scrollTop = reset ? 0 : (surfaceHeight - height) / 2 + oldY * scale - y;
      status.textContent = `${Number((scale * 100).toFixed(1))}% · ${image.naturalWidth} × ${image.naturalHeight} 像素`;
      smaller.disabled = scale <= Math.min(0.1, fitScale());
      larger.disabled = scale >= 8;
    }
    function step(factor) { fitted = false; zoom(scale * factor); }
    function fit() { fitted = true; zoom(fitScale(), undefined, undefined, true); }
    function close() {
      window.removeEventListener("resize", resize);
      dialog.close(); dialog.remove();
      if (trigger?.isConnected) trigger.focus({ preventScroll: true });
      window.scrollTo({ top: pageTop, left: pageLeft });
    }
    function resize() { if (loaded) { if (fitted) fit(); else zoom(scale); } }
    const smaller = button("缩小 −", () => step(1 / 1.25), "secondary small", true);
    const larger = button("放大 ＋", () => step(1.25), "secondary small", true);
    const actual = button("100%", () => { fitted = false; zoom(1); }, "secondary small", true);
    const fitButton = button("适应窗口", fit, "secondary small", true);
    viewport.addEventListener("wheel", (event) => {
      if (!loaded || !event.deltaY) return;
      event.preventDefault();
      const rect = viewport.getBoundingClientRect();
      fitted = false;
      zoom(scale * (event.deltaY < 0 ? 1.1 : 1 / 1.1), event.clientX - rect.left, event.clientY - rect.top);
    }, { passive: false });
    viewport.addEventListener("pointerdown", (event) => {
      if (!loaded || event.button !== 0 || (event.target !== image && event.target !== surface)) return;
      event.preventDefault();
      viewport.focus({ preventScroll: true });
      drag = { id: event.pointerId, x: event.clientX, y: event.clientY,
        left: viewport.scrollLeft, top: viewport.scrollTop };
      viewport.setPointerCapture(event.pointerId);
      viewport.classList.add("panning");
    });
    viewport.addEventListener("pointermove", (event) => {
      if (!drag || drag.id !== event.pointerId) return;
      viewport.scrollLeft = drag.left + drag.x - event.clientX;
      viewport.scrollTop = drag.top + drag.y - event.clientY;
    });
    for (const name of ["pointerup", "pointercancel", "lostpointercapture"])
      viewport.addEventListener(name, () => { drag = null; viewport.classList.remove("panning"); });
    dialog.addEventListener("keydown", (event) => {
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      if (["+", "=", "-"].includes(event.key)) {
        event.preventDefault(); step(event.key === "-" ? 1 / 1.25 : 1.25);
      }
    });
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); });
    image.addEventListener("load", () => {
      if (!dialog.isConnected) return;
      loaded = true;
      actual.disabled = false; fitButton.disabled = false;
      fit();
    }, { once: true });
    image.addEventListener("error", () => {
      if (dialog.isConnected) status.textContent = "照片加载失败，请检查网络连接并重试。";
    }, { once: true });
    dialog.append(el("div", { class: "photo-viewer-heading" }, el("h2", {}, title),
      button("关闭", close, "secondary small")),
      el("div", { class: "toolbar" }, smaller, larger, actual, fitButton, status),
      el("p", { class: "muted" }, "滚轮或＋／−缩放，按住照片拖动查看细节，Esc关闭。仅查看，不改变照片编辑及报告输出尺寸。"),
      viewport);
    document.body.append(dialog); dialog.showModal();
    window.addEventListener("resize", resize);
    if (cloudPhoto)
      loadCloudPhoto(image, cloudPhoto.identity, cloudPhoto.edit).catch(error => {
        if (dialog.isConnected) status.textContent = `照片加载失败：${error.message}。关闭后可重试。`;
      });
    else image.src = source;
  }
  function photoUploadTarget(task, slot) {
    return JSON.stringify([task.report_type, task.outline_id, task.display_types,
      task.photo_components, task.photos[slot], task.photo_edits?.[slot]]);
  }
  function cancelPhotoUpload() {
    state.photoUpload = null;
    state.photoUploadFailure = null;
  }
  function refreshPhotoUpload() {
    if (state.step === 5) render();
  }
  function renderPhotoUploadStatus() {
    const operation = state.photoUpload || state.photoUploadFailure;
    const pending = Boolean(state.photoUpload);
    return el("section", {class: "card", "data-photo-upload": "", role: "status"},
      notice(pending
        ? `正在上传「${operation.title}」：已完成 ${operation.completed}/${operation.files.length} 张，当前 ${operation.files[operation.completed].name}。可切换页面并编辑其他内容，保存和生成需等待完成。`
        : operation.error, pending ? "warning" : "error"),
      ...(!pending ? [button("重试未完成照片", () => {
        if (state.task !== operation.task || photoUploadTarget(state.task, operation.slot) !== operation.target) {
          feedback("任务或照片位置已改变，请结束本次上传后，在正确位置重新选择照片。", true); return;
        }
        return uploadPhotos(operation.slot, operation.title, operation.files.slice(operation.completed), operation.multiple, operation.mode);
      }, "secondary small")] : []),
      button(pending ? "取消后续上传及当前关联" : "结束本次上传，保留已成功照片", () => {
        if (pending && !window.confirm("取消后不再关联正在上传的照片，也不再上传后续照片；已成功关联的照片保留。已发出的请求可能仍会在后台保存照片副本。是否继续？")) return;
        cancelPhotoUpload();
        refreshPhotoUpload();
        feedback("本次上传已结束，已成功关联的照片保留，请核对并保存任务。");
      }, "secondary small"));
  }
  async function uploadPhotos(slot, title, files, multiple = false, mode = state.cloudPhotoMode || "standard") {
    if (state.busy || state.photoUpload || state.materialSelection) {
      feedback("当前有后台修改尚未完成，请等待完成或取消后再上传照片。", true); return;
    }
    if (!files.length || (!multiple && files.length !== 1) ||
        files.some(file => !/\.(jpe?g|png)$/i.test(file.name) || file.size > 32 * 1024 * 1024)) {
      feedback("请选择 JPG、JPEG、PNG 照片，单张不超过32 MB。", true); return;
    }
    const task = state.task;
    const target = photoUploadTarget(task, slot);
    const operation = {slot, title, files: [...files], multiple, mode, completed: 0, task, target};
    const optimized = document.querySelector('meta[name="workbench-photo-optimization"]')?.content === "enabled";
    state.photoUploadFailure = null;
    state.photoUpload = operation;
    refreshPhotoUpload();
    feedback("附件2照片正在后台上传，可继续浏览和编辑其他内容；请勿刷新或关闭页面。");
    try {
      for (const file of operation.files) {
        if (state.photoUpload !== operation) return;
        const form = new FormData();
        let prepared;
        if (optimized) {
          if (!window.workbenchPhotos) throw new Error("云端照片优化组件未加载，请保存任务后刷新页面重试。");
          prepared = await window.workbenchPhotos.prepare(file, mode);
          if (state.photoUpload !== operation) return;
          if (state.task !== task || photoUploadTarget(task, slot) !== target)
            throw new Error("任务或照片位置已改变，压缩后的照片未上传。");
          form.append("photo_mode", mode);
        }
        form.append("file", prepared ? prepared.file : file);
        const uploaded = await api("/api/photos", form);
        if (!uploaded || typeof uploaded.id !== "string" || !uploaded.id)
          throw new Error("后台未返回有效照片编号，请核对照片库后重试。");
        if (!state.photoFiles.some(photo => photo.id === uploaded.id)) state.photoFiles.push(uploaded);
        if (state.photoUpload !== operation) return;
        if (state.task !== task || photoUploadTarget(task, slot) !== target)
          throw new Error("任务或照片位置已改变，上传副本未关联到当前任务。请在正确任务和位置重新选择照片。");
        const key = multiple ? `${slot}:${crypto.randomUUID().replaceAll("-", "")}` : slot;
        if (multiple) task.photo_order = [...new Set([
          ...(task.photo_order || []), ...Object.keys(task.photos), key])];
        task.photos[key] = uploaded.id;
        if (task.photo_edits) delete task.photo_edits[key];
        operation.completed++;
        dirty();
        if (operation.completed < operation.files.length) refreshPhotoUpload();
      }
      feedback(`「${title}」已上传并关联 ${operation.completed} 张照片。${optimized ? "已按云端优化模式保存，电脑原文件未改变；请放大核对小字清晰度。" : ""}其他编辑已保留，请核对并保存任务。`);
    } catch (error) {
      if (state.photoUpload === operation) {
        const message = `「${title}」上传未完成：${error.message || String(error)}。已成功关联 ${operation.completed}/${operation.files.length} 张；已成功照片和其他编辑保留。`;
        if (state.task === task && photoUploadTarget(task, slot) === target)
          state.photoUploadFailure = {...operation, error: message};
        feedback(message, true);
      }
    } finally {
      if (state.photoUpload === operation) {
        state.photoUpload = null;
        refreshPhotoUpload();
      }
    }
  }
  function photoCard(slot, title, width, editable) {
    editable = editable && !state.photoUpload && state.photoUploadFailure?.task !== state.task;
    const identity = state.task.photos[slot];
    const edit = state.task.photo_edits?.[slot] || {};
    width = edit.width_cm ?? width;
    const info = state.photoFiles.find((photo) => photo.id === identity);
    const content = el("section", { class: "photo-card", "data-photo-slot": slot },
      el("h3", {}, title), el("small", {}, `报告宽度 ${width}cm · 保持裁剪后长宽比 · 居中`));
    if (identity) {
      const optimized = window.workbenchPhotos?.enabled();
      const image = el("img", { alt: title, loading: "lazy" });
      image.addEventListener("error", () => content.append(notice("照片预览失败，请检查网络连接并重试；生成时会校验照片副本。", "error")), { once: true });
      const thumbnail = el("button", { type: "button", class: "photo-thumbnail",
        "aria-label": `查看并放大${title}`, title: "点击查看大图，可放大",
        disabled: true, onclick: () => openPhotoViewer(title, image.src, thumbnail,
          optimized ? {identity, edit: clone(edit)} : null) }, image);
      image.addEventListener("load", () => { thumbnail.disabled = false; }, { once: true });
      if (optimized)
        window.workbenchPhotos.lazy(image, () => loadCloudPhoto(image, identity, edit, true), error => {
          content.append(notice(`缩略图加载失败：${error.message}。`, "error"),
            button("重试照片", () => render(), "secondary small"));
        });
      else if (document.querySelector('meta[name="workbench-photo-transport"]')?.content === "json")
        loadCloudPhoto(image, identity, edit).catch(error => {
          content.append(notice(`照片预览失败：${error.message}。当前照片和编辑仍保留，可重试。`, "error"));
        });
      else image.src = photoImageUrl(identity, edit);
      content.append(thumbnail, el("small", { class: "muted" }, "点击照片查看大图／放大"),
        el("p", { class: "muted" }, info
        ? `${info.name} · ${info.optimization ? "云端副本" : "原图"} ${info.width} × ${info.height} 像素${info.bytes ? ` · ${Math.ceil(info.bytes / 1024)} KB` : ""}`
        : "已关联照片（文件信息可重新载入）"));
    } else content.append(el("div", { class: "photo-empty" }, "尚未上传 · 草稿保留空白位置"));
    if (editable) {
      if (identity) content.append(button("编辑照片（旋转／裁剪／宽度）", () => {
        const defaultWidth = state.catalog.photo_slots.find((item) => item.id === slot.split(":")[0])?.width_cm || 8;
        openPhotoEditor(slot, title, defaultWidth);
      }, "secondary small"));
      const input = el("input", { type: "file", accept: ".jpg,.jpeg,.png", "aria-label": `上传${title}` });
      content.append(input, el("div", { class: "toolbar" },
        button(identity ? "上传并替换" : "上传照片", () => {
          const file = input.files[0];
          if (!file) { feedback(`请先选择「${title}」的照片。`, true); return; }
          if (!/\.(jpe?g|png)$/i.test(file.name) || file.size > 32 * 1024 * 1024) {
            feedback("照片支持 JPG、JPEG、PNG，单张不超过 32 MB。", true); return;
          }
          return uploadPhotos(slot, title, [file]);
        }, "secondary small"),
        identity ? button("移除此处照片", () => {
          if (!window.confirm(`移除「${title}」的本次照片关联？保存任务后生效，原照片及历史报告不删除。`)) return;
          delete state.task.photos[slot];
          if (state.task.photo_edits) delete state.task.photo_edits[slot];
          state.task.photo_order = (state.task.photo_order || []).filter((key) => key !== slot);
          if (state.task.photo_layout) delete state.task.photo_layout[`photo:${slot}`];
          dirty(); render();
        }, "danger small") : null));
    }
    return content;
  }
  function photoGroupKeys(group) {
    return [...new Set([...(state.task.photo_order || []), ...Object.keys(state.task.photos)])]
      .filter((key) => state.task.photos[key] && key.split(":")[0] === group);
  }
  function multiPhotoCard(slot, editable) {
    editable = editable && !state.photoUpload && state.photoUploadFailure?.task !== state.task;
    const keys = photoGroupKeys(slot.id);
    const content = el("section", { class: "multi-photo-group" }, el("h3", {}, slot.name));
    keys.forEach((key, index) => {
      const item = photoCard(key, `${slot.name} ${index + 1}`, slot.width_cm, editable);
      if (editable) {
        const move = (offset) => {
          const order = [...keys];
          [order[index], order[index + offset]] = [order[index + offset], order[index]];
          state.task.photo_order = [...(state.task.photo_order || []).filter((entry) => !keys.includes(entry)), ...order];
          dirty(); render();
        };
        item.append(el("div", { class: "toolbar" },
          button("上移", () => move(-1), "secondary small", index === 0),
          button("下移", () => move(1), "secondary small", index === keys.length - 1)));
      }
      content.append(item);
    });
    if (editable) {
      const input = el("input", { type: "file", multiple: true, accept: ".jpg,.jpeg,.png",
        "aria-label": `批量添加${slot.name}` });
      content.append(input, button("添加照片（可多选）", () => {
        const files = Array.from(input.files || []);
        if (!files.length) { feedback("请先选择需要添加的照片。", true); return; }
        if (files.some((file) => !/\.(jpe?g|png)$/i.test(file.name) || file.size > 32 * 1024 * 1024)) {
          feedback("照片支持 JPG、JPEG、PNG，单张不超过32 MB。", true); return;
        }
        return uploadPhotos(slot.id, slot.name, files, true);
      }, "secondary small"));
    } else if (!keys.length) content.append(el("div", { class: "photo-empty" }, "尚未上传"));
    return content;
  }
  function photoPreviewSignature(includeLayout = true) {
    const task = state.task;
    return JSON.stringify([task.id, task.outline_id, task.photos, task.photo_order, task.photo_edits,
      task.display_types, task.photo_components, task.overview?.find((item) => item.key === "report_number")?.value, task.report_type,
      includeLayout ? task.photo_layout : null]);
  }
  function renderPhotoLayoutPreview() {
    const panel = el("section", { class: "card photo-layout-preview" });
    const status = el("div", { role: "status" }), pages = el("div", { class: "photo-layout-pages" });
    const editor = el("div", { class: "photo-layout-editor" });
    let selectedTarget = null;
    const stages = { queued: "排队中", opening: "打开Word工作副本", waiting_for_document: "等待Word文档",
      repaginating: "Word正在分页", exporting_pdf: "生成PDF", reading_pdf: "读取PDF",
      closing_copy: "关闭Word工作副本", rendering_pages: "生成页面预览", reading_page_boundaries: "定位照片与标题" };
    function paintEditor(result) {
      editor.replaceChildren();
      const target = result.targets?.find((item) => item.id === selectedTarget);
      if (!target) {
        editor.append(notice("点击预览页中的照片或标题区域，设置其前方空行。附件2自然分页，标题与照片保持同页；旧的强制分页设置不再生效。"));
        return;
      }
      const value = state.task.photo_layout?.[target.id] || {};
      const blanks = value.blank_lines ?? target.default_blank_lines;
      const change = (patch) => {
        state.task.photo_layout ||= {};
        state.task.photo_layout[target.id] = {
          blank_lines: null, ...state.task.photo_layout[target.id], ...patch, page_break_before: null
        };
        dirty(); paint();
      };
      editor.append(el("strong", {}, `已选：${target.label}`),
        el("p", {}, `前方空行：${blanks}行（${target.blank_line_spacing ?? 1.5}倍行距）　自然分页，标题与照片保持同页`),
        el("div", { class: "toolbar" },
          button("增加空行", () => change({ blank_lines: blanks + 1 }), "secondary small"),
          button("减少空行", () => change({ blank_lines: blanks - 1 }), "secondary small", blanks === 0),
          button("恢复此处默认排版", () => {
            if (state.task.photo_layout) delete state.task.photo_layout[target.id];
            dirty(); paint();
          }, "secondary small")),
        notice("修改已应用到本次任务，请保存任务；点击“生成／更新排版预览”查看新的分页。"));
    }
    function paint() {
      const preview = state.photoLayoutPreview;
      status.replaceChildren(); pages.replaceChildren(); editor.replaceChildren();
      if (!preview) return;
      const stale = preview.signature !== photoPreviewSignature();
      const layoutOnly = stale && preview.contentSignature === photoPreviewSignature(false);
      if (stale && !layoutOnly) {
        status.append(notice("照片、顺序、尺寸或大纲已变化，原预览已过期，请重新生成。", "warning"));
        return;
      }
      const result = preview.result;
      if (layoutOnly) status.append(notice("排版设置已修改。下方仍是旧版页面，可继续选择位置；请更新预览查看新排版。", "warning"));
      if (result.status === "running") status.append(notice(`正在${stages[result.stage] || "处理"}，Word排版最长等待90秒。`));
      if (result.status === "running" && result.stage === "waiting_for_document") status.append(notice("请留意Microsoft Word窗口；若工作副本弹出是否更新文档域的提示，请选择“否 / No”以继续。", "warning"));
      if (result.status === "failed") status.append(notice(`排版预览失败：${result.error}`, "error"));
      if (result.warnings?.length) status.append(notice(result.warnings, "warning"));
      if (result.word_url && !stale) status.append(downloadLink(result.word_url, "下载附件2 Word工作副本"));
      if (result.status === "ready") {
        if (!stale) status.append(downloadLink(result.pdf_url, `下载排版PDF（${result.pages}页）`));
        paintEditor(result);
        result.page_urls.forEach((url, index) => {
          const image = el("img", { src: url, alt: `附件2排版第${index + 1}页`, loading: "lazy" });
          image.addEventListener("error", () => status.append(notice("页面预览加载失败，请重新生成预览或下载PDF。", "error")), { once: true });
          const sheet = el("div", { class: "photo-layout-sheet" }, image);
          const size = result.page_sizes?.[index];
          if (size) sheet.style.aspectRatio = `${size.width} / ${size.height}`;
          (result.targets || []).filter((target) => target.page === index + 1).forEach((target) => {
            const hit = button(`选择：${target.label}`, () => {
              selectedTarget = target.id; paint();
            }, `photo-layout-target${selectedTarget === target.id ? " selected" : ""}`);
            hit.style.top = `${target.top * 100}%`;
            hit.style.height = `${target.height * 100}%`;
            hit.setAttribute("aria-label", `选择排版位置：${target.label}`);
            hit.setAttribute("aria-pressed", String(selectedTarget === target.id));
            sheet.append(hit);
          });
          pages.append(el("figure", {}, el("figcaption", {}, `第 ${index + 1} 页${layoutOnly ? "（旧版位置）" : ""}`), sheet));
        });
      }
    }
    async function poll() {
      const preview = state.photoLayoutPreview;
      if (!panel.isConnected || !preview || preview.result.status !== "running") return;
      try {
        const result = await api(`/api/photo-previews/${preview.result.id}`);
        if (state.photoLayoutPreview !== preview) return;
        preview.result = result; paint();
        if (result.status === "running") setTimeout(poll, 1000);
      } catch (error) {
        if (state.photoLayoutPreview !== preview) return;
        preview.result = { ...preview.result, status: "failed", error: error.message };
        paint();
      }
    }
    panel.append(el("h2", {}, "附件2 · Word排版预览"),
      notice("按当前照片、顺序和编辑尺寸生成附件2工作副本，由本机Microsoft Word实际分页。保留照片章节页面设置；页码仅为附件内页码，不是完整报告页码。预览不保存任务、不新增报告版本，也不代表人工审版通过。"),
      button("生成／更新排版预览", () => run("正在准备附件2预览…", async () => {
        const signature = photoPreviewSignature();
        const contentSignature = photoPreviewSignature(false);
        const result = await api("/api/photo-previews", clone(state.task));
        state.photoLayoutPreview = { signature, contentSignature, result };
        paint(); setTimeout(poll, 500);
      }, "已开始Word排版，可继续编辑；修改后须重新生成预览。"), "secondary"),
      status, editor, pages);
    state.photoPreviewPaint = () => { if (panel.isConnected) paint(); };
    paint(); setTimeout(poll, 1000);
    return panel;
  }
  function renderPhotos(editable = true) {
    const root = el("div", { class: "attachment-photos" });
    if (editable && window.workbenchPhotos?.enabled()) {
      root.append(card("云端照片优化（仅新上传照片）",
        field("上传清晰度", state.cloudPhotoMode || "standard", value => {state.cloudPhotoMode = value;}, {
          choices: [["standard", "标准：长边1600像素，JPEG质量约85%"],
            ["detail", "高清保留：长边2400像素，JPEG质量约92%"]]}),
        notice("上传前在浏览器后台缩小；小图不放大，PNG保留透明/线条格式，JPG保持JPEG存储。200～500KB仅为目标，不为凑大小强行降低清晰度。电脑原图、已有照片和历史报告不改动。铭牌、小字或大幅裁剪请选择高清保留，上传后放大核对。")));
    }
    if (editable && (state.photoUpload || state.photoUploadFailure?.task === state.task)) root.append(renderPhotoUploadStatus());
    const gas = ["燃气表", "流量计（气体）"].includes(state.task.report_type);
    if (editable) root.append(notice(gas
      ? "燃气表和流量计（气体）共用附件2照片排版：样机、封样和显示照片默认宽12cm，内部结构默认10cm，零部件默认8cm，均可在线调整宽度、旋转和裁剪，高度按比例计算，不拉伸。样机、封样和显示照片支持多张，同一照片栏的相邻照片间默认一个1.5倍行距空行。可在排版预览中调整空行，标题与照片保持同页、自然分页。修改后请保存任务；照片过高时可调整宽度或裁剪并核对Word，不自动缩小指定宽度。"
      : "支持 JPG、JPEG、PNG，单张最多32 MB、4000万像素。同一照片栏上传多张时，相邻照片间默认一个1.5倍行距空行。支持左右旋转90°、自由裁剪及按厘米设置输出宽度，高度按裁剪后比例自动计算。原图不改动，可恢复。上传、替换、编辑、移除后请保存任务。照片过高时需调整宽度或裁剪并核对 Word，不自动缩小指定宽度。"));
    if (state.task.photo_issues.length) root.append(notice(state.task.photo_issues, "error"));
    const slots = new Map((state.catalog.photo_slots || []).map((slot) => [slot.id, gas ? { ...slot, width_cm: slot.id === "internal" ? 10 : 12 } : slot]));
    const displayTypes = state.task.display_types ?? state.catalog.basic_options?.display_types ?? [];
    const paired = displayTypes.includes("机械") && displayTypes.includes("电子");
    const displaySlots = paired ? ["display_mechanical", "display_electronic"] : ["display"];
    [
      ["一、样机的整体外形照片", gas ? ["sample", "sealed_overall"] : ["sample", "nameplate", "sealed_overall", "sealed_detail"],
        gas ? "样机照片和封样照片均支持多张，默认宽12cm（可调整）、等比例居中，同一照片栏的相邻照片间留一个1.5倍行距空行。封样组接续上文、自然分页，不显示“封样”题注。本类型不显示、不输出铭牌和封样特写；旧关联及原照片保留。"
          : "样机照片后空一行放铭牌照片；整体封样照片接续上文、自然分页，空一行放封样特写，其下一行为“封样”。题注与对应照片保持同页。"],
      ["二、样机的内部结构照片", ["internal"], gas
        ? "内部结构标题不强制另起一页，接续上一部分；照片默认宽10cm（可调整）、等比例居中，空间不足时标题与照片一起自然分页。"
        : "接续上一部分，不强制另起一页；标题下一行放居中、等比例的内部结构照片（默认宽8cm，可调整），空间不足时自然分页。"],
      ["三、样机的显示部分照片", displaySlots, gas
        ? "显示部分支持多张照片，默认宽12cm（可调整）、等比例居中，相邻照片间默认一个1.5倍行距空行；标题接续上文，空间不足时随照片移至下一页。机械／电子组合仍分组输出，保留各自题注。"
        : paired
        ? "标题接续上文，空间不足时随照片移至下一页；先机械部分照片、后电子部分照片，每张下方有同名居中题注并与照片保持同页。各自默认宽8cm，支持独立旋转、裁剪和宽度设置。"
        : "标题接续上文，空间不足时随照片移至下一页；下一行放居中、等比例的显示照片（默认宽8cm，可调整），照片下接“四、样机的关键零部件照片”。"],
    ].forEach(([title, identifiers, description]) => {
      root.append(card(title, notice(description),
        el("div", { class: "photo-grid" }, identifiers.map((id) => {
          const slot = slots.get(id);
          return slot ? ((["sample", "display", "display_mechanical", "display_electronic"].includes(slot.id) || (gas && slot.id === "sealed_overall"))
            ? multiPhotoCard(slot, editable) : photoCard(slot.id, slot.name, slot.width_cm, editable))
            : notice(`照片位置「${id}」未由当前服务提供。请保存当前任务，重启本机服务后刷新页面或素材目录；不能使用旧版服务上传此照片。`, "error");
        }))));
    });
    root.append(card("四、样机的关键零部件照片",
      el("p", { class: "muted" }, "名称来自当前大纲关键零部件和材料表中“名称”的最右侧子列，按原顺序编号，不显示左侧组名。每个名称下一行放对应照片，再接下一个名称。" + (gas ? "本类型零部件照片默认宽8cm，可调整，保持长宽比。" : "")),
      state.task.photo_components.length
        ? el("div", { class: "photo-grid" }, state.task.photo_components.map((component, index) =>
          photoCard(component.id, `${index + 1}）${component.name}`, 8, editable)))
        : empty("尚无已识别零部件", "请先在“资料与任务”上传并选择本次大纲；无法识别时会明确提醒，不沿用其他报告的零部件名称。")));
    const allowed = new Set([...(state.catalog.photo_slots || []).map((slot) => slot.id), ...state.task.photo_components.map((component) => component.id)]);
    const outdated = Object.keys(state.task.photos).filter((slot) => !allowed.has(slot) &&
      !["sample", "display", "display_mechanical", "display_electronic", ...(gas ? ["sealed_overall"] : [])].includes(slot.split(":")[0]));
    if (outdated.length) root.append(card("需要重新核对的旧照片关联",
      notice("以下照片已不对应当前大纲，不能自动转给其他名称。请移除旧关联后重新上传到正确位置。", "error"),
      el("div", { class: "photo-grid" }, outdated.map((slot) => photoCard(slot, `旧零部件位置 ${slot}`, 8, editable)))));
    if (editable) root.append(renderPhotoLayoutPreview());
    return root;
  }
  function localIssues() {
    const task = state.task;
    const errors = [];
    errors.push(...task.photo_issues);
    errors.push(...task.construction_issues);
    errors.push(...task.sealing_issues);
    errors.push(...(task.observation_certificates?.issues || []));
    ["explosion", "protection"].forEach((key) => {
      const item = task.observation_certificates?.[key];
      if (item?.targets?.length && !item.value.trim()) errors.push(`观察/项目一览表备注/${key === "explosion" ? "防爆" : "防护"}证书编号尚未填写。`);
    });
    if (state.sealingPending) errors.push("封样信息正在更新。");
    if (state.sealingError) errors.push(state.sealingError);
    errors.push(...task.supplementary_issues);
    errors.push(...task.report_environment.issues);
    errors.push(...task.evaluation_summary.issues, ...(task.production_basis_issues || []));
    (task.measurement_parameters || []).forEach((row, index) => {
      measurementColumns.forEach(([key, label]) => {
        if (!row[key]?.trim()) errors.push(`计量器具测量参数第${index + 1}行「${label}」尚未填写。`);
      });
      if (task.include_volume_display && !row.volume_display?.trim())
        errors.push(`计量器具测量参数第${index + 1}行「${task.volume_display_label}」尚未填写。`);
    });
    if (!task.volume_display_label?.trim()) errors.push("体积量显示列名不能为空。");
    if (task.include_volume_display && !task.measurement_parameters)
      errors.push(`计量器具测量参数「${task.volume_display_label}」尚未填写。`);
    const displayTypes = task.display_types ?? state.catalog.basic_options?.display_types ?? [];
    const paired = displayTypes.includes("机械") && displayTypes.includes("电子");
    const inactiveDisplay = paired ? ["display"] : ["display_mechanical", "display_electronic"];
    (state.catalog.photo_slots || []).filter((slot) => !inactiveDisplay.includes(slot.id)
      && !(["燃气表", "流量计（气体）"].includes(task.report_type) && ["nameplate", "sealed_detail"].includes(slot.id))).forEach((slot) => {
      if (!Object.keys(task.photos).some((key) => key === slot.id || key.split(":")[0] === slot.id))
        errors.push(`附件2缺少${slot.name}。`);
    });
    task.photo_components.forEach((component) => {
      if (!task.photos[component.id]) errors.push(`附件2缺少零部件照片：${component.name}。`);
    });
    if (!task.name.trim()) errors.push("任务名称未填写。");
    if (!task.overview.length) errors.push("总览字段未载入。");
    if (!task.samples.length || task.samples.some((s) => !s.number.trim()))
      errors.push("样机清单编号未填齐。");
    if (task.samples.length && !task.samples.some((s) => s.disassembly)) errors.push("未选择拆样样机。");
    if (task.legacy_sample_fields.length) errors.push("旧任务样机字段尚未核对。");
    applicableOverviewFields(task).forEach((f) => {
      if (f.key in environmentFields || automaticOverviewFields.has(f.key)) return;
      if (!f.value.trim()) errors.push(`总览「${f.label}」未填写。`);
    });
    if (!task.items_asset_id) errors.push("未选择项目一览表。");
    else if (!state.orderVerified) errors.push("项目顺序尚未按当前一览表核对成功，请解决未匹配或歧义后重新核对。");
    if (!task.observe_asset_id) errors.push("未选择观察模块。");
    if (!task.outline_id) errors.push("未选择本次大纲。");
    if (!task.projects.length) errors.push("未选择结果项目。");
    if (!task.equipment.length) errors.push("未添加本次设备快照。");
    task.equipment.forEach((equipment, index) => {
      [["name", "名称"], ["model", "型号"], ["performance", "测量性能"], ["number", "设备编号"],
        ["valid_until", "证书有效期"], ["storage_location", "存放地点"], ["trial_group", "试验小组"]].forEach(([key, label]) => {
        if (!equipment[key]?.trim()) errors.push(`本次设备 ${index + 1}「${equipment.name || "未命名"}」缺少${label}；请完善设备库并明确刷新快照。`);
      });
    });
    task.projects.forEach((p, index) => {
      const prefix = `第 ${index + 1} 项「${p.name}」`;
      [["start", "开始日期 / 时间"], ["end", "结束日期 / 时间"], ["temperature", "环境温度"], ["humidity", "环境湿度"]].forEach(([key, name]) => { if (!p[key].trim()) errors.push(`${prefix}缺少${name}。`); });
      [["start", "开始日期"], ["end", "结束日期"]].forEach(([key, label]) => {
        if (p[key].trim() && !/^\d{4}-\d{2}-\d{2}(?:$|[ T])/.test(p[key].trim())) errors.push(`${prefix}${label}需包含 YYYY-MM-DD 格式的完整年月日。`);
      });
      if ((p.water_output_labels == null && !p.rows.length) ||
          outputRows(p).some((r) => !r.sample_id?.trim() || !r.value?.trim())) errors.push(`${prefix}缺少样机行、真实样机编号或实测结果。`);
      if (!p.equipment_ids.length) errors.push(`${prefix}未选择设备。`);
      if (p.equipment_ids.some((id) => !task.equipment.some((e) => e.id === id))) errors.push(`${prefix}存在失效设备引用。`);
      staffIssues(p).forEach((issue) => errors.push(`${prefix}：${issue}`));
    });
    return errors;
  }
  function renderPreview() {
    const root = el("div");
    root.append(button("查看当前问题并定位", () => run("正在核对当前内容…", async () => {
      const epoch = state.epoch;
      const result = await api("/api/validate", state.task);
      if (epoch !== state.epoch) throw new Error("内容已改变，请重新检查。");
      state.validation = result; render();
    }, "已列出当前问题，可点击定位；缺项仍允许生成草稿。"), "secondary"));
    if (state.validation?.issues) root.append(issueList(state.validation.issues, true));
    const exportList = el("div");
    state.exports.forEach((output) => {
      const stale = output.epoch !== state.epoch || output.revision !== state.task.revision;
      exportList.append(el("div", { class: "export-card" },
        el("div", { class: "row-heading" }, el("strong", {}, `${output.mode === "draft" ? "待补录草稿" : "数据已校验 · 排版待核验报告"} · ${output.report_type ?? "水表"} · 版本 ${output.revision ?? "未返回"}`),
          el("span", { class: `pill${stale ? " warning" : ""}` }, stale ? "任务版本已变更" : "本次生成")),
        el("p", {}, exportGenerationTime(output)),
        el("p", {}, output.native_word === false ? "生成方式：直接生成Word"
          : output.pagination_strategy === "optimized" ? "生成方式：优化自动重排V1"
          : output.pagination_strategy === "optimized_v2" ? "生成方式：优化自动重排V2"
          : output.pagination_strategy === "optimized_v3" ? "生成方式：优化自动重排V3"
          : output.pagination_strategy === "optimized_v4" ? "生成方式：优化自动重排V4（统一排版）"
          : output.native_word === true || output.result_pagination?.report?.pages
            ? "生成方式：自动重排后生成Word" : "生成方式：历史版本未记录"),
        output.generation_timing ? el("details", {},
          el("summary", {}, `生成耗时：${output.generation_timing.total_seconds.toFixed(1)}秒（后端处理）`),
          el("ul", {}, Object.entries(output.generation_timing.stage_seconds).map(([name, seconds]) =>
            el("li", {}, `${name}：${seconds.toFixed(1)}秒`))),
          ["report", "record"].map((key) => {
            const detail = output.result_pagination?.[key];
            return detail?.stage_seconds ? el("div", {},
              el("strong", {}, `${key === "report" ? "型评报告" : "配套记录"}：${detail.passes}轮，最终PDF校验${detail.pdf_exports}次`),
              detail.reuse ? el("p", {}, detail.reuse.mode === "shared_structure"
                ? `V2复用观察及结果结构（原观察续表${detail.reuse.observation_tables_available}张、结果表${detail.reuse.result_tables_available}张）；释放自动分页，配套独立核验，不复用页码。`
                : `V2回退独立重排：${detail.reuse.fallback_reason}`) : null,
              el("ul", {}, Object.entries(detail.stage_seconds).map(([name, seconds]) =>
                el("li", {}, `${name}：${seconds.toFixed(1)}秒`)))) : null;
          })) : null,
        stale ? notice("此文件不再代表当前数据，请重新生成。保留链接仅用于核对历史内容。", "warning") : null,
        el("div", { class: "toolbar" }, downloadLink(output.url, stale ? "下载旧版 Word" : output.mode === "draft" ? "下载待补录 Word" : "下载待核验 Word"),
          output.record_url ? downloadLink(output.record_url, "下载配套原始记录（待填写）") : null,
          output.pdf_url ? downloadLink(output.pdf_url, "查看核对 PDF") : null,
          output.checked_url ? downloadLink(output.checked_url, "下载已分页 Word（仍需人工核对）") : null),
        output.localError ? notice(output.localError, "error") : null,
        renderExportComparison(output),
        output.issues?.length ? el("details", {}, el("summary", {}, "按位置查看生成时的问题"),
          issueList(output.issues, !stale)) : null,
        output.warnings?.length ? el("details", {}, el("summary", {}, `生成时的问题清单（${output.warnings.length}项）`),
          notice(output.warnings, "warning")) : null));
    });
    root.append(card("生成可编辑报告",
      el("fieldset", {},
        el("legend", {}, "生成方式"),
        [[false, "standard", "直接生成Word（较快，不等待Word重排）"],
          [true, "optimized_v4", "优化自动重排V4（统一排版）"]].map(([value, strategy, label]) =>
          el("label", { class: "check" },
            el("input", { type: "radio", name: "generation-method",
              checked: state.nativeWord === value && state.paginationStrategy === strategy,
              onchange: () => { state.nativeWord = value; state.paginationStrategy = strategy; } }), label)),
        el("p", {}, "两种方式均保留素材与样机筛选、追加记录及页底整组约束。默认直接生成Word，不等待Word重排，需手工更新页码。V4共用试验拆分的观察与结果分页器，按本次型号/规格和真实参试样机重建观察页，执行同页合并、编号叉号对齐及容量回填；保留完整报告样机组及实测值，检查页底/签名，最终重新打开Word并复核PDF。两份文件分别核验，不执行占位型号替换或清空报告实测值。耗时以实测为准。")),
      el("div", { class: "toolbar" },
        button("生成 Word 报告（允许缺项）", () => generate("auto"), "primary"),
        button("生成待核对草稿", () => generate("draft"), "secondary")),
      generationStatus(),
      exportList));
    return root;
  }
  function exportGenerationTime(output) {
    if (!output.generated_at) return "生成时间：历史版本未记录";
    const date = new Date(output.generated_at);
    if (Number.isNaN(date.getTime())) return "生成时间：记录格式异常，请核对";
    const pad = value => String(value).padStart(2, "0");
    return `生成时间：${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}（本机时间）`;
  }
  function downloadLink(url, label) {
    if (window.workbenchDownloads) return window.workbenchDownloads.link(url, label);
    if (document.querySelector('meta[name="workbench-file-transport"]')?.content === "json")
      return el("span", {class: "notice error", role: "alert"}, "下载组件未载入，请强制刷新页面后重试。");
    try {
      const safe = new URL(url, window.location.href);
      if (!url || safe.origin !== window.location.origin || !["http:", "https:"].includes(safe.protocol)) throw new Error();
      return el("a", { class: "button secondary", href: safe.href, target: "_blank", rel: "noopener" }, label);
    } catch { return el("span", { class: "unconfirmed" }, "服务返回的下载地址无效，未提供链接。"); }
  }
  function generate(mode) {
    if (state.task.projects.length && !state.orderVerified) { feedback("请先选择一览表并成功核对项目顺序，再生成报告；不输出未经确认的项目排序。", true); return; }
    const nativeWord = state.nativeWord;
    const paginationStrategy = state.paginationStrategy;
    run("正在保存生成快照并提交后台任务…", async () => {
      await refreshCurvePreview();
      const output = await api("/api/generation-jobs", { task: state.task, mode, native_word: nativeWord,
        pagination_strategy: paginationStrategy });
      if (!output.saved_task?.id || !output.id) throw new Error("后台未返回有效任务快照，当前编辑未清除，请核对任务列表。");
      state.task = normalizeTask(output.saved_task); state.dirty = false; clearDraft();
      state.generationJobs.unshift(output); paintGenerationJobs(); render();
      pollGenerationJobs();
    }, "已提交后台生成，可以继续编辑下一版。生成使用刚才保存的快照；进度在“预览与生成”中显示。");
  }

  function issueList(issues, current) {
    return el("div", {}, !current ? notice("这是历史快照的问题；当前内容已变化，请使用“查看当前问题并定位”重新检查。") : null,
      steps.map(([name], step) => {
        const group = issues.filter(issue => issue.target?.step === step);
        return group.length ? el("details", {}, el("summary", {}, `${name}（${group.length}项）`),
          group.map(issue => el("p", {}, `${issue.severity === "error" ? "需补齐/核对" : "提醒"}：${issue.message} `,
            current ? button("定位", () => locateIssue(issue.target), "small secondary") : null))) : null;
      }));
  }
  function locateIssue(target) {
    navigate(target.step);
    let element = $("panel");
    if (target.overview_key)
      element = [...document.querySelectorAll("[data-overview-key]")].find(node => node.dataset.overviewKey === target.overview_key);
    if (target.overview_group)
      element = [...document.querySelectorAll("[data-overview-group]")].find(node => node.dataset.overviewGroup === target.overview_group);
    if (target.certificate_key)
      element = [...document.querySelectorAll("[data-certificate-field]")].find(node => node.dataset.certificateField === target.certificate_key);
    if (target.photo_slot)
      element = [...document.querySelectorAll("[data-photo-slot]")].find(node => node.dataset.photoSlot === target.photo_slot);
    if (target.project_id) {
      const project = state.task.projects.find(p => p.id === target.project_id);
      if (!project) { feedback("对应项目已移除，请重新检查当前问题。", true); return; }
      element = [...document.querySelectorAll("[data-project-row]")].find(node => node.dataset.projectRow === target.project_id);
      if (target.row_id) {
        const index = project.rows.findIndex(row => row.id === target.row_id);
        if (index < 0) { feedback("结果行已改变，请重新检查。", true); return; }
        openResultsEditor(project);
        element = document.querySelector(".results-editor")?.querySelectorAll("tbody tr")[index];
      } else if (target.field && element) {
        const columns = {start:1, end:2, temperature:3, humidity:4, atmospheric_pressure:5};
        element = element.children[columns[target.field]] || element;
      }
    }
    if (!element) { feedback("对应填写位置已改变，请重新检查当前问题。", true); return; }
    element.scrollIntoView({block:"center", inline:"nearest"});
    element.classList.add("issue-target");
    (element.querySelector("input,textarea,select,button") || element).focus();
    setTimeout(() => element.classList.remove("issue-target"), 3000);
  }
  function renderExportComparison(output) {
    const comparison = output.comparison;
    if (!comparison) return button("查看本次相比上次生成的改动", () => loadExportComparison(output),
      "secondary", state.loadingComparisons.has(output.id));
    const count = comparison.sections.reduce((total, section) => total + section.changes.length, 0);
    const label = comparison.status === "first" ? "首次生成，无上一版本"
      : `相比上次生成（版本 ${comparison.previous_revision ?? "未知"}）的改动：${count}项${comparison.status === "unavailable" ? " · 比较不完整" : ""}`;
    return el("details", { open: !!output.comparisonExpanded,
      ontoggle: (event) => { output.comparisonExpanded = event.target.open; } }, el("summary", {}, label),
      comparison.warnings.length ? notice(comparison.warnings, "warning") : null,
      comparison.status === "unavailable" ? button("重新比较", () => loadExportComparison(output),
        "secondary", state.loadingComparisons.has(output.id)) : null,
      comparison.status === "compared" && !count ? el("p", {}, "任务字段及两份 Word 的可比较内容未发现变化。") : null,
      comparison.sections.map((section) => el("section", {},
        el("h3", {}, section.title),
        section.changes.map((change) => el("div", { class: "export-change" },
          el("strong", {}, change.location),
          el("div", { class: "preformatted" }, `上次：${change.before || "（无）"}`),
          el("div", { class: "preformatted" }, `本次：${change.after || "（无）"}`))))));
  }

  async function loadExportComparison(output) {
    if (state.loadingComparisons.has(output.id)) return;
    state.loadingComparisons.add(output.id);
    output.localError = "";
    const refresh = () => { if (state.step === steps.length - 1 && state.exports.includes(output)) render(); };
    refresh();
    try {
      output.comparison = await api(`/api/exports/${encodeURIComponent(output.id)}/comparison`);
      output.comparisonExpanded = true;
    } catch (error) {
      output.localError = error.message || String(error);
      if (state.exports.includes(output)) feedback(output.localError, true);
    } finally {
      state.loadingComparisons.delete(output.id);
      refresh();
    }
  }
  function discardAllowed() {
    const photos = state.photoUpload || state.photoUploadFailure?.task === state.task;
    if ((state.dirty || state.materialSelection || photos) && !window.confirm(photos
        ? "附件2照片上传尚未完成。继续会取消未完成照片的关联及后续上传，并丢弃当前未保存修改；已发出的请求可能仍保存照片副本。是否继续？"
        : state.materialSelection
          ? "当前任务有未完成的素材选择或未保存修改。继续会取消选择并丢弃未保存修改，是否继续？"
          : "当前任务有未保存修改。继续会丢弃这些修改，是否继续？")) return false;
    cancelResultSelection("", false);
    cancelPhotoUpload();
    return true;
  }
  function mergedWorkspaceEnabled() {
    return document.querySelector('meta[name="workbench-workspace-transport"]')?.content === "merged";
  }
  async function loadWorkspaceData() {
    const names = ["catalog", "equipment", "uploads", "photos", "principles", "sample_schemes", "material_schemes"];
    if (!mergedWorkspaceEnabled())
      return Promise.allSettled(names.map(name => api(`/api/${name === "sample_schemes" ? "sample-schemes" :
        name === "material_schemes" ? "material-schemes" : name}`)));
    const sections = await api("/api/workspace");
    return names.map(name => {
      const section = sections?.[name];
      const hasData = section && Object.prototype.hasOwnProperty.call(section, "data");
      const validData = hasData && (name === "catalog"
        ? section.data && Array.isArray(section.data.assets) && Array.isArray(section.data.overview)
        : Array.isArray(section.data));
      if (validData && !Object.prototype.hasOwnProperty.call(section, "error"))
        return {status: "fulfilled", value: section.data};
      return {status: "rejected", reason: new Error(!hasData && typeof section?.error === "string" && section.error
        ? section.error : `后台返回的${name}资料格式无效，请重试。`)};
    });
  }
  async function loadTaskWorkspace(identity) {
    const path = `/api/tasks/${encodeURIComponent(identity)}`;
    if (!mergedWorkspaceEnabled()) {
      const [task, exports, photos] = await Promise.all([api(path), api(`${path}/exports`), api("/api/photos")]);
      return {task, exports, photos, order_verified: false, order_changed: false, merged: false};
    }
    const result = await api(`${path}/workspace`);
    if (!result?.task || result.task.id !== identity || !Array.isArray(result.task.projects) ||
        !Array.isArray(result.exports) || !Array.isArray(result.photos) ||
        typeof result.order_verified !== "boolean" || typeof result.order_changed !== "boolean" ||
        result.order_verified !== Boolean(result.task.items_asset_id) ||
        (result.order_changed && !result.order_verified))
      throw new Error("后台返回的任务或核对结果无效，当前任务未替换，请重试。");
    return {...result, merged: true};
  }
  async function initialize() {
    await run("正在连接本机服务并读取素材、设备和上传记录…", async () => {
      const results = await loadWorkspaceData();
      const failures = [];
      const names = ["素材目录", "设备库", "已上传资料", "照片资料", "原理选项", "样机方案库", "素材方案库"];
      results.forEach((result, index) => {
        if (result.status === "fulfilled") state[["catalog", "equipment", "uploads", "photoFiles", "principles", "sampleSchemes", "materialSchemes"][index]] = result.value;
        else failures.push(`${names[index]}：${result.reason.message}`);
      });
      state.ready = results[0].status === "fulfilled";
      state.loadErrors = failures;
      if (!state.task) state.task = newTask();
      else if (!state.task.overview.length && state.catalog.overview.length) {
        state.task.overview = newTask().overview;
        dirty();
      }
      render();
      if (failures.length) {
        throw new Error(`部分资料未载入：${failures.join("；")}。不要直接刷新未保存任务。`);
      }
    }, "已连接本机工作空间。开始新任务，或打开已保存任务。");
    pollGenerationJobs();
    if (state.ready) preloadResultMaterials();
    try {
      if (Object.keys(localStorage).some(key => key.startsWith(recoveryPrefix)) && $("recovery-state"))
        $("recovery-state").textContent = "发现未保存草稿，可点击顶部“恢复暂存”查看。";
    } catch (error) { feedback("浏览器暂存不可用：" + error.message, true); }
  }
  $("save-task").addEventListener("click", () => {
    if (!state.task) return;
    run("正在保存本次任务到本机…", async () => {
      const saved = await api("/api/tasks", state.task);
      state.task = normalizeTask(saved); state.dirty = false; clearDraft(); render();
    }, "任务已保存到本机服务。设备快照与项目顺序随任务保留。");
  });
  $("new-task").addEventListener("click", () => {
    if (!discardAllowed()) return;
    clearDraft();
    state.task = newTask(); state.dirty = false; state.epoch = 0; state.exports = []; state.validation = null; state.step = 0; state.orderVerified = false;
    state.clockDrafts.clear(); state.autoEndTimes.clear(); state.autoStartTimes.clear();
    render(); feedback("已建立空白任务，未继承上次素材、数据或人员。");
  });
  async function refreshTaskList() {
    const [tasks, reports] = await Promise.all([api("/api/tasks"), api("/api/retained-reports")]);
    $("task-list").replaceChildren(...tasks.map((task) => el("div", { class: "task-list-row", "data-task-id": task.id }, button(
      el("span", {}, el("strong", {}, task.name), el("small", {}, ` · ${task.report_type ?? "水表"} · 版本 ${task.revision} · ${task.updated_at || "更新时间未提供"}`)),
      () => {
        if (!discardAllowed()) return;
        run("正在打开任务…", async () => {
          const loaded = await loadTaskWorkspace(task.id);
          state.photoFiles = loaded.photos;
          clearDraft();
          state.task = normalizeTask(loaded.task); state.dirty = false; state.epoch = 0;
          state.clockDrafts.clear(); state.autoEndTimes.clear(); state.autoStartTimes.clear();
          state.exports = loaded.exports.map((output) => ({ ...output, epoch: 0 }));
          state.validation = null; state.step = 0; state.orderVerified = loaded.order_verified;
          if (loaded.order_changed) dirty();
          $("task-dialog").close(); render();
          if (!loaded.merged && state.task.items_asset_id) {
            const previousOrder = state.task.projects.map((p) => p.id).join("\n");
            state.task = await orderedTask(state.task);
            if (state.task.projects.map((p) => p.id).join("\n") !== previousOrder) dirty();
            render();
          }
        }, "已打开已保存任务。当前设备快照不会被最新设备库自动替换。");
      }, "secondary task-entry"),
      el("button", { type: "button", class: "button danger small task-delete",
        "aria-label": `删除任务 ${task.name}（版本 ${task.revision}，ID ${task.id}）`,
        onclick: () => deleteTask(task) }, "删除"))));
    if (!tasks.length) $("task-list").append(empty("暂无已保存任务", "填写内容后点击右上角「保存任务」。"));
    if (reports.length) $("task-list").append(el("details", { class: "retained-reports" },
      el("summary", {}, `已删除任务的报告（${reports.length} 份）`),
      notice("任务已删除，以下导出文件及生成时快照仍保留。可继续下载，不代表已经人工审版或正式签发。"),
      reports.map((report) => el("section", { class: "export-card" },
        el("strong", {}, report.task_name), el("small", {}, ` · ${report.report_type ?? "水表"} · 导出版本 ${report.revision}`),
        el("p", {}, exportGenerationTime(report)),
        el("div", { class: "toolbar" },
          downloadLink(report.url, "下载 Word"),
          report.record_url ? downloadLink(report.record_url, "下载配套原始记录") : null,
          report.checked_url ? downloadLink(report.checked_url, "下载已分页 Word") : null,
          report.pdf_url ? downloadLink(report.pdf_url, "下载 PDF") : null,
          downloadLink(`/api/exports/${encodeURIComponent(report.id)}/manifest.json`, "生成清单"))))));
  }
  function deleteTask(task) {
    const current = state.task?.id === task.id;
    const localWarning = current ? "\n这也是当前打开的任务，删除后会清空当前页面，包括尚未保存的修改。" : "";
    if (!window.confirm(`永久删除任务「${task.name}」及编辑历史？\n版本：${task.revision}\n任务ID：${task.id}\n已生成报告、上传资料和设备库均保留。${localWarning}`)) return;
    run("正在删除任务，保留已生成报告…", async () => {
      await api(`/api/tasks/${encodeURIComponent(task.id)}?revision=${task.revision}`, undefined, "DELETE");
      if (current) {
        clearDraft();
        state.task = newTask(); state.dirty = false; state.epoch = 0; state.exports = [];
        state.validation = null; state.step = 0; state.orderVerified = false;
        state.clockDrafts.clear(); state.autoEndTimes.clear(); state.autoStartTimes.clear();
        render();
      }
      [...$("task-list").querySelectorAll("[data-task-id]")].find((row) => row.dataset.taskId === task.id)?.remove();
      try { await refreshTaskList(); }
      catch (error) { throw new Error(`任务已删除，但列表刷新失败：${error.message}。请重新打开任务列表。`); }
    }, "任务已删除。已生成报告可在任务列表下方的「已删除任务的报告」中下载；上传资料和设备库未改动。");
  }
  $("open-task").addEventListener("click", () => {
    if (state.materialSelection && !window.confirm("取消尚未完成的素材选择并打开任务列表？当前填写内容保持不变。")) return;
    cancelResultSelection("", false);
    run("正在读取已保存任务…", async () => {
    await refreshTaskList();
    $("task-dialog").showModal();
    }, "任务列表已载入。");
  });
  $("close-dialog").addEventListener("click", () => $("task-dialog").close());
  $("version-notes").addEventListener("click", () => $("version-dialog").showModal());
  $("close-version-dialog").addEventListener("click", () => $("version-dialog").close());
  $("version-dialog").addEventListener("keydown", (event) => {
    if (event.key === "Escape") { event.preventDefault(); $("version-dialog").close(); }
  });
  $("previous-step").addEventListener("click", () => navigate(state.step - 1));
  $("next-step").addEventListener("click", () => navigate(state.step + 1));
  window.addEventListener("beforeunload", (event) => {
    persistDraft();
    if (state.dirty || state.busy || state.materialSelection || state.photoUpload || state.photoUploadFailure?.task === state.task || state.equipmentEditorDirty || state.resultsEditorDirty) { event.preventDefault(); event.returnValue = ""; }
  });
  $("recovery")?.addEventListener("click", openRecovery);
  $("backups")?.addEventListener("click", openBackups);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") persistDraft();
  });
  initialize();
})();
