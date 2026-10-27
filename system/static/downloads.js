"use strict";

(() => {
  const maxBytes = 256 * 1024 * 1024;
  const chunkBytes = 1024 * 1024;

  function safeURL(value) {
    const url = new URL(value, window.location.href);
    if (!value || url.origin !== window.location.origin || !["http:", "https:"].includes(url.protocol) ||
        url.username || url.password || url.hash)
      throw new Error("下载地址无效或不是本站地址。");
    return url;
  }

  async function download(value, progress = () => {}) {
    const url = safeURL(value);
    const parts = [];
    let offset = 0, metadata = null;
    while (true) {
      url.searchParams.set("download_encoding", "base64");
      url.searchParams.set("download_offset", String(offset));
      if (metadata) url.searchParams.set("download_version", metadata.version);
      else url.searchParams.delete("download_version");
      const response = await fetch(url.pathname + url.search, {
        credentials: "same-origin", cache: "no-store", redirect: "error"
      });
      let data;
      try { data = await response.json(); }
      catch { throw new Error(`下载接口未返回有效JSON（HTTP ${response.status}），请刷新页面后重试。`); }
      if (!response.ok)
        throw new Error(data && typeof data.detail === "string" ? data.detail :
          `文件下载失败（HTTP ${response.status}），请重试。`);
      if (!data || typeof data.filename !== "string" || !data.filename ||
          typeof data.media_type !== "string" || !data.media_type ||
          !/^[0-9a-f]{64}$/.test(data.version) || !/^[0-9a-f]{64}$/.test(data.sha256) ||
          !Number.isSafeInteger(data.total_bytes) || data.total_bytes < 0 ||
          !Number.isSafeInteger(data.next_offset) || data.offset !== offset ||
          typeof data.eof !== "boolean" || typeof data.data_base64 !== "string" ||
          data.data_base64.length > Math.ceil(chunkBytes / 3) * 4 ||
          data.data_base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(data.data_base64))
        throw new Error("下载数据格式无效，未保存不完整文件。");
      if (data.total_bytes > maxBytes)
        throw new Error("文件超过云端浏览器下载的256 MB上限，请使用本机直接下载。");
      if (metadata && (data.version !== metadata.version || data.total_bytes !== metadata.total_bytes ||
          data.filename !== metadata.filename || data.media_type !== metadata.media_type))
        throw new Error("下载文件已变化，未混用不同版本，请重新下载。");
      const decoded = atob(data.data_base64);
      const bytes = Uint8Array.from(decoded, character => character.charCodeAt(0));
      if (bytes.length !== Math.min(chunkBytes, data.total_bytes - offset) ||
          data.next_offset !== offset + bytes.length || data.next_offset > data.total_bytes ||
          data.eof !== (data.next_offset === data.total_bytes) || (!data.eof && !bytes.length))
        throw new Error("下载文件分块不完整，未保存，请重试。");
      const digest = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)),
        byte => byte.toString(16).padStart(2, "0")).join("");
      if (digest !== data.sha256) throw new Error("下载数据校验失败，未保存，请重试。");
      metadata = data;
      parts.push(bytes);
      offset = data.next_offset;
      progress(offset, data.total_bytes);
      if (data.eof) break;
    }
    return {blob: new Blob(parts, {type: metadata.media_type}),
      filename: metadata.filename.replace(/[/\\\x00-\x1f]/g, "_")};
  }

  function link(value, label, filename = "") {
    let url;
    try { url = safeURL(value); }
    catch (error) {
      const invalid = document.createElement("span");
      invalid.className = "unconfirmed"; invalid.textContent = error.message;
      return invalid;
    }
    const anchor = document.createElement("a");
    anchor.className = "button secondary"; anchor.href = url.href;
    anchor.target = "_blank"; anchor.rel = "noopener";
    if (filename) anchor.download = filename;
    anchor.textContent = label;
    if (document.querySelector('meta[name="workbench-file-transport"]')?.content !== "json") return anchor;
    let busy = false, errorBox = null;
    anchor.addEventListener("click", async event => {
      event.preventDefault();
      if (busy) return;
      busy = true;
      errorBox?.remove(); errorBox = null;
      anchor.setAttribute("aria-busy", "true");
      anchor.setAttribute("aria-disabled", "true");
      anchor.textContent = "正在准备下载…";
      try {
        const result = await download(value, (loaded, total) => {
          anchor.textContent = `正在下载 ${total ? Math.floor(loaded * 100 / total) : 100}%…`;
        });
        const objectURL = URL.createObjectURL(result.blob);
        const save = document.createElement("a");
        save.href = objectURL; save.download = result.filename;
        document.body.append(save); save.click(); save.remove();
        setTimeout(() => URL.revokeObjectURL(objectURL), 60000);
        anchor.textContent = label;
      } catch (error) {
        anchor.textContent = label + "（失败，点击重试）";
        errorBox = document.createElement("span");
        errorBox.className = "notice error"; errorBox.setAttribute("role", "alert");
        errorBox.textContent = `下载失败：${error.message || String(error)}。已有报告未修改。`;
        anchor.insertAdjacentElement("afterend", errorBox);
      } finally {
        busy = false;
        anchor.removeAttribute("aria-busy"); anchor.removeAttribute("aria-disabled");
      }
    });
    return anchor;
  }

  window.workbenchDownloads = {download, link};
})();
