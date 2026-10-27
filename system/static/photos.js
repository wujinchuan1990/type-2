"use strict";

(() => {
  const workerURL = new URL("photo-upload-worker.js", document.currentScript.src);
  function enabled() {
    return document.querySelector('meta[name="workbench-photo-optimization"]')?.content === "enabled";
  }
  function prepare(file, mode) {
    return new Promise((resolve, reject) => {
      const worker = new Worker(workerURL);
      const finish = (error, data) => {
        clearTimeout(timer); worker.terminate();
        if (error) reject(new Error(error)); else resolve(data);
      };
      const timer = setTimeout(() => finish("照片处理超时，未上传，请重试或选择较小照片。"), 60000);
      worker.onerror = () => finish("照片处理失败，未上传。请使用新版浏览器后重试。");
      worker.onmessage = ({data}) => {
        if (data.error) { finish(data.error); return; }
        if (!(data.blob instanceof Blob) || !["image/jpeg", "image/png"].includes(data.blob.type) ||
            !data.blob.size || !Number.isInteger(data.width) || !Number.isInteger(data.height) ||
            data.width < 1 || data.height < 1 || Math.max(data.width, data.height) > (mode === "detail" ? 2400 : 1600)) {
          finish("照片处理结果无效，未上传。"); return;
        }
        finish(null, {...data, file: new File([data.blob], file.name, {type: data.blob.type})});
      };
      worker.postMessage({file, mode});
    });
  }
  const waiting = new Map(), queue = [];
  let active = 0, observer, changes;
  function pump() {
    while (active < 2 && queue.length) {
      const {image, load, failed} = queue.shift();
      if (!image.isConnected) continue;
      active++;
      Promise.resolve().then(load).catch(error => {
        if (image.isConnected) failed(error);
      }).finally(() => {active--; pump();});
    }
  }
  function lazy(image, load, failed) {
    if (!observer) {
      observer = new IntersectionObserver(entries => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const operation = waiting.get(entry.target);
          if (!operation) continue;
          waiting.delete(entry.target); observer.unobserve(entry.target);
          queue.push(operation);
        }
        pump();
      }, {rootMargin: "200px"});
      changes = new MutationObserver(() => {
        for (const image of waiting.keys()) {
          if (!image.isConnected) {observer.unobserve(image); waiting.delete(image);}
        }
      });
      changes.observe(document.body, {childList: true, subtree: true});
    }
    waiting.set(image, {image, load, failed});
    observer.observe(image);
  }
  window.workbenchPhotos = {enabled, prepare, lazy};
})();
