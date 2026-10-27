"use strict";

self.onmessage = async ({data: {file, mode}}) => {
  let bitmap;
  try {
    if (!["standard", "detail"].includes(mode) || !file.size || file.size > 32 * 1024 * 1024)
      throw Error("照片模式或文件大小无效，单张须非空且不超过32 MB。");
    const bytes = new Uint8Array(await file.arrayBuffer());
    const png = [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value);
    const jpeg = bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
    if (!png && !jpeg) throw Error("照片内容不是有效的 JPG 或 PNG。");
    if (png) {
      const view = new DataView(bytes.buffer);
      let offset = 8;
      while (offset + 12 <= bytes.length) {
        const length = view.getUint32(offset);
        if (offset + 12 + length > bytes.length) throw Error("PNG文件不完整。");
        const type = String.fromCharCode(...bytes.slice(offset + 4, offset + 8));
        if (type === "acTL") throw Error("不支持动态照片，请上传单张静态照片。");
        offset += length + 12;
        if (type === "IEND") break;
      }
    }
    bitmap = await createImageBitmap(file, {imageOrientation: "from-image"});
    if (bitmap.width * bitmap.height > 40000000) throw Error("单张照片最多支持4000万像素。");
    const limit = mode === "detail" ? 2400 : 1600;
    const scale = Math.min(1, limit / Math.max(bitmap.width, bitmap.height));
    const width = Math.max(1, Math.round(bitmap.width * scale));
    const height = Math.max(1, Math.round(bitmap.height * scale));
    const canvas = new OffscreenCanvas(width, height);
    const context = canvas.getContext("2d");
    if (!context) throw Error("浏览器无法处理照片，请使用新版浏览器。");
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = "high";
    context.drawImage(bitmap, 0, 0, width, height);
    const type = png ? "image/png" : "image/jpeg";
    const blob = await canvas.convertToBlob({type, quality: mode === "detail" ? 0.92 : 0.85});
    if (blob.type !== type || !blob.size) throw Error("浏览器照片编码失败。");
    self.postMessage({blob, width, height, originalWidth: bitmap.width, originalHeight: bitmap.height,
      originalBytes: file.size});
  } catch (error) {
    self.postMessage({error: error.message || String(error)});
  } finally {
    bitmap?.close();
  }
};
