import hashlib
from io import BytesIO
from pathlib import Path
import re
import uuid
import warnings

from PIL import Image, ImageOps, UnidentifiedImageError
from contracts import PhotoEdit

MAX_PHOTO_BYTES = 32 * 1024 * 1024
MAX_PHOTO_PIXELS = 40_000_000


def edited_photo_size(size, edit: PhotoEdit):
    width, height = size
    if edit.rotation in (90, 270):
        width, height = height, width
    if edit.crop is not None:
        left, top, right, bottom = edit.crop
        if right > width or bottom > height:
            raise ValueError("裁剪区域超出旋转后的照片边界，请重新裁剪。")
        return right - left, bottom - top
    return width, height


def edited_photo_bytes(path, edit: PhotoEdit, *, thumbnail=False):
    """Apply task-local edits to verified pixels, never to the stored upload."""
    with Image.open(path) as source:
        edited_photo_size(source.size, edit)
        image = source
        if edit.rotation:
            image = image.transpose({
                90: Image.Transpose.ROTATE_270,
                180: Image.Transpose.ROTATE_180,
                270: Image.Transpose.ROTATE_90,
            }[edit.rotation])
        if edit.crop is not None:
            image = image.crop(edit.crop)
        if thumbnail:
            image.thumbnail((400, 400), Image.Resampling.LANCZOS)
        output = BytesIO()
        if source.format == "JPEG":
            image.convert("RGB").save(output, format="JPEG", quality=90, optimize=True, dpi=(96, 96))
        else:
            image.save(output, format="PNG", dpi=(96, 96))
        return output.getvalue()


def save_photo(store, content, filename, *, cloud_mode=None):
    if cloud_mode not in (None, "standard", "detail"):
        raise ValueError("无效的照片优化模式。")
    name = Path(filename or "").name
    if Path(name).suffix.lower() not in (".jpg", ".jpeg", ".png"):
        raise ValueError("照片支持 JPG、JPEG、PNG；请先将 HEIC 等格式转换后上传。")
    if not content or len(content) > MAX_PHOTO_BYTES:
        raise ValueError("单张照片须非空且不超过 32 MB。")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(BytesIO(content), formats=["JPEG", "PNG"]) as image:
                if image.width * image.height > MAX_PHOTO_PIXELS:
                    raise ValueError("单张照片最多支持 4000 万像素，请降低分辨率后上传。")
                if getattr(image, "n_frames", 1) != 1:
                    raise ValueError("不支持多帧或动态照片，请上传单张静态照片。")
                image.verify()
            with Image.open(BytesIO(content), formats=["JPEG", "PNG"]) as image:
                oriented = ImageOps.exif_transpose(image)
                pixels = oriented.convert("RGBA" if "A" in oriented.getbands() or "transparency" in oriented.info else "RGB")
                pixels.info.clear()
                if cloud_mode:
                    limit = 2400 if cloud_mode == "detail" else 1600
                    pixels.thumbnail((limit, limit), Image.Resampling.LANCZOS)
                width, height = pixels.size
                normalized = BytesIO()
                extension = ".jpg" if cloud_mode and image.format == "JPEG" else ".png"
                if extension == ".jpg":
                    pixels.save(normalized, format="JPEG", quality=90, optimize=True, dpi=(96, 96))
                else:
                    pixels.save(normalized, format="PNG", dpi=(96, 96))
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise ValueError("照片无法安全读取，请重新导出为有效的 JPG 或 PNG。") from exc
    identity = uuid.uuid4().hex
    path = store.directory / "photos" / (identity + extension)
    data = normalized.getvalue()
    created = False
    try:
        with path.open("xb") as target:
            created = True
            target.write(data)
        return store.save("photos", {
            "id": identity, "name": name, "width": width, "height": height,
            "sha256": hashlib.sha256(data).hexdigest(),
            "source_sha256": hashlib.sha256(content).hexdigest(),
            "url": f"/api/photos/{identity}/image",
            **({"extension": extension, "optimization": cloud_mode, "bytes": len(data),
                "uploaded_bytes": len(content)} if cloud_mode else {}),
        })
    except BaseException:
        if created:
            path.unlink(missing_ok=True)
        raise


def photo_path(store, identity):
    if not re.fullmatch(r"[0-9a-f]{32}", identity):
        raise ValueError("无效的照片标识。")
    info = store.get("photos", identity)
    extension = info.get("extension", ".png")
    if extension not in (".png", ".jpg"):
        raise ValueError("照片存储格式无效，请核对照片副本。")
    path = store.directory / "photos" / (identity + extension)
    if path.is_symlink() or not path.is_file():
        raise ValueError("照片副本丢失或不是普通文件，请重新上传：" + info["name"])
    if hashlib.sha256(path.read_bytes()).hexdigest() != info["sha256"]:
        raise ValueError("照片副本已被外部修改，请重新上传：" + info["name"])
    return path
