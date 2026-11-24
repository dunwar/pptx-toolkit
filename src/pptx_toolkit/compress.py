"""Compress a .pptx by re-encoding the images embedded in it.

A .pptx is a zip archive; the pictures live under ``ppt/media/``. Decks get
bloated when full-resolution photos are dropped in, so for each embedded
image this module:

* downscales it if either dimension exceeds ``max_dimension`` (aspect ratio
  is preserved, images are never enlarged), and
* re-encodes it: JPEGs at the requested quality; PNGs without transparency
  are converted to JPEG (relationship targets and content types are patched
  to match); PNGs with transparency stay PNG, just downscaled/optimized.

Vector formats (EMF/WMF/SVG) and GIFs are passed through untouched, as is
everything else in the archive.
"""

from __future__ import annotations

import io
import re
import zipfile

from PIL import Image

_MEDIA_PREFIX = "ppt/media/"
_JPEG_EXTS = {".jpg", ".jpeg"}
_PNG_EXTS = {".png"}
_RASTER_EXTS = _JPEG_EXTS | _PNG_EXTS | {".bmp", ".tif", ".tiff"}


def _has_alpha(img: Image.Image) -> bool:
    if img.mode in ("RGBA", "LA"):
        return True
    if img.mode == "P" and "transparency" in img.info:
        return True
    return False


def _reencode(data: bytes, max_dimension: int, quality: int) -> tuple[bytes, str] | None:
    """Return (new_bytes, new_format) or None if the image should be kept.

    new_format is "JPEG" or "PNG". None means: leave the original bytes
    alone (already small enough, or a format we do not touch).
    """
    try:
        img = Image.open(io.BytesIO(data))
        img.load()
    except Exception:
        return None

    fmt = img.format
    if fmt not in ("JPEG", "PNG", "BMP", "TIFF"):
        return None

    needs_resize = max(img.size) > max_dimension
    if fmt == "JPEG" and not needs_resize and quality >= 95:
        return None

    if needs_resize:
        img.thumbnail((max_dimension, max_dimension), Image.LANCZOS)

    out = io.BytesIO()
    if fmt == "PNG" and _has_alpha(img):
        img.save(out, format="PNG", optimize=True)
        return out.getvalue(), "PNG"

    # Everything else becomes (or stays) JPEG.
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.save(out, format="JPEG", quality=quality, optimize=True)
    return out.getvalue(), "JPEG"


def compress_presentation(
    input_path: str,
    output_path: str,
    max_dimension: int = 1920,
    quality: int = 80,
) -> tuple[int, int, int]:
    """Write a compressed copy of *input_path* to *output_path*.

    Returns (original_size, new_size, images_reencoded).
    """
    with zipfile.ZipFile(input_path, "r") as zin:
        entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        infos = {item.filename: item for item in zin.infolist()}

    renames: dict[str, str] = {}  # old media name -> new media name (png -> jpg)
    reencoded = 0

    for name in list(entries):
        if not name.startswith(_MEDIA_PREFIX):
            continue
        ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if ext not in _RASTER_EXTS:
            continue
        result = _reencode(entries[name], max_dimension, quality)
        if result is None:
            continue
        new_data, new_fmt = result
        if len(new_data) >= len(entries[name]):
            continue  # never grow the file
        reencoded += 1
        if new_fmt == "JPEG" and ext != ".jpg" and ext != ".jpeg":
            new_name = name.rsplit(".", 1)[0] + ".jpg"
            renames[name] = new_name
            entries[new_name] = new_data
            del entries[name]
        else:
            entries[name] = new_data

    if renames:
        _patch_references(entries, renames)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            info = infos.get(name) or infos.get(
                next((old for old, new in renames.items() if new == name), name)
            )
            new_info = zipfile.ZipInfo(name, date_time=info.date_time if info else (1980, 1, 1, 0, 0, 0))
            new_info.compress_type = zipfile.ZIP_DEFLATED
            new_info.external_attr = info.external_attr if info else 0
            zout.writestr(new_info, data)

    import os

    return os.path.getsize(input_path), os.path.getsize(output_path), reencoded


def _patch_references(entries: dict[str, bytes], renames: dict[str, str]) -> None:
    """Update .rels targets and [Content_Types].xml after png->jpg renames."""
    for name in list(entries):
        if name.endswith(".rels"):
            data = entries[name]
            changed = False
            for old, new in renames.items():
                old_target = old.encode()  # e.g. ppt/media/image1.png
                rel_target = old[len("ppt/"):].encode()  # media/image1.png
                new_rel_target = new[len("ppt/"):].encode()
                if rel_target in data:
                    data = data.replace(rel_target, new_rel_target)
                    changed = True
                elif old_target in data:
                    data = data.replace(old_target, new.encode())
                    changed = True
            if changed:
                entries[name] = data

    ct_name = "[Content_Types].xml"
    if ct_name in entries:
        data = entries[ct_name]
        if b'Extension="jpg"' not in data and b"Extension='jpg'" not in data:
            default = b'<Default Extension="jpg" ContentType="image/jpeg"/>'
            data = re.sub(rb"<Types[^>]*>", lambda m: m.group(0) + default, data, count=1)
            entries[ct_name] = data
