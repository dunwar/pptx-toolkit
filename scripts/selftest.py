#!/usr/bin/env python3
"""End-to-end self-test for pptx-toolkit.

Builds two small decks from scratch (deck A: title + text slide; deck B:
title + a slide with one deliberately oversized photo), then verifies:

1. ``merge`` produces a deck whose slide count is the sum of both inputs and
   which python-pptx can re-open;
2. ``compress`` produces a smaller deck that python-pptx can re-open.

Run:  python scripts/selftest.py [workdir]
Exits non-zero on the first failed check.
"""

from __future__ import annotations

import io
import os
import sys

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from pptx_toolkit.compress import compress_presentation  # noqa: E402
from pptx_toolkit.merge import merge_presentations  # noqa: E402


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f" - {detail}" if detail else ""))
    if not ok:
        raise SystemExit(1)


def make_big_photo(path: str) -> None:
    """A 3000x2000 photo-like PNG: color gradient + shapes + sensor noise.

    The noise matters: a smooth gradient PNG is unrealistically small, while
    real camera photos are noisy and several MB at this size - which is
    exactly the bloat `compress` exists to remove.
    """
    img = Image.new("RGB", (3000, 2000))
    draw = ImageDraw.Draw(img)
    for y in range(2000):
        draw.line([(0, y), (3000, y)], fill=(y * 255 // 2000, 80, 200 - y * 150 // 2000))
    for i in range(60):
        x, y = (i * 173) % 2800, (i * 389) % 1800
        draw.ellipse([x, y, x + 200, y + 140], fill=(i * 37 % 255, i * 91 % 255, i * 53 % 255))
    noise = Image.effect_noise((3000, 2000), 28).convert("RGB")
    img = Image.blend(img, Image.blend(img, noise, 0.5), 0.35)
    img.save(path, format="PNG")


def build_deck_a(path: str) -> int:
    prs = Presentation()
    title = prs.slides.add_slide(prs.slide_layouts[0])
    title.shapes.title.text = "Deck A - Quarterly Review"
    title.placeholders[1].text = "Prepared by the Finance Team"
    body = prs.slides.add_slide(prs.slide_layouts[1])
    body.shapes.title.text = "Highlights"
    tf = body.placeholders[1].text_frame
    tf.text = "Revenue grew 12% year over year"
    for line in ("New market entered in March", "Churn down to 1.8%"):
        p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(18)
    prs.save(path)
    return 2


def build_deck_b(path: str, photo_path: str) -> int:
    prs = Presentation()
    title = prs.slides.add_slide(prs.slide_layouts[0])
    title.shapes.title.text = "Deck B - Product Photos"
    title.placeholders[1].text = "Appendix"
    blank = prs.slides.add_slide(prs.slide_layouts[6])
    blank.shapes.add_picture(photo_path, Inches(0.5), Inches(0.5), width=Inches(9))
    prs.save(path)
    return 2


def slide_count(path: str) -> int:
    return len(Presentation(path).slides._sldIdLst)


def main() -> int:
    workdir = sys.argv[1] if len(sys.argv) > 1 else "/tmp/pptx-toolkit-selftest"
    os.makedirs(workdir, exist_ok=True)
    photo = os.path.join(workdir, "big-photo.png")
    deck_a = os.path.join(workdir, "deck-a.pptx")
    deck_b = os.path.join(workdir, "deck-b.pptx")
    merged = os.path.join(workdir, "merged.pptx")
    compressed = os.path.join(workdir, "deck-b-compressed.pptx")

    print(f"Workdir: {workdir}")
    make_big_photo(photo)
    print(f"Generated test photo: {photo} ({os.path.getsize(photo):,} bytes)")

    n_a = build_deck_a(deck_a)
    n_b = build_deck_b(deck_b, photo)
    print(f"Built deck A ({n_a} slides, {os.path.getsize(deck_a):,} bytes)")
    print(f"Built deck B ({n_b} slides, {os.path.getsize(deck_b):,} bytes)")

    # --- merge ---
    total = merge_presentations([deck_a, deck_b], merged)
    check("merge returns summed slide count", total == n_a + n_b, f"{total} == {n_a + n_b}")
    reopened = slide_count(merged)
    check("merged deck re-opens in python-pptx", reopened == n_a + n_b, f"{reopened} slides")

    # The picture from deck B must have survived the merge.
    pics = sum(
        1
        for slide in Presentation(merged).slides
        for shape in slide.shapes
        if shape.shape_type == 13  # PICTURE
    )
    check("merged deck contains the copied picture", pics >= 1, f"{pics} picture(s)")

    # --- compress ---
    before, after, count = compress_presentation(deck_b, compressed)
    check("compress re-encoded at least one image", count >= 1, f"{count} image(s)")
    check("compressed deck is smaller", after < before, f"{before:,} -> {after:,} bytes")
    reopened_c = slide_count(compressed)
    check("compressed deck re-opens in python-pptx", reopened_c == n_b, f"{reopened_c} slides")

    print(
        f"SUMMARY merge: {n_a}+{n_b}={total} slides | "
        f"compress: {before:,} -> {after:,} bytes ({(before - after) / before * 100:.1f}% saved)"
    )
    print("ALL CHECKS PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
