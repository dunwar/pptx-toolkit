# pptx-toolkit

A tiny command-line toolkit for PowerPoint files: **merge** several `.pptx` decks into one, and **compress** a deck by shrinking the images embedded in it.

No GUI, no uploads, no account — just two small commands that run locally on your own files.

## Features

- `merge` — combine two or more decks into a single file, slides in the order you list the inputs. Text, shapes and embedded pictures are carried over.
- `compress` — downscale oversized embedded images (default: nothing wider or taller than 1920 px) and re-encode them, which typically cuts photo-heavy decks by 80–97%. Layout, text and slide order are never touched.

## Installation

Requires Python 3.10 or newer.

```bash
pip install pptx-toolkit
```

Or from a clone of this repository:

```bash
git clone https://github.com/dunwar/pptx-toolkit.git
cd pptx-toolkit
pip install .
```

The only runtime dependencies are [python-pptx](https://python-pptx.readthedocs.io/) and [Pillow](https://pillow.readthedocs.io/).

## Usage

### Merge

```bash
pptx-toolkit merge part1.pptx part2.pptx part3.pptx -o combined.pptx
# Merged 3 files (24 slides) -> combined.pptx (3.1 MB)
```

The first file provides the theme, slide size and master layouts for the combined deck (see Limitations).

### Compress

```bash
pptx-toolkit compress keynote.pptx -o keynote-small.pptx
# Re-encoded 14 image(s)
# Before: 18.6 MB
# After:  2.3 MB
# Saved:  16.3 MB (87.6%)
```

Options:

```bash
pptx-toolkit compress deck.pptx -o deck-small.pptx --max-dimension 1600 --quality 75
```

- `--max-dimension N` — downscale any image whose width or height exceeds `N` pixels (default `1920`). Aspect ratio is preserved; images are never enlarged.
- `--quality 1-95` — JPEG quality used when re-encoding (default `80`).

PNG images with transparency stay PNG (downscaled and optimized); opaque PNGs, BMPs and TIFFs are converted to JPEG. Vector images (EMF/WMF/SVG) and GIFs are left as they are.

Run `pptx-toolkit --help`, `pptx-toolkit merge --help` or `pptx-toolkit compress --help` for the full reference.

## Prefer a no-code version?

Prefer no-code? The same tools run free in your browser at https://aippt.org/merge-ppt and https://aippt.org/compress-ppt — no sign-up, no watermark. Files are processed locally in the browser there, too.

## Limitations

- **Merge adopts the first deck's theme.** Copied slides are re-created on the first file's slide master and layouts (matched by layout position). A source deck built on a very different theme may come out with the first deck's fonts and colors, and slide size always follows the first file.
- **Exotic slide content may not survive a merge.** Animations, transitions, SmartArt, embedded video/audio and speaker notes are not guaranteed to copy; plain text, shapes, tables and pictures are.
- **Compress is lossy.** Re-encoding JPEGs and downscaling large images discards pixels. Keep your original if you need print-resolution images. Images are only ever replaced when the result is actually smaller — the output file never grows.
- Only `.pptx` (Office Open XML) files are supported — not legacy `.ppt`.

## Development

```bash
pip install -e .
python scripts/selftest.py   # builds sample decks and verifies merge + compress end to end
```

## Contributing

Issues and pull requests are welcome. Please keep the dependency list at python-pptx + Pillow, keep the two commands working exactly as documented, and run `scripts/selftest.py` before submitting — it must print `ALL CHECKS PASSED`. For anything bigger than a bug fix, opening an issue to discuss the approach first is appreciated.

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 AIPPT.org.
