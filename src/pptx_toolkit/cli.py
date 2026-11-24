"""Command-line interface: ``pptx-toolkit merge`` and ``pptx-toolkit compress``."""

from __future__ import annotations

import argparse
import os
import sys

from . import __version__
from .compress import compress_presentation
from .merge import merge_presentations


def _human(num_bytes: int) -> str:
    value = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{value:.1f} GB"


def _check_inputs(paths: list[str]) -> None:
    for path in paths:
        if not os.path.isfile(path):
            raise SystemExit(f"error: input file not found: {path}")
        if not path.lower().endswith(".pptx"):
            raise SystemExit(f"error: not a .pptx file: {path}")


def cmd_merge(args: argparse.Namespace) -> int:
    _check_inputs(args.inputs)
    total = merge_presentations(args.inputs, args.output)
    size = os.path.getsize(args.output)
    print(
        f"Merged {len(args.inputs)} files ({total} slides) -> {args.output} ({_human(size)})"
    )
    return 0


def cmd_compress(args: argparse.Namespace) -> int:
    _check_inputs([args.input])
    before, after, count = compress_presentation(
        args.input,
        args.output,
        max_dimension=args.max_dimension,
        quality=args.quality,
    )
    saved = before - after
    pct = (saved / before * 100) if before else 0.0
    print(f"Re-encoded {count} image(s)")
    print(f"Before: {_human(before)}")
    print(f"After:  {_human(after)}")
    print(f"Saved:  {_human(saved)} ({pct:.1f}%)")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-toolkit",
        description="Merge PowerPoint decks and compress the images inside them.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p_merge = sub.add_parser(
        "merge",
        help="Merge two or more .pptx files into one, in the order given.",
        description="Merge two or more .pptx files into one deck. Slides appear "
        "in the order the files are listed; the first file provides the theme "
        "and slide size.",
    )
    p_merge.add_argument("inputs", nargs="+", metavar="INPUT.pptx", help="Input decks, in merge order.")
    p_merge.add_argument("-o", "--output", required=True, help="Path of the merged .pptx to write.")
    p_merge.set_defaults(func=cmd_merge)

    p_compress = sub.add_parser(
        "compress",
        help="Compress a .pptx by downscaling and re-encoding its embedded images.",
        description="Compress a .pptx by downscaling oversized embedded images "
        "and re-encoding them (JPEG quality / PNG optimization). Layout, text "
        "and slide order are untouched.",
    )
    p_compress.add_argument("input", metavar="INPUT.pptx", help="The deck to compress.")
    p_compress.add_argument("-o", "--output", required=True, help="Path of the compressed .pptx to write.")
    p_compress.add_argument(
        "--max-dimension",
        type=int,
        default=1920,
        help="Downscale images whose width or height exceeds this many pixels (default: 1920).",
    )
    p_compress.add_argument(
        "--quality",
        type=int,
        default=80,
        choices=range(1, 96),
        metavar="1-95",
        help="JPEG quality for re-encoded images (default: 80).",
    )
    p_compress.set_defaults(func=cmd_compress)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:  # surface a clean error, not a traceback
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
