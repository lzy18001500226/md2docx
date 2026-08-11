from __future__ import annotations

import argparse
from pathlib import Path

from .pipeline import build_document, load_profile


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a verified DOCX from Markdown.")
    subcommands = parser.add_subparsers(dest="command", required=True)
    build = subcommands.add_parser("build", help="convert Markdown into a checked DOCX")
    build.add_argument("source", type=Path)
    build.add_argument("--profile", type=Path, required=True)
    build.add_argument("--out", type=Path, required=True)
    build.add_argument("--overwrite", action="store_true")
    build.add_argument("--word", action="store_true", help="update fields through Microsoft Word")
    build.add_argument("--pdf", type=Path, help="export PDF during the optional Word pass")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    profile = load_profile(args.profile)
    result = build_document(
        args.source,
        profile,
        args.out,
        overwrite=args.overwrite,
        use_word=args.word or args.pdf is not None,
        pdf_output=args.pdf,
    )
    print(f"Built {result.output}")
    print(f"Captions: {result.caption_count}; images: {result.image_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
