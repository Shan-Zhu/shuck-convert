"""CLI entry point for shuck-convert."""

import argparse
import sys


def main():
    parser = argparse.ArgumentParser(
        prog="shuck-convert",
        description="Convert between document formats — DOCX/PDF to Markdown, Markdown to DOCX.",
    )
    subparsers = parser.add_subparsers(dest="command")

    p_doc = subparsers.add_parser("doc2md", help="Convert DOCX/PDF to Markdown")
    p_doc.add_argument("file", help="Path to .docx or .pdf file")

    p_md = subparsers.add_parser("md2docx", help="Convert Markdown to DOCX")
    p_md.add_argument("file", help="Path to .md file")

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "doc2md":
        from .core.doc_to_md import convert_doc_to_markdown
        try:
            result = convert_doc_to_markdown(args.file)
            print(f"Output: {result['output_path']}")
            if result["image_count"] > 0:
                print(f"Images: {result['image_count']} extracted to {result['image_dir']}")
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

    elif args.command == "md2docx":
        from .core.md_to_docx import convert_markdown_to_docx
        try:
            result = convert_markdown_to_docx(args.file)
            print(f"Output: {result['output_path']}")
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
