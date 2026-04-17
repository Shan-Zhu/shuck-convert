"""MCP Server for shuck-convert: bidirectional document format conversion."""

from pathlib import Path
from fastmcp import FastMCP

from . import __version__

mcp = FastMCP(
    "shuck-convert",
    instructions=(
        "Convert between document formats. "
        "DOCX/PDF to Markdown (with image extraction), "
        "or Markdown to styled DOCX (Times New Roman + SimSun, three-line tables). "
        f"v{__version__}"
    ),
)


@mcp.tool()
def doc_to_markdown(file_path: str) -> str:
    """Convert a DOCX or PDF file to Markdown.

    Extracts text, formatting, tables, and images.
    Output .md file and extracted images are saved next to the source file.

    Args:
        file_path: Absolute path to a .docx or .pdf file.
    """
    try:
        from .core.doc_to_md import convert_doc_to_markdown

        result = convert_doc_to_markdown(file_path)

        lines = [
            "## Conversion Complete",
            "",
            f"- **Source**: {Path(file_path).name}",
            f"- **Output**: {result['output_path']}",
        ]
        if result["image_count"] > 0:
            lines.append(f"- **Images**: {result['image_count']} extracted to {result['image_dir']}")

        # Include markdown preview (first 200 lines)
        md_lines = result["markdown"].splitlines()
        if len(md_lines) > 200:
            preview = "\n".join(md_lines[:200])
            lines.extend(["", f"### Preview (first 200 of {len(md_lines)} lines)", "", preview, "", "..."])
        else:
            lines.extend(["", "### Content", "", result["markdown"]])

        return "\n".join(lines)

    except Exception as e:
        return f"Error: {e}"


@mcp.tool()
def markdown_to_docx(file_path: str) -> str:
    """Convert a Markdown file to a styled DOCX document.

    Uses pandoc for conversion, then applies academic styling:
    Times New Roman (English) + SimSun/宋体 (Chinese), 12pt, double spacing,
    justified alignment, three-line tables. Handles Obsidian image syntax
    (![[image|size]]) and footnote preprocessing.

    Requires pandoc installed (https://pandoc.org/installing.html).

    Args:
        file_path: Absolute path to a .md file.
    """
    try:
        from .core.md_to_docx import convert_markdown_to_docx

        result = convert_markdown_to_docx(file_path)

        return "\n".join([
            "## Conversion Complete",
            "",
            f"- **Source**: {Path(file_path).name}",
            f"- **Output**: {result['output_path']}",
            "",
            "Styling applied: Times New Roman + SimSun, 12pt, double spacing, three-line tables.",
        ])

    except Exception as e:
        return f"Error: {e}"


def main():
    mcp.run()


if __name__ == "__main__":
    main()
