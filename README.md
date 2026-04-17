# shuck-convert

Convert between document formats — DOCX/PDF to Markdown, Markdown to DOCX.

An MCP server for AI agents, with CLI support.

## Tools

| Tool | Direction | Description |
|------|-----------|-------------|
| `doc_to_markdown` | DOCX/PDF → MD | Extract text, formatting, tables, images |
| `markdown_to_docx` | MD → DOCX | Academic styling: Times New Roman + SimSun, three-line tables |

## Prerequisites

- Python 3.10+
- [Pandoc](https://pandoc.org/installing.html) (required for `markdown_to_docx`)

## Install

```bash
pip install shuck-convert
```

Or from source:

```bash
git clone https://github.com/Shan-Zhu/shuck-convert.git
cd shuck-convert
pip install -e .
```

## Usage

### MCP Server (for AI agents)

Add to your MCP client config:

```json
{
  "mcpServers": {
    "shuck-convert": {
      "command": "shuck-convert",
      "args": [],
      "transportType": "stdio"
    }
  }
}
```

Or with uvx:

```json
{
  "mcpServers": {
    "shuck-convert": {
      "command": "uvx",
      "args": ["shuck-convert"],
      "transportType": "stdio"
    }
  }
}
```

### CLI

```bash
# DOCX/PDF to Markdown
shuck-convert doc2md report.docx
shuck-convert doc2md paper.pdf

# Markdown to DOCX
shuck-convert md2docx paper.md
```

### Development

```bash
pip install -e .
fastmcp dev src/shuck_convert/server.py
```

## License

MIT
