"""DOCX/PDF → Markdown conversion core logic.

Extracted from doc2md/convert.py, stripped of Windows UI (toast notifications).
"""

from pathlib import Path


# ── Image helpers ────────────────────────────────────────────────────────

_CONTENT_TYPE_EXT = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/tiff": ".tiff",
    "image/svg+xml": ".svg",
    "image/x-emf": ".emf",
    "image/x-wmf": ".wmf",
}


def _ext_from_content_type(content_type: str) -> str:
    return _CONTENT_TYPE_EXT.get(content_type, ".png")


# ── Word (.docx) conversion ─────────────────────────────────────────────


def convert_docx(filepath: Path, img_dir: Path, img_folder: str) -> str:
    from docx import Document
    from docx.oxml.ns import qn  # noqa: F811

    doc = Document(str(filepath))
    lines: list[str] = []
    img_counter = [0]

    for element in doc.element.body:
        tag = element.tag.split("}")[-1]

        if tag == "p":
            lines.append(_convert_paragraph(element, doc, img_dir, img_folder, img_counter))
        elif tag == "tbl":
            lines.append(_convert_table(element, doc))

    return "\n".join(lines).strip() + "\n"


def _extract_paragraph_images(element, doc, img_dir: Path, img_folder: str, img_counter: list) -> list[str]:
    """Extract all images from a paragraph, save to disk, return markdown refs."""
    from docx.oxml.ns import qn

    images = []
    # Inline images via w:drawing
    for drawing in element.findall(f".//{qn('w:drawing')}"):
        for blip in drawing.findall(f".//{qn('a:blip')}"):
            rId = blip.get(qn("r:embed"))
            if not rId:
                continue
            try:
                rel = doc.part.rels[rId]
                img_data = rel.target_part.blob
                ext = _ext_from_content_type(rel.target_part.content_type)
                img_counter[0] += 1
                img_name = f"image_{img_counter[0]:03d}{ext}"
                img_dir.mkdir(parents=True, exist_ok=True)
                (img_dir / img_name).write_bytes(img_data)
                images.append(f"![{img_name}]({img_folder}/{img_name})")
            except (KeyError, AttributeError):
                continue

    # Legacy VML images via w:pict > v:imagedata
    for imagedata in element.findall(".//{urn:schemas-microsoft-com:vml}imagedata"):
        rId = imagedata.get(qn("r:id"))
        if not rId:
            continue
        try:
            rel = doc.part.rels[rId]
            img_data = rel.target_part.blob
            ext = _ext_from_content_type(rel.target_part.content_type)
            img_counter[0] += 1
            img_name = f"image_{img_counter[0]:03d}{ext}"
            img_dir.mkdir(parents=True, exist_ok=True)
            (img_dir / img_name).write_bytes(img_data)
            images.append(f"![{img_name}]({img_folder}/{img_name})")
        except (KeyError, AttributeError):
            continue

    return images


def _run_text(run) -> str:
    """Convert a single run to Markdown with bold/italic."""
    from docx.oxml.ns import qn

    text = run.text or ""
    if not text.strip():
        return text

    rpr = run.find(qn("w:rPr"))
    bold = False
    italic = False
    if rpr is not None:
        bold = rpr.find(qn("w:b")) is not None and rpr.find(qn("w:b")).get(qn("w:val"), "true") != "false"
        italic = rpr.find(qn("w:i")) is not None and rpr.find(qn("w:i")).get(qn("w:val"), "true") != "false"

    if bold and italic:
        return f"***{text}***"
    elif bold:
        return f"**{text}**"
    elif italic:
        return f"*{text}*"
    return text


def _paragraph_text(element) -> str:
    """Extract full text from a paragraph element with inline formatting."""
    from docx.oxml.ns import qn

    parts = []
    for run in element.findall(qn("w:r")):
        parts.append(_run_text(run))
    return "".join(parts)


def _get_style_name(element, doc) -> str:
    """Get the style name of a paragraph element."""
    from docx.oxml.ns import qn

    ppr = element.find(qn("w:pPr"))
    if ppr is not None:
        style_el = ppr.find(qn("w:pStyle"))
        if style_el is not None:
            return style_el.get(qn("w:val"), "")
    return ""


def _get_numpr(element):
    """Get numbering properties (list info) from a paragraph."""
    from docx.oxml.ns import qn

    ppr = element.find(qn("w:pPr"))
    if ppr is not None:
        return ppr.find(qn("w:numPr"))
    return None


def _convert_paragraph(element, doc, img_dir: Path, img_folder: str, img_counter: list) -> str:
    from docx.oxml.ns import qn  # noqa: F811

    style = _get_style_name(element, doc).lower()
    text = _paragraph_text(element)

    # Extract images from this paragraph
    img_refs = _extract_paragraph_images(element, doc, img_dir, img_folder, img_counter)

    # Headings
    for level in range(1, 7):
        if style in (f"heading{level}", f"heading {level}"):
            result = f"\n{'#' * level} {text}\n"
            if img_refs:
                result += "\n" + "\n\n".join(img_refs) + "\n"
            return result

    # Title / Subtitle
    if style == "title":
        result = f"\n# {text}\n"
        if img_refs:
            result += "\n" + "\n\n".join(img_refs) + "\n"
        return result
    if style == "subtitle":
        result = f"\n## {text}\n"
        if img_refs:
            result += "\n" + "\n\n".join(img_refs) + "\n"
        return result

    # Lists
    numpr = _get_numpr(element)
    if numpr is not None or "list" in style:
        indent_level = 0
        if numpr is not None:
            ilvl = numpr.find(qn("w:ilvl"))
            if ilvl is not None:
                indent_level = int(ilvl.get(qn("w:val"), "0"))
        prefix = "  " * indent_level + "- "
        result = prefix + text
        if img_refs:
            result += "\n" + "\n\n".join(img_refs)
        return result

    # Image-only paragraph (no text, just images)
    if not text.strip() and img_refs:
        return "\n".join(img_refs) + "\n"

    # Normal paragraph
    if not text.strip():
        return ""

    result = text + "\n"
    if img_refs:
        result += "\n" + "\n\n".join(img_refs) + "\n"
    return result


def _convert_table(element, doc) -> str:
    from docx.oxml.ns import qn

    rows = element.findall(qn("w:tr"))
    if not rows:
        return ""

    table_data: list[list[str]] = []
    for tr in rows:
        cells = []
        for tc in tr.findall(qn("w:tc")):
            cell_texts = []
            for p in tc.findall(qn("w:p")):
                cell_texts.append(_paragraph_text(p).strip())
            cells.append(" ".join(cell_texts).replace("|", "\\|"))
        table_data.append(cells)

    if not table_data:
        return ""

    # Normalize column count
    max_cols = max(len(row) for row in table_data)
    for row in table_data:
        while len(row) < max_cols:
            row.append("")

    md_lines = ["\n"]
    # Header row
    md_lines.append("| " + " | ".join(table_data[0]) + " |")
    md_lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    # Data rows
    for row in table_data[1:]:
        md_lines.append("| " + " | ".join(row) + " |")
    md_lines.append("")

    return "\n".join(md_lines)


# ── PDF conversion ───────────────────────────────────────────────────────


def convert_pdf(filepath: Path, img_dir: Path, img_folder: str) -> str:
    import pdfplumber
    try:
        import fitz  # PyMuPDF — used for image extraction
        has_fitz = True
    except ImportError:
        has_fitz = False

    lines: list[str] = []
    img_counter = 0

    # Extract images with PyMuPDF if available
    page_images: dict[int, list[str]] = {}
    if has_fitz:
        fitz_doc = fitz.open(str(filepath))
        for page_num in range(len(fitz_doc)):
            page = fitz_doc[page_num]
            page_imgs = []
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                try:
                    base_image = fitz_doc.extract_image(xref)
                    if not base_image or not base_image.get("image"):
                        continue
                    img_bytes = base_image["image"]
                    img_ext = "." + base_image.get("ext", "png")
                    img_counter += 1
                    img_name = f"image_{img_counter:03d}{img_ext}"
                    img_dir.mkdir(parents=True, exist_ok=True)
                    (img_dir / img_name).write_bytes(img_bytes)
                    page_imgs.append(f"![{img_name}]({img_folder}/{img_name})")
                except Exception:
                    continue
            if page_imgs:
                page_images[page_num] = page_imgs
        fitz_doc.close()

    # Extract text with pdfplumber
    with pdfplumber.open(str(filepath)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                lines.append(text)
            # Append images for this page
            if i in page_images:
                lines.append("\n".join(page_images[i]))
            # Page separator
            if i < len(pdf.pages) - 1:
                lines.append("\n---\n")

    return "\n\n".join(lines).strip() + "\n"


# ── Public API ───────────────────────────────────────────────────────────


def convert_doc_to_markdown(file_path: str) -> dict:
    """Convert a DOCX or PDF file to Markdown.

    Returns:
        dict with keys: markdown, output_path, image_dir (str|None), image_count (int)
    """
    src = Path(file_path).resolve()
    if not src.exists():
        raise FileNotFoundError(f"File not found: {src}")

    ext = src.suffix.lower()
    if ext not in (".docx", ".pdf"):
        raise ValueError(f"Unsupported format: {ext}. Supported: .docx, .pdf")

    # Output path — same directory as source
    out_dir = src.parent
    out_path = out_dir / (src.stem + ".md")

    # Handle name collision
    counter = 1
    while out_path.exists():
        out_path = out_dir / f"{src.stem} ({counter}).md"
        counter += 1

    # Image output folder
    img_folder_name = src.stem + "_images"
    img_dir = out_dir / img_folder_name

    if ext == ".docx":
        md = convert_docx(src, img_dir, img_folder_name)
    else:
        md = convert_pdf(src, img_dir, img_folder_name)

    out_path.write_text(md, encoding="utf-8")

    img_count = len(list(img_dir.glob("*"))) if img_dir.exists() else 0

    return {
        "markdown": md,
        "output_path": str(out_path),
        "image_dir": str(img_dir) if img_count > 0 else None,
        "image_count": img_count,
    }
