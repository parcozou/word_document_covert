"""Convert a Markdown report into a formatted DOCX with native Word tables."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import os
import re
from urllib.parse import urlparse

from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.shared import Inches, Pt, RGBColor
import markdown
from PIL import Image, UnidentifiedImageError
import requests


TABLE_WIDTH_IN = 6.5
TABLE_WIDTH_DXA = 9360
TABLE_INDENT_DXA = 120
DEFAULT_IMAGE_HOSTS = "lf-bot-studio-plugin-resource.coze.cn"
TABLE_DELIMITER_CELL = re.compile(r"^:?-{3,}:?$")


@dataclass
class ConversionResult:
    output_path: Path
    table_count: int = 0
    image_count: int = 0
    skipped_image_count: int = 0
    heading_count: int = 0


def _table_cells(line: str) -> list[str] | None:
    stripped = line.strip()
    if "|" not in stripped:
        return None
    cells = stripped.strip("|").split("|")
    return [cell.strip() for cell in cells]


def _is_table_delimiter(line: str) -> bool:
    cells = _table_cells(line)
    return bool(
        cells
        and len(cells) >= 2
        and all(TABLE_DELIMITER_CELL.fullmatch(cell) for cell in cells)
    )


def _normalise_table_boundaries(markdown_text: str) -> str:
    """Close LLM-written pipe tables even when a following blank line is missing."""
    lines = markdown_text.splitlines()
    normalised: list[str] = []
    index = 0
    while index < len(lines):
        starts_table = (
            index + 1 < len(lines)
            and _table_cells(lines[index]) is not None
            and _is_table_delimiter(lines[index + 1])
        )
        if not starts_table:
            normalised.append(lines[index])
            index += 1
            continue

        if normalised and normalised[-1].strip():
            normalised.append("")

        expected_columns = len(_table_cells(lines[index + 1]) or [])
        normalised.extend([lines[index], lines[index + 1]])
        index += 2
        while index < len(lines):
            cells = _table_cells(lines[index])
            if cells is None or len(cells) != expected_columns:
                break
            normalised.append(lines[index])
            index += 1

        if normalised[-1].strip():
            normalised.append("")

    return "\n".join(normalised)


def _normalise_horizontal_rules(markdown_text: str) -> str:
    """Prevent section dividers from converting the prior paragraph to an H2."""
    lines = markdown_text.splitlines()
    normalised: list[str] = []
    for index, line in enumerate(lines):
        if line.strip() == "---":
            if normalised and normalised[-1].strip():
                normalised.append("")
            normalised.append(line)
            if index + 1 < len(lines) and lines[index + 1].strip():
                normalised.append("")
            continue
        normalised.append(line)
    return "\n".join(normalised)


def _normalise_report_markdown(markdown_text: str) -> str:
    with_closed_tables = _normalise_table_boundaries(markdown_text)
    return _normalise_horizontal_rules(with_closed_tables)


def _safe_hosts() -> set[str]:
    configured = os.getenv("IMAGE_HOST_ALLOWLIST", DEFAULT_IMAGE_HOSTS)
    return {host.strip().lower() for host in configured.split(",") if host.strip()}


def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def _set_cell_margins(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in {"top": 80, "bottom": 80, "start": 120, "end": 120}.items():
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_table_geometry(table, widths: list[float]) -> None:
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    tbl_width = tbl_pr.find(qn("w:tblW"))
    if tbl_width is None:
        tbl_width = OxmlElement("w:tblW")
        tbl_pr.append(tbl_width)
    tbl_width.set(qn("w:type"), "dxa")
    tbl_width.set(qn("w:w"), str(TABLE_WIDTH_DXA))

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    indent = tbl_pr.find(qn("w:tblInd"))
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tbl_pr.append(indent)
    indent.set(qn("w:type"), "dxa")
    indent.set(qn("w:w"), str(TABLE_INDENT_DXA))

    for row in table.rows:
        for index, cell in enumerate(row.cells):
            width = Inches(widths[min(index, len(widths) - 1)])
            cell.width = width
            _set_cell_margins(cell)


def _column_widths(column_count: int) -> list[float]:
    if column_count <= 1:
        return [TABLE_WIDTH_IN]
    if column_count == 2:
        return [1.9, 4.6]
    if column_count == 3:
        return [1.55, 2.475, 2.475]
    if column_count == 4:
        return [1.55, 1.65, 1.65, 1.65]
    first = 1.55
    other = (TABLE_WIDTH_IN - first) / (column_count - 1)
    return [first] + [other] * (column_count - 1)


def _add_hyperlink(paragraph, text: str, url: str) -> None:
    relationship_id = paragraph.part.relate_to(
        url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend([color, underline])
    run.append(properties)
    text_element = OxmlElement("w:t")
    text_element.text = text
    run.append(text_element)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def _render_inline(paragraph, node, bold: bool = False, italic: bool = False) -> None:
    if isinstance(node, NavigableString):
        text = str(node)
        if text:
            run = paragraph.add_run(text)
            run.bold = bold
            run.italic = italic
        return
    if not isinstance(node, Tag):
        return
    if node.name in {"strong", "b"}:
        for child in node.children:
            _render_inline(paragraph, child, True, italic)
    elif node.name in {"em", "i"}:
        for child in node.children:
            _render_inline(paragraph, child, bold, True)
    elif node.name == "code":
        run = paragraph.add_run(node.get_text())
        run.font.name = "Consolas"
        run.font.size = Pt(9)
        run.bold = bold
        run.italic = italic
    elif node.name == "a":
        url = node.get("href", "")
        label = node.get_text() or url
        if url:
            _add_hyperlink(paragraph, label, url)
        else:
            paragraph.add_run(label)
    elif node.name == "br":
        paragraph.add_run().add_break()
    elif node.name != "img":
        for child in node.children:
            _render_inline(paragraph, child, bold, italic)


def _style_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.1

    for style_name, size, color, before, after in [
        ("Heading 1", 16, "2E74B5", 16, 8),
        ("Heading 2", 13, "2E74B5", 12, 6),
        ("Heading 3", 12, "1F4D78", 8, 4),
    ]:
        style = document.styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    header = section.header.paragraphs[0]
    header.text = "Corporate Valuation Analysis Report"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.runs[0].font.size = Pt(8)
    header.runs[0].font.color.rgb = RGBColor(100, 100, 100)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("Page ")
    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = "PAGE"
    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")
    footer.runs[-1]._r.extend([field_begin, instruction, field_end])
    for run in footer.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(100, 100, 100)


def _add_rule(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(7)
    p_pr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "D9E2F3")
    borders.append(bottom)
    p_pr.append(borders)


def _download_image(url: str) -> BytesIO | None:
    parsed = urlparse(url)
    if parsed.scheme != "https" or (parsed.hostname or "").lower() not in _safe_hosts():
        return None
    try:
        response = requests.get(
            url,
            timeout=(4, 15),
            stream=True,
            headers={"User-Agent": "Coze-DOCX-Report-Renderer/1.0"},
        )
        response.raise_for_status()
        if not response.headers.get("Content-Type", "").lower().startswith("image/"):
            return None
        maximum_bytes = 10 * 1024 * 1024
        content = bytearray()
        for chunk in response.iter_content(chunk_size=64 * 1024):
            content.extend(chunk)
            if len(content) > maximum_bytes:
                return None
        source_image = BytesIO(bytes(content))
        try:
            with Image.open(source_image) as image:
                image.load()
                normalized = BytesIO()
                image.save(normalized, format="PNG")
                normalized.seek(0)
                return normalized
        except (UnidentifiedImageError, OSError):
            return None
    except requests.RequestException:
        return None


def _render_image(
    document: Document, image_tag: Tag, include_images: bool, result: ConversionResult
) -> None:
    caption = image_tag.get("alt", "").strip() or "Chart"
    source = image_tag.get("src", "")
    content = _download_image(source) if include_images else None
    if content is None:
        result.skipped_image_count += 1
        note = document.add_paragraph()
        note.paragraph_format.space_after = Pt(4)
        run = note.add_run(f"[Figure unavailable: {caption}]")
        run.italic = True
        run.font.color.rgb = RGBColor(100, 100, 100)
        return
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(content, width=Inches(6.25))
    caption_para = document.add_paragraph(caption)
    caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_para.runs[0].italic = True
    caption_para.runs[0].font.size = Pt(9)
    result.image_count += 1


def _render_table(document: Document, source_table: Tag, result: ConversionResult) -> None:
    source_rows = source_table.find_all("tr")
    if not source_rows:
        return
    column_count = max(len(row.find_all(["th", "td"])) for row in source_rows)
    table = document.add_table(rows=len(source_rows), cols=column_count)
    table.style = "Table Grid"
    widths = _column_widths(column_count)
    _set_table_geometry(table, widths)

    for row_index, source_row in enumerate(source_rows):
        source_cells = source_row.find_all(["th", "td"])
        for column_index, source_cell in enumerate(source_cells):
            target = table.cell(row_index, column_index)
            paragraph = target.paragraphs[0]
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.paragraph_format.line_spacing = 1.0
            for child in source_cell.children:
                _render_inline(paragraph, child, bold=(row_index == 0))
            for run in paragraph.runs:
                run.font.size = Pt(9)
                if row_index == 0:
                    run.bold = True
            if row_index == 0:
                _set_cell_shading(target, "F2F4F7")

    after_table = document.add_paragraph()
    after_table.paragraph_format.space_after = Pt(3)
    result.table_count += 1


def _render_list(document: Document, list_tag: Tag, result: ConversionResult, level: int = 0) -> None:
    style = "List Number" if list_tag.name == "ol" else "List Bullet"
    for item in list_tag.find_all("li", recursive=False):
        paragraph = document.add_paragraph(style=style)
        if level:
            paragraph.paragraph_format.left_indent = Inches(0.25 * level)
        for child in item.children:
            if isinstance(child, Tag) and child.name in {"ul", "ol"}:
                continue
            _render_inline(paragraph, child)
        for child_list in item.find_all(["ul", "ol"], recursive=False):
            _render_list(document, child_list, result, level + 1)


def _render_block(document: Document, node: Tag, include_images: bool, result: ConversionResult) -> None:
    if node.name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
        level = min(int(node.name[1]), 3)
        document.add_heading(node.get_text(" ", strip=True), level=level)
        result.heading_count += 1
    elif node.name == "p":
        images = node.find_all("img")
        text_without_images = node.get_text(" ", strip=True)
        if text_without_images:
            paragraph = document.add_paragraph()
            for child in node.children:
                _render_inline(paragraph, child)
        for image in images:
            _render_image(document, image, include_images, result)
    elif node.name == "table":
        _render_table(document, node, result)
    elif node.name in {"ul", "ol"}:
        _render_list(document, node, result)
    elif node.name == "hr":
        _add_rule(document)
    elif node.name == "blockquote":
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.left_indent = Inches(0.25)
        run = paragraph.add_run(node.get_text(" ", strip=True))
        run.italic = True
    elif node.name == "pre":
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.left_indent = Inches(0.2)
        run = paragraph.add_run(node.get_text())
        run.font.name = "Consolas"
        run.font.size = Pt(9)


def convert_markdown_to_docx(
    markdown_text: str,
    title: str,
    output_path: Path,
    include_images: bool = True,
) -> ConversionResult:
    """Create a DOCX report; Markdown tables become editable Word tables."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    _style_document(document)
    document.core_properties.title = title
    document.core_properties.subject = "Generated from Coze report Markdown"

    html = markdown.markdown(
        _normalise_report_markdown(markdown_text),
        extensions=["tables", "fenced_code", "sane_lists"],
        output_format="html5",
    )
    body = BeautifulSoup(html, "html.parser")
    if title and body.find("h1") is None:
        document.add_heading(title, level=1)

    result = ConversionResult(output_path=output_path)
    for block in body.contents:
        if isinstance(block, Tag):
            _render_block(document, block, include_images, result)

    document.save(output_path)
    return result


def safe_file_stem(title: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 _-]+", "", title).strip().replace(" ", "_")
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned[:80] or "generated_report"
