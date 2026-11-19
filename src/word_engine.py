"""
Core Word manipulation engine.

Manages Word documents using python-docx (offline .docx manipulation)
and optionally pywin32 COM automation (live Word application control on Windows).
"""

import os
import re
import sys
from typing import Optional

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as PP_ALIGN, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

from utils.helpers import normalize_path, ensure_docx_extension, is_windows


class WordEngine:
    """Core engine for Word document manipulation."""

    def __init__(self):
        self._documents: dict[str, Document] = {}  # path -> Document
        self._active_doc: Optional[str] = None
        self._com_app = None  # COM Word.Application (Windows only)

    # ── Helpers ──────────────────────────────────────────────────────

    def _get_doc(self, path: Optional[str] = None) -> tuple[str, Document]:
        """Get a document by path, or the active document."""
        if path:
            path = normalize_path(path)
            if path not in self._documents:
                raise ValueError(f"Document not open: {path}")
            return path, self._documents[path]
        if self._active_doc and self._active_doc in self._documents:
            return self._active_doc, self._documents[self._active_doc]
        raise ValueError("No active document. Open or create a document first.")

    def _resolve_alignment(self, alignment: str) -> PP_ALIGN:
        mapping = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY,
        }
        val = mapping.get(alignment.lower())
        if val is None:
            raise ValueError(f"Unknown alignment: {alignment}. Use: left, center, right, justify")
        return val

    # ── Document Management ──────────────────────────────────────────

    def create_document(self, path: Optional[str] = None) -> str:
        """Create a new blank Word document."""
        doc = Document()
        if path:
            path = normalize_path(ensure_docx_extension(path))
        else:
            import tempfile
            fd, path = tempfile.mkstemp(suffix=".docx")
            os.close(fd)
            path = normalize_path(path)
        doc.save(path)
        self._documents[path] = doc
        self._active_doc = path
        return path

    def open_document(self, path: str) -> str:
        """Open an existing Word document."""
        path = normalize_path(path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        doc = Document(path)
        self._documents[path] = doc
        self._active_doc = path
        return path

    def save_document(self, path: Optional[str] = None) -> str:
        """Save the document."""
        doc_path, doc = self._get_doc(path)
        doc.save(doc_path)
        return doc_path

    def save_document_as(self, new_path: str, source_path: Optional[str] = None) -> str:
        """Save the document to a new path."""
        old_path, doc = self._get_doc(source_path)
        new_path = normalize_path(ensure_docx_extension(new_path))
        os.makedirs(os.path.dirname(new_path), exist_ok=True)
        doc.save(new_path)
        self._documents[new_path] = doc
        if old_path in self._documents:
            del self._documents[old_path]
        self._active_doc = new_path
        return new_path

    def close_document(self, path: Optional[str] = None) -> str:
        """Close a document."""
        doc_path, _ = self._get_doc(path)
        del self._documents[doc_path]
        if self._active_doc == doc_path:
            self._active_doc = next(iter(self._documents), None)
        return doc_path

    def list_documents(self) -> list[dict]:
        """List all open documents."""
        result = []
        for p in self._documents:
            result.append({
                "path": p,
                "active": p == self._active_doc,
                "paragraphs": len(self._documents[p].paragraphs),
                "tables": len(self._documents[p].tables),
            })
        return result

    # ── Content Operations ───────────────────────────────────────────

    def read_document(self, path: Optional[str] = None) -> dict:
        """Read the full document content."""
        doc_path, doc = self._get_doc(path)
        paragraphs = []
        for i, para in enumerate(doc.paragraphs):
            paragraphs.append({
                "index": i,
                "text": para.text,
                "style": para.style.name if para.style else None,
            })
        tables_info = []
        for i, table in enumerate(doc.tables):
            rows_data = []
            for row in table.rows:
                rows_data.append([cell.text for cell in row.cells])
            tables_info.append({"index": i, "rows": len(table.rows), "cols": len(table.columns), "data": rows_data})
        return {
            "path": doc_path,
            "paragraphs": paragraphs,
            "tables": tables_info,
            "sections": len(doc.sections),
        }

    def read_paragraph(self, index: int, path: Optional[str] = None) -> dict:
        """Read a specific paragraph by index."""
        doc_path, doc = self._get_doc(path)
        if index < 0 or index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {index} out of range (0-{len(doc.paragraphs)-1})")
        para = doc.paragraphs[index]
        runs = []
        for run in para.runs:
            runs.append({
                "text": run.text,
                "bold": run.bold,
                "italic": run.italic,
                "underline": run.underline,
                "font_name": run.font.name,
                "font_size": str(run.font.size) if run.font.size else None,
            })
        return {
            "index": index,
            "text": para.text,
            "style": para.style.name if para.style else None,
            "alignment": str(para.alignment) if para.alignment else None,
            "runs": runs,
        }

    def add_paragraph(self, text: str, style: Optional[str] = None, path: Optional[str] = None) -> dict:
        """Add a new paragraph at the end of the document."""
        doc_path, doc = self._get_doc(path)
        para = doc.add_paragraph(text, style=style)
        idx = len(doc.paragraphs) - 1
        return {"index": idx, "text": text, "style": style}

    def add_heading(self, text: str, level: int = 1, path: Optional[str] = None) -> dict:
        """Add a heading."""
        doc_path, doc = self._get_doc(path)
        if level < 0 or level > 9:
            raise ValueError("Heading level must be 0-9")
        doc.add_heading(text, level=level)
        idx = len(doc.paragraphs) - 1
        return {"index": idx, "text": text, "level": level}

    def insert_text(self, text: str, paragraph_index: int, position: int = -1, path: Optional[str] = None) -> dict:
        """Insert text into a specific paragraph at a position."""
        doc_path, doc = self._get_doc(path)
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {paragraph_index} out of range")
        para = doc.paragraphs[paragraph_index]
        if position == -1 or position >= len(para.text):
            para.add_run(text)
        else:
            # Insert at position by manipulating runs
            current_pos = 0
            for run in para.runs:
                run_end = current_pos + len(run.text)
                if current_pos <= position < run_end:
                    offset = position - current_pos
                    run.text = run.text[:offset] + text + run.text[offset:]
                    break
                current_pos = run_end
            else:
                para.add_run(text)
        return {"paragraph_index": paragraph_index, "inserted_text": text}

    def delete_paragraph(self, index: int, path: Optional[str] = None) -> dict:
        """Delete a paragraph by index."""
        doc_path, doc = self._get_doc(path)
        if index < 0 or index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {index} out of range")
        para = doc.paragraphs[index]
        deleted_text = para.text
        p_element = para._element
        p_element.getparent().remove(p_element)
        return {"deleted_index": index, "deleted_text": deleted_text}

    def find_and_replace(self, find_text: str, replace_text: str, path: Optional[str] = None) -> dict:
        """Find and replace text in the document."""
        doc_path, doc = self._get_doc(path)
        count = 0
        for para in doc.paragraphs:
            if find_text in para.text:
                for run in para.runs:
                    if find_text in run.text:
                        run.text = run.text.replace(find_text, replace_text)
                        count += 1
        # Also check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if find_text in para.text:
                            for run in para.runs:
                                if find_text in run.text:
                                    run.text = run.text.replace(find_text, replace_text)
                                    count += 1
        return {"find": find_text, "replace": replace_text, "replacements": count}

    def search_text(self, search_text: str, path: Optional[str] = None) -> list[dict]:
        """Search for text occurrences in the document."""
        doc_path, doc = self._get_doc(path)
        results = []
        for i, para in enumerate(doc.paragraphs):
            if search_text.lower() in para.text.lower():
                # Find all positions
                text_lower = para.text.lower()
                search_lower = search_text.lower()
                start = 0
                while True:
                    pos = text_lower.find(search_lower, start)
                    if pos == -1:
                        break
                    results.append({
                        "paragraph_index": i,
                        "position": pos,
                        "context": para.text[max(0, pos-20):pos+len(search_text)+20],
                    })
                    start = pos + 1
        return results

    def add_page_break(self, path: Optional[str] = None) -> dict:
        """Add a page break."""
        doc_path, doc = self._get_doc(path)
        doc.add_page_break()
        return {"status": "page_break_added"}

    # ── Formatting ───────────────────────────────────────────────────

    def format_text(
        self,
        paragraph_index: int,
        run_index: int = 0,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        strikethrough: Optional[bool] = None,
        font_name: Optional[str] = None,
        font_size: Optional[float] = None,
        font_color: Optional[str] = None,
        highlight: Optional[bool] = None,
        path: Optional[str] = None,
    ) -> dict:
        """Apply formatting to a text run."""
        doc_path, doc = self._get_doc(path)
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {paragraph_index} out of range")
        para = doc.paragraphs[paragraph_index]
        if not para.runs:
            # Create a run if none exist
            para.add_run(para.text)
        if run_index < 0 or run_index >= len(para.runs):
            raise IndexError(f"Run index {run_index} out of range (0-{len(para.runs)-1})")
        run = para.runs[run_index]
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline
        if strikethrough is not None:
            run.font.strike = strikethrough
        if font_name is not None:
            run.font.name = font_name
        if font_size is not None:
            run.font.size = Pt(font_size)
        if font_color is not None:
            # Accept hex color like "#FF0000" or "FF0000"
            color = font_color.lstrip("#")
            run.font.color.rgb = RGBColor(
                int(color[0:2], 16),
                int(color[2:4], 16),
                int(color[4:6], 16),
            )
        return {"paragraph_index": paragraph_index, "run_index": run_index, "formatted": True}

    def set_paragraph_alignment(self, paragraph_index: int, alignment: str, path: Optional[str] = None) -> dict:
        """Set paragraph alignment."""
        doc_path, doc = self._get_doc(path)
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {paragraph_index} out of range")
        doc.paragraphs[paragraph_index].alignment = self._resolve_alignment(alignment)
        return {"paragraph_index": paragraph_index, "alignment": alignment}

    def set_paragraph_spacing(
        self,
        paragraph_index: int,
        space_before: Optional[float] = None,
        space_after: Optional[float] = None,
        line_spacing: Optional[float] = None,
        path: Optional[str] = None,
    ) -> dict:
        """Set paragraph spacing in points."""
        doc_path, doc = self._get_doc(path)
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {paragraph_index} out of range")
        pf = doc.paragraphs[paragraph_index].paragraph_format
        if space_before is not None:
            pf.space_before = Pt(space_before)
        if space_after is not None:
            pf.space_after = Pt(space_after)
        if line_spacing is not None:
            pf.line_spacing = Pt(line_spacing)
        return {"paragraph_index": paragraph_index, "spacing_set": True}

    # ── Tables ───────────────────────────────────────────────────────

    def create_table(self, rows: int, cols: int, data: Optional[list[list[str]]] = None, style: Optional[str] = None, path: Optional[str] = None) -> dict:
        """Create a table."""
        doc_path, doc = self._get_doc(path)
        table = doc.add_table(rows=rows, cols=cols)
        if style:
            table.style = style
        if data:
            for r_idx, row_data in enumerate(data):
                if r_idx >= rows:
                    break
                for c_idx, cell_text in enumerate(row_data):
                    if c_idx >= cols:
                        break
                    table.cell(r_idx, c_idx).text = str(cell_text)
        table_index = len(doc.tables) - 1
        return {"table_index": table_index, "rows": rows, "cols": cols}

    def read_table(self, table_index: int, path: Optional[str] = None) -> dict:
        """Read table data."""
        doc_path, doc = self._get_doc(path)
        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"Table index {table_index} out of range")
        table = doc.tables[table_index]
        data = []
        for row in table.rows:
            data.append([cell.text for cell in row.cells])
        return {"table_index": table_index, "rows": len(table.rows), "cols": len(table.columns), "data": data}

    def add_table_row(self, table_index: int, data: Optional[list[str]] = None, path: Optional[str] = None) -> dict:
        """Add a row to a table."""
        doc_path, doc = self._get_doc(path)
        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"Table index {table_index} out of range")
        table = doc.tables[table_index]
        row = table.add_row()
        if data:
            for i, text in enumerate(data):
                if i < len(row.cells):
                    row.cells[i].text = str(text)
        return {"table_index": table_index, "new_row_count": len(table.rows)}

    def add_table_column(self, table_index: int, data: Optional[list[str]] = None, path: Optional[str] = None) -> dict:
        """Add a column to a table."""
        doc_path, doc = self._get_doc(path)
        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"Table index {table_index} out of range")
        table = doc.tables[table_index]
        # python-docx doesn't have a direct add_column, so we manipulate XML
        for i, row in enumerate(table.rows):
            new_cell = row._tr.add_tc()
            if data and i < len(data):
                from docx.opc.constants import RELATIONSHIP_TYPE as RT
                cell = row.cells[-1]
                cell.text = str(data[i])
        return {"table_index": table_index, "new_col_count": len(table.columns)}

    def write_table_cell(self, table_index: int, row: int, col: int, text: str, path: Optional[str] = None) -> dict:
        """Write text to a specific table cell."""
        doc_path, doc = self._get_doc(path)
        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"Table index {table_index} out of range")
        table = doc.tables[table_index]
        if row < 0 or row >= len(table.rows):
            raise IndexError(f"Row {row} out of range")
        if col < 0 or col >= len(table.columns):
            raise IndexError(f"Column {col} out of range")
        table.cell(row, col).text = text
        return {"table_index": table_index, "row": row, "col": col, "text": text}

    def style_table(self, table_index: int, style_name: str, path: Optional[str] = None) -> dict:
        """Apply a style to a table."""
        doc_path, doc = self._get_doc(path)
        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"Table index {table_index} out of range")
        doc.tables[table_index].style = style_name
        return {"table_index": table_index, "style": style_name}

    # ── Images ───────────────────────────────────────────────────────

    def insert_image(self, image_path: str, width_inches: Optional[float] = None, height_inches: Optional[float] = None, path: Optional[str] = None) -> dict:
        """Insert an image into the document."""
        doc_path, doc = self._get_doc(path)
        image_path = normalize_path(image_path)
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Image not found: {image_path}")
        kwargs = {}
        if width_inches:
            kwargs["width"] = Inches(width_inches)
        if height_inches:
            kwargs["height"] = Inches(height_inches)
        doc.add_picture(image_path, **kwargs)
        return {"image_path": image_path, "inserted": True}

    # ── Headers & Footers ────────────────────────────────────────────

    def add_header(self, text: str, section_index: int = 0, path: Optional[str] = None) -> dict:
        """Add or edit a header."""
        doc_path, doc = self._get_doc(path)
        if section_index < 0 or section_index >= len(doc.sections):
            raise IndexError(f"Section index {section_index} out of range")
        section = doc.sections[section_index]
        header = section.header
        header.is_linked_to_previous = False
        if header.paragraphs:
            header.paragraphs[0].text = text
        else:
            header.add_paragraph(text)
        return {"section_index": section_index, "header_text": text}

    def add_footer(self, text: str, section_index: int = 0, path: Optional[str] = None) -> dict:
        """Add or edit a footer."""
        doc_path, doc = self._get_doc(path)
        if section_index < 0 or section_index >= len(doc.sections):
            raise IndexError(f"Section index {section_index} out of range")
        section = doc.sections[section_index]
        footer = section.footer
        footer.is_linked_to_previous = False
        if footer.paragraphs:
            footer.paragraphs[0].text = text
        else:
            footer.add_paragraph(text)
        return {"section_index": section_index, "footer_text": text}

    # ── Styles ───────────────────────────────────────────────────────

    def apply_style(self, paragraph_index: int, style_name: str, path: Optional[str] = None) -> dict:
        """Apply a style to a paragraph."""
        doc_path, doc = self._get_doc(path)
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            raise IndexError(f"Paragraph index {paragraph_index} out of range")
        doc.paragraphs[paragraph_index].style = style_name
        return {"paragraph_index": paragraph_index, "style": style_name}

    def list_styles(self, path: Optional[str] = None) -> list[dict]:
        """List available styles in the document."""
        doc_path, doc = self._get_doc(path)
        styles = []
        for style in doc.styles:
            styles.append({
                "name": style.name,
                "type": str(style.type),
                "builtin": style.builtin,
            })
        return styles

    # ── Export ────────────────────────────────────────────────────────

    def export_to_pdf(self, output_path: str, source_path: Optional[str] = None) -> dict:
        """Export document to PDF (requires Word on Windows via COM)."""
        doc_path, doc = self._get_doc(source_path)
        output_path = normalize_path(output_path)
        if not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"
        # Save first
        doc.save(doc_path)
        if is_windows():
            try:
                import win32com.client
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_doc = word_app.Documents.Open(doc_path)
                word_doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
                word_doc.Close()
                word_app.Quit()
                return {"pdf_path": output_path, "exported": True}
            except Exception as e:
                return {"error": f"COM PDF export failed: {str(e)}. Ensure Microsoft Word is installed."}
        else:
            return {"error": "PDF export via COM is only available on Windows with Microsoft Word installed."}

    def export_to_html(self, output_path: str, source_path: Optional[str] = None) -> dict:
        """Export document to HTML."""
        doc_path, doc = self._get_doc(source_path)
        output_path = normalize_path(output_path)
        if not output_path.lower().endswith(".html"):
            output_path += ".html"
        # Build simple HTML from document content
        html_parts = [
            "<!DOCTYPE html>",
            "<html><head><meta charset='utf-8'><title>Document</title></head><body>",
        ]
        for para in doc.paragraphs:
            style = para.style.name if para.style else "Normal"
            if style.startswith("Heading"):
                try:
                    level = int(style.split()[-1])
                except (ValueError, IndexError):
                    level = 1
                html_parts.append(f"<h{level}>{para.text}</h{level}>")
            elif style == "Title":
                html_parts.append(f"<h1 class='title'>{para.text}</h1>")
            else:
                if para.text.strip():
                    html_parts.append(f"<p>{para.text}</p>")
        # Tables
        for table in doc.tables:
            html_parts.append("<table border='1' cellpadding='4' cellspacing='0'>")
            for row in table.rows:
                html_parts.append("<tr>")
                for cell in row.cells:
                    html_parts.append(f"<td>{cell.text}</td>")
                html_parts.append("</tr>")
            html_parts.append("</table>")
        html_parts.append("</body></html>")
        html_content = "\n".join(html_parts)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return {"html_path": output_path, "exported": True}

    # ── Page Setup ───────────────────────────────────────────────────

    def set_page_margins(
        self,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
        section_index: int = 0,
        path: Optional[str] = None,
    ) -> dict:
        """Set page margins in inches."""
        doc_path, doc = self._get_doc(path)
        if section_index < 0 or section_index >= len(doc.sections):
            raise IndexError(f"Section index {section_index} out of range")
        section = doc.sections[section_index]
        if top is not None:
            section.top_margin = Inches(top)
        if bottom is not None:
            section.bottom_margin = Inches(bottom)
        if left is not None:
            section.left_margin = Inches(left)
        if right is not None:
            section.right_margin = Inches(right)
        return {"section_index": section_index, "margins_set": True}

    def set_page_orientation(self, orientation: str, section_index: int = 0, path: Optional[str] = None) -> dict:
        """Set page orientation: 'portrait' or 'landscape'."""
        doc_path, doc = self._get_doc(path)
        if section_index < 0 or section_index >= len(doc.sections):
            raise IndexError(f"Section index {section_index} out of range")
        section = doc.sections[section_index]
        if orientation.lower() == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            # Swap dimensions
            new_width = section.page_height
            new_height = section.page_width
            section.page_width = new_width
            section.page_height = new_height
        elif orientation.lower() == "portrait":
            section.orientation = WD_ORIENT.PORTRAIT
            new_width = min(section.page_width, section.page_height)
            new_height = max(section.page_width, section.page_height)
            section.page_width = new_width
            section.page_height = new_height
        else:
            raise ValueError("Orientation must be 'portrait' or 'landscape'")
        return {"section_index": section_index, "orientation": orientation}

    def set_page_size(self, width_inches: float, height_inches: float, section_index: int = 0, path: Optional[str] = None) -> dict:
        """Set page size in inches."""
        doc_path, doc = self._get_doc(path)
        if section_index < 0 or section_index >= len(doc.sections):
            raise IndexError(f"Section index {section_index} out of range")
        section = doc.sections[section_index]
        section.page_width = Inches(width_inches)
        section.page_height = Inches(height_inches)
        return {"section_index": section_index, "width": width_inches, "height": height_inches}
