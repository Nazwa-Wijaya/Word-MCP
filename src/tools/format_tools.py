"""Formatting tools: font, alignment, spacing."""

from typing import Optional


def register_format_tools(mcp, engine):
    @mcp.tool()
    def format_text(
        paragraph_index: int,
        run_index: int = 0,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font_name: str = "",
        font_size: float = 0,
        font_color: str = "",
        path: str = "",
    ) -> dict:
        """Apply font formatting to a text run within a paragraph.

        Args:
            paragraph_index: Zero-based paragraph index.
            run_index: Zero-based run index within the paragraph (default: 0).
            bold: Set bold (True/False) or leave unchanged (None).
            italic: Set italic (True/False) or leave unchanged (None).
            underline: Set underline (True/False) or leave unchanged (None).
            strikethrough: Set strikethrough (True/False) or leave unchanged (None).
            font_name: Font name (e.g., 'Arial', 'Times New Roman'). Empty = unchanged.
            font_size: Font size in points (e.g., 12.0). 0 = unchanged.
            font_color: Hex color (e.g., '#FF0000' for red). Empty = unchanged.
            path: Optional document path.
        """
        try:
            return {
                "success": True,
                **engine.format_text(
                    paragraph_index=paragraph_index,
                    run_index=run_index,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    strikethrough=strikethrough,
                    font_name=font_name if font_name else None,
                    font_size=font_size if font_size > 0 else None,
                    font_color=font_color if font_color else None,
                    path=path if path else None,
                ),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def set_paragraph_alignment(paragraph_index: int, alignment: str, path: str = "") -> dict:
        """Set the alignment of a paragraph.

        Args:
            paragraph_index: Zero-based paragraph index.
            alignment: Alignment value: 'left', 'center', 'right', or 'justify'.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.set_paragraph_alignment(paragraph_index, alignment, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def set_paragraph_spacing(
        paragraph_index: int,
        space_before: float = 0,
        space_after: float = 0,
        line_spacing: float = 0,
        path: str = "",
    ) -> dict:
        """Set paragraph spacing.

        Args:
            paragraph_index: Zero-based paragraph index.
            space_before: Space before paragraph in points. 0 = unchanged.
            space_after: Space after paragraph in points. 0 = unchanged.
            line_spacing: Line spacing in points. 0 = unchanged.
            path: Optional document path.
        """
        try:
            return {
                "success": True,
                **engine.set_paragraph_spacing(
                    paragraph_index=paragraph_index,
                    space_before=space_before if space_before > 0 else None,
                    space_after=space_after if space_after > 0 else None,
                    line_spacing=line_spacing if line_spacing > 0 else None,
                    path=path if path else None,
                ),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
