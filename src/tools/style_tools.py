"""Style management tools."""


def register_style_tools(mcp, engine):
    @mcp.tool()
    def apply_style(paragraph_index: int, style_name: str, path: str = "") -> dict:
        """Apply a named Word style to a paragraph.

        Args:
            paragraph_index: Zero-based paragraph index.
            style_name: Style name (e.g., 'Normal', 'Heading 1', 'Title', 'List Bullet', 'Quote').
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.apply_style(paragraph_index, style_name, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def list_styles(path: str = "") -> dict:
        """List all available styles in the document.

        Args:
            path: Optional document path.
        """
        try:
            styles = engine.list_styles(path if path else None)
            return {"success": True, "styles": styles, "count": len(styles)}
        except Exception as e:
            return {"success": False, "error": str(e)}
