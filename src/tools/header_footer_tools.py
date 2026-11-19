"""Header and footer tools."""


def register_header_footer_tools(mcp, engine):
    @mcp.tool()
    def add_header(text: str, section_index: int = 0, path: str = "") -> dict:
        """Add or update the header of a document section.

        Args:
            text: Header text content.
            section_index: Zero-based section index (default: 0, the first section).
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.add_header(text, section_index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def add_footer(text: str, section_index: int = 0, path: str = "") -> dict:
        """Add or update the footer of a document section.

        Args:
            text: Footer text content.
            section_index: Zero-based section index (default: 0, the first section).
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.add_footer(text, section_index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}
