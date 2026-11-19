"""Page setup tools: margins, orientation, paper size."""


def register_page_setup_tools(mcp, engine):
    @mcp.tool()
    def set_page_margins(
        top: float = 0,
        bottom: float = 0,
        left: float = 0,
        right: float = 0,
        section_index: int = 0,
        path: str = "",
    ) -> dict:
        """Set page margins in inches.

        Args:
            top: Top margin in inches. 0 = unchanged.
            bottom: Bottom margin in inches. 0 = unchanged.
            left: Left margin in inches. 0 = unchanged.
            right: Right margin in inches. 0 = unchanged.
            section_index: Zero-based section index (default: 0).
            path: Optional document path.
        """
        try:
            return {
                "success": True,
                **engine.set_page_margins(
                    top=top if top > 0 else None,
                    bottom=bottom if bottom > 0 else None,
                    left=left if left > 0 else None,
                    right=right if right > 0 else None,
                    section_index=section_index,
                    path=path if path else None,
                ),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def set_page_orientation(orientation: str, section_index: int = 0, path: str = "") -> dict:
        """Set page orientation.

        Args:
            orientation: 'portrait' or 'landscape'.
            section_index: Zero-based section index (default: 0).
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.set_page_orientation(orientation, section_index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def set_page_size(width_inches: float, height_inches: float, section_index: int = 0, path: str = "") -> dict:
        """Set page paper size in inches.

        Common sizes: Letter (8.5 x 11), A4 (8.27 x 11.69), Legal (8.5 x 14).

        Args:
            width_inches: Page width in inches.
            height_inches: Page height in inches.
            section_index: Zero-based section index (default: 0).
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.set_page_size(width_inches, height_inches, section_index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}
