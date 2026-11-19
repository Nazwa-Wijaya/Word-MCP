"""Export tools: PDF, HTML."""


def register_export_tools(mcp, engine):
    @mcp.tool()
    def export_to_pdf(output_path: str, source_path: str = "") -> dict:
        """Export a Word document to PDF format.

        Requires Microsoft Word to be installed on Windows (uses COM automation).

        Args:
            output_path: File path for the output PDF.
            source_path: Optional source document path. If empty, uses the active document.
        """
        try:
            return {"success": True, **engine.export_to_pdf(output_path, source_path if source_path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def export_to_html(output_path: str, source_path: str = "") -> dict:
        """Export a Word document to HTML format.

        Args:
            output_path: File path for the output HTML file.
            source_path: Optional source document path. If empty, uses the active document.
        """
        try:
            return {"success": True, **engine.export_to_html(output_path, source_path if source_path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}
