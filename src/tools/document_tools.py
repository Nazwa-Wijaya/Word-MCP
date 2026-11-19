"""Document management tools: create, open, save, close, list."""


def register_document_tools(mcp, engine):
    @mcp.tool()
    def create_document(path: str = "") -> dict:
        """Create a new blank Word document.

        Args:
            path: Optional file path to save the document. If empty, creates a temp file.
        """
        try:
            result_path = engine.create_document(path if path else None)
            return {"success": True, "path": result_path, "message": f"Document created: {result_path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def open_document(path: str) -> dict:
        """Open an existing Word (.docx) document.

        Args:
            path: Full path to the .docx file to open.
        """
        try:
            result_path = engine.open_document(path)
            return {"success": True, "path": result_path, "message": f"Document opened: {result_path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def save_document(path: str = "") -> dict:
        """Save the current document (or a specific one by path).

        Args:
            path: Optional path of the document to save. If empty, saves the active document.
        """
        try:
            result_path = engine.save_document(path if path else None)
            return {"success": True, "path": result_path, "message": f"Document saved: {result_path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def save_document_as(new_path: str, source_path: str = "") -> dict:
        """Save the document to a new file path.

        Args:
            new_path: New file path to save the document to.
            source_path: Optional path of the source document. If empty, uses the active document.
        """
        try:
            result_path = engine.save_document_as(new_path, source_path if source_path else None)
            return {"success": True, "path": result_path, "message": f"Document saved as: {result_path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def close_document(path: str = "") -> dict:
        """Close a document.

        Args:
            path: Optional path of the document to close. If empty, closes the active document.
        """
        try:
            result_path = engine.close_document(path if path else None)
            return {"success": True, "path": result_path, "message": f"Document closed: {result_path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def list_documents() -> dict:
        """List all currently open documents with their info."""
        try:
            docs = engine.list_documents()
            return {"success": True, "documents": docs, "count": len(docs)}
        except Exception as e:
            return {"success": False, "error": str(e)}
