"""Content manipulation tools: read, add, insert, delete, search, replace."""


def register_content_tools(mcp, engine):
    @mcp.tool()
    def read_document(path: str = "") -> dict:
        """Read the full content of a Word document including paragraphs and tables.

        Args:
            path: Optional document path. If empty, reads the active document.
        """
        try:
            return {"success": True, **engine.read_document(path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def read_paragraph(index: int, path: str = "") -> dict:
        """Read a specific paragraph by index, including run-level detail.

        Args:
            index: Zero-based paragraph index.
            path: Optional document path. If empty, uses the active document.
        """
        try:
            return {"success": True, **engine.read_paragraph(index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def add_paragraph(text: str, style: str = "", path: str = "") -> dict:
        """Add a new paragraph at the end of the document.

        Args:
            text: The text content of the paragraph.
            style: Optional Word style name (e.g., 'Normal', 'List Bullet', 'Quote').
            path: Optional document path. If empty, uses the active document.
        """
        try:
            return {"success": True, **engine.add_paragraph(text, style if style else None, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def add_heading(text: str, level: int = 1, path: str = "") -> dict:
        """Add a heading to the document.

        Args:
            text: The heading text.
            level: Heading level 0-9 (0 = Title, 1 = Heading 1, etc.).
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.add_heading(text, level, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def insert_text(text: str, paragraph_index: int, position: int = -1, path: str = "") -> dict:
        """Insert text into a specific paragraph at a character position.

        Args:
            text: Text to insert.
            paragraph_index: Zero-based paragraph index.
            position: Character position within the paragraph (-1 = append at end).
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.insert_text(text, paragraph_index, position, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def delete_paragraph(index: int, path: str = "") -> dict:
        """Delete a paragraph by its index.

        Args:
            index: Zero-based paragraph index to delete.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.delete_paragraph(index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def find_and_replace(find_text: str, replace_text: str, path: str = "") -> dict:
        """Find and replace text throughout the document (including tables).

        Args:
            find_text: Text to find.
            replace_text: Text to replace with.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.find_and_replace(find_text, replace_text, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def search_text(search_text: str, path: str = "") -> dict:
        """Search for text occurrences in the document.

        Args:
            search_text: Text to search for (case-insensitive).
            path: Optional document path.
        """
        try:
            results = engine.search_text(search_text, path if path else None)
            return {"success": True, "results": results, "count": len(results)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def add_page_break(path: str = "") -> dict:
        """Insert a page break at the end of the document.

        Args:
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.add_page_break(path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}
