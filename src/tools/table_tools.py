"""Table tools: create, read, modify, style tables."""

from typing import Optional


def register_table_tools(mcp, engine):
    @mcp.tool()
    def create_table(rows: int, cols: int, data: list[list[str]] | None = None, style: str = "", path: str = "") -> dict:
        """Create a new table in the document.

        Args:
            rows: Number of rows.
            cols: Number of columns.
            data: Optional 2D list of cell values to populate the table.
            style: Optional table style name (e.g., 'Table Grid', 'Light Shading').
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.create_table(rows, cols, data, style if style else None, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def read_table(table_index: int, path: str = "") -> dict:
        """Read all data from a table.

        Args:
            table_index: Zero-based table index.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.read_table(table_index, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def add_table_row(table_index: int, data: list[str] | None = None, path: str = "") -> dict:
        """Add a row to an existing table.

        Args:
            table_index: Zero-based table index.
            data: Optional list of cell values for the new row.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.add_table_row(table_index, data, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def add_table_column(table_index: int, data: list[str] | None = None, path: str = "") -> dict:
        """Add a column to an existing table.

        Args:
            table_index: Zero-based table index.
            data: Optional list of cell values for the new column.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.add_table_column(table_index, data, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def write_table_cell(table_index: int, row: int, col: int, text: str, path: str = "") -> dict:
        """Write text to a specific table cell.

        Args:
            table_index: Zero-based table index.
            row: Zero-based row index.
            col: Zero-based column index.
            text: Text to write to the cell.
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.write_table_cell(table_index, row, col, text, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @mcp.tool()
    def style_table(table_index: int, style_name: str, path: str = "") -> dict:
        """Apply a built-in Word style to a table.

        Args:
            table_index: Zero-based table index.
            style_name: Table style name (e.g., 'Table Grid', 'Light List', 'Medium Shading 1').
            path: Optional document path.
        """
        try:
            return {"success": True, **engine.style_table(table_index, style_name, path if path else None)}
        except Exception as e:
            return {"success": False, "error": str(e)}
