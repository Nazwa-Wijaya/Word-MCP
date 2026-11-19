"""
MCP Word Server – Main entry point.

Provides AI agents with full control over Microsoft Word documents
through the Model Context Protocol (MCP).
"""

import sys
import os

# Add src to path so imports work
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp.server.mcpserver import MCPServer

from word_engine import WordEngine
from tools import register_all_tools


def main():
    # Create the MCP server
    mcp = MCPServer(
        name="MCP Word Server",
        version="1.0.0",
        description=(
            "Full-featured MCP server for Microsoft Word document manipulation. "
            "Create, read, edit, format, and export Word documents. "
            "Supports tables, images, headers/footers, styles, page setup, "
            "and PDF/HTML export."
        ),
    )

    # Create the Word engine
    engine = WordEngine()

    # Register all tools
    register_all_tools(mcp, engine)

    # Run the server
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
