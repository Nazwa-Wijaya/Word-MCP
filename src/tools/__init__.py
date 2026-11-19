from .document_tools import register_document_tools
from .content_tools import register_content_tools
from .format_tools import register_format_tools
from .table_tools import register_table_tools
from .image_tools import register_image_tools
from .header_footer_tools import register_header_footer_tools
from .style_tools import register_style_tools
from .export_tools import register_export_tools
from .page_setup_tools import register_page_setup_tools


def register_all_tools(mcp_server, engine):
    """Register all Word MCP tools."""
    register_document_tools(mcp_server, engine)
    register_content_tools(mcp_server, engine)
    register_format_tools(mcp_server, engine)
    register_table_tools(mcp_server, engine)
    register_image_tools(mcp_server, engine)
    register_header_footer_tools(mcp_server, engine)
    register_style_tools(mcp_server, engine)
    register_export_tools(mcp_server, engine)
    register_page_setup_tools(mcp_server, engine)
