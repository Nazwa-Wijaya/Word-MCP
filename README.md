# MCP Word Server 📝

A **Model Context Protocol (MCP)** server that gives AI agents full control over Microsoft Word documents. Create, read, edit, format, and manage `.docx` files programmatically through a standardized MCP interface.

## ✨ Features

### Document Management
- **Create** new Word documents
- **Open** existing `.docx` files
- **Save / Save As** documents
- **Close** documents
- **List** open documents

### Content Operations
- **Read** full document text or specific paragraphs
- **Add** paragraphs, headings, page breaks
- **Insert** text at specific positions
- **Delete** paragraphs
- **Replace** text (find & replace)
- **Search** text within documents

### Formatting
- **Font styling** – bold, italic, underline, strikethrough, color, size, font name
- **Paragraph alignment** – left, center, right, justify
- **Paragraph spacing** – before/after spacing, line spacing
- **Heading levels** – H1 through H9

### Tables
- **Create** tables with custom rows and columns
- **Read** table data
- **Add rows/columns** to existing tables
- **Write** to specific cells
- **Style** tables with built-in Word table styles

### Images
- **Insert** images into documents
- **Resize** images (width & height in inches)

### Headers & Footers
- **Add/Edit** headers and footers
- **Different first page** headers/footers

### Styles & Templates
- **Apply** built-in Word styles
- **List** available styles

### Export
- **Export to PDF**
- **Export to HTML**

### Page Setup
- **Set** page margins, orientation, paper size

### Live Word Application Control (COM)
- **Launch** Microsoft Word application
- **Control** Word via COM automation on Windows
- **Works offline** with `python-docx` for `.docx` file manipulation

## 🚀 Quick Start

### Prerequisites
- Python 3.10+
- Microsoft Word (for COM automation features, Windows only)
- `python-docx` and `pywin32` libraries

### Installation

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/MCP_Word.git
cd MCP_Word

# Install dependencies
pip install -r requirements.txt
```

### Running the MCP Server

```bash
# Run with stdio transport (default for MCP)
python src/server.py
```

### Configuration

#### OpenCode

Tambahkan di file `opencode.json` di root project, atau di `~/.config/opencode/config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "py",
      "args": ["src/server.py"],
      "cwd": "C:\\path\\to\\MCP_Word"
    }
  }
}
```

Lalu jalankan OpenCode di folder project:
```bash
cd C:\path\to\MCP_Word
opencode
```

#### Claude Desktop

Tambahkan ke `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "word": {
      "command": "py",
      "args": ["src/server.py"],
      "cwd": "C:\\path\\to\\MCP_Word"
    }
  }
}
```

#### Cursor / Windsurf / Other MCP Clients

```json
{
  "mcpServers": {
    "word": {
      "command": "py",
      "args": ["src/server.py"],
      "cwd": "C:\\path\\to\\MCP_Word"
    }
  }
}
```

## 📁 Project Structure

```
MCP_Word/
├── src/
│   ├── server.py              # MCP server entry point
│   ├── word_engine.py         # Core Word manipulation engine
│   ├── tools/
│   │   ├── __init__.py
│   │   ├── document_tools.py  # Document CRUD operations
│   │   ├── content_tools.py   # Text/content manipulation
│   │   ├── format_tools.py    # Formatting operations
│   │   ├── table_tools.py     # Table operations
│   │   ├── image_tools.py     # Image operations
│   │   ├── header_footer_tools.py  # Headers & footers
│   │   ├── style_tools.py     # Style management
│   │   ├── export_tools.py    # Export (PDF, HTML)
│   │   └── page_setup_tools.py    # Page layout settings
│   └── utils/
│       ├── __init__.py
│       └── helpers.py         # Shared utilities
├── tests/
│   └── test_word_engine.py
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

## 🛠 Available MCP Tools

| Tool | Description |
|------|-------------|
| `create_document` | Create a new blank Word document |
| `open_document` | Open an existing .docx file |
| `save_document` | Save the current document |
| `save_document_as` | Save document to a new path |
| `close_document` | Close a document |
| `list_documents` | List all open documents |
| `read_document` | Read all text from a document |
| `read_paragraph` | Read a specific paragraph |
| `add_paragraph` | Add a new paragraph |
| `add_heading` | Add a heading (level 1-9) |
| `insert_text` | Insert text at a position |
| `delete_paragraph` | Delete a specific paragraph |
| `find_and_replace` | Find and replace text |
| `search_text` | Search for text occurrences |
| `format_text` | Apply font formatting |
| `set_paragraph_alignment` | Set paragraph alignment |
| `set_paragraph_spacing` | Set paragraph spacing |
| `create_table` | Create a new table |
| `read_table` | Read table data |
| `add_table_row` | Add a row to a table |
| `add_table_column` | Add a column to a table |
| `write_table_cell` | Write to a table cell |
| `style_table` | Apply a table style |
| `insert_image` | Insert an image |
| `add_header` | Add/edit document header |
| `add_footer` | Add/edit document footer |
| `apply_style` | Apply a Word style |
| `list_styles` | List available styles |
| `export_to_pdf` | Export document to PDF |
| `export_to_html` | Export document to HTML |
| `set_page_margins` | Set page margins |
| `set_page_orientation` | Set page orientation |
| `set_page_size` | Set paper size |
| `add_page_break` | Insert a page break |

## 📄 License

MIT License – see [LICENSE](LICENSE) for details.

## 🤝 Contributing

Contributions are welcome! Please open an issue or submit a pull request.
