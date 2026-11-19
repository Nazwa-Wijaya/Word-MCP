"""Basic tests for the Word engine."""

import os
import sys
import tempfile

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from word_engine import WordEngine


def test_create_and_read():
    engine = WordEngine()

    # Create document
    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test.docx")
        created_path = engine.create_document(path)
        assert os.path.exists(created_path), "Document file should exist"

        # Add content
        engine.add_heading("Test Document", level=1)
        engine.add_paragraph("This is a test paragraph.")
        engine.add_paragraph("Second paragraph with more text.")

        # Save
        engine.save_document()

        # Read
        content = engine.read_document()
        assert len(content["paragraphs"]) >= 3, "Should have at least 3 paragraphs"

        # Search
        results = engine.search_text("test")
        assert len(results) > 0, "Should find 'test' in the document"

        # Find and replace
        result = engine.find_and_replace("test", "demo")
        assert result["replacements"] > 0, "Should have made replacements"

        print("[PASS] test_create_and_read passed")


def test_tables():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test_table.docx")
        engine.create_document(path)

        # Create table
        data = [["Name", "Age", "City"], ["Alice", "30", "NYC"], ["Bob", "25", "LA"]]
        result = engine.create_table(3, 3, data=data, style="Table Grid")
        assert result["table_index"] == 0

        # Read table
        table_data = engine.read_table(0)
        assert table_data["rows"] == 3
        assert table_data["cols"] == 3

        # Write cell
        engine.write_table_cell(0, 1, 1, "31")

        # Add row
        engine.add_table_row(0, data=["Charlie", "28", "Chicago"])

        engine.save_document()
        print("[PASS] test_tables passed")


def test_formatting():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test_format.docx")
        engine.create_document(path)

        engine.add_paragraph("Bold and beautiful text")
        # Paragraph index 0 (the one we just added)
        idx = len(engine._documents[engine._active_doc].paragraphs) - 1
        engine.format_text(idx, 0, bold=True, font_size=14.0, font_color="#FF0000")
        engine.set_paragraph_alignment(idx, "center")

        engine.save_document()
        print("[PASS] test_formatting passed")


def test_headers_footers():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test_hf.docx")
        engine.create_document(path)

        engine.add_header("My Document Header")
        engine.add_footer("Page Footer - Confidential")

        engine.save_document()
        print("[PASS] test_headers_footers passed")


def test_page_setup():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test_page.docx")
        engine.create_document(path)

        engine.set_page_margins(top=1.0, bottom=1.0, left=1.25, right=1.25)
        engine.set_page_orientation("landscape")
        engine.set_page_size(11.0, 8.5)  # Letter landscape

        engine.save_document()
        print("[PASS] test_page_setup passed")


def test_export_html():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test_export.docx")
        engine.create_document(path)

        engine.add_heading("Export Test", level=1)
        engine.add_paragraph("This document will be exported to HTML.")
        engine.save_document()

        html_path = os.path.join(tmpdir, "output.html")
        result = engine.export_to_html(html_path)
        assert result["exported"] is True
        assert os.path.exists(html_path)

        print("[PASS] test_export_html passed")


def test_styles():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        path = os.path.join(tmpdir, "test_styles.docx")
        engine.create_document(path)

        engine.add_paragraph("A styled paragraph")
        styles = engine.list_styles()
        assert len(styles) > 0, "Should have available styles"

        engine.save_document()
        print("[PASS] test_styles passed")


def test_document_lifecycle():
    engine = WordEngine()

    with tempfile.TemporaryDirectory() as tmpdir:
        # Create
        path1 = os.path.join(tmpdir, "doc1.docx")
        engine.create_document(path1)
        engine.add_paragraph("Document 1")

        path2 = os.path.join(tmpdir, "doc2.docx")
        engine.create_document(path2)
        engine.add_paragraph("Document 2")

        # List
        docs = engine.list_documents()
        assert len(docs) == 2, "Should have 2 open documents"

        # Save as
        path3 = os.path.join(tmpdir, "doc2_copy.docx")
        engine.save_document_as(path3)
        assert os.path.exists(path3)

        # Close
        engine.close_document()
        docs = engine.list_documents()
        assert len(docs) == 1, "Should have 1 open document after close"

        print("[PASS] test_document_lifecycle passed")


if __name__ == "__main__":
    test_create_and_read()
    test_tables()
    test_formatting()
    test_headers_footers()
    test_page_setup()
    test_export_html()
    test_styles()
    test_document_lifecycle()
    print("\nAll tests passed!")
