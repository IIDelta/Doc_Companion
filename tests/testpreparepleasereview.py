import tempfile
import unittest
from pathlib import Path

from macros.PreparePleaseReview import (
    build_audit_path,
    build_output_path,
    normalize_word_text,
)


class PleaseReviewPureFunctionTests(unittest.TestCase):
    """Tests that do not require Windows, Word, or pywin32."""

    def test_normalize_word_text_preserves_content(self) -> None:
        self.assertEqual(
            normalize_word_text("Line one\r\nLine two"),
            "Line one\rLine two",
        )

    def test_normalize_word_text_handles_none(self) -> None:
        self.assertEqual(normalize_word_text(None), "")

    def test_output_path_uses_please_review_marker(self) -> None:
        with tempfile.TemporaryDirectory() as temp_directory:
            source = Path(temp_directory) / "Protocol.docx"
            source.write_bytes(b"placeholder")

            output = build_output_path(source)

            self.assertEqual(
                output.name,
                "Protocol_PleaseReview.docx",
            )

    def test_output_path_avoids_existing_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_directory:
            source = Path(temp_directory) / "Protocol.docx"
            source.write_bytes(b"placeholder")

            existing_output = (
                Path(temp_directory) / "Protocol_PleaseReview.docx"
            )
            existing_output.write_bytes(b"placeholder")

            output = build_output_path(source)

            self.assertEqual(
                output.name,
                "Protocol_PleaseReview_2.docx",
            )

    def test_audit_path_replaces_word_extension(self) -> None:
        with tempfile.TemporaryDirectory() as temp_directory:
            output = (
                Path(temp_directory)
                / "Protocol_PleaseReview.docx"
            )

            audit_path = build_audit_path(output)

            self.assertEqual(
                audit_path.name,
                "Protocol_PleaseReview.audit.json",
            )


if __name__ == "__main__":
    unittest.main()
