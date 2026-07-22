import unittest

from app import build_editor_sections


class EditorSectionsTests(unittest.TestCase):
    def test_heading_lines_do_not_create_empty_duplicate_section(self) -> None:
        sections = build_editor_sections(
            "Bab 1: Pendahuluan\nIsi bab satu.\n\nBab 2: Isi\nIsi bab dua.",
            fallback_title="Bab 1",
        )

        self.assertEqual(
            sections,
            [
                {"title": "Bab 1: Pendahuluan", "body": "Isi bab satu."},
                {"title": "Bab 2: Isi", "body": "Isi bab dua."},
            ],
        )


if __name__ == "__main__":
    unittest.main()
