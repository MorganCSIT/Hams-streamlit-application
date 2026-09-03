import unittest

from audit_webfleet_rda import audit_pdf_filename


class AuditPdfFilenameTests(unittest.TestCase):
    def test_uses_short_name_and_visible_id(self):
        self.assertEqual(audit_pdf_filename("Jean Dupont", "12345"), "Jean_Dupont_12345.pdf")

    def test_removes_display_metadata_and_repeated_collab_prefix(self):
        self.assertEqual(
            audit_pdf_filename(
                "Abastanotti Anita (ID: collab-collab-0226, RDA: 876, WF: 309)",
                "collab-collab-0226",
            ),
            "Abastanotti_Anita_0226.pdf",
        )

    def test_filename_is_limited_to_sixty_characters(self):
        filename = audit_pdf_filename("A very long collaborator name " * 8, "987654")

        self.assertLessEqual(len(filename), 50)
        self.assertTrue(filename.endswith("_987654.pdf"))

    def test_sanitizes_unsupported_characters(self):
        self.assertEqual(audit_pdf_filename("Jean / Dupont", "12:34"), "Jean___Dupont_12_34.pdf")

    def test_same_names_with_different_ids_are_unique(self):
        first = audit_pdf_filename("Jean Dupont", "12345")
        second = audit_pdf_filename("Jean Dupont", "67890")

        self.assertNotEqual(first, second)

    def test_pathological_long_ids_remain_limited_and_unique(self):
        first = audit_pdf_filename("Jean Dupont", "1" * 100)
        second = audit_pdf_filename("Jean Dupont", "1" * 99 + "2")

        self.assertLessEqual(len(first), 50)
        self.assertLessEqual(len(second), 50)
        self.assertNotEqual(first, second)


if __name__ == "__main__":
    unittest.main()
