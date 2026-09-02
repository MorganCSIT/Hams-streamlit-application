import unittest

from audit_webfleet_rda import audit_pdf_filename


class AuditPdfFilenameTests(unittest.TestCase):
    def test_uses_short_name_and_visible_id(self):
        self.assertEqual(audit_pdf_filename("Jean Dupont", "12345"), "Jean_Dupont__id_12345.pdf")

    def test_filename_is_limited_to_sixty_characters(self):
        filename = audit_pdf_filename("A very long collaborator name " * 8, "987654")

        self.assertLessEqual(len(filename), 60)
        self.assertTrue(filename.endswith("__id_987654.pdf"))

    def test_sanitizes_unsupported_characters(self):
        self.assertEqual(audit_pdf_filename("Jean / Dupont", "12:34"), "Jean___Dupont__id_12_34.pdf")

    def test_same_names_with_different_ids_are_unique(self):
        first = audit_pdf_filename("Jean Dupont", "12345")
        second = audit_pdf_filename("Jean Dupont", "67890")

        self.assertNotEqual(first, second)

    def test_pathological_long_ids_remain_limited_and_unique(self):
        first = audit_pdf_filename("Jean Dupont", "1" * 100)
        second = audit_pdf_filename("Jean Dupont", "1" * 99 + "2")

        self.assertLessEqual(len(first), 60)
        self.assertLessEqual(len(second), 60)
        self.assertNotEqual(first, second)


if __name__ == "__main__":
    unittest.main()
