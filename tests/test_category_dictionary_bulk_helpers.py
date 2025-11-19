import json
import tempfile
import unittest
from pathlib import Path

from core.category_dictionary import CategoryDictionary


class CategoryDictionaryBulkHelpersTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tempdir.cleanup)
        self.dictionary_path = Path(self._tempdir.name) / "category_dictionary.json"
        with open(self.dictionary_path, "w", encoding="utf-8") as handle:
            json.dump({"mappings": [], "categories": []}, handle)

        self.dictionary = CategoryDictionary(self.dictionary_path)
        # Start each test with a clean slate
        self.dictionary.mappings.clear()
        self.dictionary.categories.clear()

    def test_list_mappings_returns_sorted_copy(self) -> None:
        payloads = [
            {"description": "beta item", "category": "Category B", "confidence": 0.75},
            {"description": "Alpha Item", "category": "Category A", "confidence": 0.5},
        ]
        self.dictionary.upsert_mappings(payloads)

        snapshot = self.dictionary.list_mappings()
        descriptions = [row["description"] for row in snapshot]

        self.assertEqual(descriptions, ["Alpha Item", "beta item"])

        snapshot[0]["category"] = "Changed Category"
        normalized_key = "alpha item"
        self.assertIn(normalized_key, self.dictionary.mappings)
        self.assertEqual(self.dictionary.mappings[normalized_key].category, "Category A")

    def test_delete_prunes_unused_categories(self) -> None:
        payloads = [
            {"description": "Alpha", "category": "Category A"},
            {"description": "Beta", "category": "Category B"},
        ]
        self.dictionary.upsert_mappings(payloads)

        removed = self.dictionary.delete_mappings(["Alpha"])
        self.assertEqual(removed, 1)

        categories = self.dictionary.get_all_categories()
        self.assertNotIn("Category A", categories)
        self.assertIn("Category B", categories)

    def test_rename_category_updates_entries_and_prunes(self) -> None:
        payloads = [
            {"description": "Alpha", "category": "Legacy"},
            {"description": "Beta", "category": "Legacy"},
        ]
        self.dictionary.upsert_mappings(payloads)

        updated = self.dictionary.rename_category_for_descriptions(
            ["Alpha", "Beta"], "Modern"
        )
        self.assertEqual(updated, 2)

        categories = self.dictionary.get_all_categories()
        self.assertIn("Modern", categories)
        self.assertNotIn("Legacy", categories)

    def test_backup_current_file_creates_timestamp_copy(self) -> None:
        # Ensure dictionary has been saved at least once
        self.dictionary.upsert_mappings(
            [{"description": "Sample", "category": "Example"}]
        )
        self.dictionary.save_dictionary()

        backup_path = self.dictionary.backup_current_file()
        self.assertIsNotNone(backup_path)
        self.assertTrue(backup_path.exists())


if __name__ == "__main__":
    unittest.main()




