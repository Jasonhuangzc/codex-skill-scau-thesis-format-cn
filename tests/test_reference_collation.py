import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scau-thesis-format-cn" / "scripts"))
import reference_order_utils as refs


class ReferenceCollationTests(unittest.TestCase):
    def test_surname_pinyin_order_differs_from_unicode(self):
        entries = ["张三. 示例[J]. 2024.", "王四. 示例[J]. 2024.", "Brown A. Example[J]. 2024."]
        self.assertEqual(refs.sort_reference_entries(entries), [entries[1], entries[0], entries[2]])

    def test_missing_pinyin_does_not_silently_sort_by_codepoint(self):
        with patch.object(refs, "_pinyin_key", return_value=("raw", "fallback:raw")):
            with self.assertRaisesRegex(RuntimeError, "pypinyin"):
                refs.sort_reference_entries(["张三. 示例[J]. 2024."])
            report = refs.inspect_reference_sequence(["张三. 示例[J]. 2024."])
            self.assertEqual(report["status"], "manual_confirm")

    def test_foreign_only_needs_no_pinyin(self):
        with patch.object(refs, "_pinyin_key", side_effect=AssertionError("must not run")):
            self.assertEqual(refs.sort_reference_entries(["Zhang Z. Test.", "Brown A. Test."]),
                             ["Brown A. Test.", "Zhang Z. Test."])


if __name__ == "__main__":
    unittest.main()
