import sys
from pathlib import Path
from types import SimpleNamespace as NS
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scau-thesis-format-cn" / "scripts"))
import inspect_word_format_signatures as signals


def body_item(chars, points):
    paragraph = NS(Range=NS(Start=1, Text="正文样例", Tables=NS(Count=0), Font=NS(Size=12),
        ParagraphFormat=NS(CharacterUnitFirstLineIndent=chars, FirstLineIndent=points)))
    return {"index": 1, "text": "正文样例", "paragraph": paragraph}


class FormatSignalTests(unittest.TestCase):
    def test_two_character_indent_accepts_equivalent_point_units(self):
        report = signals.body_first_line_indent_check([body_item(None, 24)], toc_end=0, references_title_item=None)
        self.assertEqual(report["mismatch_count"], 0)
        self.assertEqual(report["paragraphs_checked"], 1)

    def test_wrong_indent_remains_a_candidate(self):
        report = signals.body_first_line_indent_check([body_item(1.8, 21.6)], toc_end=0, references_title_item=None)
        self.assertEqual(report["mismatch_count"], 1)
        self.assertEqual(report["status"], "suggested")

    def test_missing_or_mixed_page_break_property_needs_visual_check(self):
        for page_break in (0, 9999999):
            item = body_item(0, 0)
            item["paragraph"].Range.ParagraphFormat.PageBreakBefore = page_break
            report = signals.english_abstract_page_break_check(NS(), [item], english_title_item=item)
            self.assertEqual(report["status"], "manual_confirm")


if __name__ == "__main__":
    unittest.main()
