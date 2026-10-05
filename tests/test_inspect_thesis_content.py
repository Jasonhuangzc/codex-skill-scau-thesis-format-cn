"""Regression fixtures for content boundaries and evidence-safe findings."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from xml.sax.saxutils import escape
import zipfile


SCRIPTS = Path(__file__).resolve().parents[1] / "scau-thesis-format-cn" / "scripts"
sys.path.insert(0, str(SCRIPTS))
from inspect_thesis_content import inspect_content  # noqa: E402
from inspect_word_report import build_section_audit  # noqa: E402

W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def p(text: str, style: str = "") -> str:
    properties = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{properties}<w:r><w:t>{escape(text)}</w:t></w:r></w:p>"


class ContentInspectionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.doc = self.root / "thesis.docx"
        self.rules = self.root / "skill"
        official = self.rules / "assets/official-2024"
        official.mkdir(parents=True)
        (official / "manifest.json").write_text(json.dumps({"required_files": [{"filename": "source.pdf", "sha256": "00"}]}), encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def write_doc(self, paragraphs: list[str], styles: str = ""):
        body = f'<w:document xmlns:w="{W_URI}"><w:body>{"".join(paragraphs)}<w:sectPr/></w:body></w:document>'
        with zipfile.ZipFile(self.doc, "w") as archive:
            archive.writestr("word/document.xml", body)
            if styles:
                archive.writestr("word/styles.xml", f'<w:styles xmlns:w="{W_URI}">{styles}</w:styles>')

    def report(self):
        return inspect_content(self.doc, self.rules)

    @staticmethod
    def base() -> list[str]:
        return [p("本科毕业论文"), p("原创性声明"), p("使用授权声明"), p("摘        要"),
                p("研究结果" * 100), p("关键词：水稻；光合作用；产量"),
                p("Abstract: We measured rice growth and identified a response."),
                p("Key words: Rice; Photosynthesis; Yield"), p("目 录"),
                p("1 绪论"), p("研究背景。"), p("2 材料与方法"), p("3 结果与分析"), p("4 讨论与结论")]

    def test_input_is_read_only_and_positive_signals_are_not_compliance(self):
        self.write_doc(self.base())
        original = self.doc.read_bytes()
        report = self.report()
        self.assertEqual(self.doc.read_bytes(), original)
        self.assertTrue(report["read_only"])
        self.assertEqual(report["summary"]["confirmed_issues"], 0)
        self.assertEqual(report["sections"]["chinese_abstract"]["status"], "detected")
        self.assertEqual(report["judgement_basis"]["school_compliance"], "not_determined")

    def test_counting_discloses_han_and_visible_counts_and_excludes_keywords(self):
        self.write_doc([p("摘要"), p("汉字ABC 123。"), p("关键词：甲；乙；丙"), p("Abstract: One sentence."), p("Key words: One; Two; Three")])
        report = self.report()
        stats = report["abstract_stats"]["chinese_abstract"]
        self.assertEqual(stats["han_characters"], 2)
        self.assertEqual(stats["visible_characters_excluding_whitespace"], 9)
        finding = next(f for f in report["findings"] if f["rule_id"] == "abstract.cn_length")
        self.assertEqual(finding["status"], "manual_confirm")
        self.assertEqual(finding["source"]["comment_ids"], [29])

    def test_year_and_normal_numbered_prose_remain_in_abstract(self):
        body = "2024 年，研究以水稻为材料开展观察。" + "本研究结果与实验观察保持一致。" * 25
        self.write_doc([p("摘要"), p(body), p("1 观察结果"), p("关键词：水稻；观察；生长"),
                        p("Abstract: Rice was studied."), p("Key words: Rice; Observation; Growth"),
                        p("目录"), p("1 绪论")])
        report = self.report()
        self.assertEqual(report["abstract_stats"]["chinese_abstract"]["body_paragraphs"], [2, 3])
        self.assertGreater(report["abstract_stats"]["chinese_abstract"]["visible_characters_excluding_whitespace"], 300)
        self.assertEqual(report["abstract_stats"]["chinese_abstract"]["boundary"], "keyword_label")
        ids = {f["rule_id"] for f in report["findings"]}
        self.assertNotIn("abstract.empty.chinese_abstract", ids)
        self.assertNotIn("abstract.keyword_label_missing.chinese_abstract", ids)
        self.assertEqual([location["paragraph"] for location in report["sections"]["body"]["locations"]], [8])

    def test_missing_abstract_boundary_does_not_confirm_empty_content(self):
        self.write_doc([p("摘要"), p("第一章 绪论", "Heading1"), p("正文内容。")])
        report = self.report()
        empty = next(f for f in report["findings"] if f["rule_id"] == "abstract.empty.chinese_abstract")
        self.assertEqual(empty["status"], "manual_confirm")
        self.assertEqual(report["sections"]["body"]["status"], "detected")

    def test_real_empty_abstract_between_its_label_and_keywords_is_confirmed(self):
        self.write_doc([p("摘要"), p("关键词：甲；乙；丙")])
        report = self.report()
        empty = next(f for f in report["findings"] if f["rule_id"] == "abstract.empty.chinese_abstract")
        self.assertEqual(empty["status"], "confirmed")

    def test_explicit_outline_heading_with_terminal_tab_is_not_plain_toc(self):
        styles = '<w:style w:type="paragraph" w:styleId="CustomChapter"><w:name w:val="Chapter"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
        self.write_doc([p("目录"), p("1 绪论\t1", "Heading1"), p("第二章 方法\t2", "CustomChapter")], styles)
        report = self.report()
        self.assertEqual([location["paragraph"] for location in report["sections"]["body"]["locations"]], [2, 3])

    def test_chinese_chapter_notation_is_detected_after_abstract_keywords(self):
        self.write_doc([p("摘要"), p("研究" * 180), p("关键词：甲；乙；丙"),
                        p("目录"), p("第1章 绪论"), p("第二章材料与方法")])
        report = self.report()
        self.assertEqual([location["paragraph"] for location in report["sections"]["body"]["locations"]], [5, 6])

    def test_cn_keywords_apply_count_separators_and_terminal_rule(self):
        self.write_doc([p("摘要"), p("研究" * 180), p("关键词：甲;乙;"), p("Abstract: Results."), p("Key words: Rice; Yield; Light")])
        report = self.report()
        ids = {f["rule_id"]: f for f in report["findings"]}
        for key in ("keywords.cn_count", "keywords.separator.chinese_abstract", "keywords.trailing_punctuation.chinese_abstract"):
            self.assertEqual(ids[key]["status"], "confirmed")
            self.assertEqual(ids[key]["source"]["comment_ids"], [30])
            self.assertEqual(ids[key]["location"]["paragraph"], 3)

    def test_comma_keyword_boundary_does_not_invent_count(self):
        self.write_doc([p("摘要"), p("研究" * 180), p("关键词：甲，乙，丙")])
        report = self.report()
        ids = {f["rule_id"] for f in report["findings"]}
        self.assertIn("keywords.ambiguous_separator.chinese_abstract", ids)
        self.assertNotIn("keywords.cn_count", ids)

    def test_no_unsupported_english_length_or_keyword_count_limit(self):
        self.write_doc([p("摘要"), p("研究" * 180), p("关键词：甲；乙；丙"), p("Abstract: Short."), p("Key words: Rice")])
        report = self.report()
        self.assertEqual(report["keyword_stats"]["english_abstract"]["count"], 1)
        self.assertFalse(any("en_count" in f["rule_id"] or "en_length" in f["rule_id"] for f in report["findings"]))
        bilingual = next(f for f in report["findings"] if f["rule_id"] == "abstract.bilingual_consistency")
        self.assertEqual(bilingual["status"], "manual_confirm")

    def test_citation_like_years_are_only_manual_candidates(self):
        self.write_doc([p("摘要"), p("实验分两阶段（2024）开展；参见[1]。"), p("关键词：甲；乙；丙")])
        report = self.report()
        candidates = [f for f in report["findings"] if f["rule_id"] == "abstract.cn_citation_candidate"]
        self.assertTrue(candidates)
        self.assertTrue(all(f["status"] == "manual_confirm" for f in candidates))

    def test_toc_styles_and_inherited_styles_cannot_supply_real_sections(self):
        styles = '<w:style w:type="paragraph" w:styleId="CustomContents"><w:name w:val="Contents Custom"/><w:basedOn w:val="TOC1"/></w:style>'
        self.write_doc([p("目录"), p("参考文献", "TOC1"), p("致谢", "CustomContents"), p("1 绪论", "CustomContents")], styles)
        report = self.report()
        for key in ("references", "acknowledgements", "body"):
            self.assertEqual(report["sections"][key]["status"], "not_detected")

    def test_toc_sdt_gallery_cannot_supply_body(self):
        control = '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent>' + p("1 绪论") + p("参考文献") + "</w:sdtContent></w:sdt>"
        self.write_doc([p("目录"), control, p("摘 要"), p("研究" * 180), p("关键词：甲；乙；丙")])
        report = self.report()
        self.assertEqual(report["sections"]["body"]["status"], "not_detected")
        self.assertEqual(report["sections"]["references"]["status"], "not_detected")

    def test_multiline_toc_field_ends_before_real_reference_heading(self):
        begin = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
        end = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
        self.write_doc([p("目录"), begin, p("参考文献"), p("致谢"), end, p("参考文献"), p("真实文献.")])
        report = self.report()
        self.assertEqual(len(report["sections"]["references"]["locations"]), 1)
        self.assertEqual(report["sections"]["references"]["locations"][0]["paragraph"], 6)
        self.assertEqual(report["sections"]["acknowledgements"]["status"], "not_detected")

    def test_plain_toc_page_suffix_does_not_supply_body_or_acknowledgements(self):
        self.write_doc([p("目 录"), p("1 绪论\t1"), p("致谢......28")])
        report = self.report()
        self.assertEqual(report["sections"]["body"]["status"], "not_detected")
        self.assertEqual(report["sections"]["acknowledgements"]["status"], "not_detected")

    def test_duplicate_sections_are_exposed_and_optional_sections_are_not_missing_errors(self):
        self.write_doc(self.base() + [p("摘 要"), p("第二份摘要")])
        report = self.report()
        duplicate = next(f for f in report["findings"] if f["rule_id"] == "structure.duplicate.chinese_abstract")
        self.assertEqual(duplicate["status"], "manual_confirm")
        self.assertEqual(duplicate["observed"], 2)
        self.assertFalse(any(f["rule_id"] == "structure.appendix" for f in report["findings"]))

    def test_placeholder_author_and_reference_content_are_not_declared_errors(self):
        self.write_doc(self.base() + [p("附录A 标题标题"), p("English Title"), p("Song Nianxiu"), p("XXXX")])
        report = self.report()
        placeholders = [f for f in report["findings"] if f["category"] == "template_placeholder"]
        self.assertGreaterEqual(len(placeholders), 4)
        self.assertTrue(all(f["status"] == "manual_confirm" for f in placeholders))

    def test_removed_and_textbox_labels_are_not_structure(self):
        deletion = '<w:del>' + p("原创性声明") + "</w:del>"
        textbox = '<w:p><w:r><w:drawing><w:txbxContent>' + p("使用授权声明") + "</w:txbxContent></w:drawing></w:r></w:p>"
        self.write_doc([deletion, textbox, p("目录")])
        report = self.report()
        self.assertEqual(report["sections"]["originality_statement"]["status"], "not_detected")
        self.assertEqual(report["sections"]["authorization_statement"]["status"], "not_detected")
        self.assertGreater(report["extraction"]["tracked_change_elements"], 0)

    def test_science_structure_is_recommendation_and_title_length_is_not_invented(self):
        self.write_doc([p("本科毕业论文"), p("这是一个很长的题目" * 10), p("1 系统设计"), p("2 性能评估")])
        report = self.report()
        item = next(f for f in report["findings"] if f["rule_id"] == "body.recommended_science_structure")
        self.assertEqual(item["status"], "suggested")
        self.assertEqual(item["source"]["rule_strength"], "template_example")
        self.assertFalse(any("title_length" in f["rule_id"] for f in report["findings"]))

    def test_english_title_capitalization_is_only_a_manual_candidate(self):
        self.write_doc([p("摘要"), p("研究" * 180), p("关键词：甲；乙；丙"),
                        p("Effects of light on rice"), p("Song Nianxiu"), p("(College, Guangzhou, China)"),
                        p("Abstract: Results."), p("Key words: Rice; Yield; Light")])
        report = self.report()
        title = next(f for f in report["findings"] if f["rule_id"] == "title.en_capitalization_candidate")
        self.assertEqual(title["status"], "manual_confirm")
        self.assertEqual(title["source"]["comment_ids"], [31])
        self.assertEqual(title["location"]["paragraph"], 4)

    def test_legacy_word_report_does_not_equate_detection_with_completion(self):
        report = build_section_audit({"chinese_abstract": True, "references": True},
                                     {"chinese_abstract": "", "references": "陈爱东. 真实引用."})
        self.assertEqual(report["chinese_abstract"]["status"], "detected")
        self.assertEqual(report["references"]["status"], "template_placeholder_candidate")

    def test_missing_official_files_and_hash_mismatch_are_explicit(self):
        self.write_doc(self.base())
        report = self.report()
        self.assertEqual(report["rule_provenance"]["official_files"][0]["status"], "not_imported")
        original = self.rules / "assets/official-2024/source.pdf"
        original.write_bytes(b"source identity fixture")
        self.assertEqual(self.report()["rule_provenance"]["official_files"][0]["status"], "hash_mismatch")
        manifest = original.parent / "manifest.json"
        manifest.write_text(json.dumps({"required_files": [{"filename": original.name, "sha256": hashlib.sha256(original.read_bytes()).hexdigest()}]}), encoding="utf-8")
        provenance = self.report()["rule_provenance"]
        self.assertEqual(provenance["official_package_status"], "hash_matches_manifest")
        self.assertEqual(provenance["rule_verification"], "repository_transcription_only")

    def test_cli_writes_report_and_refuses_input_overwrite(self):
        self.write_doc(self.base())
        report_path = self.root / "audit.json"
        result = subprocess.run([sys.executable, str(SCRIPTS / "inspect_thesis_content.py"), str(self.doc), "--output", str(report_path)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(json.loads(report_path.read_text(encoding="utf-8"))["read_only"])
        original = self.doc.read_bytes()
        result = subprocess.run([sys.executable, str(SCRIPTS / "inspect_thesis_content.py"), str(self.doc), "--output", str(self.doc)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn("must not overwrite", result.stderr)
        self.assertEqual(self.doc.read_bytes(), original)

    def test_cli_refuses_hardlink_to_input(self):
        self.write_doc(self.base())
        output = self.root / "same-inode.json"
        os.link(self.doc, output)
        original = self.doc.read_bytes()
        result = subprocess.run([sys.executable, str(SCRIPTS / "inspect_thesis_content.py"), str(self.doc), "--output", str(output)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertEqual(self.doc.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
