# -*- coding: utf-8 -*-
"""Built-in unit tests for Word Formatter Pro v2.7.4."""

from __future__ import annotations

import os
import tempfile
import unittest
import json
import tkinter as tk
from pathlib import Path
from unittest import mock

from docx import Document
from docx.oxml import OxmlElement

from wfp_config import DEFAULT_CONFIG
from wfp_core import (
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_PRESERVE,
    LegacyConversionUnavailable,
    WordProcessor,
    check_first_line_indent,
    is_body_paragraph,
)
from wfp_gui import WordFormatterGUI


class TextNormalizationTests(unittest.TestCase):
    def test_symbol_normalization_keeps_decimal_numbers(self):
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("你好,世界."),
            "你好，世界。",
        )
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("3.14 是 pi"),
            "3.14 是 pi",
        )

    def test_ellipsis_and_quotes(self):
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text('他说"你好"...'),
            "他说“你好”……",
        )
        self.assertEqual(
            WordProcessor._normalize_symbols_in_text("version 1.2"),
            "version 1.2",
        )

    def test_markdown_cleaning(self):
        raw = "# 标题\n**粗体** 和 [链接](https://example.com)\n![图片](a.png)\n> 引用"
        cleaned = WordProcessor._clean_markdown(raw)
        self.assertEqual(cleaned, "标题\n粗体 和 链接\n图片\n引用")

    def test_markdown_cleaning_preserves_source_numeric_numbering(self):
        raw = (
            "一、登录\n\n"
            "1. 打开软件\n"
            "2. 完成登录\n\n"
            "（一）首页搜索\n\n"
            "1. 输入关键词\n"
            "2. 点击院校\n"
            "1.2.3 小节号保持原样"
        )
        cleaned = WordProcessor._clean_markdown(raw)
        self.assertEqual(cleaned, raw)


class BlankLineTests(unittest.TestCase):
    def test_delete_single_blank_line_and_compress_multiple(self):
        text = "a\n\nb\n\n\nc"
        self.assertEqual(
            WordProcessor._remove_blank_lines_from_text(text),
            "a\nb\n\nc",
        )

    def test_keep_single_blank_line_and_compress_multiple(self):
        text = "a\n\nb\n\n\nc"
        self.assertEqual(
            WordProcessor._remove_blank_lines_from_text(text, keep_single_blank_lines=True),
            "a\n\nb\n\nc",
        )

    def test_blank_line_mode_aliases(self):
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("preserve"),
            BLANK_LINE_MODE_PRESERVE,
        )
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("delete_single"),
            BLANK_LINE_MODE_DELETE_SINGLE,
        )
        self.assertEqual(
            WordProcessor._normalize_blank_line_mode("keep_single"),
            BLANK_LINE_MODE_KEEP_SINGLE,
        )

    def test_processor_preserve_mode_leaves_text_unchanged(self):
        processor = WordProcessor({"blank_line_mode": BLANK_LINE_MODE_PRESERVE})
        text = "a\n\nb\n\n\nc"
        self.assertEqual(processor._normalize_text_blank_lines(text), text)


class TableHelperTests(unittest.TestCase):
    def test_numeric_table_text(self):
        self.assertTrue(WordProcessor._is_numeric_table_text("1,234.56"))
        self.assertTrue(WordProcessor._is_numeric_table_text("¥100元"))
        self.assertTrue(WordProcessor._is_numeric_table_text("12.5%"))
        self.assertFalse(WordProcessor._is_numeric_table_text("abc"))

    def test_short_table_text(self):
        self.assertTrue(WordProcessor._is_short_table_text("合计", max_len=4))
        self.assertFalse(WordProcessor._is_short_table_text("较长文本内容", max_len=4))

    def test_table_percentage_normalization(self):
        pcts = WordProcessor._normalize_table_pcts([1, 9], 20, 80)
        self.assertAlmostEqual(sum(pcts), 100.0)
        self.assertEqual(pcts, [20.0, 80.0])


class OoxmlProtectionTests(unittest.TestCase):
    def test_ooxml_element_detection(self):
        doc = Document()
        para = doc.add_paragraph()
        run = para.add_run()
        run._r.append(OxmlElement("w:drawing"))
        self.assertTrue(WordProcessor._has_drawing_or_pict(para))
        self.assertFalse(WordProcessor._has_field_codes(para))

        para_field = doc.add_paragraph()
        field_run = para_field.add_run()
        field_run._r.append(OxmlElement("w:fldChar"))
        self.assertTrue(WordProcessor._has_field_codes(para_field))

    def test_strip_leading_whitespace_removes_plain_blank_run(self):
        doc = Document()
        para = doc.add_paragraph()
        para.add_run("   ")
        para.add_run("正文")
        WordProcessor({})._strip_leading_whitespace(para)
        self.assertEqual(para.text, "正文")

    def test_strip_leading_whitespace_preserves_special_run(self):
        doc = Document()
        para = doc.add_paragraph()
        special_run = para.add_run()
        special_run._r.append(OxmlElement("w:fldChar"))
        para.add_run("正文")

        WordProcessor({})._strip_leading_whitespace(para)

        self.assertTrue(WordProcessor._has_field_codes(para))
        self.assertEqual(len(para.runs), 2)
        self.assertEqual(para.text, "正文")


class FirstLineIndentTests(unittest.TestCase):
    def _format_doc(self, paragraphs, config=None, table_text=None):
        with tempfile.TemporaryDirectory(prefix="wfp_indent_test_") as tmpdir:
            source = Path(tmpdir) / "source.docx"
            output = Path(tmpdir) / "output.docx"
            doc = Document()
            for text in paragraphs:
                doc.add_paragraph(text)
            if table_text is not None:
                table = doc.add_table(rows=1, cols=1)
                table.cell(0, 0).text = table_text
            doc.save(source)

            merged_config = DEFAULT_CONFIG.copy()
            if config:
                merged_config.update(config)
            WordProcessor(merged_config).format_document(str(source), str(output))
            return Document(output)

    def test_body_paragraph_gets_first_line_indent(self):
        output = self._format_doc(["这是一个普通正文段落，用于验证首行缩进。"])

        self.assertTrue(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))

    def test_h1_paragraph_does_not_get_first_line_indent(self):
        output = self._format_doc(["一、绪论"])

        self.assertFalse(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))

    def test_h2_paragraph_does_not_get_first_line_indent(self):
        output = self._format_doc(["（一）研究背景"])

        self.assertFalse(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))

    def test_figure_caption_does_not_get_first_line_indent(self):
        output = self._format_doc(["图 1 系统结构图"])

        self.assertFalse(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))

    def test_table_caption_does_not_get_first_line_indent(self):
        output = self._format_doc(["表 1 实验结果"])

        self.assertFalse(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))

    def test_reference_entry_does_not_get_first_line_indent(self):
        output = self._format_doc(["[1] 张三. 文献标题. 期刊, 2024."])

        self.assertFalse(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))

    def test_table_cell_paragraph_is_not_body_paragraph(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        para = table.cell(0, 0).paragraphs[0]
        para.text = "表格内正文"

        self.assertFalse(is_body_paragraph(para))

    def test_old_config_without_first_line_indent_fields_is_compatible(self):
        old_config = DEFAULT_CONFIG.copy()
        for key in (
            "enable_first_line_indent",
            "first_line_indent_chars",
            "first_line_indent_tolerance_chars",
            "first_line_indent_scope",
        ):
            old_config.pop(key, None)

        output = self._format_doc(["旧配置正文段落。"], config=old_config)

        self.assertTrue(check_first_line_indent(output.paragraphs[0], 2.0, 0.2))


class GUIFirstLineIndentConfigTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.message_patches = [
            mock.patch("wfp_gui.messagebox.showinfo"),
            mock.patch("wfp_gui.messagebox.showerror"),
            mock.patch("wfp_gui.messagebox.showwarning"),
        ]
        for patcher in self.message_patches:
            patcher.start()
        self.app = WordFormatterGUI(self.root)

    def tearDown(self):
        for patcher in reversed(self.message_patches):
            patcher.stop()
        self.root.destroy()

    def _entry_value(self, key):
        return self.app.entries[key].get().strip()

    def test_gui_first_line_indent_defaults_initialize(self):
        self.assertTrue(self.app.enable_first_line_indent_var.get())
        self.assertEqual(self._entry_value("first_line_indent_chars"), "2.0")
        self.assertEqual(self._entry_value("first_line_indent_tolerance_chars"), "0.2")
        self.assertEqual(self._entry_value("first_line_indent_scope"), "body_only")

    def test_save_config_includes_first_line_indent_fields(self):
        with tempfile.TemporaryDirectory(prefix="wfp_gui_config_") as tmpdir:
            output = Path(tmpdir) / "config.json"
            with mock.patch("wfp_gui.filedialog.asksaveasfilename", return_value=str(output)):
                self.app.save_config()

            saved = json.loads(output.read_text(encoding="utf-8"))

        self.assertTrue(saved["enable_first_line_indent"])
        self.assertEqual(saved["first_line_indent_chars"], 2.0)
        self.assertEqual(saved["first_line_indent_tolerance_chars"], 0.2)
        self.assertEqual(saved["first_line_indent_scope"], "body_only")

    def test_old_config_load_uses_first_line_indent_defaults(self):
        old_config = DEFAULT_CONFIG.copy()
        for key in (
            "enable_first_line_indent",
            "first_line_indent_chars",
            "first_line_indent_tolerance_chars",
            "first_line_indent_scope",
        ):
            old_config.pop(key, None)

        self.app._apply_config(old_config)
        collected = self.app.collect_config()

        self.assertTrue(collected["enable_first_line_indent"])
        self.assertEqual(collected["first_line_indent_chars"], 2.0)
        self.assertEqual(collected["first_line_indent_tolerance_chars"], 0.2)
        self.assertEqual(collected["first_line_indent_scope"], "body_only")

    def test_invalid_first_line_indent_chars_is_blocked_on_save(self):
        self.app.entries["first_line_indent_chars"].delete(0, tk.END)
        self.app.entries["first_line_indent_chars"].insert(0, "5.1")

        with tempfile.TemporaryDirectory(prefix="wfp_gui_config_") as tmpdir:
            output = Path(tmpdir) / "config.json"
            with mock.patch("wfp_gui.filedialog.asksaveasfilename", return_value=str(output)):
                self.app.save_config()

            self.assertFalse(output.exists())

    def test_invalid_first_line_indent_tolerance_is_blocked_on_save(self):
        self.app.entries["first_line_indent_tolerance_chars"].delete(0, tk.END)
        self.app.entries["first_line_indent_tolerance_chars"].insert(0, "1.1")

        with tempfile.TemporaryDirectory(prefix="wfp_gui_config_") as tmpdir:
            output = Path(tmpdir) / "config.json"
            with mock.patch("wfp_gui.filedialog.asksaveasfilename", return_value=str(output)):
                self.app.save_config()

            self.assertFalse(output.exists())


class TempAndConversionTests(unittest.TestCase):
    def test_temp_docx_path_uses_system_temp_and_safe_name(self):
        processor = WordProcessor(DEFAULT_CONFIG.copy())
        temp_path = processor._make_temp_docx_path("copy", 'bad<name>:"?')
        try:
            self.assertEqual(
                os.path.normcase(os.path.abspath(os.path.dirname(temp_path))),
                os.path.normcase(os.path.abspath(tempfile.gettempdir())),
            )
            self.assertRegex(os.path.basename(temp_path), r"^~temp_copy_bad_name_[0-9]+_[0-9a-f]{8}\.docx$")
            self.assertIn(temp_path, processor.temp_files)
        finally:
            processor._cleanup_temp_files()

    def test_txt_conversion_uses_blank_line_mode_and_cleans_temp_file(self):
        with tempfile.TemporaryDirectory(prefix="wfp273_test_") as tmpdir:
            source = Path(tmpdir) / "sample.txt"
            source.write_text("标题\n\n正文一\n\n\n正文二", encoding="utf-8")

            processor = WordProcessor(DEFAULT_CONFIG.copy())
            temp_docx, is_from_txt = processor.convert_to_docx(str(source))
            try:
                self.assertTrue(is_from_txt)
                self.assertTrue(os.path.exists(temp_docx))
                self.assertEqual(
                    os.path.normcase(os.path.abspath(os.path.dirname(temp_docx))),
                    os.path.normcase(os.path.abspath(tempfile.gettempdir())),
                )
                self.assertFalse(list(Path(tmpdir).glob("~temp_*.docx")))

                converted = Document(temp_docx)
                self.assertEqual(
                    [p.text for p in converted.paragraphs],
                    ["标题", "正文一", "", "正文二"],
                )
            finally:
                processor._cleanup_temp_files()
            self.assertFalse(os.path.exists(temp_docx))

    def test_missing_soffice_marks_legacy_conversion_skipped(self):
        class MissingConverter:
            available = False

        processor = WordProcessor(DEFAULT_CONFIG.copy())
        processor.soffice_converter = MissingConverter()
        with self.assertRaises(LegacyConversionUnavailable):
            processor._convert_legacy_with_soffice("legacy.doc", "unused.docx")


def main(argv=None):
    suite = unittest.defaultTestLoader.loadTestsFromModule(__import__(__name__))
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    if result.wasSuccessful():
        print("所有单元测试通过")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
