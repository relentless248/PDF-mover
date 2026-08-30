# -*- coding: utf-8 -*-
"""core 纯逻辑单元测试（标准库 unittest，无需第三方依赖）。

运行：python -m unittest discover -s tests -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.matching import (extract_year_from_project_id, folder_matches,
                           unique_path, unique_target_path)
from core.settings import (DEFAULT_SETTINGS, Settings, deep_merge,
                           segments_to_pattern, pattern_to_segments)


class TestYearExtraction(unittest.TestCase):
    def test_two_digit_year_after_prefix(self):
        # 年份取前缀之后的头两位数字
        self.assertEqual(extract_year_from_project_id("AB20251234"), "25")
        self.assertEqual(extract_year_from_project_id("ABCD24123456"), "24")

    def test_century_branch(self):
        # 数字段以世纪前缀开头时取其后两位
        self.assertEqual(extract_year_from_project_id("AB20251234", prefix_length=2), "25")
        self.assertEqual(extract_year_from_project_id("XY99201234", century_prefix="20"), "12")

    def test_custom_rules(self):
        self.assertEqual(
            extract_year_from_project_id("XYZ2025123", prefix_length=3, century_prefix="20"),
            "25")

    def test_invalid_input(self):
        self.assertIsNone(extract_year_from_project_id(None))
        self.assertIsNone(extract_year_from_project_id(""))
        self.assertIsNone(extract_year_from_project_id("A"))


class TestFolderMatching(unittest.TestCase):
    def test_prefix_match(self):
        self.assertTrue(folder_matches("AB123456 Data", "AB123456"))
        self.assertTrue(folder_matches("ab123456 data", "AB123456"))  # 不区分大小写

    def test_word_boundary_match(self):
        self.assertTrue(folder_matches("01 AB123456 Data", "AB123456"))
        self.assertFalse(folder_matches("XAB123456Y", "AB123456"))  # 无边界不算命中

    def test_no_match(self):
        self.assertFalse(folder_matches("CD999999 Data", "AB123456"))

    def test_flags_off(self):
        self.assertFalse(folder_matches("AB123456 Data", "AB123456",
                                        prefix_match=False, word_boundary_match=False))


class TestUniquePath(unittest.TestCase):
    def setUp(self):
        self.existing = set()

    def exists(self, p):
        return p in self.existing

    def test_no_conflict(self):
        self.assertEqual(unique_path(r"D:\a\TDS.pdf", exists=self.exists),
                         r"D:\a\TDS.pdf")

    def test_conflict_appends_counter(self):
        self.existing.update({r"D:\a\TDS.pdf", r"D:\a\TDS_1.pdf"})
        self.assertEqual(unique_path(r"D:\a\TDS.pdf", exists=self.exists),
                         r"D:\a\TDS_2.pdf")

    def test_conflict_without_ext(self):
        self.existing.add(r"D:\a\AB123456")
        self.assertEqual(unique_path(r"D:\a\AB123456", exists=self.exists),
                         r"D:\a\AB123456_1")

    def test_unique_target_path(self):
        self.existing.add(os.path.join("D:", "a", "TDS.pdf"))
        self.assertEqual(
            unique_target_path(os.path.join("D:", "a"), "TDS", exists=self.exists),
            os.path.join("D:", "a", "TDS_1.pdf"))


class TestSettings(unittest.TestCase):
    def test_deep_merge_nested(self):
        base = {"a": {"x": 1, "y": 2}, "b": 3}
        override = {"a": {"y": 20, "z": 30}}
        merged = deep_merge(base, override)
        self.assertEqual(merged, {"a": {"x": 1, "y": 20, "z": 30}, "b": 3})
        # 原 dict 不应被修改
        self.assertEqual(base, {"a": {"x": 1, "y": 2}, "b": 3})

    def test_default_segments_compile_to_default_pattern(self):
        pattern = segments_to_pattern(DEFAULT_SETTINGS["project_pattern_segments"])
        self.assertEqual(pattern, DEFAULT_SETTINGS["project_pattern"])

    def test_pattern_round_trip(self):
        pattern = DEFAULT_SETTINGS["project_pattern"]
        segments = pattern_to_segments(pattern)
        self.assertIsNotNone(segments)
        self.assertEqual(segments_to_pattern(segments), pattern)

    def test_fixed_values_round_trip(self):
        pattern = r"(?:A|B)[A-Z]{2,3}\d{6}"
        segments = pattern_to_segments(pattern)
        self.assertIsNotNone(segments)
        self.assertEqual(segments_to_pattern(segments), pattern)

    def test_unsupported_pattern_returns_none(self):
        self.assertIsNone(pattern_to_segments(r"\d{4}(?=X)"))

    def test_build_filename(self):
        s = Settings({})
        self.assertEqual(s.build_filename('not_found', project_id="AB123456"),
                         "未找到项目_AB123456.pdf")
        self.assertEqual(s.build_filename('see_relation', see_id="CD789012", project_id="AB123456"),
                         "关联CD789012_未找到项目_AB123456.pdf")

    def test_subfolder_matches_ignores_case_and_separators(self):
        s = Settings({})
        self.assertTrue(s.subfolder_matches("01 Data"))
        self.assertTrue(s.subfolder_matches("data_files"))
        self.assertTrue(s.subfolder_matches("TEST-DATA"))
        self.assertFalse(s.subfolder_matches("Misc"))  # 不含子目录名则不命中

    def test_ocr_lang_string(self):
        self.assertEqual(Settings({}).ocr_lang_string(), "eng+chi_sim")
        self.assertEqual(Settings({"ocr": {"languages": ["eng"]}}).ocr_lang_string(), "eng")

    def test_invalid_pattern_raises(self):
        with self.assertRaises(ValueError):
            Settings({"project_pattern": "("})

    def test_derived_regexes(self):
        s = Settings({})
        m = s.analyze_no_see_pattern.match("未找到项目AB123456.pdf")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "AB123456")
        m3 = s.see_from_filename_regex.search("关联CD789012_其他.pdf")
        self.assertIsNotNone(m3)
        self.assertEqual(m3.group(1), "CD789012")

    def test_generated_filenames_are_parseable(self):
        """工具自己重命名的文件必须能被 analyze 正则解析（回归：分隔符不一致）"""
        s = Settings({})
        name = s.build_filename('not_found', project_id="AB123456")
        m = s.analyze_no_see_pattern.match(name)
        self.assertIsNotNone(m, f"analyze_no_see_pattern 无法解析自己生成的文件名: {name}")
        self.assertEqual(m.group(1), "AB123456")

        name2 = s.build_filename('see_relation', see_id="CD789012", project_id="AB123456")
        m2 = s.analyze_see_pattern.match(name2)
        self.assertIsNotNone(m2, f"analyze_see_pattern 无法解析自己生成的文件名: {name2}")
        self.assertEqual(m2.group(1), "CD789012")
        self.assertEqual(m2.group(2), "AB123456")


class TestAppPathValidation(unittest.TestCase):
    """is_target_path_valid 的白名单校验（重点：Windows 盘符大小写）。"""

    @classmethod
    def setUpClass(cls):
        import pdf_mover  # 导入期不应有弹窗/退出等副作用
        cls.pdf_mover = pdf_mover

    def _make_app(self, roots):
        app = self.pdf_mover.PDFMoverApp.__new__(self.pdf_mover.PDFMoverApp)
        app.target_roots = roots
        return app

    def test_valid_under_root(self):
        app = self._make_app([{'path': r'C:\Projects', 'depth': 1}])
        self.assertTrue(app.is_target_path_valid(r'C:\Projects\AB123456\Data'))

    def test_drive_letter_case_insensitive(self):
        app = self._make_app([{'path': r'C:\Projects', 'depth': 1}])
        self.assertTrue(app.is_target_path_valid(r'c:\projects\AB123456\Data'))

    def test_root_itself_is_valid(self):
        app = self._make_app([{'path': r'C:\Projects', 'depth': 1}])
        self.assertTrue(app.is_target_path_valid(r'C:\Projects'))

    def test_prefix_collision_is_rejected(self):
        app = self._make_app([{'path': r'C:\Projects', 'depth': 1}])
        self.assertFalse(app.is_target_path_valid(r'C:\ProjectsOther\AB123456\Data'))

    def test_other_drive_is_rejected(self):
        app = self._make_app([{'path': r'C:\Projects', 'depth': 1}])
        self.assertFalse(app.is_target_path_valid(r'D:\Projects\AB123456\Data'))


class TestAppPureDelegation(unittest.TestCase):
    """主程序的年份提取 / 文件夹匹配方法应与 core.matching 行为一致。"""

    @classmethod
    def setUpClass(cls):
        import pdf_mover
        cls.pdf_mover = pdf_mover

    def test_folder_matches_delegates(self):
        app = self.pdf_mover.PDFMoverApp.__new__(self.pdf_mover.PDFMoverApp)
        self.assertTrue(app.folder_matches("01 AB123456 Data", "AB123456"))
        self.assertFalse(app.folder_matches("XAB123456Y", "AB123456"))

    def test_extract_year_delegates(self):
        app = self.pdf_mover.PDFMoverApp.__new__(self.pdf_mover.PDFMoverApp)
        self.assertEqual(app.extract_year_from_project_id("AB20251234"), "25")


if __name__ == "__main__":
    unittest.main()
