# -*- coding: utf-8 -*-
"""
集中式配置模块（core.settings）

职责：
1. 从 settings.json 读取用户配置，与内置默认值做深度合并；
2. 对项目编号正则等关键项做合法性校验；
3. 由「项目编号格式」统一推导所有派生正则（关联、文件名、目标匹配等）；
4. 为可视化配置界面提供「编号分段 <-> 正则」互转能力；
5. 向业务代码暴露只读 Settings 对象。

兼容性：
- settings.json 不存在或某项缺失时，自动回退默认值，行为与原版一致；
- 默认值严格等于改造前的硬编码值，保证向后兼容。
"""

import copy
import json
import os
import re

# 项目根目录（core 的上一级）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 与业务运行相关的文件路径（保持原样）
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')
INDEX_CACHE_FILE = os.path.join(BASE_DIR, 'index_cache.json')

# 通用设置文件（settings.json 用户实际配置，不提交；settings.example.json 模板，提交）
SETTINGS_FILE = os.path.join(BASE_DIR, 'settings.json')
SETTINGS_EXAMPLE_FILE = os.path.join(BASE_DIR, 'settings.example.json')


# ---------------------------------------------------------------------------
# 默认配置（严格等于改造前的硬编码值）
# ---------------------------------------------------------------------------
DEFAULT_SETTINGS = {
    # 项目编号正则（通用占位默认：2-5位大写字母 + 6-10位数字，不含任何企业特征；
    # 实际使用请务必在设置中改为贵组织自己的编号规则）
    "project_pattern": r"[A-Z]{2,5}\d{6,10}",

    # 项目编号可视化分段（供界面建模；None 表示使用上述高级正则）
    "project_pattern_segments": [
        {"type": "letters", "case": "upper", "min_len": 2, "max_len": 5},
        {"type": "digits", "min_len": 6, "max_len": 10},
    ],

    # 关联编号识别关键词（OCR 文本中触发关联提取）
    "see_keywords": ["see"],
    # 文件名中「关联」前缀（用于解析已重命名文件）
    "relation_keyword": "关联",
    # 文件名中「未找到项目」标记
    "not_found_keyword": "未找到项目",

    # 重命名 / 文件名规则模板（占位符：{original} 原名含扩展名；{project_id}；{see_id}）
    "filename_rules": {
        "unidentified": "未识别_{original}",
        "not_found": "未找到项目_{project_id}.pdf",
        "see_relation": "关联{see_id}_未找到项目_{project_id}.pdf",
        "multi_project": "多项目_{original}",
        "skip_identified": "{project_id}.pdf",
    },

    # 归档目标文件名（不含扩展名）
    "target_filename": "TDS",

    # 目标子目录名匹配：通用占位默认（实际使用请在设置中修改为贵组织实际名称）
    "target_subfolder": {"name": "Data"},

    # 文件夹匹配规则
    "folder_match": {"prefix_match": True, "word_boundary_match": True},

    # 年份提取规则（仅从编号内嵌数字段提取；unknown_folder 为未识别年份的归档目录名）
    "year_rules": {"prefix_length": 4, "century_prefix": "20", "unknown_folder": "未分类"},

    # 年份分类管理（enabled 控制是否启用年份筛选功能；available 用于预置主程序年份下拉）
    "years": {"enabled": True, "available": []},

    # OCR 识别参数（languages 为语言代码列表，如 ["eng"], ["chi_sim"], ["eng","chi_sim"]）
    "ocr": {
        "languages": ["eng", "chi_sim"],
        "dpi": 400,
        "max_pages": 5,
        "config": "--psm 3 --oem 3 -c preserve_interword_spaces=1",
        "page_timeout_sec": 20,
        "convert_timeout_sec": 30,
        "preview_timeout_sec": 15,
        "auto_rotate": True,      # 未识别到编号时是否自动旋转重试
        "rotate_step": 90,
    },

    # 缓存与性能
    "cache": {
        "ocr_max": 50,
        "search_max": 100,
        "search_depth": 1,
        "chunk_size": 10000,
        "cache_ttl_days": 7,
    },

    # 文件操作
    "file_ops": {
        "retries": 5,
        "retry_interval_sec": 2,
        "skip_dirs": ["$RECYCLE.BIN", "$WINDOWS.~BT", "System Volume Information", "Temp", "tmp"],
    },

    # 记录
    "records": {"auto_save_count": 10},

    # 依赖目录（相对项目根目录）。poppler 目录名不再写死版本号。
    "deps": {
        "tesseract_dir": "Tesseract-OCR",
        "poppler_dir": "poppler-25.12.0",
    },
}


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def deep_merge(base, override):
    """深度合并两个 dict：override 覆盖 base；子 dict 递归合并。"""
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result.get(key, {}), value)
        else:
            result[key] = copy.deepcopy(value)
    return result


# ---------------------------------------------------------------------------
# 编号分段 <-> 正则 互转
# ---------------------------------------------------------------------------
def segments_to_pattern(segments):
    """将可视化分段列表编译为正则字符串。

    每段字段：
      - type: "fixed" | "digits" | "letters" | "any"
      - fixed:   values(list)            -> (?:A|B)
      - digits:  min_len, max_len        -> \\d{n} 或 \\d{n,m}
      - letters: case(upper/lower/any), min_len, max_len -> [A-Z]{n} 等
      - any:     char_set(str 如 [A-Z0-9]), min_len, max_len
      - optional: bool -> 追加 ?
    """
    parts = []
    for seg in segments or []:
        stype = seg.get("type", "fixed")
        optional = bool(seg.get("optional", False))

        if stype == "fixed":
            values = [re.escape(str(v)) for v in seg.get("values", []) if v is not None and str(v) != ""]
            if not values:
                continue
            if len(values) == 1:
                part = values[0]
            else:
                part = "(?:" + "|".join(values) + ")"
        elif stype == "digits":
            char_set = r"\d"
            min_len = int(seg.get("min_len", 1))
            max_len = int(seg.get("max_len", min_len))
            part = _quantify(char_set, min_len, max_len)
        elif stype == "letters":
            case = seg.get("case", "upper")
            char_set = {"upper": "[A-Z]", "lower": "[a-z]", "any": "[A-Za-z]"}.get(case, "[A-Z]")
            min_len = int(seg.get("min_len", 1))
            max_len = int(seg.get("max_len", min_len))
            part = _quantify(char_set, min_len, max_len)
        elif stype == "any":
            char_set = seg.get("char_set") or "[A-Za-z0-9]"
            min_len = int(seg.get("min_len", 1))
            max_len = int(seg.get("max_len", min_len))
            part = _quantify(char_set, min_len, max_len)
        else:
            continue

        if optional:
            part += "?"
        parts.append(part)

    return "".join(parts)


def _quantify(char_set, min_len, max_len):
    if min_len <= 0:
        return char_set + "*"
    if min_len == max_len:
        if min_len == 1:
            return char_set
        return "%s{%d}" % (char_set, min_len)
    return "%s{%d,%d}" % (char_set, min_len, max_len)


def pattern_to_segments(pattern):
    """将常见、本工具可生成的正则解析回分段列表；无法解析时返回 None。

    仅支持本工具能生成的安全子集，用于在界面中预填/回显分段。
    """
    segments = []
    i = 0
    n = len(pattern)
    while i < n:
        # (?:A|B|C)  固定值列表
        if pattern.startswith("(?:", i):
            j = pattern.find(")", i)
            if j == -1:
                return None
            inner = pattern[i + 3:j]
            if not inner or "|" not in inner:
                return None
            values = []
            for item in inner.split("|"):
                if re.fullmatch(r"[A-Za-z0-9\\]+", item) is None:
                    return None
                values.append(re.sub(r"\\(.)", r"\1", item))
            segments.append({"type": "fixed", "values": values})
            i = j + 1
            continue

        # [A-Z]{n} / [A-Z]{n,m} / [A-Z]
        m = re.match(r"\[([^\]]+)\](?:\{(\d+)(?:,(\d+))?\})?", pattern[i:])
        if m:
            char_set = "[" + m.group(1) + "]"
            min_len, max_len = 1, 1
            if m.group(2):
                min_len = int(m.group(2))
                max_len = int(m.group(3) or m.group(2))
            if "\\d" in m.group(1):
                segments.append({"type": "digits", "min_len": min_len, "max_len": max_len})
            elif m.group(1) in ("A-Z", "a-z", "A-Za-z"):
                case = {"A-Z": "upper", "a-z": "lower", "A-Za-z": "any"}[m.group(1)]
                segments.append({"type": "letters", "case": case, "min_len": min_len, "max_len": max_len})
            else:
                segments.append({"type": "any", "char_set": char_set, "min_len": min_len, "max_len": max_len})
            i += m.end()
            continue

        # \d{10} / \d{10,12} / \d
        m = re.match(r"\\d(?:\{(\d+)(?:,(\d+))?\})?", pattern[i:])
        if m:
            min_len, max_len = 1, 1
            if m.group(1):
                min_len = int(m.group(1))
                max_len = int(m.group(2) or m.group(1))
            segments.append({"type": "digits", "min_len": min_len, "max_len": max_len})
            i += m.end()
            continue

        # 无法识别
        return None

    return segments


# ---------------------------------------------------------------------------
# Settings 对象
# ---------------------------------------------------------------------------
class Settings:
    def __init__(self, data=None):
        self.data = deep_merge(DEFAULT_SETTINGS, data or {})
        self._migrate()
        self._compiled = {}
        self._recompile()

    # ---- 重载 ----
    def reload(self, data=None):
        """根据新数据（或重新读取 settings.json）刷新配置与派生正则。"""
        if data is None:
            data = load_raw_settings()
        self.data = deep_merge(DEFAULT_SETTINGS, data or {})
        self._migrate()
        self._recompile()

    def _migrate(self):
        """旧版配置迁移：只要存在 ocr.lang 字符串，就用它覆盖 languages 并移除 lang。

        注意：不能用『languages 为空才迁移』作条件——deep_merge 总会注入默认
        languages，导致旧版 lang 永远无法触发迁移。
        """
        o = self.data.get("ocr", {})
        if isinstance(o.get("lang"), str):
            o["languages"] = [x for x in o.get("lang", "").split("+") if x]
            o.pop("lang", None)

    def _recompile(self):
        self._compiled = {}
        pp = self.project_pattern

        # 校验主正则
        try:
            self._compiled["project_regex"] = re.compile(pp, re.IGNORECASE)
        except re.error as e:
            raise ValueError("项目编号正则不合法: %s" % e)

        self._compiled["see_regex"] = self._build_see_regex()
        self._compiled["see_from_filename_regex"] = re.compile(
            re.escape(self.relation_keyword) + r'(' + pp + r')', re.IGNORECASE
        )
        self._compiled["analyze_see_pattern"] = re.compile(
            r'^' + re.escape(self.relation_keyword) + r'(' + pp + r')_'
            + re.escape(self.not_found_keyword) + r'_?(' + pp + r')\.pdf$',
            re.IGNORECASE,
        )
        # _? 兼容带/不带下划线两种命名，确保能解析 build_filename 自己生成的文件名
        self._compiled["analyze_no_see_pattern"] = re.compile(
            r'^' + re.escape(self.not_found_keyword) + r'_?(' + pp + r')\.pdf$',
            re.IGNORECASE,
        )

    def _build_see_regex(self):
        pp = self.project_pattern
        kws = self.see_keywords or ["see"]
        kw_alt = "|".join(re.escape(str(k)) for k in kws)
        if len(kws) == 1:
            prefix = re.escape(str(kws[0])) + r'\s*'
        else:
            prefix = r'(?:' + kw_alt + r')\s*'
        return re.compile(r'(?:' + prefix + r'|\(\s*)(' + pp + r')\s*\)?', re.IGNORECASE)

    # ---- 原始配置直接访问 ----
    @property
    def project_pattern(self):
        return self.data["project_pattern"]

    @property
    def see_keywords(self):
        return self.data.get("see_keywords", ["see"])

    @property
    def relation_keyword(self):
        return self.data.get("relation_keyword", "关联")

    @property
    def not_found_keyword(self):
        return self.data.get("not_found_keyword", "未找到项目")

    @property
    def filename_rules(self):
        return self.data["filename_rules"]

    @property
    def target_filename(self):
        return self.data["target_filename"]

    @property
    def target_subfolder(self):
        return self.data["target_subfolder"]

    @property
    def folder_match(self):
        return self.data["folder_match"]

    @property
    def year_rules(self):
        return self.data["year_rules"]

    @property
    def ocr(self):
        return self.data["ocr"]

    def ocr_lang_string(self):
        """返回传给 Tesseract 的语言字符串，如 'eng+chi_sim'。"""
        langs = self.ocr.get("languages") or []
        if isinstance(langs, str):
            langs = [langs]
        # 兼容旧版 ocr.lang
        if not langs and isinstance(self.ocr.get("lang"), str):
            langs = [x for x in self.ocr.get("lang", "").split("+") if x]
        return "+".join(langs) or "eng"

    @property
    def cache(self):
        return self.data["cache"]

    @property
    def years_config(self):
        return self.data.get("years", {"enabled": True, "available": []})

    @property
    def year_enabled(self):
        """是否启用年份筛选功能（主界面勾选开关，关闭时跳过一切年份过滤）。"""
        return bool(self.data.get("years", {}).get("enabled", True))

    @property
    def file_ops(self):
        return self.data["file_ops"]

    @property
    def records(self):
        return self.data["records"]

    @property
    def deps(self):
        return self.data["deps"]

    # ---- 派生正则 ----
    @property
    def project_regex(self):
        return self._compiled["project_regex"]

    @property
    def see_regex(self):
        return self._compiled["see_regex"]

    @property
    def see_from_filename_regex(self):
        return self._compiled["see_from_filename_regex"]

    @property
    def analyze_see_pattern(self):
        return self._compiled["analyze_see_pattern"]

    @property
    def analyze_no_see_pattern(self):
        return self._compiled["analyze_no_see_pattern"]

    # ---- 子目录名（不区分大小写、忽略空格/分隔符差异） ----
    @property
    def subfolder_name(self):
        return self.target_subfolder.get("name", "Test Data") or "Test Data"

    @property
    def subfolder_name_folded(self):
        return self.subfolder_name.casefold()

    @staticmethod
    def _normalize_name(name):
        """归一化名称：小写并去除非字母数字（空格、下划线、横线等）。"""
        if not name:
            return ""
        return re.sub(r'[^a-z0-9]+', '', name.casefold())

    def subfolder_matches(self, name):
        """判断 name 是否命中目标子目录名（忽略大小写与分隔符差异）。"""
        target = self._normalize_name(self.subfolder_name)
        if not target:
            return False
        return target in self._normalize_name(name)

    # ---- 文件名模板 ----
    def build_filename(self, rule, **kwargs):
        """按 filename_rules 模板生成文件名。"""
        return self.filename_rules.get(rule, "").format(**kwargs)

    # ---- 持久化 ----
    def save(self, path=None):
        path = path or SETTINGS_FILE
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, indent=4, ensure_ascii=False)


# ---------------------------------------------------------------------------
# 加载
# ---------------------------------------------------------------------------
def load_raw_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


# 全局单例（业务代码统一引用）
settings = Settings(load_raw_settings())


def reload_settings(data=None):
    """刷新全局 settings（供可视化配置保存后调用）。"""
    settings.reload(data)
    return settings