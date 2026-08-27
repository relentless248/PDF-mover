# -*- coding: utf-8 -*-
"""
可视化配置对话框（core.settings_dialog）

提供统一的「设置」界面，涵盖四类配置：
1. 项目编号格式（段式建模，可视化生成正则，无需手写）；
2. 识别参数（OCR 中/英语言、DPI、页数、自动旋转、引擎配置、超时、缓存）；
3. 归档与子目录（目标子目录名、归档文件名、文件夹匹配）；
4. 年份管理（年份提取规则、未识别年份目录、可用年份分类管理）。

交互要点：
- 鼠标点击单个数值输入框时，底部实时显示推荐范围与过大/过小的后果；
- 语言通过「中文 / 英文」独立勾选，支持单选或同选；
- 「恢复默认」仅作用于当前设置页，不影响其它页面；
- 关闭窗口自动保存并返回，便于主程序立即刷新生效。
"""

import copy
import re
import tkinter as tk
from tkinter import messagebox, ttk

from core.settings import (settings, DEFAULT_SETTINGS,
                           segments_to_pattern, pattern_to_segments)

# 语言代码映射
LANG_ENG_CODE = "eng"
LANG_CHI_CODE = "chi_sim"

# 数值参数的提示信息：范围 + 过小/过大后果
OCR_HINTS = {
    "dpi": {
        "range": "150 ~ 600（推荐 300 ~ 400）",
        "small": "过小(<150)：文字模糊，小字号/字符识别率明显下降，易漏读编号。",
        "big": "过大(>600)：渲染极慢、内存剧增，易触发超时；超过 400 收益很低。",
    },
    "max_pages": {
        "range": "1 ~ 20（推荐 1 ~ 5）",
        "small": "过小：若编号出现在靠后页，将无法识别。",
        "big": "过大：每多一页都增加转换与OCR耗时，批量处理显著变慢。",
    },
    "page_timeout_sec": {
        "range": "5 ~ 60（推荐 10 ~ 30）",
        "small": "过小：单页OCR可能被强制中断，造成『未识别』假象。",
        "big": "过大：异常页会长期占用线程，拖慢整体进度。",
    },
    "convert_timeout_sec": {
        "range": "10 ~ 120（推荐 20 ~ 60）",
        "small": "过小：大体积PDF转换会被强制中断。",
        "big": "过大：卡住的页面会挂起线程。",
    },
    "preview_timeout_sec": {
        "range": "5 ~ 60（推荐 10 ~ 20）",
        "small": "过小：预览图未渲染完成就被中断。",
        "big": "过大：切换文件时响应变慢。",
    },
    "ocr_max": {
        "range": "10 ~ 200（推荐 20 ~ 100）",
        "small": "过小：缓存命中率低，同一文件被反复识别。",
        "big": "过大：占用较多内存。",
    },
    "search_max": {
        "range": "10 ~ 500（推荐 50 ~ 200）",
        "small": "过小：频繁重新搜索文件夹，响应变慢。",
        "big": "过大：占用较多内存。",
    },
    "rotate_step": {
        "range": "1 ~ 360（常用 90，建议取能整除 360 的值）",
        "small": "过小：需要旋转很多次才能回到原位。",
        "big": "过大的非整除值可能导致旋转后角度错乱。",
    },
    "config": {
        "range": "Tesseract 命令行参数",
        "small": "参数错误可能导致识别失败或乱码。",
        "big": "-",
    },
}

# 识别参数数值字段的默认值（用于内存/界面）
OCR_INT_DEFAULTS = {
    "dpi": 400, "max_pages": 5, "page_timeout_sec": 20,
    "convert_timeout_sec": 30, "preview_timeout_sec": 15,
    "ocr_max": 50, "search_max": 100, "rotate_step": 90,
}

# 识别参数字段顺序与来源（'ocr' 写入 ocr；'cache' 写入 cache）
OCR_FIELDS = [
    ("dpi", "识别分辨率 DPI", "ocr"),
    ("max_pages", "最大识别页数", "ocr"),
    ("page_timeout_sec", "每页 OCR 超时(秒)", "ocr"),
    ("convert_timeout_sec", "PDF 转换超时(秒)", "ocr"),
    ("preview_timeout_sec", "预览渲染超时(秒)", "ocr"),
    ("ocr_max", "OCR 缓存条数", "cache"),
    ("search_max", "搜索缓存条数", "cache"),
    ("rotate_step", "自动旋转步长(度)", "ocr"),
]


class SegmentDialog(tk.Toplevel):
    """编辑单个编号分段。"""

    TYPE_LABELS = [
        ("固定文本", "fixed"),
        ("数字", "digits"),
        ("字母(大写)", "letters_upper"),
        ("字母(小写)", "letters_lower"),
        ("字母(任意大小写)", "letters_any"),
        ("自定义字符集", "any"),
    ]

    def __init__(self, parent, segment=None):
        super().__init__(parent)
        self.segment = segment or {}
        self.result = None

        self.title("编辑编号段")
        self.geometry("420x320")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        stype = self.segment.get("type", "fixed")

        tk.Label(self, text="段类型:").grid(row=0, column=0, sticky="w", padx=8, pady=(10, 2))
        self.type_codes = [c for _, c in self.TYPE_LABELS]
        self.type_labels = [lbl for lbl, _ in self.TYPE_LABELS]
        self.type_var = tk.StringVar()
        self.type_combo = ttk.Combobox(self, textvariable=self.type_var, values=self.type_labels,
                                       state="readonly", width=20)
        self.type_combo.grid(row=0, column=1, sticky="w", padx=8, pady=(10, 2))
        self._init_type_from_segment(stype)
        self.type_combo.bind("<<ComboboxSelected>>", self._on_type_change)

        self.fixed_label = tk.Label(self, text="固定值(多个用逗号分隔):")
        self.fixed_entry = tk.Entry(self, width=30)

        self.len_label = tk.Label(self, text="长度(最小-最大, 相同时填一个):")
        self.min_entry = tk.Entry(self, width=8)
        self.max_entry = tk.Entry(self, width=8)

        self.charset_label = tk.Label(self, text="字符集(如 A-Z0-9):")
        self.charset_entry = tk.Entry(self, width=30)

        self.optional_var = tk.BooleanVar(value=bool(self.segment.get("optional", False)))
        self.optional_check = tk.Checkbutton(self, text="该段可省略(可选)", variable=self.optional_var)

        self._place_fields()

        if stype == "fixed":
            self.fixed_entry.insert(0, ",".join(self.segment.get("values", [])))
        else:
            self.min_entry.insert(0, str(self.segment.get("min_len", 1)))
            self.max_entry.insert(0, str(self.segment.get("max_len", self.segment.get("min_len", 1))))
        if stype == "any":
            self.charset_entry.insert(0, self.segment.get("char_set", "[A-Za-z0-9]"))

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=15)
        tk.Button(btn_frame, text="确定", command=self._ok, width=10).pack(side="left", padx=6)
        tk.Button(btn_frame, text="取消", command=self.destroy, width=10).pack(side="left", padx=6)

    def _init_type_from_segment(self, stype):
        code = stype
        if stype == "letters":
            case = self.segment.get("case", "upper")
            code = "letters_" + case
        for idx, c in enumerate(self.type_codes):
            if c == code:
                self.type_var.set(self.type_labels[idx])
                return
        self.type_var.set(self.type_labels[0])

    def _current_type_code(self):
        idx = self.type_labels.index(self.type_var.get())
        return self.type_codes[idx]

    def _on_type_change(self, event=None):
        self._place_fields()

    def _place_fields(self):
        code = self._current_type_code()
        is_fixed = (code == "fixed")
        is_len = code in ("digits", "letters_upper", "letters_lower", "letters_any", "any")
        is_charset = (code == "any")

        self.fixed_label.grid_forget()
        self.fixed_entry.grid_forget()
        self.len_label.grid_forget()
        self.min_entry.grid_forget()
        self.max_entry.grid_forget()
        self.charset_label.grid_forget()
        self.charset_entry.grid_forget()
        self.optional_check.grid_forget()

        if is_fixed:
            self.fixed_label.grid(row=1, column=0, sticky="w", padx=8, pady=4)
            self.fixed_entry.grid(row=1, column=1, sticky="w", padx=8, pady=4)
        if is_len:
            self.len_label.grid(row=2, column=0, sticky="w", padx=8, pady=4)
            len_frame = tk.Frame(self)
            len_frame.grid(row=2, column=1, sticky="w", padx=8, pady=4)
            self.min_entry.grid(row=0, column=0)
            tk.Label(len_frame, text=" 到 ").grid(row=0, column=1)
            self.max_entry.grid(row=0, column=2)
            self.min_entry.grid(in_=len_frame)
            self.max_entry.grid(in_=len_frame)
        if is_charset:
            self.charset_label.grid(row=3, column=0, sticky="w", padx=8, pady=4)
            self.charset_entry.grid(row=3, column=1, sticky="w", padx=8, pady=4)
        self.optional_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=6)

    def _ok(self):
        code = self._current_type_code()
        seg = {"optional": self.optional_var.get()}

        if code == "fixed":
            values = [v.strip() for v in self.fixed_entry.get().split(",") if v.strip()]
            if not values:
                messagebox.showwarning("提示", "请填写至少一个固定值", parent=self)
                return
            seg["type"] = "fixed"
            seg["values"] = values
        elif code.startswith("letters_"):
            case = code.split("_", 1)[1]
            mn, mx = self._read_len()
            if mn is None:
                return
            seg["type"] = "letters"
            seg["case"] = case
            seg["min_len"], seg["max_len"] = mn, mx
        elif code == "digits":
            mn, mx = self._read_len()
            if mn is None:
                return
            seg["type"] = "digits"
            seg["min_len"], seg["max_len"] = mn, mx
        else:
            cs = self.charset_entry.get().strip() or "[A-Za-z0-9]"
            mn, mx = self._read_len()
            if mn is None:
                return
            seg["type"] = "any"
            seg["char_set"] = cs
            seg["min_len"], seg["max_len"] = mn, mx

        self.result = seg
        self.destroy()

    def _read_len(self):
        try:
            mn = int(self.min_entry.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "最小长度需为整数", parent=self)
            return None
        max_text = self.max_entry.get().strip()
        try:
            mx = int(max_text) if max_text else mn
        except ValueError:
            messagebox.showwarning("提示", "最大长度需为整数", parent=self)
            return None
        if mn < 1 or mx < mn:
            messagebox.showwarning("提示", "长度需 >=1 且最大 >= 最小", parent=self)
            return None
        return mn, mx


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("项目设置")
        self.geometry("760x600")
        self.minsize(700, 520)
        self.transient(parent)
        self.grab_set()

        self._saved = False

        # 工作副本
        self.segments = list(settings.data.get("project_pattern_segments") or
                             pattern_to_segments(settings.project_pattern) or [])
        self.advanced = not bool(settings.data.get("project_pattern_segments"))
        self.year_available = list(settings.years_config.get("available", []))

        self._build_ui()
        self._refresh_preview()

        # 关闭窗口(X) = 自动保存
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---- UI 骨架 ----
    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        self.tab_pattern = ttk.Frame(nb)
        self.tab_ocr = ttk.Frame(nb)
        self.tab_archive = ttk.Frame(nb)
        self.tab_years = ttk.Frame(nb)
        nb.add(self.tab_pattern, text="项目编号格式")
        nb.add(self.tab_ocr, text="识别参数")
        nb.add(self.tab_archive, text="归档与子目录")
        nb.add(self.tab_years, text="年份管理")

        self._build_pattern_tab()
        self._build_ocr_tab()
        self._build_archive_tab()
        self._build_years_tab()

        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", padx=8, pady=8)
        tk.Button(btn_frame, text="保存并关闭", command=self._on_save_close, width=14,
                  bg="#d4e6ff").pack(side="right", padx=6)
        tk.Button(btn_frame, text="关闭(自动保存)", command=self._on_close, width=14).pack(side="right", padx=6)

        tip = "提示：修改后关闭本窗口会自动保存并生效；每页上方有独立的『恢复默认』。"
        tk.Label(btn_frame, text=tip, fg="gray", font=("微软雅黑", 8)).pack(side="left", padx=6)

    # ---- 项目编号格式页 ----
    def _build_pattern_tab(self):
        f = self.tab_pattern
        f.grid_columnconfigure(0, weight=1)

        top = tk.Frame(f)
        top.grid(row=0, column=0, sticky="ew")
        tk.Button(top, text="恢复本页默认", command=self._reset_pattern_tab, width=14).pack(side="right", pady=4, padx=4)

        self.advanced_var = tk.BooleanVar(value=self.advanced)
        tk.Checkbutton(f, text="高级模式（直接编辑正则）", variable=self.advanced_var,
                       command=self._toggle_advanced).grid(row=1, column=0, sticky="w", pady=(4, 2))

        self.list_frame = tk.LabelFrame(f, text="编号分段（从上到下依次拼接）")
        self.list_frame.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        self.list_frame.grid_columnconfigure(0, weight=1)
        self.seg_listbox = tk.Listbox(self.list_frame, height=6, font=("微软雅黑", 9))
        self.seg_listbox.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        seg_btns = tk.Frame(self.list_frame)
        seg_btns.grid(row=0, column=1, sticky="n", padx=4, pady=4)
        tk.Button(seg_btns, text="添加段", command=self._add_segment, width=8).pack(pady=2)
        tk.Button(seg_btns, text="编辑", command=self._edit_segment, width=8).pack(pady=2)
        tk.Button(seg_btns, text="删除", command=self._remove_segment, width=8).pack(pady=2)
        tk.Button(seg_btns, text="上移", command=lambda: self._move_segment(-1), width=8).pack(pady=2)
        tk.Button(seg_btns, text="下移", command=lambda: self._move_segment(1), width=8).pack(pady=2)

        preview_frame = tk.LabelFrame(f, text="正则预览（自动生成）")
        preview_frame.grid(row=3, column=0, sticky="ew", padx=4, pady=4)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.pattern_text = tk.Text(preview_frame, height=2, font=("Consolas", 9))
        self.pattern_text.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        self.pattern_text.config(state="disabled")

        test_frame = tk.LabelFrame(f, text="测试（粘贴一个编号，实时验证是否匹配）")
        test_frame.grid(row=4, column=0, sticky="ew", padx=4, pady=4)
        test_frame.grid_columnconfigure(0, weight=1)
        self.test_entry = tk.Entry(test_frame, font=("微软雅黑", 10))
        self.test_entry.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        self.test_entry.bind("<KeyRelease>", lambda e: self._refresh_preview())
        self.test_result = tk.Label(test_frame, text="", fg="gray", anchor="w")
        self.test_result.grid(row=1, column=0, sticky="w", padx=4, pady=(0, 4))

        self._refresh_seg_list()
        self._toggle_advanced()

    def _reset_pattern_tab(self):
        d = DEFAULT_SETTINGS
        default_segs = list(d.get("project_pattern_segments") or [])
        if default_segs:
            self.segments = copy.deepcopy(default_segs)
        else:
            self.segments = []
        self.advanced_var.set(not bool(d.get("project_pattern_segments")))
        if self.advanced_var.get():
            self.pattern_text.config(state="normal")
            self.pattern_text.delete("1.0", "end")
            self.pattern_text.insert("1.0", d.get("project_pattern", ""))
            self.pattern_text.config(state="disabled")
        self._refresh_seg_list()
        self._toggle_advanced()
        messagebox.showinfo("提示", "已恢复「项目编号格式」本页默认值。", parent=self)

    # ---- 识别参数页 ----
    def _build_ocr_tab(self):
        f = self.tab_ocr
        f.grid_columnconfigure(1, weight=1)
        o = settings.ocr
        c = settings.cache

        top = tk.Frame(f)
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Button(top, text="恢复本页默认", command=self._reset_ocr_tab, width=14).pack(side="right", pady=4, padx=4)

        # 语言：中文 / 英文 独立勾选
        lang_frame = tk.LabelFrame(f, text="识别语言（可单选或多选）")
        lang_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=4)
        self.lang_var_chi = tk.BooleanVar()
        self.lang_var_eng = tk.BooleanVar()
        self._init_lang_vars(settings.ocr_lang_string())
        self.chi_check = tk.Checkbutton(lang_frame, text="简体中文", variable=self.lang_var_chi,
                                        command=lambda: self._refresh_lang_label())
        self.eng_check = tk.Checkbutton(lang_frame, text="英文", variable=self.lang_var_eng,
                                        command=lambda: self._refresh_lang_label())
        self.chi_check.pack(side="left", padx=10, pady=6)
        self.eng_check.pack(side="left", padx=10, pady=6)
        self.lang_result_label = tk.Label(lang_frame, text="", fg="gray")
        self.lang_result_label.pack(side="left", padx=10)
        self._refresh_lang_label()

        # 数值参数网格
        num_frame = tk.LabelFrame(f, text="识别 / 性能参数（点击输入框查看范围与后果）")
        num_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=4, pady=4)
        num_frame.grid_columnconfigure(1, weight=1)

        self.ocr_entry_map = {}
        self.ocr_source_map = {}
        for i, (key, label, source) in enumerate(OCR_FIELDS):
            default = OCR_INT_DEFAULTS[key]
            val = str((c if source == "cache" else o).get(key, default))
            tk.Label(num_frame, text=label + ":").grid(row=i, column=0, sticky="w", padx=6, pady=3)
            e = tk.Entry(num_frame, width=18)
            e.insert(0, val)
            e.grid(row=i, column=1, sticky="ew", padx=6, pady=3)
            e.bind("<FocusIn>", lambda ev, k=key: self._show_ocr_hint(k))
            self.ocr_entry_map[key] = e
            self.ocr_source_map[key] = source

        # OCR 引擎配置（高级）
        tk.Label(f, text="OCR 引擎配置(高级):").grid(row=3, column=0, sticky="w", padx=8, pady=3)
        self.config_entry = tk.Entry(f, width=30)
        self.config_entry.insert(0, o.get("config", ""))
        self.config_entry.grid(row=3, column=1, sticky="ew", padx=8, pady=3)
        self.config_entry.bind("<FocusIn>", lambda ev: self._show_ocr_hint("config"))

        # 自动旋转开关
        self.auto_rotate_var = tk.BooleanVar(value=o.get("auto_rotate", True))
        tk.Checkbutton(f, text="未识别到编号时自动旋转重试", variable=self.auto_rotate_var
                       ).grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 2))

        # 底部提示区（实时显示）
        hint_frame = tk.LabelFrame(f, text="参数提示")
        hint_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=4, pady=6)
        hint_frame.grid_columnconfigure(0, weight=1)
        self.ocr_hint_var = tk.StringVar(value="点击任意参数输入框，此处会显示推荐范围与过大/过小的后果。")
        tk.Label(hint_frame, textvariable=self.ocr_hint_var, anchor="w", justify="left",
                 fg="#1f4e79", font=("微软雅黑", 9), wraplength=680).grid(row=0, column=0, sticky="w", padx=6, pady=4)

    def _init_lang_vars(self, lang_str):
        langs = [x for x in str(lang_str).split("+") if x]
        self.lang_var_eng.set(LANG_ENG_CODE in langs)
        self.lang_var_chi.set(LANG_CHI_CODE in langs)

    def _refresh_lang_label(self):
        codes = self._current_lang_codes()
        self.lang_result_label.config(text="→ " + ("+".join(codes) if codes else "（未选择任何语言）"))

    def _current_lang_codes(self):
        codes = []
        if self.lang_var_eng.get():
            codes.append(LANG_ENG_CODE)
        if self.lang_var_chi.get():
            codes.append(LANG_CHI_CODE)
        return codes

    def _show_ocr_hint(self, key):
        hint = OCR_HINTS.get(key)
        if not hint:
            self.ocr_hint_var.set("（该字段无额外提示）")
            return
        self.ocr_hint_var.set(
            f"推荐/{hint['range']}\n• 值过小：{hint['small']}\n• 值过大：{hint['big']}")

    def _reset_ocr_tab(self):
        d = DEFAULT_SETTINGS
        for key, entry in self.ocr_entry_map.items():
            entry.delete(0, tk.END)
            defv = OCR_INT_DEFAULTS.get(key)
            if key == "ocr_max":
                defv = d["cache"].get("ocr_max", OCR_INT_DEFAULTS[key])
            elif key == "search_max":
                defv = d["cache"].get("search_max", OCR_INT_DEFAULTS[key])
            entry.insert(0, str(defv))
        self.config_entry.delete(0, tk.END)
        self.config_entry.insert(0, d["ocr"].get("config", ""))
        self.auto_rotate_var.set(d["ocr"].get("auto_rotate", True))
        self._init_lang_vars("+".join(d["ocr"].get("languages", ["eng"])))
        self._refresh_lang_label()
        messagebox.showinfo("提示", "已恢复「识别参数」本页默认值。", parent=self)

    # ---- 归档与子目录页 ----
    def _build_archive_tab(self):
        f = self.tab_archive
        f.grid_columnconfigure(1, weight=1)

        top = tk.Frame(f)
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Button(top, text="恢复本页默认", command=self._reset_archive_tab, width=14).pack(side="right", pady=4, padx=4)

        tk.Label(f, text="目标子文件夹名:").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.subfolder_entry = tk.Entry(f, width=30)
        self.subfolder_entry.insert(0, settings.subfolder_name)
        self.subfolder_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        tk.Label(f, text="⚠ 匹配规则：不区分大小写，忽略空格/下划线/横线等分隔符差异。\n"
                         "   如填 Data 可匹配 data、Data、01 Data、data_files 等。",
                 fg="#b8860b", justify="left", font=("微软雅黑", 8)
                 ).grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

        tk.Label(f, text="归档文件名(不含扩展名):").grid(row=3, column=0, sticky="w", padx=8, pady=4)
        self.target_filename_entry = tk.Entry(f, width=30)
        self.target_filename_entry.insert(0, settings.target_filename)
        self.target_filename_entry.grid(row=3, column=1, sticky="ew", padx=8, pady=4)

        self.prefix_match_var = tk.BooleanVar(value=settings.folder_match.get("prefix_match", True))
        self.boundary_match_var = tk.BooleanVar(value=settings.folder_match.get("word_boundary_match", True))
        tk.Checkbutton(f, text="匹配规则1：文件夹名以编号开头", variable=self.prefix_match_var
                       ).grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=3)
        tk.Checkbutton(f, text="匹配规则2：文件夹名包含完整编号(按单词边界)", variable=self.boundary_match_var
                       ).grid(row=5, column=0, columnspan=2, sticky="w", padx=8, pady=3)

    def _reset_archive_tab(self):
        d = DEFAULT_SETTINGS
        self.subfolder_entry.delete(0, tk.END)
        self.subfolder_entry.insert(0, d["target_subfolder"].get("name", "Data"))
        self.target_filename_entry.delete(0, tk.END)
        self.target_filename_entry.insert(0, d.get("target_filename", "TDS"))
        self.prefix_match_var.set(d["folder_match"].get("prefix_match", True))
        self.boundary_match_var.set(d["folder_match"].get("word_boundary_match", True))
        messagebox.showinfo("提示", "已恢复「归档与子目录」本页默认值。", parent=self)

    # ---- 年份管理页 ----
    def _build_years_tab(self):
        f = self.tab_years
        f.grid_columnconfigure(1, weight=1)

        top = tk.Frame(f)
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Button(top, text="恢复本页默认", command=self._reset_years_tab, width=14).pack(side="right", pady=4, padx=4)

        tk.Label(f, text="编号前缀长度(年份提取):").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.prefix_len_entry = tk.Entry(f, width=30)
        self.prefix_len_entry.insert(0, str(settings.year_rules.get("prefix_length", 4)))
        self.prefix_len_entry.grid(row=1, column=1, sticky="ew", padx=8, pady=4)

        tk.Label(f, text="世纪前缀(如 20 表示20xx年):").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        self.century_entry = tk.Entry(f, width=30)
        self.century_entry.insert(0, str(settings.year_rules.get("century_prefix", "20")))
        self.century_entry.grid(row=2, column=1, sticky="ew", padx=8, pady=4)

        tk.Label(f, text="未识别年份归档目录名:").grid(row=3, column=0, sticky="w", padx=8, pady=4)
        self.unknown_folder_entry = tk.Entry(f, width=30)
        self.unknown_folder_entry.insert(0, settings.year_rules.get("unknown_folder", "未分类"))
        self.unknown_folder_entry.grid(row=3, column=1, sticky="ew", padx=8, pady=4)

        # 可用年份分类管理
        ylist_frame = tk.LabelFrame(f, text="可用年份分类（预置到主程序年份下拉；留空则只显示扫描发现）")
        ylist_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=4, pady=6)
        ylist_frame.grid_columnconfigure(0, weight=1)
        f.grid_rowconfigure(4, weight=1)

        self.year_listbox = tk.Listbox(ylist_frame, height=6, font=("微软雅黑", 9))
        self.year_listbox.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self._refresh_year_list()

        ybtn = tk.Frame(ylist_frame)
        ybtn.grid(row=0, column=1, sticky="n", padx=4, pady=4)
        self.year_add_entry = tk.Entry(ybtn, width=8)
        self.year_add_entry.pack(pady=2)
        tk.Button(ybtn, text="添加", command=self._add_year, width=8).pack(pady=2)
        tk.Button(ybtn, text="删除选中", command=self._remove_year, width=8).pack(pady=2)

        tk.Label(f, text="说明：可用年份用于主界面「分拣年份」下拉预置，例如填写 24、25、26 等。\n"
                         "若年份内嵌于编号，程序仍会自动识别，此项仅用于快捷筛选。",
                 fg="gray", justify="left", font=("微软雅黑", 8)
                 ).grid(row=5, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 4))

    def _refresh_year_list(self):
        self.year_listbox.delete(0, tk.END)
        for y in self.year_available:
            self.year_listbox.insert(tk.END, y)

    def _add_year(self):
        val = self.year_add_entry.get().strip()
        if not val:
            return
        if val not in self.year_available:
            self.year_available.append(val)
            self._refresh_year_list()
        self.year_add_entry.delete(0, tk.END)

    def _remove_year(self):
        sel = self.year_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选中要删除的年份")
            return
        self.year_available.pop(sel[0])
        self._refresh_year_list()

    def _reset_years_tab(self):
        d = DEFAULT_SETTINGS
        self.prefix_len_entry.delete(0, tk.END)
        self.prefix_len_entry.insert(0, str(d["year_rules"].get("prefix_length", 4)))
        self.century_entry.delete(0, tk.END)
        self.century_entry.insert(0, str(d["year_rules"].get("century_prefix", "20")))
        self.unknown_folder_entry.delete(0, tk.END)
        self.unknown_folder_entry.insert(0, d["year_rules"].get("unknown_folder", "未分类"))
        self.year_available = list(d.get("years", {}).get("available", []))
        self._refresh_year_list()
        messagebox.showinfo("提示", "已恢复「年份管理」本页默认值。", parent=self)

    # ---- 编号段交互 ----
    def _toggle_advanced(self):
        if self.advanced_var.get():
            self.list_frame.grid_remove()
            self.pattern_text.config(state="normal")
        else:
            self.list_frame.grid()
            self.pattern_text.config(state="disabled")
        self._refresh_preview()

    def _seg_desc(self, seg):
        stype = seg.get("type", "fixed")
        if stype == "fixed":
            return "固定: " + "/".join(seg.get("values", []))
        if stype == "digits":
            return "数字: %d~%d" % (seg.get("min_len", 1), seg.get("max_len", seg.get("min_len", 1)))
        if stype == "letters":
            case = {"upper": "大写", "lower": "小写", "any": "任意"}.get(seg.get("case"), "大写")
            return "字母(%s): %d~%d" % (case, seg.get("min_len", 1), seg.get("max_len", seg.get("min_len", 1)))
        return "自定义: %d~%d" % (seg.get("min_len", 1), seg.get("max_len", seg.get("min_len", 1)))

    def _refresh_seg_list(self):
        self.seg_listbox.delete(0, tk.END)
        for seg in self.segments:
            self.seg_listbox.insert(tk.END, self._seg_desc(seg))

    def _current_pattern(self):
        if self.advanced_var.get():
            return self.pattern_text.get("1.0", "end").strip()
        return segments_to_pattern(self.segments)

    def _refresh_preview(self):
        pattern = self._current_pattern()
        if not self.advanced_var.get():
            self.pattern_text.config(state="normal")
            self.pattern_text.delete("1.0", "end")
            self.pattern_text.insert("1.0", pattern)
            self.pattern_text.config(state="disabled")
        try:
            re.compile(pattern, re.IGNORECASE)
            valid = True
        except re.error:
            valid = False
        if not pattern:
            self.test_result.config(text="（正则不能为空）", fg="red")
            return
        if not valid:
            self.test_result.config(text="（正则不合法）", fg="red")
            return
        sample = self.test_entry.get().strip()
        if sample:
            if re.fullmatch(pattern, sample, re.IGNORECASE):
                self.test_result.config(text=f"✔ 匹配：{sample}", fg="green")
            else:
                self.test_result.config(text=f"✘ 不匹配：{sample}", fg="red")
        else:
            self.test_result.config(text="（输入编号即可测试）", fg="gray")

    def _add_segment(self):
        dlg = SegmentDialog(self)
        self.wait_window(dlg)
        if dlg.result:
            self.segments.append(dlg.result)
            self._refresh_seg_list()
            self._refresh_preview()

    def _edit_segment(self):
        sel = self.seg_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选中要编辑的段")
            return
        idx = sel[0]
        dlg = SegmentDialog(self, self.segments[idx])
        self.wait_window(dlg)
        if dlg.result:
            self.segments[idx] = dlg.result
            self._refresh_seg_list()
            self._refresh_preview()

    def _remove_segment(self):
        sel = self.seg_listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请先选中要删除的段")
            return
        self.segments.pop(sel[0])
        self._refresh_seg_list()
        self._refresh_preview()

    def _move_segment(self, delta):
        sel = self.seg_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + delta
        if 0 <= new_idx < len(self.segments):
            self.segments[idx], self.segments[new_idx] = self.segments[new_idx], self.segments[idx]
            self._refresh_seg_list()
            self.seg_listbox.selection_set(new_idx)
            self._refresh_preview()

    # ---- 保存 ----
    def _read_int(self, entry, default):
        try:
            return int(entry.get().strip())
        except ValueError:
            return default

    def _save(self):
        """将界面内容写入 settings.data 并持久化到 settings.json；成功返回 True，失败返回 False。"""
        data = settings.data

        # 项目编号格式
        pattern = self._current_pattern()
        if not pattern:
            messagebox.showwarning("提示", "项目编号正则不能为空", parent=self)
            return False
        try:
            re.compile(pattern, re.IGNORECASE)
        except re.error as e:
            messagebox.showerror("错误", f"项目编号正则不合法：\n{e}", parent=self)
            return False
        data["project_pattern"] = pattern
        data["project_pattern_segments"] = None if self.advanced_var.get() else self.segments

        # 归档与子目录
        subfolder_name = self.subfolder_entry.get().strip()
        if not subfolder_name:
            messagebox.showwarning("提示", "目标子文件夹名不能为空", parent=self)
            return False
        data["target_subfolder"]["name"] = subfolder_name
        tf = self.target_filename_entry.get().strip()
        data["target_filename"] = tf if tf else "TDS"

        # 文件夹匹配
        data["folder_match"]["prefix_match"] = self.prefix_match_var.get()
        data["folder_match"]["word_boundary_match"] = self.boundary_match_var.get()

        # 年份管理
        data["year_rules"]["prefix_length"] = self._read_int(self.prefix_len_entry, 4)
        data["year_rules"]["century_prefix"] = self.century_entry.get().strip() or "20"
        data["year_rules"]["unknown_folder"] = self.unknown_folder_entry.get().strip() or "未分类"
        year_list = []
        for y in self.year_available:
            y = y.strip()
            if y and y not in year_list:
                year_list.append(y)
        data.setdefault("years", {})["available"] = sorted(year_list, key=lambda x: (len(x), x))

        # OCR 与缓存
        o = data["ocr"]
        o["languages"] = self._current_lang_codes()
        if not o["languages"]:
            messagebox.showwarning("提示", "请至少选择一种识别语言（中文/英文）", parent=self)
            return False
        o["config"] = self.config_entry.get().strip()
        o["auto_rotate"] = self.auto_rotate_var.get()
        for key, entry in self.ocr_entry_map.items():
            val = self._read_int(entry, OCR_INT_DEFAULTS.get(key, 0))
            if self.ocr_source_map[key] == "cache":
                data["cache"][key] = val
            else:
                o[key] = val

        try:
            settings.save()
        except Exception as e:
            messagebox.showerror("保存失败", str(e), parent=self)
            return False

        self._saved = True
        self.destroy()
        return True

    def _on_save_close(self):
        self._save()

    def _on_close(self):
        # 关闭窗口(X) 自动保存；保存失败则保持窗口打开以便修正
        self._save()


def open_settings_dialog(parent):
    """打开设置对话框；返回是否保存成功。"""
    dlg = SettingsDialog(parent)
    parent.wait_window(dlg)
    return dlg._saved