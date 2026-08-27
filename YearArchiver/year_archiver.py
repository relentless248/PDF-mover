import os
import re
import shutil
import sys
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 允许从项目根目录导入 core 配置模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.settings import settings, reload_settings  # noqa: E402

UNKNOWN_FOLDER = settings.year_rules.get('unknown_folder', '未分类')


def extract_project_ids_from_filename(filename):
    # 关联 + 未找到：关联A_未找到项目B.pdf
    see_match = settings.analyze_see_pattern.match(filename)
    if see_match:
        ids = [see_match.group(1), see_match.group(2)]
        return tuple(dict.fromkeys(ids))

    # 未找到：未找到项目A.pdf
    no_see_match = settings.analyze_no_see_pattern.match(filename)
    if no_see_match:
        return (no_see_match.group(1),)

    # 兜底：直接提取项目编号
    search_result = settings.project_regex.findall(filename)
    if search_result:
        return tuple(dict.fromkeys(search_result))

    return None


def extract_year_from_project_id(project_id):
    prefix_len = settings.year_rules.get('prefix_length', 4)
    century = settings.year_rules.get('century_prefix', '20')
    number_part = project_id[prefix_len:]
    if len(number_part) >= 4 and number_part[:2] == century:
        return number_part[2:4]
    elif len(number_part) >= 2:
        return number_part[:2]
    return None


def get_full_year(project_id):
    year_suffix = extract_year_from_project_id(project_id)
    if year_suffix and len(year_suffix) == 2:
        return "20" + year_suffix
    return None


class YearArchiverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("未找到项目按年份归档工具")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.source_folder = tk.StringVar()
        self.export_folder = tk.StringVar()
        self.pdf_files = []
        self.analysis_results = []

        self.create_widgets()

    def create_widgets(self):
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_container.grid_rowconfigure(3, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        src_frame = tk.LabelFrame(main_container, text="导入路径（源文件夹）", padx=8, pady=8)
        src_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        src_frame.grid_columnconfigure(1, weight=1)
        tk.Label(src_frame, text="源文件夹:").grid(row=0, column=0, sticky="w")
        tk.Entry(src_frame, textvariable=self.source_folder, width=60).grid(row=0, column=1, sticky="ew", padx=5)
        tk.Button(src_frame, text="浏览...", command=self.browse_source, width=10).grid(row=0, column=2, padx=2)
        tk.Button(src_frame, text="扫描文件", command=self.scan_files, width=10, bg='lightblue').grid(row=0, column=3, padx=2)

        dst_frame = tk.LabelFrame(main_container, text="导出路径（目标文件夹）", padx=8, pady=8)
        dst_frame.grid(row=1, column=0, sticky="ew", pady=(0, 5))
        dst_frame.grid_columnconfigure(1, weight=1)
        tk.Label(dst_frame, text="目标文件夹:").grid(row=0, column=0, sticky="w")
        tk.Entry(dst_frame, textvariable=self.export_folder, width=60).grid(row=0, column=1, sticky="ew", padx=5)
        tk.Button(dst_frame, text="浏览...", command=self.browse_export, width=10).grid(row=0, column=2, padx=2)

        info_label = tk.Label(dst_frame, text="将在目标文件夹下自动创建年份子文件夹（如 2025、2026...）",
                              fg="gray", font=("微软雅黑", 8))
        info_label.grid(row=1, column=1, sticky="w", padx=5, pady=(3, 0))

        list_frame = tk.LabelFrame(main_container, text="待归档文件预览", padx=8, pady=8)
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 5))
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(list_frame, columns=("filename", "projects", "years", "action"),
                                 show="headings", height=10)
        self.tree.heading("filename", text="文件名")
        self.tree.heading("projects", text="识别项目号")
        self.tree.heading("years", text="归档年份")
        self.tree.heading("action", text="操作类型")
        self.tree.column("filename", width=300, minwidth=150)
        self.tree.column("projects", width=280, minwidth=120)
        self.tree.column("years", width=120, minwidth=80)
        self.tree.column("action", width=150, minwidth=90)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        btn_frame = tk.Frame(main_container)
        btn_frame.grid(row=3, column=0, sticky="ew", pady=(5, 0))
        btn_frame.grid_columnconfigure(0, weight=1)

        btn_container = tk.Frame(btn_frame)
        btn_container.pack(anchor="center")
        self.btn_execute = tk.Button(btn_container, text="开始归档", command=self.execute_archive,
                                     bg='lightgreen', width=15, height=2, font=("微软雅黑", 10, "bold"))
        self.btn_execute.pack(side=tk.LEFT, padx=10)

        tk.Button(btn_container, text="退出", command=self.on_closing,
                  width=10, height=2).pack(side=tk.LEFT, padx=10)

        self.status_var = tk.StringVar(value="就绪")
        self.progress_var = tk.IntVar(value=0)
        self.progress_max_var = tk.IntVar(value=0)

        status_frame = tk.Frame(main_container)
        status_frame.grid(row=4, column=0, sticky="ew", pady=(5, 0))
        status_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var,
                                            maximum=100, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.progress_label = tk.Label(status_frame, text="", font=("微软雅黑", 8))
        self.progress_label.grid(row=0, column=1, sticky="e")

        status_bar = tk.Label(status_frame, textvariable=self.status_var, bd=1,
                              relief=tk.SUNKEN, anchor=tk.W, font=("微软雅黑", 9))
        status_bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(3, 0))

    def browse_source(self):
        folder = filedialog.askdirectory(title="选择包含\"未找到项目\"PDF的源文件夹")
        if folder:
            self.source_folder.set(folder)

    def browse_export(self):
        folder = filedialog.askdirectory(title="选择导出目标文件夹（将在此创建年份子文件夹）")
        if folder:
            self.export_folder.set(folder)

    def scan_files(self):
        src = self.source_folder.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showwarning("提示", "请先选择有效的源文件夹")
            return

        self.tree.delete(*self.tree.get_children())
        self.pdf_files = []
        self.analysis_results = []

        for f in sorted(os.listdir(src)):
            if not f.lower().endswith('.pdf'):
                continue
            full_path = os.path.join(src, f)
            self.pdf_files.append(full_path)

            project_ids = extract_project_ids_from_filename(f)

            if not project_ids:
                self.tree.insert("", tk.END, values=(f, "未识别", "未分类", "跳过"))
                self.analysis_results.append((full_path, None, None, "skip"))
                continue

            years = []
            for pid in project_ids:
                fy = get_full_year(pid)
                if fy:
                    years.append(fy)
                else:
                    years.append("未知")

            n = len(project_ids)
            if n == 1:
                action = "单归档"
            elif n == 2:
                action = "双归档（关联）"
            else:
                action = f"多项目归档({n}个)"

            projects_str = " / ".join(project_ids)
            years_str = " / ".join(years)
            self.tree.insert("", tk.END, values=(f, projects_str, years_str, action))
            self.analysis_results.append((full_path, project_ids, years, "archive"))

        self.status_var.set(f"扫描完成：找到 {len(self.pdf_files)} 个PDF文件")

    def execute_archive(self):
        if not self.analysis_results:
            messagebox.showwarning("提示", "请先点击「扫描文件」分析源文件夹")
            return

        export_root = self.export_folder.get().strip()
        if not export_root:
            messagebox.showwarning("提示", "请先选择导出目标文件夹")
            return

        os.makedirs(export_root, exist_ok=True)

        total = len(self.analysis_results)
        success_count = 0
        skip_count = 0
        fail_count = 0
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        report_lines = [f"归档报告 - {timestamp}", f"源文件夹: {self.source_folder.get()}",
                        f"目标文件夹: {export_root}", "=" * 60]

        self.progress_max_var.set(total)
        self.progress_var.set(0)
        self.progress_bar['maximum'] = total

        for idx, (full_path, project_ids, years, action) in enumerate(self.analysis_results):
            self.status_var.set(f"正在处理: {os.path.basename(full_path)} ...")
            self.progress_var.set(idx + 1)
            self.progress_label.config(text=f"{idx + 1}/{total}")
            self.root.update()

            if action == "skip":
                skip_count += 1
                self.tree.item(self.tree.get_children()[idx], tags=("skip",))
                report_lines.append(f"[跳过] {os.path.basename(full_path)} - 未识别到项目号/年份")
                continue

            file_ok = True
            processed_set = set()
            for i, pid in enumerate(project_ids):
                year = years[i] if i < len(years) else "未知"
                if year == "未知":
                    year_folder = UNKNOWN_FOLDER
                else:
                    year_folder = year

                target_dir = os.path.join(export_root, year_folder)
                os.makedirs(target_dir, exist_ok=True)

                target_name = f"{pid}.pdf"
                target_path = os.path.join(target_dir, target_name)
                counter = 1
                while os.path.exists(target_path):
                    target_path = os.path.join(target_dir, f"{pid}_{counter}.pdf")
                    counter += 1

                key = (target_path, pid)
                if key in processed_set:
                    continue
                processed_set.add(key)

                try:
                    shutil.copy2(full_path, target_path)
                except Exception as e:
                    fail_count += 1
                    file_ok = False
                    report_lines.append(f"[失败] {os.path.basename(full_path)} -> {target_path} : {e}")
                    break

                n = len(project_ids)
                if n >= 2:
                    report_lines.append(f"[多项目归档] {os.path.basename(full_path)} -> {year_folder}/{os.path.basename(target_path)} ({pid})")
                else:
                    report_lines.append(f"[归档] {os.path.basename(full_path)} -> {year_folder}/{os.path.basename(target_path)}")

            if file_ok:
                success_count += 1

        self.progress_label.config(text=f"{total}/{total}")

        report_lines.append("=" * 60)
        report_lines.append(f"总计: {total} | 成功: {success_count} | 跳过: {skip_count} | 失败: {fail_count}")
        report_text = "\n".join(report_lines)

        report_file = os.path.join(export_root, f"归档报告_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(report_file, 'w', encoding='utf-8') as rf:
            rf.write(report_text)

        self.tree.tag_configure("skip", background="#FFF3CD")

        summary = f"归档完成！\n\n成功: {success_count} 个\n跳过: {skip_count} 个（未识别）\n失败: {fail_count} 个\n\n报告已保存至:\n{report_file}"
        self.status_var.set(f"归档完成：成功 {success_count}，跳过 {skip_count}，失败 {fail_count}")
        messagebox.showinfo("归档完成", summary)

    def on_closing(self):
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = YearArchiverApp(root)
    root.mainloop()
