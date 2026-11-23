"""
数据概览报告工具 - 图形界面版 v3.0
支持格式: CSV / JSON / Excel (.xlsx, .xls)
打包: pyinstaller --onefile --windowed --name DataReportGUI data_report_gui.py
"""

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import pandas as pd
import numpy as np
import json
import os
import sys
import threading

SUPPORTED_EXT = {
    '.csv': 'CSV',
    '.json': 'JSON',
    '.xlsx': 'Excel',
    '.xls': 'Excel',
}


def load_file(file_path):
    """加载数据文件，返回 (df, file_type, extra_info)"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in SUPPORTED_EXT:
        raise ValueError(f"不支持的文件格式: {ext}\n支持的格式: {', '.join(SUPPORTED_EXT.keys())}")

    file_type = SUPPORTED_EXT[ext]
    extra_info = ""

    if ext == '.csv':
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding='gbk')
            extra_info = "（使用 GBK 编码读取）"

    elif ext == '.json':
        try:
            df = pd.read_json(file_path)
        except ValueError:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                df = pd.DataFrame(data)
            elif isinstance(data, dict):
                found = False
                for key, val in data.items():
                    if isinstance(val, list):
                        df = pd.DataFrame(val)
                        found = True
                        break
                if not found:
                    df = pd.DataFrame([data])
            else:
                df = pd.DataFrame([data])

    else:  # Excel
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        if len(sheet_names) > 1:
            extra_info = f"（检测到 {len(sheet_names)} 个工作表，读取第一个: '{sheet_names[0]}'）"
        df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl' if ext == '.xlsx' else None)

    return df, file_type, extra_info


def infer_type_label(series):
    clean = series.dropna()
    if len(clean) == 0:
        return "全部缺失"
    if pd.api.types.is_numeric_dtype(series):
        return "数字"
    if pd.api.types.is_bool_dtype(series):
        return "布尔"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "日期"
    if series.dtype == 'object':
        has_number = False
        has_string = False
        for val in clean.head(100):
            if isinstance(val, (int, float, np.integer, np.floating)):
                has_number = True
            elif isinstance(val, str):
                try:
                    float(val)
                    has_number = True
                except ValueError:
                    has_string = True
            else:
                has_string = True
        if has_number and has_string:
            return "混合"
        elif has_number:
            return "数字(文本)"
        else:
            return "文本"
    return str(series.dtype)


def detect_issue(series):
    clean = series.dropna()
    if len(clean) == 0:
        return "全部缺失"
    if series.dtype == 'object':
        has_number = False
        has_string = False
        for val in clean.head(100):
            if isinstance(val, (int, float, np.integer, np.floating)):
                has_number = True
            elif isinstance(val, str):
                try:
                    float(val)
                    has_number = True
                except ValueError:
                    has_string = True
            else:
                has_string = True
        if has_number and has_string:
            return "⚠️ 混合数据类型"
    if pd.api.types.is_numeric_dtype(series) and len(clean) > 0:
        try:
            q1 = clean.quantile(0.25)
            q3 = clean.quantile(0.75)
            iqr = q3 - q1
            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr
            outliers = ((clean < lower) | (clean > upper)).sum()
            if outliers > 0:
                return f"⚠️ {outliers} 个疑似异常值"
        except Exception:
            pass
    return "✓ 正常"


def truncate_value(val, max_len=30):
    if isinstance(val, str) and len(val) > max_len:
        return val[:max_len] + "…"
    return val


def build_report_text(df, file_type, file_name, extra_info=""):
    """生成纯文本报告"""
    total_rows = len(df)
    total_cols = len(df.columns)
    rows_with_missing = df.isnull().any(axis=1).sum()

    lines = []
    lines.append("=" * 70)
    lines.append(f"  数据概览报告: {file_name} {extra_info}")
    lines.append("=" * 70)

    # 基本信息
    lines.append("\n【基本信息】")
    lines.append(f"  文件格式   : {file_type}")
    lines.append(f"  行数       : {total_rows:,}")
    lines.append(f"  列数       : {total_cols:,}")
    if total_rows > 0:
        lines.append(f"  含缺失行数 : {rows_with_missing:,}  ({rows_with_missing / total_rows * 100:.1f}%)")
    else:
        lines.append(f"  含缺失行数 : 0")

    # 列详细信息
    lines.append("\n【列详细信息】")
    lines.append("-" * 75)
    lines.append(f"{'列名':<22} {'类型':<12} {'缺失数':>8} {'缺失率':>8}   {'潜在问题'}")
    lines.append("-" * 75)

    high_missing = []
    for col in df.columns:
        missing_count = df[col].isnull().sum()
        missing_pct = (missing_count / total_rows * 100) if total_rows > 0 else 0
        dtype = infer_type_label(df[col])
        issue = detect_issue(df[col])
        col_display = str(col)[:20] + "…" if len(str(col)) > 22 else str(col)
        lines.append(f"{col_display:<22} {dtype:<12} {missing_count:>8} {missing_pct:>7.1f}%   {issue}")
        if missing_pct >= 30:
            high_missing.append(str(col))

    lines.append("-" * 75)

    if high_missing:
        lines.append(f"\n⚠️ 警告: 以下列缺失率超过 30% — {', '.join(high_missing)}")

    # 前10行预览
    lines.append("\n【数据预览 — 前 10 行】")
    preview_df = df.head(10).copy()
    for col in preview_df.columns:
        preview_df[col] = preview_df[col].map(truncate_value)

    with pd.option_context('display.max_columns', None,
                           'display.width', 200,
                           'display.max_colwidth', 30):
        lines.append(preview_df.to_string(index=False))

    lines.append("\n" + "=" * 70)
    lines.append("  报告结束")
    lines.append("=" * 70)

    return "\n".join(lines)


class DataReportGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("数据概览报告工具 v3.0")
        self.root.geometry("1000x750")
        self.root.configure(bg="#f7f9fc")
        self.df = None
        self.file_type = None
        self.current_file = None

        # 标题
        title_frame = tk.Frame(root, bg="#f7f9fc")
        title_frame.pack(pady=(20, 10), fill=tk.X, padx=30)
        tk.Label(title_frame, text="📋 数据概览报告工具", font=("Segoe UI", 20, "bold"),
                 bg="#f7f9fc", fg="#0b1e33").pack(anchor="w")
        tk.Label(title_frame, text="支持 CSV / JSON / Excel — 快速查看格式、行数、缺失值、潜在错误",
                 font=("Segoe UI", 10), bg="#f7f9fc", fg="#4a5e74").pack(anchor="w", pady=(4, 0))

        # 文件选择区
        file_frame = tk.Frame(root, bg="#f0f4fa", bd=2, relief=tk.RIDGE)
        file_frame.pack(pady=15, padx=30, fill=tk.X)

        btn_frame = tk.Frame(file_frame, bg="#f0f4fa")
        btn_frame.pack(pady=15)

        self.select_btn = tk.Button(btn_frame, text="📂 选择文件", font=("Segoe UI", 12),
                                    bg="#1e3d5b", fg="white", padx=30, pady=8,
                                    activebackground="#0f2b41", activeforeground="white",
                                    relief=tk.FLAT, cursor="hand2", command=self.select_file)
        self.select_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.export_btn = tk.Button(btn_frame, text="💾 导出报告", font=("Segoe UI", 12),
                                    bg="#2a7faa", fg="white", padx=30, pady=8,
                                    activebackground="#1e6a8f", activeforeground="white",
                                    relief=tk.FLAT, cursor="hand2", command=self.export_report,
                                    state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT)

        self.file_label = tk.Label(file_frame, text="尚未选择文件", font=("Segoe UI", 10),
                                   bg="#f0f4fa", fg="#6c839b")
        self.file_label.pack(pady=(0, 10))

        # 报告显示区（可滚动文本框）
        report_frame = tk.Frame(root, bg="#f7f9fc")
        report_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 20))

        self.report_text = scrolledtext.ScrolledText(report_frame, wrap=tk.NONE,
                                                      font=("Consolas", 10),
                                                      bg="white", fg="#203e59",
                                                      relief=tk.SOLID, bd=1)
        self.report_text.pack(fill=tk.BOTH, expand=True)
        self.report_text.insert(tk.END, "⬆️ 请点击上方按钮选择数据文件，报告将显示在此处。\n")
        self.report_text.config(state=tk.DISABLED)

        # 底部状态栏
        self.status_label = tk.Label(root, text="就绪", font=("Segoe UI", 9),
                                     bg="#f7f9fc", fg="#6c839b", anchor="w")
        self.status_label.pack(fill=tk.X, padx=30, pady=(0, 10))

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[
                ("所有支持的格式", "*.csv *.json *.xlsx *.xls"),
                ("CSV 文件", "*.csv"),
                ("JSON 文件", "*.json"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件", "*.*"),
            ]
        )
        if not file_path:
            return

        self.file_label.config(text=f"⏳ 正在加载: {os.path.basename(file_path)} …")
        self.status_label.config(text="正在加载文件…")
        self.select_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)

        # 使用线程避免界面卡顿（大文件读取）
        threading.Thread(target=self._load_and_render, args=(file_path,), daemon=True).start()

    def _load_and_render(self, file_path):
        try:
            df, file_type, extra_info = load_file(file_path)
            if df.empty:
                self.root.after(0, lambda: self._show_error("文件内容为空或无法解析出数据列。"))
                return

            self.df = df
            self.file_type = file_type
            self.current_file = file_path

            report = build_report_text(df, file_type, os.path.basename(file_path), extra_info)
            self.root.after(0, lambda: self._update_ui(report, file_path))

        except Exception as e:
            self.root.after(0, lambda: self._show_error(f"读取文件失败: {str(e)}"))

    def _update_ui(self, report, file_path):
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)
        self.report_text.config(state=tk.DISABLED)

        self.file_label.config(text=f"📄 {os.path.basename(file_path)}")
        self.status_label.config(text=f"✅ 报告已生成 — {os.path.basename(file_path)}")
        self.select_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.NORMAL)

    def _show_error(self, msg):
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, f"❌ {msg}\n")
        self.report_text.config(state=tk.DISABLED)
        self.file_label.config(text="加载失败")
        self.status_label.config(text=f"❌ {msg}")
        self.select_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        messagebox.showerror("错误", msg)

    def export_report(self):
        """将当前报告保存为 txt 文件"""
        if self.df is None:
            return

        default_name = os.path.splitext(os.path.basename(self.current_file))[0] + "_report.txt"
        save_path = filedialog.asksaveasfilename(
            title="保存报告",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not save_path:
            return

        try:
            content = self.report_text.get(1.0, tk.END).strip()
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write(content)
            self.status_label.config(text=f"✅ 报告已保存: {save_path}")
            messagebox.showinfo("成功", f"报告已保存到:\n{save_path}")
        except Exception as e:
            messagebox.showerror("保存失败", f"保存报告时出错:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataReportGUI(root)
    root.mainloop()
