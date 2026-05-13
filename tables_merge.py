#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格汇总工具
支持两种汇总模式：
1. 按地区行汇总：将各地区数据填充到模板对应行，保持模板行顺序
2. 按A字段汇总：按指定字段分组汇总数字列，地区列自动前向填充

新增功能：支持多选Excel文件和选择文件夹，自由组合数据源
"""

import os
import sys
import platform
import warnings
warnings.filterwarnings('ignore')

# ===== 提前检查依赖 =====
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError as e:
    print("未安装 tkinter，请用以下命令安装依赖：")
    print("Linux: sudo apt-get update && sudo apt-get install python3-tk")
    print("MacOS: brew install python-tk")
    sys.exit(1)
try:
    import pandas as pd
    import numpy as np
except ImportError as e:
    print("缺少 pandas 或 numpy，请先执行：pip3 install pandas numpy openpyxl xlrd")
    sys.exit(1)
try:
    import openpyxl, xlrd
except ImportError as e:
    print("缺少 openpyxl 或 xlrd，请先执行：pip3 install openpyxl xlrd")
    sys.exit(1)
from pathlib import Path

class TableMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格汇总工具")
        self.root.geometry("800x750")
        self.root.resizable(True, True)

        # 存储所有待处理的Excel文件路径
        self.excel_files = []

        # 样式 - 自动平台字体兼容
        syst = platform.system()
        if syst == "Windows":
            default_font = ('微软雅黑', 10)
        elif syst == "Darwin":
            default_font = ('Heiti SC', 12)    # MacOS 常见中文字体
        else:
            default_font = ('Arial', 10)

        self.style = ttk.Style()
        self.style.configure('TButton', font=default_font)
        self.style.configure('TLabel', font=default_font)
        self.style.configure('TEntry', font=default_font)

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        row = 0

        ttk.Label(main_frame, text="[数据源] 请添加Excel文件或包含Excel的文件夹：").grid(row=row, column=0, sticky=tk.W, pady=5)
        row += 1

        source_btn_frame = ttk.Frame(main_frame)
        source_btn_frame.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Button(source_btn_frame, text="添加Excel文件...", command=self.add_files).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(source_btn_frame, text="添加文件夹...", command=self.add_folder).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(source_btn_frame, text="清空所有", command=self.clear_files).grid(row=0, column=2)
        row += 1

        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.files_listbox = tk.Listbox(list_frame, height=6, selectmode=tk.EXTENDED, font=('Consolas', 9))
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        del_btn_frame = ttk.Frame(main_frame)
        del_btn_frame.grid(row=row+1, column=0, columnspan=3, sticky=tk.W, pady=2)
        ttk.Button(del_btn_frame, text="删除选中", command=self.delete_selected).grid(row=0, column=0)
        row += 2

        ttk.Label(main_frame, text="[输出] 输出文件路径：").grid(row=row, column=0, sticky=tk.W, pady=(15, 5))
        row += 1

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        output_frame.columnconfigure(0, weight=1)
        self.output_path_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path_var).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(output_frame, text="浏览...", command=self.browse_output).grid(row=0, column=1)
        row += 1

        ttk.Label(main_frame, text="[表头] 表头所在行（从1开始计数）：").grid(row=row, column=0, sticky=tk.W, pady=(15, 5))
        row += 1

        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=2)
        self.header_row_var = tk.StringVar(value="1")
        ttk.Spinbox(header_frame, from_=1, to=20, textvariable=self.header_row_var, width=8).grid(row=0, column=0)
        ttk.Label(header_frame, text="  （例如：第1行是表头就填1，第2行是表头就填2）", foreground="gray").grid(row=0, column=1, padx=(10, 0))
        row += 1

        ttk.Label(main_frame, text="[标签] 代表地区的字段名称：").grid(row=row, column=0, sticky=tk.W, pady=(15, 5))
        row += 1
        self.region_col_var = tk.StringVar(value="地区")
        ttk.Entry(main_frame, textvariable=self.region_col_var, width=30).grid(row=row, column=0, sticky=tk.W, pady=2)
        row += 1

        ttk.Label(main_frame, text="[模式] 汇总模式：").grid(row=row, column=0, sticky=tk.W, pady=(15, 5))
        row += 1

        mode_frame = ttk.Frame(main_frame)
        mode_frame.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=5)
        self.mode_var = tk.IntVar(value=1)
        ttk.Radiobutton(mode_frame, text="模式1：按地区行汇总（各地区数据填充到对应行，支持文本+数字）",
                        variable=self.mode_var, value=1).grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Radiobutton(mode_frame, text="模式2：按A字段汇总（按指定字段分组求和+保留原始数据）",
                        variable=self.mode_var, value=2).grid(row=1, column=0, sticky=tk.W, pady=2)
        row += 1

        self.a_field_frame = ttk.LabelFrame(main_frame, text="模式2设置", padding="10")
        self.a_field_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        ttk.Label(self.a_field_frame, text="A字段名称（如：任务批次、月份、类别等）：").grid(row=0, column=0, sticky=tk.W)
        self.a_field_var = tk.StringVar(value="任务批次")
        ttk.Entry(self.a_field_frame, textvariable=self.a_field_var, width=25).grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        ttk.Label(self.a_field_frame, text="数字列汇总方式：").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.agg_var = tk.StringVar(value="sum")
        ttk.Combobox(self.a_field_frame, textvariable=self.agg_var, values=["sum", "mean", "count"],
                     state="readonly", width=15).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=(10, 0))
        row += 1

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=20)
        ttk.Button(btn_frame, text="[开始] 开始汇总", command=self.start_merge, width=20).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="[日志] 查看日志", command=self.show_log, width=20).grid(row=0, column=1, padx=5)
        row += 1

        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

        self.log_messages = []

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        for f in files:
            if f not in self.excel_files:
                self.excel_files.append(f)
                self.files_listbox.insert(tk.END, f)
        self.update_status()

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        if not folder:
            return
        found = 0
        for ext in ['*.xlsx', '*.xls']:
            for f in Path(folder).glob(ext):
                f_str = str(f)
                if f_str not in self.excel_files:
                    self.excel_files.append(f_str)
                    self.files_listbox.insert(tk.END, f_str)
                    found += 1
        self.log(f"从文件夹添加了 {found} 个Excel文件")
        self.update_status()

    def delete_selected(self):
        selected = self.files_listbox.curselection()
        for idx in reversed(selected):
            self.files_listbox.delete(idx)
            del self.excel_files[idx]
        self.update_status()

    def clear_files(self):
        self.files_listbox.delete(0, tk.END)
        self.excel_files.clear()
        self.update_status()

    def update_status(self):
        self.status_var.set(f"已添加 {len(self.excel_files)} 个Excel文件")

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.output_path_var.set(path)

    def log(self, msg):
        self.log_messages.append(msg)
        self.status_var.set(msg)
        self.root.update_idletasks()
        print(msg)  # 兼容无界面情况也可追踪日志

    def read_excel_with_header(self, filepath, header_row):
        pd_header = header_row - 1
        try:
            df = pd.read_excel(filepath, header=pd_header, engine='openpyxl')
        except Exception as e:
            try:
                df = pd.read_excel(filepath, header=pd_header, engine='xlrd')
            except:
                raise e
        return df

    def clean_dataframe(self, df):
        df = df.dropna(how='all').reset_index(drop=True)
        df = df.dropna(axis=1, how='all')
        return df

    def mode1_merge(self, file_paths, output_path, header_row, region_col):
        self.log("正在读取输入文件...")
        all_data = {}
        all_columns = set()

        for fp in file_paths:
            fname = Path(fp).name
            try:
                df = self.read_excel_with_header(fp, header_row)
                df = self.clean_dataframe(df)
                if region_col not in df.columns:
                    self.log(f"[警告] 跳过 {fname}：未找到字段 '{region_col}'")
                    continue
                df = df[df[region_col].notna()].copy()
                if len(df) == 0:
                    self.log(f"[警告] 跳过 {fname}：没有有效数据行")
                    continue
                all_data[fname] = df
                all_columns.update(df.columns)
                self.log(f"[成功] 读取 {fname}：{len(df)} 行")
            except Exception as e:
                self.log(f"[失败] 读取 {fname} 失败: {str(e)}")
        if not all_data:
            raise ValueError("没有成功读取任何有效数据文件")

        base_df = list(all_data.values())[0]
        final_columns = list(base_df.columns)
        for col in all_columns:
            if col not in final_columns:
                final_columns.append(col)
        all_regions = base_df[region_col].dropna().astype(str).str.strip().tolist()
        seen = set()
        ordered_regions = []
        for r in all_regions:
            if r not in seen:
                seen.add(r)
                ordered_regions.append(r)
        self.log(f"共涉及 {len(ordered_regions)} 个地区，按模板顺序")
        result_data = []
        for region in ordered_regions:
            row_data = {region_col: region}
            for name, df in all_data.items():
                mask = df[region_col].astype(str).str.strip() == region
                matched = df[mask]
                if len(matched) > 0:
                    source_row = matched.iloc[0]
                    for col in final_columns:
                        if col == region_col:
                            continue
                        if col in source_row.index and pd.notna(source_row[col]):
                            if col not in row_data or pd.isna(row_data.get(col)):
                                row_data[col] = source_row[col]
                            else:
                                try:
                                    existing = float(row_data[col])
                                    new_val = float(source_row[col])
                                    row_data[col] = existing + new_val
                                except (ValueError, TypeError):
                                    pass
            result_data.append(row_data)
        result_df = pd.DataFrame(result_data)
        for col in final_columns:
            if col not in result_df.columns:
                result_df[col] = np.nan
        result_df = result_df[final_columns]
        result_df.to_excel(output_path, index=False, engine='openpyxl')
        self.log(f"[完成] 汇总完成！已保存至: {output_path}")
        self.log(f"   共 {len(result_df)} 行，{len(result_df.columns)} 列")
        return result_df

    def mode2_merge(self, file_paths, output_path, header_row, region_col, a_field, agg_func='sum'):
        self.log("正在读取输入文件...")
        all_raw_data = []
        all_columns = set()

        for fp in file_paths:
            fname = Path(fp).name
            try:
                df = self.read_excel_with_header(fp, header_row)
                df = self.clean_dataframe(df)
                if a_field not in df.columns:
                    self.log(f"[警告] 跳过 {fname}：未找到A字段 '{a_field}'")
                    continue
                if region_col in df.columns:
                    df[region_col] = df[region_col].ffill()
                df['_数据来源'] = fname
                if region_col in df.columns:
                    df['_地区'] = df[region_col].fillna('未指定')
                else:
                    df['_地区'] = fname
                all_raw_data.append(df)
                all_columns.update(df.columns)
                self.log(f"[成功] 读取 {fname}：{len(df)} 行")
            except Exception as e:
                self.log(f"[失败] 读取 {fname} 失败: {str(e)}")
        if not all_raw_data:
            raise ValueError("没有成功读取任何有效数据文件")

        final_columns = []
        for df in all_raw_data:
            for col in df.columns:
                if col not in final_columns:
                    final_columns.append(col)
        aligned_data = []
        for df in all_raw_data:
            for col in final_columns:
                if col not in df.columns:
                    df[col] = np.nan
            aligned_data.append(df[final_columns])

        raw_combined = pd.concat(aligned_data, ignore_index=True)
        self.log(f"原始数据共 {len(raw_combined)} 行")
        numeric_cols = []
        for col in raw_combined.columns:
            if col in [a_field, '_数据来源', '_地区', region_col]:
                continue
            try:
                converted = pd.to_numeric(raw_combined[col], errors='coerce')
                if converted.notna().sum() > 0:
                    numeric_cols.append(col)
            except:
                pass
        self.log(f"识别到 {len(numeric_cols)} 个数字列: {numeric_cols}")
        if agg_func == 'sum':
            agg_dict = {col: 'sum' for col in numeric_cols}
        elif agg_func == 'mean':
            agg_dict = {col: 'mean' for col in numeric_cols}
        elif agg_func == 'count':
            agg_dict = {col: 'count' for col in numeric_cols}
        else:
            agg_dict = {col: 'sum' for col in numeric_cols}
        summary_df = raw_combined[raw_combined[a_field].notna()].copy()
        if len(summary_df) == 0:
            raise ValueError(f"没有有效数据包含A字段 '{a_field}'")
        grouped = summary_df.groupby(a_field, sort=False).agg(agg_dict).reset_index()
        count_df = summary_df.groupby(a_field, sort=False).size().reset_index(name='数据行数')
        grouped = grouped.merge(count_df, on=a_field, how='left')
        ordered_cols = [a_field, '数据行数'] + numeric_cols
        for col in grouped.columns:
            if col not in ordered_cols:
                ordered_cols.append(col)
        grouped = grouped[ordered_cols]
        self.log(f"汇总结果共 {len(grouped)} 个{a_field}")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            grouped.to_excel(writer, sheet_name='汇总数据', index=False)
            raw_output = raw_combined.drop(columns=['_数据来源', '_地区'], errors='ignore')
            raw_output.to_excel(writer, sheet_name='原始数据', index=False)

        self.log(f"[完成] 汇总完成！已保存至: {output_path}")
        self.log(f"   Sheet1 '汇总数据': {len(grouped)} 行")
        self.log(f"   Sheet2 '原始数据': {len(raw_combined)} 行")
        return grouped, raw_combined

    def start_merge(self):
        if not self.excel_files:
            messagebox.showerror("错误", "请先添加Excel文件或包含Excel的文件夹")
            return
        output_path = self.output_path_var.get().strip()
        if not output_path:
            messagebox.showerror("错误", "请选择输出文件路径")
            return
        try:
            header_row = int(self.header_row_var.get())
            if header_row < 1:
                raise ValueError()
        except:
            messagebox.showerror("错误", "表头行必须是大于等于1的整数")
            return
        region_col = self.region_col_var.get().strip()
        if not region_col:
            messagebox.showerror("错误", "请填写地区字段名称")
            return
        mode = self.mode_var.get()
        try:
            self.log("=" * 40)
            self.log("开始汇总...")
            if mode == 1:
                self.mode1_merge(self.excel_files, output_path, header_row, region_col)
            else:
                a_field = self.a_field_var.get().strip()
                if not a_field:
                    messagebox.showerror("错误", "模式2需要填写A字段名称")
                    return
                agg_func = self.agg_var.get()
                self.mode2_merge(self.excel_files, output_path, header_row, region_col, a_field, agg_func)
            messagebox.showinfo("完成", "汇总完成！\n" + self.status_var.get())
        except Exception as e:
            self.log(f"[错误] {str(e)}")
            import traceback
            print(traceback.format_exc())
            messagebox.showerror("汇总失败", str(e))

    def show_log(self):
        log_window = tk.Toplevel(self.root)
        log_window.title("运行日志")
        log_window.geometry("600x400")
        text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD, font=('Consolas', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        for msg in self.log_messages:
            text.insert(tk.END, msg + "\n")
        text.see(tk.END)

def main():
    try:
        root = tk.Tk()
        app = TableMergeApp(root)
        root.mainloop()
    except Exception as e:
        print('[致命错误]', str(e))
        import traceback
        print(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()