import os
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from excel import adjust_excel_format


def normalize_path(path_text):
    if not path_text:
        return ""
    return os.path.normpath(path_text)


def is_valid_yyyymmdd(date_text):
    if len(date_text) != 8 or not date_text.isdigit():
        return False
    try:
        datetime.datetime.strptime(date_text, "%Y%m%d")
    except ValueError:
        return False
    return True


def build_output_file_name(pro_num, pro_date):
    return f"{pro_num}_BOM_{pro_date}.xlsx"


class ExcelFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BOM Excel 格式化工具")
        self.root.geometry("800x480")

        self.input_var = tk.StringVar(value="")
        self.output_var = tk.StringVar(value="")
        self.name_var = tk.StringVar(value="")
        self.num_var = tk.StringVar(value="")
        self.date_var = tk.StringVar(value=datetime.datetime.now().strftime("%Y%m%d"))

        # 实时生成输出文件名：编号或日期变化时自动更新
        self.num_var.trace_add("write", self._auto_fill_output)
        self.date_var.trace_add("write", self._auto_fill_output)
        self.input_var.trace_add("write", self._auto_fill_output)

        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        # Input file row
        ttk.Label(container, text="输入文件").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.input_var, width=72).grid(row=0, column=1, padx=8, pady=6, sticky="ew")
        ttk.Button(container, text="选择...", command=self.choose_input).grid(row=0, column=2, pady=6)

        # Project params
        ttk.Label(container, text="项目名称").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.name_var, width=72).grid(row=1, column=1, padx=8, pady=6, sticky="ew")

        ttk.Label(container, text="项目编号").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.num_var, width=72).grid(row=2, column=1, padx=8, pady=6, sticky="ew")

        ttk.Label(container, text="项目日期").grid(row=3, column=0, sticky="w", pady=6)
        date_frame = ttk.Frame(container)
        date_frame.grid(row=3, column=1, padx=8, pady=6, sticky="w")
        ttk.Entry(date_frame, textvariable=self.date_var, width=14).pack(side=tk.LEFT)
        ttk.Label(date_frame, text="  格式: YYYYMMDD").pack(side=tk.LEFT)

        # Output file row
        ttk.Label(container, text="输出文件").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.output_var, width=72).grid(row=4, column=1, padx=8, pady=6, sticky="ew")
        ttk.Button(container, text="开始处理", command=self.run_format).grid(row=4, column=2, pady=6)

        # Log area
        ttk.Label(container, text="日志").grid(row=5, column=0, sticky="w", pady=(10, 6))
        self.log_text = tk.Text(container, height=12, width=92)
        self.log_text.grid(row=6, column=0, columnspan=3, sticky="nsew")

        container.rowconfigure(6, weight=1)
        container.columnconfigure(1, weight=1)

    def choose_input(self):
        path = filedialog.askopenfilename(
            title="选择输入 Excel",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xltx *.xltm"), ("All Files", "*.*")],
        )
        if path:
            self.input_var.set(normalize_path(path))

    def _auto_fill_output(self, *_):
        pro_num = self.num_var.get().strip()
        pro_date = self.date_var.get().strip()
        if not pro_num or not is_valid_yyyymmdd(pro_date):
            return
        output_name = build_output_file_name(pro_num, pro_date)
        input_path = normalize_path(self.input_var.get().strip())
        base_dir = os.path.dirname(os.path.abspath(input_path)) if input_path else os.getcwd()
        self.output_var.set(normalize_path(os.path.join(base_dir, output_name)))

    def run_format(self):
        input_file = normalize_path(self.input_var.get().strip())
        output_file = normalize_path(self.output_var.get().strip())
        pro_name = self.name_var.get().strip()
        pro_num = self.num_var.get().strip()
        pro_date = self.date_var.get().strip()

        if not input_file or not os.path.isfile(input_file):
            messagebox.showerror("参数错误", "输入文件不存在")
            return
        if not pro_name:
            messagebox.showerror("参数错误", "项目名称不能为空")
            return
        if not pro_num:
            messagebox.showerror("参数错误", "项目编号不能为空")
            return
        if not is_valid_yyyymmdd(pro_date):
            messagebox.showerror("参数错误", "项目日期格式错误，请使用 YYYYMMDD")
            return

        if not output_file:
            output_file = build_output_file_name(pro_num, pro_date)
            output_file = os.path.join(os.path.dirname(os.path.abspath(input_file)), output_file)
            self.output_var.set(output_file)

        try:
            max_row, max_col = adjust_excel_format(
                input_file,
                output_file,
                pro_name=pro_name,
                pro_num=pro_num,
                pro_date=pro_date,
            )
            self.append_log("处理完成")
            self.append_log(f"输入文件: {input_file}")
            self.append_log(f"输出文件: {output_file}")
            self.append_log(f"项目名称: {pro_name}")
            self.append_log(f"项目编号: {pro_num}")
            self.append_log(f"项目日期: {pro_date}")
            self.append_log(f"工作表规模: {max_row} 行, {max_col} 列")
            self.append_log("-" * 60)
            messagebox.showinfo("成功", "Excel 处理完成")
        except Exception as exc:
            self.append_log(f"处理失败: {exc}")
            messagebox.showerror("失败", str(exc))

    def append_log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)


def main():
    root = tk.Tk()
    app = ExcelFormatterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
