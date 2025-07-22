import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from core.convert import convert_word_to_pdf
import glob

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word转PDF工具")

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()

        frame_input = tk.Frame(root)
        frame_input.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_input, text="选择 Word 文件：").pack(anchor="w")
        input_entry = tk.Entry(frame_input, textvariable=self.input_path, state="readonly")
        input_entry.pack(side="left", fill="x", expand=True)
        # 浏览按钮弹出菜单
        def browse_menu(event=None):
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="选择文件", command=self.browse_file)
            menu.add_command(label="选择文件夹", command=self.browse_folder)
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        browse_btn = tk.Button(frame_input, text="浏览")
        browse_btn.pack(side="right")
        browse_btn.bind("<Button-1>", browse_menu)

        frame_output = tk.Frame(root)
        frame_output.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_output, text="输出 PDF 文件路径：").pack(anchor="w")
        output_entry = tk.Entry(frame_output, textvariable=self.output_path, state="readonly")
        output_entry.pack(fill="x", expand=True)

        tk.Button(root, text="开始转换", command=self.start_convert_thread).pack(pady=10)

        # 日志输出框
        frame_log = tk.Frame(root)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)
        tk.Label(frame_log, text="转换日志：").pack(anchor="w")
        self.log_text = ScrolledText(frame_log, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True)

        # 进度条
        frame_progress = tk.Frame(root)
        frame_progress.pack(fill="x", padx=10, pady=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame_progress, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", expand=True, side="left")
        self.progress_label = tk.Label(frame_progress, text="0/0")
        self.progress_label.pack(side="right")

        # 清空日志按钮
        tk.Button(root, text="清空日志", command=self.clear_log).pack(pady=2)

    def log_message(self, msg):
        def append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, append)

    def clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, "end")
        self.log_text.config(state="disabled")

    def set_progress(self, value, max_value):
        def update():
            self.progress_bar.config(maximum=max_value)
            self.progress_var.set(value)
            self.progress_label.config(text=f"{int(value)}/{int(max_value)}")
        self.root.after(0, update)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word 文件", "*.docx *.doc")])
        if file_path:
            self.input_path.set(file_path)
            output = os.path.splitext(file_path)[0] + ".pdf"
            self.output_path.set(output)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.input_path.set(folder_path)
            # 输出目录为所选文件夹同级 pdfDir
            output_dir = os.path.join(os.path.dirname(folder_path), "pdfDir")
            self.output_path.set(output_dir)

    def start_convert_thread(self):
        t = threading.Thread(target=self.convert, daemon=True)
        t.start()

    def convert(self):
        in_path = self.input_path.get()
        out_path = self.output_path.get()
        if not in_path:
            self.log_message("[错误] 请选择 Word 文件路径")
            return
        if os.path.isfile(in_path):
            # 单文件转换
            self.set_progress(0, 1)
            self.log_message(f"[开始] 转换文件: {in_path}")
            try:
                result = convert_word_to_pdf(in_path, out_path)
                self.set_progress(1, 1)
                self.log_message(f"[成功] {in_path} -> {result}")
            except Exception as e:
                self.set_progress(0, 1)
                self.log_message(f"[失败] {in_path} -> {out_path} ({e})")
        elif os.path.isdir(in_path):
            # 批量转换
            word_files = [
                f for f in glob.glob(os.path.join(in_path, "**"), recursive=True)
                if os.path.isfile(f) and f.lower().endswith((".doc", ".docx"))
            ]
            total = len(word_files)
            if not word_files:
                self.log_message("[提示] 该文件夹下未找到 Word 文件。")
                self.set_progress(0, 0)
                return
            self.set_progress(0, total)
            success, failed = [], []
            for idx, word_file in enumerate(word_files, 1):
                rel_path = os.path.relpath(word_file, in_path)
                rel_dir = os.path.dirname(rel_path)
                pdf_dir = os.path.join(out_path, rel_dir)
                os.makedirs(pdf_dir, exist_ok=True)
                pdf_path = os.path.join(pdf_dir, os.path.splitext(os.path.basename(word_file))[0] + ".pdf")
                self.log_message(f"[进度] 正在转换第{idx}/{total}个: {word_file}")
                try:
                    convert_word_to_pdf(word_file, pdf_path)
                    success.append(pdf_path)
                    self.log_message(f"[成功] {word_file} -> {pdf_path}")
                except Exception as e:
                    failed.append(f"{word_file} -> {pdf_path} ({e})")
                    self.log_message(f"[失败] {word_file} -> {pdf_path} ({e})")
                self.set_progress(idx, total)
            msg = f"批量转换完成！\n成功：{len(success)} 个\n失败：{len(failed)} 个"
            if failed:
                msg += "\n\n失败列表：\n" + "\n".join(failed)
            self.log_message("-" * 50)
            self.log_message("\n[结果] " + msg + "\n")
            self.log_message("-" * 50)
        else:
            self.log_message("[错误] 请选择有效的文件或文件夹路径")

def run_app():
    root = tk.Tk()
    app = App(root)
    root.mainloop() 