import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
try:
    # running as module
    from app.docx_template_apply import apply_template_to_docx
except Exception:
    # running as script/frozen
    from docx_template_apply import apply_template_to_docx

def _default_output_dir() -> str:
    # Frozen one-file exe: use current working directory
    if hasattr(sys, '_MEIPASS'):
        return os.path.abspath(os.path.join(os.getcwd(), 'output'))
    # Dev/runtime: project-root/output
    return os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'output'))

OUTPUT_DIR_DEFAULT = _default_output_dir()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("竹心Word作业模板批量套用")
        self.geometry("720x520")
        # 设置窗口图标（优先 ICO，其次 PNG）
        try:
            ico = self._resource_path(os.path.join('assets', 'icon.ico'))
            png = self._resource_path(os.path.join('assets', 'logo.png'))
            if os.path.exists(ico):
                try:
                    self.iconbitmap(ico)
                except Exception:
                    pass
            if os.path.exists(png):
                try:
                    self.iconphoto(True, tk.PhotoImage(file=png))
                except Exception:
                    pass
        except Exception:
            pass

        self.file_list = []
        self.template_path = None
        self.output_dir = os.path.abspath(OUTPUT_DIR_DEFAULT)

        self._build_ui()

    def _resource_path(self, relative: str) -> str:
        # 兼容 PyInstaller 打包后的资源路径
        if hasattr(sys, '_MEIPASS'):
            base = getattr(sys, '_MEIPASS')  # type: ignore[attr-defined]
        else:
            base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
        return os.path.join(base, relative)

    def _build_ui(self):
        # 文件列表
        frm_top = tk.Frame(self)
        frm_top.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        tk.Label(frm_top, text="待处理文档（.docx）：").pack(anchor='w')
        self.lst_files = tk.Listbox(frm_top, selectmode=tk.EXTENDED)
        self.lst_files.pack(fill=tk.BOTH, expand=True)

        btns = tk.Frame(frm_top)
        btns.pack(fill=tk.X, pady=6)
        tk.Button(btns, text="添加文档", command=self.add_files).pack(side=tk.LEFT)
        tk.Button(btns, text="移除选中", command=self.remove_selected).pack(side=tk.LEFT, padx=6)
        tk.Button(btns, text="清空", command=self.clear_files).pack(side=tk.LEFT)

        # 模板与输出
        frm_mid = tk.Frame(self)
        frm_mid.pack(fill=tk.X, padx=12, pady=4)

        tk.Label(frm_mid, text="模板文件：").grid(row=0, column=0, sticky='w')
        self.var_template = tk.StringVar(value="未选择")
        tk.Entry(frm_mid, textvariable=self.var_template, state='readonly').grid(row=0, column=1, sticky='we', padx=6)
        tk.Button(frm_mid, text="选择模板", command=self.choose_template).grid(row=0, column=2)

        tk.Label(frm_mid, text="输出目录：").grid(row=1, column=0, sticky='w', pady=(6, 0))
        self.var_outdir = tk.StringVar(value=self.output_dir)
        tk.Entry(frm_mid, textvariable=self.var_outdir).grid(row=1, column=1, sticky='we', padx=6, pady=(6, 0))
        tk.Button(frm_mid, text="浏览", command=self.choose_output_dir).grid(row=1, column=2, pady=(6, 0))
        frm_mid.columnconfigure(1, weight=1)

        # 生成按钮
        frm_bottom = tk.Frame(self)
        frm_bottom.pack(fill=tk.X, padx=12, pady=12)
        tk.Button(frm_bottom, text="开始生成", command=self.start).pack(side=tk.RIGHT)

        # 状态栏
        self.var_status = tk.StringVar(value="就绪")
        tk.Label(self, textvariable=self.var_status, anchor='w').pack(fill=tk.X, padx=12, pady=4)

    def add_files(self):
        paths = filedialog.askopenfilenames(title="选择 .docx 文档", filetypes=[("Word 文档", "*.docx")])
        if not paths:
            return
        for p in paths:
            if p not in self.file_list and p.lower().endswith('.docx'):
                self.file_list.append(p)
                self.lst_files.insert(tk.END, p)

    def remove_selected(self):
        sel = list(self.lst_files.curselection())
        sel.sort(reverse=True)
        for idx in sel:
            path = self.lst_files.get(idx)
            self.lst_files.delete(idx)
            if path in self.file_list:
                self.file_list.remove(path)

    def clear_files(self):
        self.lst_files.delete(0, tk.END)
        self.file_list.clear()

    def choose_template(self):
        p = filedialog.askopenfilename(title="选择模板 .docx", filetypes=[("Word 文档", "*.docx")])
        if p:
            self.template_path = p
            self.var_template.set(p)

    def choose_output_dir(self):
        d = filedialog.askdirectory(title="选择输出目录", mustexist=True)
        if d:
            self.output_dir = d
            self.var_outdir.set(d)

    def start(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要处理的文档")
            return
        if not self.template_path:
            messagebox.showwarning("提示", "请先选择模板文档")
            return
        outdir = self.var_outdir.get().strip() or self.output_dir
        os.makedirs(outdir, exist_ok=True)

        ok, fail = 0, 0
        self.var_status.set("处理中…")
        self.update_idletasks()
        for src in self.file_list:
            try:
                base = os.path.basename(src)
                name, ext = os.path.splitext(base)
                dst = os.path.join(outdir, f"{name}_templated{ext}")
                apply_template_to_docx(src, self.template_path, dst)
                ok += 1
            except Exception as e:
                print("处理失败:", src, e)
                fail += 1
        self.var_status.set(f"完成：成功 {ok}，失败 {fail}。")
        messagebox.showinfo("完成", f"处理完成：成功 {ok}，失败 {fail}。")


if __name__ == "__main__":
    app = App()
    app.mainloop()
