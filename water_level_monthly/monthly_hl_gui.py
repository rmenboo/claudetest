"""月別最高・最低水位 集計ツール(Tkinter GUI)。

monthly_hl.py と同じフォルダに置いて、このファイルを実行する:
    python monthly_hl_gui.py
"""
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import monthly_hl  # 集計の計算部分(同じフォルダに置く)

# 結果表に表示する列。(列名, 幅, 配置)
COLUMNS = [
    ("年月", 70, "center"),
    ("最高水位(m)", 90, "e"),
    ("最高の日時", 130, "center"),
    ("最低水位(m)", 90, "e"),
    ("最低の日時", 130, "center"),
    ("有効データ数", 90, "e"),
    ("暫定値を含む", 80, "center"),
    ("注意", 80, "center"),
]
NEED_RECALC = "再計算が必要です。「集計実行」を押してください。"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("月別最高・最低水位 集計ツール")
        self.result = None  # 最新の集計結果(DataFrame)。入力を変えるとNoneに戻す

        # --- ウィンドウを画面中央に配置 ---
        w, h = 1000, 560
        x = int((root.winfo_screenwidth() - w) / 2)
        y = int((root.winfo_screenheight() - h) / 2)
        root.geometry(f"{w}x{h}+{x}+{y}")
        root.minsize(800, 450)

        # --- 入力変数(変更されたら結果をクリアする) ---
        self.folder_var = tk.StringVar()
        self.attr_vars = {  # 集計に使う属性。0は既定でON
            "0": tk.BooleanVar(value=True),
            "2": tk.BooleanVar(value=False),
            "8": tk.BooleanVar(value=False),
        }
        self.folder_var.trace_add("write", self.clear_results)
        for v in self.attr_vars.values():
            v.trace_add("write", self.clear_results)

        # --- 2ペイン構造: 左=入力、右=結果 ---
        main = ttk.Frame(root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
        self.build_input_pane(main)
        self.build_result_pane(main)
        self.clear_results()

    # ---------------- 左ペイン: 入力 ----------------
    def build_input_pane(self, parent):
        pane = ttk.Frame(parent)
        pane.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        g1 = ttk.LabelFrame(pane, text="データの場所", padding=10)
        g1.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(g1, text="ZIPを入れたフォルダ:").grid(row=0, column=0, sticky="w")
        ttk.Entry(g1, textvariable=self.folder_var, width=38).grid(row=1, column=0, sticky="ew", pady=2)
        ttk.Button(g1, text="参照...", command=self.choose_folder).grid(row=1, column=1, padx=(5, 0))

        g2 = ttk.LabelFrame(pane, text="集計に使う属性", padding=10)
        g2.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Checkbutton(g2, text="0: 有効", variable=self.attr_vars["0"]).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(g2, text="2: 推定値", variable=self.attr_vars["2"]).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(g2, text="8: 工事中水位", variable=self.attr_vars["8"]).grid(row=2, column=0, sticky="w")
        ttk.Label(g2, text="※ 欠測(1)・閉局(3)・-9999は常に除外", foreground="gray").grid(
            row=3, column=0, sticky="w", pady=(6, 0))

        self.run_btn = ttk.Button(pane, text="集計実行", command=self.run)
        self.run_btn.grid(row=2, column=0, sticky="ew", pady=(4, 4), ipady=4)
        self.save_btn = ttk.Button(pane, text="Excelに保存...", command=self.save, state="disabled")
        self.save_btn.grid(row=3, column=0, sticky="ew", ipady=4)

        # 確認用メッセージ(読み込み件数・属性件数など)
        g3 = ttk.LabelFrame(pane, text="確認用メッセージ", padding=5)
        g3.grid(row=4, column=0, sticky="nsew", pady=(8, 0))
        pane.rowconfigure(4, weight=1)
        self.log = tk.Text(g3, width=44, height=8, wrap="word", font=("Courier New", 9), state="disabled")
        self.log.pack(fill="both", expand=True)

    # ---------------- 右ペイン: 結果 ----------------
    def build_result_pane(self, parent):
        pane = ttk.LabelFrame(parent, text="月別 最高・最低水位", padding=5)
        pane.grid(row=0, column=1, sticky="nsew")
        pane.columnconfigure(0, weight=1)
        pane.rowconfigure(1, weight=1)

        # 状態表示(再計算が必要か等)
        self.status = ttk.Label(pane, text="", font=("Arial", 10, "bold"))
        self.status.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        self.tree = ttk.Treeview(pane, columns=[c[0] for c in COLUMNS], show="headings")
        for name, width, anchor in COLUMNS:
            self.tree.heading(name, text=name)
            self.tree.column(name, width=width, anchor=anchor)
        self.tree.tag_configure("warn", foreground="red")
        sb = ttk.Scrollbar(pane, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

    # ---------------- 動作 ----------------
    def choose_folder(self):
        d = filedialog.askdirectory(title="ZIPを入れたフォルダを選択")
        if d:
            self.folder_var.set(d)

    def clear_results(self, *args):
        """入力が変わったら、古い結果を消して再計算を促す(再計算忘れ防止)。"""
        self.result = None
        if hasattr(self, "tree"):
            self.tree.delete(*self.tree.get_children())
            self.status.config(text=NEED_RECALC, foreground="red")
            self.save_btn.config(state="disabled")

    def set_log(self, lines):
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.insert("1.0", "\n".join(lines))
        self.log.config(state="disabled")

    def run(self):
        folder = Path(self.folder_var.get().strip())
        if not folder.is_dir():
            messagebox.showwarning("入力エラー", "ZIPを入れたフォルダを選択してください。")
            return
        attrs = tuple(k for k, v in self.attr_vars.items() if v.get())
        if not attrs:
            messagebox.showwarning("入力エラー", "集計に使う属性を1つ以上選んでください。")
            return

        self.root.config(cursor="watch")
        self.status.config(text="集計中...", foreground="black")
        self.root.update_idletasks()
        try:
            result, msgs = monthly_hl.aggregate(folder, attrs)
        except Exception as e:  # 想定外の形式のCSVなどもここで受け止める
            self.set_log([f"エラー: {e}"])
            self.status.config(text=NEED_RECALC, foreground="red")
            messagebox.showerror("集計エラー", f"集計に失敗しました。\n\n{e}")
            return
        finally:
            self.root.config(cursor="")

        self.result = result
        self.set_log(msgs + [f"集計に使った属性: {', '.join(attrs)}"])
        self.tree.delete(*self.tree.get_children())
        for _, r in result.iterrows():
            vals = [
                r["年月"], f"{r['最高水位(m)']:.2f}", r["最高の日時"],
                f"{r['最低水位(m)']:.2f}", r["最低の日時"],
                int(r["有効データ数"]), "*" if r["暫定値を含む"] else "", r["注意"],
            ]
            self.tree.insert("", "end", values=vals, tags=("warn",) if r["注意"] else ())
        self.status.config(text=f"集計完了: {len(result)}か月分", foreground="green")
        self.save_btn.config(state="normal")

    def save(self):
        if self.result is None:
            messagebox.showwarning("保存できません", NEED_RECALC)
            return
        path = filedialog.asksaveasfilename(
            title="Excelに保存", defaultextension=".xlsx", filetypes=[("Excelブック", "*.xlsx")],
            initialdir=self.folder_var.get(), initialfile="月別最高最低.xlsx")
        if not path:
            return
        try:
            monthly_hl.save_excel(self.result, Path(path))
        except PermissionError:
            messagebox.showerror("保存エラー", "保存できません。同名のExcelファイルを開いていないか確認してください。")
            return
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存に失敗しました。\n\n{e}")
            return
        messagebox.showinfo("保存完了", f"保存しました。\n{path}")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
