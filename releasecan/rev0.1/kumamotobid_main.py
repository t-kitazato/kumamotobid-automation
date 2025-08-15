import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import subprocess

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_var.set(folder_path)

def select_excel_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        excel_var.set(file_path)

def select_script_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Python files", "*.py")]
    )
    if file_path:
        script_var.set(file_path)

def run_process():
    folder = folder_var.get()
    excel_file = excel_var.get()
    start_date = start_date_var.get()
    end_date = end_date_var.get()
    script_file = script_choice_var.get()  # ラジオボタンの値を取得

    if not folder or not excel_file or not start_date or not end_date or not script_file:
        messagebox.showwarning("入力不足", "保存先、Excelファイル、開始日、終了日、スクリプトファイルを選択してください。")
        print(f"sys.argv: {sys.argv}")
        return

    try:
        # subprocess.run(["python", script_file, folder, excel_file, start_date, end_date], check=True)
        subprocess.run([
            "python",
            script_file,
            excel_file,         # 例: D:/work/kumamotobid/北里道路_入札候補案件.xlsx
            # start_date = sys.argv[3]
            # end_date = sys.argvv[4]
            folder_var.get()     # 例: D:/work/kumamotobid
        ],check=True)

        messagebox.showinfo("実行", "スクリプトが正常に実行されました。")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("エラー", f"スクリプト実行中にエラーが発生しました: {e}")

# GUIの構築
root = tk.Tk()
root.title("入札情報取得フォーム")

folder_var = tk.StringVar()
excel_var = tk.StringVar()
start_date_var = tk.StringVar()
end_date_var = tk.StringVar()
script_var = tk.StringVar()

tk.Label(root, text="📁 保存先フォルダ:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
tk.Entry(root, textvariable=folder_var, width=50).grid(row=0, column=1, padx=5)
tk.Button(root, text="選択", command=select_folder).grid(row=0, column=2, padx=5)

tk.Label(root, text="📋 施工番号Excel:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
tk.Entry(root, textvariable=excel_var, width=50).grid(row=1, column=1, padx=5)
tk.Button(root, text="選択", command=select_excel_file).grid(row=1, column=2, padx=5)

tk.Label(root, text="📅 開始日:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
DateEntry(root, textvariable=start_date_var, width=12).grid(row=2, column=1, padx=5)

tk.Label(root, text="📅 終了日:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
DateEntry(root, textvariable=end_date_var, width=12).grid(row=3, column=1, padx=5)

# スクリプト選択用ラジオボタンの追加
script_choice_var = tk.StringVar(value="kumamotopre.py")

tk.Label(root, text="🐍 実行スクリプト:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
tk.Radiobutton(root, text="入札情報更新", variable=script_choice_var, value="kumamotopre.py").grid(row=4, column=1, sticky="w")
tk.Radiobutton(root, text="仕様書ダウンロード", variable=script_choice_var, value="announcement_info.py").grid(row=4, column=2, sticky="w")

tk.Button(root, text="▶️ 実行", command=run_process, bg="#4CAF50", fg="white").grid(row=5, column=1, pady=15)

root.mainloop()