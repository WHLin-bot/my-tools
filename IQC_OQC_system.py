import os
import json
import csv
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime

# ==========================================
# 1. 路徑鎖定
# ==========================================
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HISTORY_DIR = os.path.join(DB_PATH, "IOQC_folder_history")
HIST_FILE = os.path.join(HISTORY_DIR, "IOQC_folder_history.txt")
DONE_FILE = os.path.join(HISTORY_DIR, "IOQC_folder_history_done.txt")

os.makedirs(HISTORY_DIR, exist_ok=True)

# ----------------- 功能邏輯 -----------------

def get_folder_status(sn_path):
    """檢查照片狀態，回傳 (文字, 顏色, 是否已完成)"""
    if not os.path.exists(sn_path): return "路徑遺失", "red", False
    has_iqc, has_oqc = False, False
    try:
        for folder in os.listdir(sn_path):
            full_path = os.path.join(sn_path, folder)
            if os.path.isdir(full_path):
                has_files = any(os.path.isfile(os.path.join(full_path, f)) for f in os.listdir(full_path))
                if "IQC" in folder and has_files: has_iqc = True
                if "OQC" in folder and has_files: has_oqc = True
    except: pass
    if has_iqc and has_oqc: return "✅ 已完成", "green", True
    return "❌ 未完成", "red", False

def load_json_data(file_path):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f: return json.load(f)
        except: return []
    return []

def refresh_all_lists(event=None):
    """同時刷新兩個分頁"""
    update_tree_display(HIST_FILE, tree_todo, entry_search_todo, is_todo=True)
    update_tree_display(DONE_FILE, tree_done, entry_search_done, is_todo=False)

def update_tree_display(file_path, target_tree, search_entry, is_todo=True):
    query = search_entry.get().strip().lower()
    for item in target_tree.get_children(): target_tree.delete(item)
    records = load_json_data(file_path)
    for r in reversed(records):
        if any(query in str(r.get(k, "")).lower() for k in ['sn', 'customer', 'cust_id']):
            status_text, color_tag, is_done = get_folder_status(r['path'])
            if is_todo:
                check_mark = "☑" if is_done else "☐"
                target_tree.insert("", "end", values=(check_mark, r['time'], r['customer'], r.get('cust_id', 'N/A'), r['model'], r['sn'], r.get('staff', 'N/A'), status_text, r['path']), tags=(color_tag,))
            else:
                target_tree.insert("", "end", values=(r['time'], r['customer'], r.get('cust_id', 'N/A'), r['model'], r['sn'], r.get('staff', 'N/A'), "📁 已結案", r['path']))

def export_csv_action():
    """匯出勾選資料並搬移至 Done 檔案"""
    selected_data = []
    for item_id in tree_todo.get_children():
        vals = tree_todo.item(item_id, 'values')
        if vals[0] == "☑": selected_data.append(vals)

    if not selected_data:
        messagebox.showwarning("提示", "請先勾選待處理清單中已完成的項目！")
        return

    out_file = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=f"IQC稽核待貼上資料_{datetime.now().strftime('%Y%m%d')}.csv")
    if not out_file: return

    try:
        with open(out_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["客戶名稱", "產品型號", "SN編號", "客戶編號", "", "建立日期", "作業人員"])
            for r in selected_data:
                # 索引：2客戶, 4型號, 5SN, 3編號, 1日期, 6人員
                writer.writerow([r[2], r[4], r[5], r[3], "", r[1], r[6]])

        # 搬移資料
        done_sns = [r[5] for r in selected_data]
        all_todo = load_json_data(HIST_FILE)
        to_keep = [r for r in all_todo if r['sn'] not in done_sns]
        to_done = [r for r in all_todo if r['sn'] in done_sns]

        with open(HIST_FILE, "w", encoding="utf-8") as f: json.dump(to_keep, f, ensure_ascii=False, indent=4)
        all_finished = load_json_data(DONE_FILE)
        all_finished.extend(to_done)
        with open(DONE_FILE, "w", encoding="utf-8") as f: json.dump(all_finished, f, ensure_ascii=False, indent=4)

        messagebox.showinfo("成功", "CSV 匯出成功！資料已移至結案清單。")
        refresh_all_lists()
    except Exception as e: messagebox.showerror("錯誤", f"匯出失敗：{e}")

def toggle_todo_check(event):
    item_id = tree_todo.identify_row(event.y)
    column = tree_todo.identify_column(event.x)
    if column == "#1" and item_id:
        vals = list(tree_todo.item(item_id, 'values'))
        vals[0] = "☑" if vals[0] == "☐" else "☐"
        tree_todo.item(item_id, values=vals)

def create_folders():
    customer, model, sn, cust_id, staff = entry_customer.get().strip(), entry_model.get().strip(), entry_sn.get().strip(), entry_cust_id.get().strip(), entry_staff.get().strip()
    if not all([customer, model, sn, cust_id, staff]):
        messagebox.showwarning("提示", "請填寫所有欄位！")
        return
    sn_path = os.path.join(DB_PATH, customer, model, sn)
    today = datetime.now().strftime('%Y%m%d')
    try:
        os.makedirs(os.path.join(sn_path, f"{today}_IQC"), exist_ok=True)
        os.makedirs(os.path.join(sn_path, f"{today}_OQC"), exist_ok=True)
        
        records = load_json_data(HIST_FILE)
        records.append({"time": datetime.now().strftime('%Y-%m-%d %H:%M'), "customer": customer, "cust_id": cust_id, "model": model, "sn": sn, "staff": staff, "path": sn_path})
        with open(HIST_FILE, "w", encoding="utf-8") as f: json.dump(records, f, ensure_ascii=False, indent=4)
        
        messagebox.showinfo("成功", "資料夾已建立！")
        os.startfile(sn_path)
        refresh_all_lists()
    except Exception as e: messagebox.showerror("錯誤", f"建立失敗：{e}")

def open_path_from_tree(tree_obj):
    selected = tree_obj.selection()
    if selected:
        path = tree_obj.item(selected, "values")[-1]
        if os.path.exists(path): os.startfile(path)

# --- UI 介面 ---
root = tk.Tk()
root.title("IQC/OQC 稽核管理系統 v5.0")
root.geometry("1200x850")

FONT_LABEL, FONT_ENTRY, FONT_BTN = ("微軟正黑體", 12), ("微軟正黑體", 12), ("微軟正黑體", 12, "bold")

# 重要：解決 EXE 打包後底色消失的問題
style = ttk.Style()
style.theme_use('clam') # 強制使用 'clam' 主題以支援標籤底色

style.configure("TNotebook.Tab", font=FONT_LABEL, padding=[10, 5])
style.configure("Treeview.Heading", font=FONT_BTN)
style.configure("Treeview", font=("微軟正黑體", 11), rowheight=35)

# 1. 輸入區
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_LABEL, padx=10, pady=10)
frame_input.pack(fill="x", padx=20, pady=10)
tk.Label(frame_input, text="客戶:", font=FONT_LABEL).grid(row=0, column=0)
entry_customer = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_customer.grid(row=0, column=1, padx=5)
tk.Label(frame_input, text="型號:", font=FONT_LABEL).grid(row=0, column=2)
entry_model = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_model.grid(row=0, column=3, padx=5)
tk.Label(frame_input, text="SN:", font=FONT_LABEL).grid(row=1, column=0)
entry_sn = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_sn.grid(row=1, column=1, padx=5)
tk.Label(frame_input, text="客戶編號:", font=FONT_LABEL).grid(row=1, column=2)
entry_cust_id = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_cust_id.grid(row=1, column=3, padx=5)
tk.Label(frame_input, text="人員:", font=FONT_LABEL).grid(row=1, column=4)
entry_staff = tk.Entry(frame_input, width=10, font=FONT_ENTRY); entry_staff.grid(row=1, column=5, padx=5)
tk.Button(frame_input, text="➕ 建立並開啟", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BTN).grid(row=0, column=6, rowspan=2, padx=15, sticky="nswe")

# 2. 分頁區
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=20, pady=10)

# --- Tab 1: 待處理 ---
tab_todo = tk.Frame(notebook)
notebook.add(tab_todo, text="  🕒 待處理清單 (History)  ")
tools_todo = tk.Frame(tab_todo, pady=5); tools_todo.pack(fill="x")
entry_search_todo = tk.Entry(tools_todo, font=FONT_ENTRY); entry_search_todo.pack(side="left", fill="x", expand=True, padx=5)
entry_search_todo.bind("<KeyRelease>", lambda e: update_tree_display(HIST_FILE, tree_todo, entry_search_todo, True))
tk.Button(tools_todo, text="📤 匯出CSV並結案", command=export_csv_action, bg="#27AE60", fg="white", font=FONT_BTN).pack(side="right", padx=5)
tk.Button(tools_todo, text="🔄 更新", command=refresh_all_lists, font=FONT_LABEL).pack(side="right", padx=5)

tree_todo = ttk.Treeview(tab_todo, columns=("選取", "建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "狀態", "路徑"), show="headings")
for col in ["選取", "建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "狀態"]: tree_todo.heading(col, text=col)
tree_todo.column("選取", width=50, anchor="center"); tree_todo.column("路徑", width=0, stretch=False)

# 定義顏色標籤 (需與更新顯示邏輯中的 color_tag 對應)
tree_todo.tag_configure("green", background="#DFF2BF", foreground="#270")
tree_todo.tag_configure("orange", background="#FEEFB3", foreground="#9F6000")
tree_todo.tag_configure("red", background="#FFBABA", foreground="#D8000C")

tree_todo.pack(fill="both", expand=True)
tree_todo.bind("<Button-1>", toggle_todo_check)
tree_todo.bind("<Double-1>", lambda e: open_path_from_tree(tree_todo))

# --- Tab 2: 已結案 ---
tab_done = tk.Frame(notebook)
notebook.add(tab_done, text="  ✅ 已結案清單 (Done)  ")
tools_done = tk.Frame(tab_done, pady=5); tools_done.pack(fill="x")
entry_search_done = tk.Entry(tools_done, font=FONT_ENTRY); entry_search_done.pack(side="left", fill="x", expand=True, padx=5)
entry_search_done.bind("<KeyRelease>", lambda e: update_tree_display(DONE_FILE, tree_done, entry_search_done, False))
tk.Button(tools_done, text="🔄 刷新", command=refresh_all_lists, font=FONT_LABEL).pack(side="right", padx=5)

tree_done = ttk.Treeview(tab_done, columns=("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "狀態", "路徑"), show="headings")
for col in ["建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "狀態"]: tree_done.heading(col, text=col)
tree_done.column("路徑", width=0, stretch=False); tree_done.pack(fill="both", expand=True)
tree_done.bind("<Double-1>", lambda e: open_path_from_tree(tree_done))

tk.Label(root, text=f"📍 目前資料庫路徑: {DB_PATH}", fg="gray", font=("微軟正黑體", 10)).pack(side="bottom", anchor="w", padx=20)

refresh_all_lists()
root.mainloop()
