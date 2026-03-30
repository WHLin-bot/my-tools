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
HIST_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
HIST_FILE = os.path.join(HIST_DIR, "IOQC_folder_history.txt")
DONE_FILE = os.path.join(HIST_DIR, "IOQC_folder_history_done.txt")

# 確保紀錄檔的目錄存在
os.makedirs(HIST_DIR, exist_ok=True)

# ----------------- 功能邏輯 -----------------

def get_folder_status(sn_path):
    """檢查路徑下是否有照片"""
    if not os.path.exists(sn_path): return "路徑遺失", "red"
    has_iqc, has_oqc = False, False
    try:
        for folder in os.listdir(sn_path):
            full_path = os.path.join(sn_path, folder)
            if os.path.isdir(full_path):
                has_files = any(os.path.isfile(os.path.join(full_path, f)) for f in os.listdir(full_path))
                if "IQC" in folder and has_files: has_iqc = True
                if "OQC" in folder and has_files: has_oqc = True
    except: pass
    if has_iqc and has_oqc: return "✅ 已完成", "green"
    if has_iqc or has_oqc: return "⚠️ 部分上傳", "orange"
    return "❌ 尚未放照片", "red"

def load_records(file_path):
    if not os.path.exists(file_path): return []
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except: return []

def save_records(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def archive_done_records():
    """轉移已完成資料並產出 CSV (自定義欄位與日期格式)"""
    records = load_records(HIST_FILE)
    done_records = load_records(DONE_FILE)
    new_active, to_archive = [], []

    for r in records:
        status, _ = get_folder_status(r['path'])
        if status == "✅ 已完成": to_archive.append(r)
        else: new_active.append(r)

    if not to_archive:
        messagebox.showinfo("提示", "目前沒有符合「已完成」狀態的資料可歸檔。")
        return

    csv_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv")],
        title="選擇 CSV 儲存位置",
        initialfile=f"IOQC_Archive_{datetime.now().strftime('%Y%m%d')}.csv"
    )
    
    if csv_path:
        try:
            export_data = []
            for item in to_archive:
                clean_date = item.get('time', '').split(' ')[0]
                export_data.append({
                    "客戶名稱": item.get('customer', ''),
                    "產品型號": item.get('model', ''),
                    "SN編號": item.get('sn', ''),
                    "客戶編號": item.get('cust_id', ''),
                    "建立日期": clean_date,
                    "作業人員": item.get('staff', '')
                })

            done_records.extend(to_archive)
            save_records(DONE_FILE, done_records)
            save_records(HIST_FILE, new_active)

            fieldnames = ["客戶名稱", "產品型號", "SN編號", "客戶編號", "建立日期", "作業人員"]
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(export_data)
                
            messagebox.showinfo("成功", f"已成功歸檔 {len(to_archive)} 筆資料並產出 CSV！")
        except Exception as e:
            messagebox.showerror("錯誤", f"歸檔失敗：{e}")

    refresh_search()
    refresh_done_tab()

def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l): tv.move(k, '', index)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def create_folders():
    customer, model, sn = entry_customer.get().strip(), entry_model.get().strip(), entry_sn.get().strip()
    cust_id, staff = entry_cust_id.get().strip(), entry_staff.get().strip()
    
    if not (customer and cust_id and model and sn and staff):
        messagebox.showwarning("提示", "請填寫所有欄位！")
        return
        
    sn_path = os.path.join(DB_PATH, customer, model, sn)
    today = datetime.now().strftime('%Y%m%d')
    
    try:
        os.makedirs(os.path.join(sn_path, f"{today}_IQC"), exist_ok=True)
        os.makedirs(os.path.join(sn_path, f"{today}_OQC"), exist_ok=True)
        records = load_records(HIST_FILE)
        records.append({
            "time": datetime.now().strftime('%Y-%m-%d %H:%M'),
            "customer": customer, "cust_id": cust_id, "model": model, "sn": sn, "staff": staff, "path": sn_path
        })
        save_records(HIST_FILE, records)
        messagebox.showinfo("成功", "資料夾已建立！")
        os.startfile(sn_path)
        refresh_search()
    except Exception as e: messagebox.showerror("錯誤", f"建立失敗：{e}")

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    records = load_records(HIST_FILE)
    for r in reversed(records):
        targets = [r.get('sn',''), r.get('customer',''), r.get('cust_id',''), r.get('model',''), r.get('staff','')]
        if any(query in str(t).lower() for t in targets):
            status_text, color_tag = get_folder_status(r['path'])
            tree.insert("", "end", values=(r['time'], r['customer'], r.get('cust_id', 'N/A'), r['model'], r['sn'], r.get('staff', 'N/A'), status_text, r['path']), tags=(color_tag,))

def refresh_done_tab():
    for item in tree_done.get_children(): tree_done.delete(item)
    records = load_records(DONE_FILE)
    for r in reversed(records):
        tree_done.insert("", "end", values=(r['time'], r['customer'], r.get('cust_id', 'N/A'), r['model'], r['sn'], r.get('staff', 'N/A'), "✅ 已歸檔", r['path']), tags=("gray",))

def open_selected(event):
    tv = event.widget
    selected = tv.selection()
    if not selected: return
    path = tv.item(selected, "values")[-1]
    if os.path.exists(path): os.startfile(path)
    else: messagebox.showerror("錯誤", "找不到路徑！")

# --- UI 介面 ---
root = tk.Tk()
root.title("IQC/OQC 管理系統 v4.3")
root.geometry("1250x850")

FONT_MAIN = ("微軟正黑體", 12)
FONT_BOLD = ("微軟正黑體", 12, "bold")

style = ttk.Style()
style.theme_use('clam')
style.configure("Treeview.Heading", font=FONT_BOLD)
style.configure("Treeview", font=FONT_MAIN, rowheight=35)
style.map("Treeview", background=[('selected', '#347083')])

# 1. 輸入區
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_BOLD, padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

labels_text = ["客戶名稱:", "產品型號:", "SN 編號:", "客戶編號:", "作業人員:"]
entries = []
for i, txt in enumerate(labels_text):
    tk.Label(frame_input, text=txt, font=FONT_MAIN).grid(row=i//3, column=(i%3)*2, sticky="w", padx=5, pady=8)
    e = tk.Entry(frame_input, width=18, font=FONT_MAIN)
    e.grid(row=i//3, column=(i%3)*2+1, padx=5)
    entries.append(e)

entry_customer, entry_model, entry_sn, entry_cust_id, entry_staff = entries
btn_main = tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BOLD, width=15)
btn_main.grid(row=0, column=6, rowspan=2, padx=20, sticky="nswe")

# 2. 分頁區
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=20, pady=10)

# --- 分頁 1: 進行中 ---
tab_active = tk.Frame(notebook)
notebook.add(tab_active, text=" 進行中資料 ")

frame_search = tk.Frame(tab_active)
frame_search.pack(fill="x", pady=10, padx=10)

tk.Label(frame_search, text="搜尋:", font=FONT_MAIN).pack(side="left")
entry_search = tk.Entry(frame_search, font=FONT_MAIN)
entry_search.pack(side="left", fill="x", expand=True, padx=10)
entry_search.bind("<KeyRelease>", refresh_search)

tk.Button(frame_search, text="🔄 更新狀態", command=refresh_search, font=FONT_MAIN).pack(side="left", padx=5)
tk.Button(frame_search, text="📦 產出CSV '稽核表單專用'", command=archive_done_records, bg="#27AE60", fg="white", font=FONT_BOLD).pack(side="right", padx=5)

columns = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "目前狀態", "路徑")
tree = ttk.Treeview(tab_active, columns=columns, show="headings")
for col in columns: 
    tree.heading(col, text=col, command=lambda _c=col: treeview_sort_column(tree, _c, False))
    tree.column(col, width=120, anchor="center")
tree.column("路徑", width=0, stretch=False)
tree.pack(fill="both", expand=True, padx=10, pady=5)

# --- 分頁 2: 已歸檔 ---
tab_done = tk.Frame(notebook)
notebook.add(tab_done, text=" 已歸檔歷史 (Done) ")

tree_done = ttk.Treeview(tab_done, columns=columns, show="headings")
for col in columns: 
    tree_done.heading(col, text=col, command=lambda _c=col: treeview_sort_column(tree_done, _c, False))
    tree_done.column(col, width=120, anchor="center")
tree_done.column("路徑", width=0, stretch=False)
tree_done.pack(fill="both", expand=True, padx=10, pady=10)

for t in [tree, tree_done]:
    t.tag_configure("green", background="#DFF2BF", foreground="#270")
    t.tag_configure("orange", background="#FEEFB3", foreground="#9F6000")
    t.tag_configure("red", background="#FFBABA", foreground="#D8000C")
    t.tag_configure("gray", background="#F2F2F2", foreground="#666")
    t.bind("<Double-1>", open_selected)

path_info = tk.Label(root, text=f"📍 資料夾根路徑: {DB_PATH}", fg="gray", font=("微軟正黑體", 10))
path_info.pack(side="bottom", anchor="w", padx=20, pady=5)

refresh_search()
refresh_done_tab()
root.mainloop()
