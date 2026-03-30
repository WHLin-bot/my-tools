import os
import json
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

# ==========================================
# 1. 路徑鎖定
# ==========================================
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_FILE = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history\IOQC_folder_history.txt"

# 確保紀錄檔的目錄存在
os.makedirs(os.path.dirname(HIST_FILE), exist_ok=True)

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

def save_record(customer, cust_id, model, sn, staff, path):
    """儲存資料到 .txt (JSON格式)"""
    records = []
    if os.path.exists(HIST_FILE):
        try:
            with open(HIST_FILE, "r", encoding="utf-8") as f: records = json.load(f)
        except: records = []
    records.append({
        "time": datetime.now().strftime('%Y-%m-%d %H:%M'),
        "customer": customer, "cust_id": cust_id, "model": model, "sn": sn, "staff": staff, "path": path
    })
    with open(HIST_FILE, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=4)

def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l): tv.move(k, '', index)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def create_folders():
    customer = entry_customer.get().strip()
    model = entry_model.get().strip()
    sn = entry_sn.get().strip()
    cust_id = entry_cust_id.get().strip()
    staff = entry_staff.get().strip()
    
    if not (customer and cust_id and model and sn and staff):
        messagebox.showwarning("提示", "請填寫所有欄位！")
        return
        
    if not os.path.exists(DB_PATH):
        messagebox.showerror("錯誤", "無法存取網路路徑，請檢查 \\fs2 連線！")
        return
        
    today = datetime.now().strftime('%Y%m%d')
    # 【修正】資料夾結構不包含客戶編號：DB_PATH / 客戶名稱 / 型號 / SN
    sn_path = os.path.join(DB_PATH, customer, model, sn)
    
    try:
        os.makedirs(os.path.join(sn_path, f"{today}_IQC"), exist_ok=True)
        os.makedirs(os.path.join(sn_path, f"{today}_OQC"), exist_ok=True)
        # 紀錄中依然保留 cust_id 供搜尋與顯示
        save_record(customer, cust_id, model, sn, staff, sn_path)
        messagebox.showinfo("成功", f"資料夾已建立！\n人員: {staff}")
        os.startfile(sn_path)
        refresh_search()
    except Exception as e: messagebox.showerror("錯誤", f"建立失敗：{e}")

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    if os.path.exists(HIST_FILE):
        try:
            with open(HIST_FILE, "r", encoding="utf-8") as f:
                records = json.load(f)
                for r in reversed(records):
                    search_targets = [r.get('sn', ''), r.get('customer', ''), r.get('cust_id', ''), r.get('model', ''), r.get('staff', '')]
                    if any(query in str(t).lower() for t in search_targets):
                        status_text, color_tag = get_folder_status(r['path'])
                        tree.insert("", "end", values=(r['time'], r['customer'], r.get('cust_id', 'N/A'), r['model'], r['sn'], r.get('staff', 'N/A'), status_text, r['path']), tags=(color_tag,))
        except: pass

def open_selected(event):
    selected = tree.selection()
    if not selected: return
    path = tree.item(selected, "values")[-1]
    if os.path.exists(path): os.startfile(path)
    else: messagebox.showerror("錯誤", "找不到資料夾路徑！")

# --- UI 介面 ---
root = tk.Tk()
root.title("IQC/OQC 管理系統 v3.1")
root.geometry("1150x750")

FONT_LABEL = ("微軟正黑體", 12)
FONT_ENTRY = ("微軟正黑體", 12)
FONT_BTN = ("微軟正黑體", 12, "bold")

style = ttk.Style()
style.configure("Treeview.Heading", font=("微軟正黑體", 12, "bold"))
style.configure("Treeview", font=("微軟正黑體", 11), rowheight=30)

# 1. 輸入區
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_LABEL, padx=10, pady=10)
frame_input.pack(fill="x", padx=20, pady=10)

# 第一排：客戶名稱、產品型號
tk.Label(frame_input, text="客戶名稱:", font=FONT_LABEL).grid(row=0, column=0, sticky="w", pady=5)
entry_customer = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_customer.grid(row=0, column=1, padx=5)

tk.Label(frame_input, text="產品型號:", font=FONT_LABEL).grid(row=0, column=2, sticky="w", pady=5)
entry_model = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_model.grid(row=0, column=3, padx=5)

# 第二排：SN 編號、客戶編號、作業人員
tk.Label(frame_input, text="SN 編號:", font=FONT_LABEL).grid(row=1, column=0, sticky="w", pady=5)
entry_sn = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_sn.grid(row=1, column=1, padx=5)

tk.Label(frame_input, text="客戶編號:", font=FONT_LABEL).grid(row=1, column=2, sticky="w", pady=5)
entry_cust_id = tk.Entry(frame_input, width=15, font=FONT_ENTRY); entry_cust_id.grid(row=1, column=3, padx=5)

tk.Label(frame_input, text="作業人員:", font=FONT_LABEL).grid(row=1, column=4, sticky="w", pady=5)
entry_staff = tk.Entry(frame_input, width=10, font=FONT_ENTRY); entry_staff.grid(row=1, column=5, padx=5)

# 建立按鈕
btn_create = tk.Button(frame_input, text="➕ 建立資料夾並開啟", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BTN)
btn_create.grid(row=0, column=6, rowspan=2, sticky="nswe", padx=15)

# 2. 搜尋列表區
frame_list = tk.LabelFrame(root, text=" 歷史紀錄 (搜尋支援: 客戶、編號、SN) ", font=FONT_LABEL, padx=10, pady=10)
frame_list.pack(fill="both", expand=True, padx=20, pady=10)

frame_search_tool = tk.Frame(frame_list)
frame_search_tool.pack(fill="x", pady=5)

entry_search = tk.Entry(frame_search_tool, font=FONT_ENTRY)
entry_search.pack(side="left", fill="x", expand=True, padx=(0, 5))
entry_search.bind("<KeyRelease>", refresh_search)

btn_refresh = tk.Button(frame_search_tool, text="🔄 更新狀態", command=refresh_search, bg="#F0F3F4", font=("微軟正黑體", 11))
btn_refresh.pack(side="right")

columns = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "目前狀態", "路徑")
tree = ttk.Treeview(frame_list, columns=columns, show="headings")

for col in columns:
    tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))

tree.column("建立時間", width=140); tree.column("客戶", width=90); tree.column("客戶編號", width=90)
tree.column("型號", width=100); tree.column("SN", width=140); tree.column("作業人員", width=90)
tree.column("目前狀態", width=120); tree.column("路徑", width=0, stretch=False)

tree.tag_configure("green", background="#DFF2BF", foreground="#270")
tree.tag_configure("orange", background="#FEEFB3", foreground="#9F6000")
tree.tag_configure("red", background="#FFBABA", foreground="#D8000C")

tree.pack(fill="both", expand=True)
tree.bind("<Double-1>", open_selected)

path_info = tk.Label(root, text=f"📍 目前資料庫路徑: {DB_PATH}", fg="gray", font=("微軟正黑體", 10))
path_info.pack(side="bottom", anchor="w", padx=20)

refresh_search()
root.mainloop()
