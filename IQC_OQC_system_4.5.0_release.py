import os
import sys
import shutil
import sqlite3
import csv
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# ==========================================
# 0. 路徑與資料庫設定
# ==========================================

CURRENT_VERSION = "4.5.0" 
UPDATE_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
SQLITE_DB = os.path.join(UPDATE_DIR, "ioqc_management.db")
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

def check_for_updates():
#    if not getattr(sys, 'frozen', False): return
    try:
        if not os.path.exists(VERSION_FILE_PATH): return
        with open(VERSION_FILE_PATH, "r", encoding="utf-8") as f:
            server_version = f.read().strip()
        if server_version != CURRENT_VERSION:
            msg = f"偵測到新版本 v{server_version}！\n是否現在下載更新檔至桌面並關閉程式？"
            if messagebox.askyesno("系統更新提示", msg):
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                target_zip_path = os.path.join(desktop_path, f"IQC_OQC_system_v{server_version}.zip")
                if os.path.exists(REMOTE_ZIP_PATH):
                    shutil.copy2(REMOTE_ZIP_PATH, target_zip_path)
                    messagebox.showinfo("下載完成", "更新檔zip已存至桌面，請先刪除舊檔案再解壓縮重新開啟。")
                    os._exit(0)
    except: pass


# ==========================================
# 1. 核心功能邏輯
# ==========================================

def treeview_sort_column(tv, col, reverse):
    """表頭排序功能"""
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l): tv.move(k, '', index)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def get_staff_list():
    try:
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT staff FROM active_records UNION SELECT DISTINCT staff FROM done_records")
        staffs = [row[0] for row in cursor.fetchall() if row[0]]
        conn.close(); return staffs
    except: return []

def count_sn_occurrence(sn):
    """計算 SN 在資料庫中出現的總次數"""
    if not sn: return 0
    try:
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM active_records WHERE sn = ? COLLATE NOCASE", (sn,))
        c1 = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM done_records WHERE sn = ? COLLATE NOCASE", (sn,))
        c2 = cursor.fetchone()[0]
        conn.close()
        return c1 + c2
    except: return 0

def refresh_staff_list(event=None):
    entry_staff['values'] = get_staff_list()

def auto_fill_data(event=None):
    sn = entry_sn.get().strip()
    if not sn: return
    # 更新次數顯示 (移除顏色，統一使用灰色)
    count = count_sn_occurrence(sn)
    label_count.config(text=f"此 SN 已出現過: {count} 次")
    
    try:
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        cursor.execute("""
            SELECT customer, model, cust_id FROM (
                SELECT customer, model, cust_id, time FROM active_records WHERE sn = ? COLLATE NOCASE
                UNION ALL
                SELECT customer, model, cust_id, time FROM done_records WHERE sn = ? COLLATE NOCASE
            ) AS all_records
            ORDER BY time DESC
            LIMIT 1
        """, (sn, sn))
        res = cursor.fetchone(); conn.close()
        if res:
            entry_customer.delete(0, tk.END); entry_customer.insert(0, res[0])
            entry_model.delete(0, tk.END); entry_model.insert(0, res[1])
            entry_cust_id.delete(0, tk.END); entry_cust_id.insert(0, res[2])
    except: pass

def get_oqc_ship_date(sn_path):
    try:
        if not sn_path or not os.path.exists(sn_path): return "N/A"
        for folder in os.listdir(sn_path):
            if "OQC" in folder:
                full_p = os.path.join(sn_path, folder)
                files = [os.path.join(full_p, f) for f in os.listdir(full_p) if os.path.isfile(os.path.join(full_p, f))]
                if files:
                    mtime = os.path.getmtime(max(files, key=os.path.getmtime))
                    return datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')
    except: pass
    return "N/A"

def get_folder_status(sn_path, date_keyword):
    if not sn_path or not os.path.exists(sn_path): return "路徑遺失", "red"
    has_iqc, has_oqc = False, False
    try:
        for folder in os.listdir(sn_path):
            if date_keyword in folder:
                f_list = os.listdir(os.path.join(sn_path, folder))
                has_f = any(os.path.isfile(os.path.join(sn_path, folder, f)) for f in f_list)
                if "IQC" in folder and has_f: has_iqc = True
                if "OQC" in folder and has_f: has_oqc = True
    except: pass
    if has_iqc and has_oqc: return "✅ 已完成", "green"
    if has_iqc or has_oqc: return "⚠️ 部分上傳", "orange"
    return "❌ 尚未放照片", "red"

def create_folders():
    customer, model, sn, cust_id, staff = entry_customer.get().strip(), entry_model.get().strip(), entry_sn.get().strip(), entry_cust_id.get().strip(), entry_staff.get()
    if not all([customer, model, sn, cust_id, staff]):
        messagebox.showwarning("提示", "請填寫所有欄位！"); return
    
    sn_path = os.path.join(DB_PATH, customer, model, sn)
    today = datetime.now().strftime('%Y%m%d')
    iqc_p = os.path.join(sn_path, f"{today}_IQC")
    try:
        os.makedirs(iqc_p, exist_ok=True); os.makedirs(os.path.join(sn_path, f"{today}_OQC"), exist_ok=True)
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        cursor.execute("INSERT INTO active_records (time, customer, cust_id, model, sn, staff, status, path) VALUES (?,?,?,?,?,?,?,?)",
            (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        conn.commit(); conn.close()
        messagebox.showinfo("成功", f"資料夾已建立！\n(此 SN 累計 {count_sn_occurrence(sn)} 次紀錄)"); os.startfile(iqc_p); refresh_search()
    except Exception as e: messagebox.showerror("錯誤", str(e))

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records ORDER BY id DESC")
    rows = cursor.fetchall(); conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r[2:7]):
            date_k = r[1][:10].replace("-", "").replace("/", "")
            status_t, color_tag = get_folder_status(r[10], date_k)
            tree.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], status_t, "📂 開啟 IQC", "📂 開啟 OQC", r[10], r[0]), tags=(color_tag,))

def refresh_done_tab(event=None):
    query = entry_done_search.get().strip().lower()
    for item in tree_done.get_children(): tree_done.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM done_records ORDER BY id DESC")
    rows = cursor.fetchall(); conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r[2:8]):
            tree_done.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], r[7], "✅ 已歸檔", "📂 開啟 IQC", "📂 開啟 OQC", r[8]), tags=("gray",))

def archive_to_csv():
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records")
    active_rows = cursor.fetchall()
    to_archive = []
    for r in active_rows:
        date_k = r[1][:10].replace("-", "").replace("/", "")
        status, _ = get_folder_status(r[10], date_k)
        if status == "✅ 已完成": to_archive.append(r)
    
    if not to_archive:
        messagebox.showinfo("提示", "目前沒有已完成的項目可歸檔。"); conn.close(); return

    csv_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=f"Archive_{datetime.now().strftime('%Y%m%d')}.csv")
    if csv_path:
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=["客戶名稱", "產品型號", "SN編號", "客戶編號", "建立日期", "作業人員", "出貨日期"])
            writer.writeheader()
            for r in to_archive:
                ship_d = get_oqc_ship_date(r[10])
                cursor.execute("INSERT INTO done_records (time, customer, cust_id, model, sn, staff, ship_date, path) VALUES (?,?,?,?,?,?,?,?)",
                               (r[1], r[2], r[3], r[4], r[5], r[6], ship_d, r[10]))
                cursor.execute("DELETE FROM active_records WHERE id = ?", (r[0],))
                writer.writerow({"客戶名稱": r[2], "產品型號": r[4], "SN編號": r[5], "客戶編號": r[3], "建立日期": r[1][:10], "作業人員": r[6], "出貨日期": ship_d})
        conn.commit(); messagebox.showinfo("成功", "歸檔與 CSV 產出成功！")
    conn.close(); refresh_search(); refresh_done_tab()

def open_selected(event):
    tv = event.widget; item = tv.selection()
    if not item: return
    col = tv.identify_column(event.x)
    vals = tv.item(item, "values")
    date_k = vals[0][:10].replace("-", "").replace("/", "")
    path = vals[9] if tv == tree else vals[10]
    iqc_btn, oqc_btn = ("#8", "#9") if tv == tree else ("#9", "#10")
    if col == iqc_btn:
        p = os.path.join(path, f"{date_k}_IQC")
        if os.path.exists(p): os.startfile(p)
    elif col == oqc_btn:
        p = os.path.join(path, f"{date_k}_OQC")
        if os.path.exists(p): os.startfile(p)

# ==========================================
# 2. UI 介面佈局
# ==========================================
root = tk.Tk()
root.title(f"IQC/OQC 管理系統 v{CURRENT_VERSION}")
root.geometry("1400x850")

style = ttk.Style(); style.theme_use('clam')
style.configure("Treeview.Heading", font=("微軟正黑體", 12, "bold"))
style.configure("Treeview", font=("微軟正黑體", 12), rowheight=35)

frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=("微軟正黑體", 12, "bold"), padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

tk.Label(frame_input, text="SN 編號:", font=("微軟正黑體", 12)).grid(row=0, column=0, sticky="w", padx=5)
entry_sn = tk.Entry(frame_input, width=20, font=("微軟正黑體", 12))
entry_sn.grid(row=0, column=1, padx=5)
entry_sn.bind("<Tab>", auto_fill_data)

tk.Label(frame_input, text="(輸入 SN編號 後按 Tab 自動帶入資訊)", font=("微軟正黑體", 9), fg="blue").grid(row=1, column=1, sticky="nw", padx=5)
# 修正：SN 次數標籤統一使用灰色
label_count = tk.Label(frame_input, text="此 SN 已出現過: 0 次", font=("微軟正黑體", 10), fg="#555555")
label_count.grid(row=1, column=2, columnspan=2, sticky="nw", padx=5)

tk.Label(frame_input, text="客戶名稱:", font=("微軟正黑體", 12)).grid(row=0, column=2, sticky="w", padx=5)
entry_customer = tk.Entry(frame_input, width=20, font=("微軟正黑體", 12))
entry_customer.grid(row=0, column=3, padx=5)

tk.Label(frame_input, text="產品型號:", font=("微軟正黑體", 12)).grid(row=0, column=4, sticky="w", padx=5)
entry_model = tk.Entry(frame_input, width=20, font=("微軟正黑體", 12))
entry_model.grid(row=0, column=5, padx=5)

tk.Label(frame_input, text="客戶編號:", font=("微軟正黑體", 12)).grid(row=2, column=0, sticky="w", padx=5, pady=(15,0))
entry_cust_id = tk.Entry(frame_input, width=20, font=("微軟正黑體", 12))
entry_cust_id.grid(row=2, column=1, padx=5, pady=(15,0))

tk.Label(frame_input, text="作業人員:", font=("微軟正黑體", 12)).grid(row=2, column=2, sticky="w", padx=5, pady=(15,0))
entry_staff = ttk.Combobox(frame_input, width=18, font=("微軟正黑體", 12))
entry_staff.grid(row=2, column=3, padx=5, pady=(15,0))
entry_staff.bind("<Button-1>", refresh_staff_list)

tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=("微軟正黑體", 12, "bold")).grid(row=0, column=6, rowspan=3, padx=20, sticky="nswe")

# --- 分頁與卷軸 ---
notebook = ttk.Notebook(root); notebook.pack(fill="both", expand=True, padx=20, pady=10)

# Tab 1
tab1 = tk.Frame(notebook); notebook.add(tab1, text=" 進行中資料 ")
f_t1 = tk.Frame(tab1); f_t1.pack(fill="x", pady=5)
tk.Label(f_t1, text="🔍 搜尋:", font=("微軟正黑體", 12)).pack(side="left", padx=5)
entry_search = tk.Entry(f_t1, font=("微軟正黑體", 12)); entry_search.pack(side="left", fill="x", expand=True, padx=5)
entry_search.bind("<KeyRelease>", refresh_search)
tk.Button(f_t1, text="🔄 刷新列表", command=refresh_search, font=("微軟正黑體", 11)).pack(side="left", padx=5)
tk.Button(f_t1, text="📦 產生紙本表單專用CSV", command=archive_to_csv, bg="#27AE60", fg="white", font=("微軟正黑體", 11, "bold")).pack(side="right", padx=10)

tree_frame1 = tk.Frame(tab1); tree_frame1.pack(fill="both", expand=True)
cols1 = ("建立時間", "客戶名稱", "客戶編號", "產品型號", "SN編號", "作業人員", "狀態", "IQC", "OQC", "Path", "ID")
tree = ttk.Treeview(tree_frame1, columns=cols1, show="headings")
vsb1 = ttk.Scrollbar(tree_frame1, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=vsb1.set); vsb1.pack(side="right", fill="y"); tree.pack(side="left", fill="both", expand=True)

for c in cols1: 
    tree.heading(c, text=c, command=lambda _c=c: treeview_sort_column(tree, _c, False))
    tree.column(c, width=110, anchor="center")
tree.column("Path", width=0, stretch=False); tree.column("ID", width=0, stretch=False)

# Tab 2
tab2 = tk.Frame(notebook); notebook.add(tab2, text=" 已歸檔歷史 ")
f_t2 = tk.Frame(tab2); f_t2.pack(fill="x", pady=5)
tk.Label(f_t2, text="🔍 搜尋歷史:", font=("微軟正黑體", 12)).pack(side="left", padx=5)
entry_done_search = tk.Entry(f_t2, font=("微軟正黑體", 12)); entry_done_search.pack(side="left", fill="x", expand=True, padx=5)
entry_done_search.bind("<KeyRelease>", refresh_done_tab)
tk.Button(f_t2, text="🔄 刷新歷史", command=refresh_done_tab, font=("微軟正黑體", 11)).pack(side="left", padx=5)

tree_frame2 = tk.Frame(tab2); tree_frame2.pack(fill="both", expand=True)
cols2 = ("建立時間", "客戶名稱", "客戶編號", "產品型號", "SN編號", "作業人員", "出貨日期", "狀態", "IQC", "OQC", "Path")
tree_done = ttk.Treeview(tree_frame2, columns=cols2, show="headings")
vsb2 = ttk.Scrollbar(tree_frame2, orient="vertical", command=tree_done.yview)
tree_done.configure(yscrollcommand=vsb2.set); vsb2.pack(side="right", fill="y"); tree_done.pack(side="left", fill="both", expand=True)

for c in cols2: 
    tree_done.heading(c, text=c, command=lambda _c=c: treeview_sort_column(tree_done, _c, False))
    tree_done.column(c, width=110, anchor="center")
tree_done.column("Path", width=0, stretch=False)

# 顏色標籤與點擊
for t in [tree, tree_done]:
    t.tag_configure("green", background="#DFF2BF", foreground="#270")
    t.tag_configure("orange", background="#FEEFB3", foreground="#9F6000")
    t.tag_configure("red", background="#FFBABA", foreground="#D8000C")
    t.tag_configure("gray", background="#F2F2F2", foreground="#666")
    t.bind("<ButtonRelease-1>", open_selected)

root.after(1000, check_for_updates)
refresh_search(); refresh_done_tab()
root.mainloop()
