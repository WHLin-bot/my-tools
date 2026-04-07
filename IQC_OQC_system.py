import os
import sys
import shutil
import sqlite3
import csv
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# ==========================================
# 0. 版本號與自動更新邏輯 (維持不動)
# ==========================================
CURRENT_VERSION = "4.4.3" 
UPDATE_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

def check_for_updates():
    if not getattr(sys, 'frozen', False): return
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
                    messagebox.showinfo("下載完成", "更新檔已存至桌面，請解壓縮覆蓋後重新開啟。")
                    os._exit(0)
    except: pass

# ==========================================
# 1. 路徑與資料庫設定 (維持不動)
# ==========================================
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = UPDATE_DIR
SQLITE_DB = os.path.join(HIST_DIR, "ioqc_management.db")

os.makedirs(HIST_DIR, exist_ok=True)

def init_db():
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS active_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, status TEXT, iqc_done INTEGER DEFAULT 0, oqc_done INTEGER DEFAULT 0, path TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS done_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, ship_date TEXT, path TEXT)''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. 核心功能邏輯 (維持不動)
# ==========================================

def get_oqc_ship_date(sn_path):
    try:
        if not sn_path or not os.path.exists(sn_path): return "N/A"
        for folder in os.listdir(sn_path):
            if "OQC" in folder:
                full_path = os.path.join(sn_path, folder)
                files = [f for f in os.listdir(full_path) if os.path.isfile(os.path.join(full_path, f))]
                if files:
                    mtime = os.path.getmtime(os.path.join(full_path, files[0]))
                    return datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')
    except: pass
    return "N/A"

def get_folder_status(sn_path, date_keyword):
    if not sn_path or not os.path.exists(sn_path): return "路徑遺失", "red"
    has_iqc, has_oqc = False, False
    try:
        for folder in os.listdir(sn_path):
            if date_keyword in folder and os.path.isdir(os.path.join(sn_path, folder)):
                files = os.listdir(os.path.join(sn_path, folder))
                has_files = any(os.path.isfile(os.path.join(sn_path, folder, f)) for f in files)
                if "IQC" in folder and has_files: has_iqc = True
                if "OQC" in folder and has_files: has_oqc = True
    except: pass
    if has_iqc and has_oqc: return "✅ 已完成", "green"
    if has_iqc or has_oqc: return "⚠️ 部分上傳", "orange"
    return "❌ 尚未放照片", "red"

def create_folders():
    customer = entry_customer.get().strip()
    model = entry_model.get().strip()
    sn = entry_sn.get().strip()
    cust_id = entry_cust_id.get().strip()
    staff = entry_staff.get().strip()
    if not all([customer, model, sn, cust_id, staff]):
        messagebox.showwarning("提示", "請填寫所有欄位！")
        return
    sn_path = os.path.join(DB_PATH, customer, model, sn)
    today = datetime.now().strftime('%Y%m%d')
    iqc_path = os.path.join(sn_path, f"{today}_IQC")
    oqc_path = os.path.join(sn_path, f"{today}_OQC")
    try:
        os.makedirs(iqc_path, exist_ok=True)
        os.makedirs(oqc_path, exist_ok=True)
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        cursor.execute("INSERT INTO active_records (time, customer, cust_id, model, sn, staff, status, path) VALUES (?,?,?,?,?,?,?,?)",
            (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        conn.commit(); conn.close()
        messagebox.showinfo("成功", f"資料夾已建立！\n位置：{model}/{sn}")
        os.startfile(iqc_path)
        refresh_search()
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

def archive_done_records():
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records")
    to_archive = []
    for r in cursor.fetchall():
        date_k = r[1][:10].replace("-", "").replace("/", "")
        status, _ = get_folder_status(r[10], date_k)
        if status == "✅ 已完成": to_archive.append(r)
    
    if not to_archive:
        messagebox.showinfo("提示", "無已完成項目。"); conn.close(); return

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
        conn.commit(); messagebox.showinfo("成功", "歸檔完成！")
    conn.close(); refresh_search(); refresh_done_tab()

def open_selected(event):
    tv = event.widget
    item = tv.selection()
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
# 3. UI 介面佈局 (修正點：字體、按鈕功能、名稱)
# ==========================================
root = tk.Tk()
root.title(f"IQC/OQC 管理系統 v{CURRENT_VERSION}")
root.geometry("1400x850")

# 樣式設定
FONT_MAIN, FONT_BOLD = ("微軟正黑體", 12), ("微軟正黑體", 12, "bold")
style = ttk.Style()
style.theme_use('clam')
style.configure("Treeview.Heading", font=FONT_BOLD)
style.configure("Treeview", font=FONT_MAIN, rowheight=35)

# 輸入區
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_BOLD, padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

# (修正 2) 確保 Label 字體維持 FONT_MAIN
labels_text = ["客戶名稱:", "產品型號:", "SN 編號:", "客戶編號:", "作業人員:"]
entries = []
for i, text in enumerate(labels_text):
    tk.Label(frame_input, text=text, font=FONT_MAIN).grid(row=i//3, column=(i%3)*2, sticky="w", padx=5)
    e = tk.Entry(frame_input, width=18, font=FONT_MAIN)
    e.grid(row=i//3, column=(i%3)*2+1, padx=5, pady=5)
    entries.append(e)

entry_customer, entry_model, entry_sn, entry_cust_id, entry_staff = entries
tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BOLD).grid(row=0, column=6, rowspan=2, padx=20, sticky="nswe")

# 分頁區
notebook = ttk.Notebook(root); notebook.pack(fill="both", expand=True, padx=20, pady=10)
tab1 = tk.Frame(notebook); tab2 = tk.Frame(notebook)
notebook.add(tab1, text=" 進行中資料 "); notebook.add(tab2, text=" 已歸檔歷史 ")

# --- 進行中 Tab ---
f_search = tk.Frame(tab1); f_search.pack(fill="x", pady=5)
tk.Label(f_search, text="🔍 搜尋:", font=FONT_MAIN).pack(side="left", padx=5)
entry_search = tk.Entry(f_search, font=FONT_MAIN); entry_search.pack(side="left", fill="x", expand=True, padx=5)
entry_search.bind("<KeyRelease>", refresh_search)

# (修正 1) 補回刷新按鈕
tk.Button(f_search, text="🔄 刷新狀態", command=refresh_search, bg="#E67E22", fg="white", font=FONT_BOLD).pack(side="left", padx=5)

# (修正 3) 修改按鈕名稱
tk.Button(f_search, text="📦 產生紙本報表前CSV檔案", command=archive_done_records, bg="#27AE60", fg="white", font=FONT_BOLD).pack(side="right", padx=10)

cols1 = ("時間", "客戶", "客戶編號", "型號", "SN", "作業員", "狀態", "IQC", "OQC", "Path", "ID")
tree = ttk.Treeview(tab1, columns=cols1, show="headings")
for c in cols1: tree.heading(c, text=c); tree.column(c, width=100, anchor="center")
tree.column("Path", width=0, stretch=False); tree.column("ID", width=0, stretch=False)
tree.pack(fill="both", expand=True)

# --- 已歸檔 Tab ---
f_done = tk.Frame(tab2); f_done.pack(fill="x", pady=5)
tk.Label(f_done, text="🔍 搜尋歷史:", font=FONT_MAIN).pack(side="left", padx=5)
entry_done_search = tk.Entry(f_done, font=FONT_MAIN); entry_done_search.pack(fill="x", padx=10)
entry_done_search.bind("<KeyRelease>", refresh_done_tab)

cols2 = ("時間", "客戶", "客戶編號", "型號", "SN", "作業員", "出貨日", "狀態", "IQC", "OQC", "Path")
tree_done = ttk.Treeview(tab2, columns=cols2, show="headings")
for c in cols2: tree_done.heading(c, text=c); tree_done.column(c, width=100, anchor="center")
tree_done.column("Path", width=0, stretch=False)
tree_done.pack(fill="both", expand=True)

for t in [tree, tree_done]:
    t.tag_configure("green", background="#DFF2BF"); t.tag_configure("orange", background="#FEEFB3")
    t.tag_configure("red", background="#FFBABA"); t.tag_configure("gray", background="#F2F2F2")
    t.bind("<ButtonRelease-1>", open_selected)

root.after(1000, check_for_updates)
refresh_search(); refresh_done_tab()
root.mainloop()
