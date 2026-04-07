import os
import sys
import shutil
import sqlite3
import csv
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# ==========================================
# 0. 版本號與手動更新
# ==========================================
CURRENT_VERSION = "4.5.1" 
UPDATE_DIR = r"\\fs2\Dept(Q)\08_p03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

def check_for_updates():
    if not getattr(sys, 'frozen', False): return
    try:
        if not os.path.exists(VERSION_FILE_PATH): return
        with open(VERSION_FILE_PATH, "r", encoding="utf-8") as f:
            server_version = f.read().strip()
        if server_version != CURRENT_VERSION:
            msg = f"偵測到新版本 v{server_version}！\n是否現在下載新版本並關閉程式？"
            if messagebox.askyesno("系統更新提示", msg):
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                target_zip_path = os.path.join(desktop_path, f"IQC_OQC_system_v{server_version}.zip")
                if os.path.exists(REMOTE_ZIP_PATH):
                    shutil.copy2(REMOTE_ZIP_PATH, target_zip_path)
                    messagebox.showinfo("下載完成", "已下載至桌面，程式即將關閉。")
                    os._exit(0)
    except: pass

# ==========================================
# 1. 路徑與資料庫設定
# ==========================================
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
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
# 2. 核心功能邏輯 (已修正路徑層級)
# ==========================================

def get_oqc_ship_date(sn_path):
    try:
        if not sn_path or not os.path.exists(sn_path): return "N/A"
        for folder in os.listdir(sn_path):
            if "OQC" in folder:
                full_path = os.path.join(sn_path, folder)
                files = [os.path.join(full_path, f) for f in os.listdir(full_path) if os.path.isfile(os.path.join(full_path, f))]
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
            # 判斷邏輯：資料夾名稱包含日期(YYYYMMDD)且包含IQC/OQC關鍵字
            if date_keyword in folder and os.path.isdir(os.path.join(sn_path, folder)):
                # 檢查資料夾內是否有任何檔案
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
    
    # 【最新修正路徑】層級：根目錄 / 客戶名稱 / 產品型號 / SN 編號
    sn_path = os.path.join(DB_PATH, customer, model, sn)
    today = datetime.now().strftime('%Y%m%d')
    iqc_path = os.path.join(sn_path, f"{today}_IQC")
    oqc_path = os.path.join(sn_path, f"{today}_OQC")
    
    try:
        os.makedirs(iqc_path, exist_ok=True)
        os.makedirs(oqc_path, exist_ok=True)
        
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        cursor.execute("""INSERT INTO active_records 
            (time, customer, cust_id, model, sn, staff, status, iqc_done, oqc_done, path) 
            VALUES (?,?,?,?,?,?,?,0,0,?)""",
            (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("成功", f"資料夾已建立！\n位置：{model} / {sn}")
        if os.path.exists(iqc_path): os.startfile(iqc_path)
        refresh_search()
    except Exception as e: 
        messagebox.showerror("錯誤", f"建立失敗：{e}")

# ==========================================
# 3. 介面重新整理與歸檔
# ==========================================

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r[2:7]):
            date_keyword = r[1][:10].replace("-", "").replace("/", "")
            status_text, color_tag = get_folder_status(r[10], date_keyword)
            tree.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], status_text, "📂 開啟 IQC", "📂 開啟 OQC", r[10], r[0]), tags=(color_tag,))

def refresh_done_tab(event=None):
    query = entry_done_search.get().strip().lower()
    for item in tree_done.get_children(): tree_done.delete(item)
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM done_records ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r[2:8]):
            tree_done.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], r[7], "✅ 已歸檔", "📂 開啟 IQC", "📂 開啟 OQC", r[8]), tags=("gray",))

def archive_done_records():
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records")
    active_rows = cursor.fetchall()
    to_archive = []
    for r in active_rows:
        date_keyword = r[1][:10].replace("-", "").replace("/", "")
        status, _ = get_folder_status(r[10], date_keyword)
        if status == "✅ 已完成": to_archive.append(r)
    
    if not to_archive:
        messagebox.showinfo("提示", "沒有符合「已完成」狀態的資料可供歸檔。")
        conn.close(); return

    csv_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=f"IOQC_Archive_{datetime.now().strftime('%Y%m%d')}.csv")
    if csv_path:
        try:
            export_data = []
            for r in to_archive:
                ship_date = get_oqc_ship_date(r[10])
                cursor.execute("INSERT INTO done_records (time, customer, cust_id, model, sn, staff, ship_date, path) VALUES (?,?,?,?,?,?,?,?)",
                               (r[1], r[2], r[3], r[4], r[5], r[6], ship_date, r[10]))
                cursor.execute("DELETE FROM active_records WHERE id = ?", (r[0],))
                export_data.append({"客戶名稱": r[2], "產品型號": r[4], "SN編號": r[5], "客戶編號": r[3], "建立日期": r[1][:10], "作業人員": r[6], "出貨日期": ship_date})
            conn.commit()
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=["客戶名稱", "產品型號", "SN編號", "客戶編號", "建立日期", "作業人員", "出貨日期"])
                writer.writeheader(); writer.writerows(export_data)
            messagebox.showinfo("成功", f"歸檔完成，共 {len(to_archive)} 筆。")
        except Exception as e: messagebox.showerror("錯誤", f"歸檔失敗：{e}")
    conn.close(); refresh_search(); refresh_done_tab()

def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l): tv.move(k, '', index)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def open_selected(event):
    tv = event.widget
    selected = tv.selection()
    if not selected: return
    col_id = tv.identify_column(event.x)
    vals = tv.item(selected, "values")
    date_keyword = vals[0][:10].replace("-", "").replace("/", "")
    path = vals[9] if tv == tree else vals[10]
    iqc_btn_id, oqc_btn_id = ("#8", "#9") if tv == tree else ("#9", "#10")

    if not path or path == "None" or not os.path.exists(path):
        messagebox.showerror("錯誤", "找不到實體資料夾路徑。")
        return

    if col_id == iqc_btn_id:
        target = os.path.join(path, f"{date_keyword}_IQC")
        if os.path.exists(target): os.startfile(target)
        else: messagebox.showwarning("提示", "找不到對應日期的 IQC 資料夾。")
    elif col_id == oqc_btn_id:
        target = os.path.join(path, f"{date_keyword}_OQC")
        if os.path.exists(target): os.startfile(target)
        else: messagebox.showwarning("提示", "找不到對應日期的 OQC 資料夾。")

# ==========================================
# 4. UI 介面佈局
# ==========================================
root = tk.Tk()
root.title(f"IQC/OQC 管理系統 v{CURRENT_VERSION}")
root.geometry("1400x850")

FONT_MAIN, FONT_BOLD = ("微軟正黑體", 12), ("微軟正黑體", 12, "bold")
style = ttk.Style()
style.theme_use('clam')
style.configure("Treeview.Heading", font=FONT_BOLD)
style.configure("Treeview", font=FONT_MAIN, rowheight=35)

frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_BOLD, padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

tk.Label(frame_input, text="客戶名稱:").grid(row=0, column=0, sticky="w", padx=5)
entry_customer = tk.Entry(frame_input, width=18, font=FONT_MAIN); entry_customer.grid(row=0, column=1, padx=5)
tk.Label(frame_input, text="產品型號:").grid(row=0, column=2, sticky="w", padx=5)
entry_model = tk.Entry(frame_input, width=18, font=FONT_MAIN); entry_model.grid(row=0, column=3, padx=5)
tk.Label(frame_input, text="SN 編號:").grid(row=0, column=4, sticky="w", padx=5)
entry_sn = tk.Entry(frame_input, width=18, font=FONT_MAIN); entry_sn.grid(row=0, column=5, padx=5)
tk.Label(frame_input, text="客戶編號:").grid(row=1, column=0, sticky="w", padx=5)
entry_cust_id = tk.Entry(frame_input, width=18, font=FONT_MAIN); entry_cust_id.grid(row=1, column=1, padx=5)
tk.Label(frame_input, text="作業人員:").grid(row=1, column=2, sticky="w", padx=5)
entry_staff = tk.Entry(frame_input, width=18, font=FONT_MAIN); entry_staff.grid(row=1, column=3, padx=5)

tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BOLD).grid(row=0, column=6, rowspan=2, padx=20, sticky="nswe")

notebook = ttk.Notebook(root); notebook.pack(fill="both", expand=True, padx=20, pady=10)

# --- 進行中分頁 ---
tab_active = tk.Frame(notebook); notebook.add(tab_active, text=" 進行中資料 ")
frame_search = tk.Frame(tab_active); frame_search.pack(fill="x", pady=10, padx=10)
entry_search = tk.Entry(frame_search, font=FONT_MAIN); entry_search.pack(side="left", fill="x", expand=True, padx=10)
entry_search.bind("<KeyRelease>", refresh_search)
tk.Button(frame_search, text="📦 產出CSV歸檔", command=archive_done_records, bg="#27AE60", fg="white", font=FONT_BOLD).pack(side="right", padx=5)

cols_active = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "目前狀態", "IQC 資料夾", "OQC 資料夾", "路徑", "ID")
tree = ttk.Treeview(tab_active, columns=cols_active, show="headings")
tree.pack(fill="both", expand=True, padx=10)
for c in cols_active: 
    tree.heading(c, text=c, command=lambda _c=c: treeview_sort_column(tree, _c, False))
    tree.column(c, width=110, anchor="center")
tree.column("路徑", width=0, stretch=False); tree.column("ID", width=0, stretch=False)

# --- 已歸檔分頁 ---
tab_done = tk.Frame(notebook); notebook.add(tab_done, text=" 已歸檔歷史 ")
frame_done_search = tk.Frame(tab_done); frame_done_search.pack(fill="x", pady=10, padx=10)
entry_done_search = tk.Entry(frame_done_search, font=FONT_MAIN); entry_done_search.pack(side="left", fill="x", expand=True, padx=10)
entry_done_search.bind("<KeyRelease>", refresh_done_tab)

cols_done = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "出貨日期", "狀態", "IQC 資料夾", "OQC 資料夾", "路徑")
tree_done = ttk.Treeview(tab_done, columns=cols_done, show="headings")
tree_done.pack(fill="both", expand=True, padx=10)
for c in cols_done: 
    tree_done.heading(c, text=c, command=lambda _c=c: treeview_sort_column(tree_done, _c, False))
    tree_done.column(c, width=110, anchor="center")
tree_done.column("路徑", width=0, stretch=False)

# 樣式配置與事件綁定
for t in [tree, tree_done]:
    t.tag_configure("green", background="#DFF2BF", foreground="#270")
    t.tag_configure("orange", background="#FEEFB3", foreground="#9F6000")
    t.tag_configure("red", background="#FFBABA", foreground="#D8000C")
    t.tag_configure("gray", background="#F2F2F2", foreground="#666")
    t.bind("<ButtonRelease-1>", open_selected)

root.after(100, check_for_updates)
refresh_search(); refresh_done_tab()
root.mainloop()
