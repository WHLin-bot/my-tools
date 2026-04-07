import os
import sys
import shutil
import sqlite3
import csv
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# ==========================================
# 1. 路徑與資料庫設定
# ==========================================
CURRENT_VERSION = "4.6.2" 
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
SQLITE_DB = os.path.join(HIST_DIR, "ioqc_management.db")

os.makedirs(HIST_DIR, exist_ok=True)

def init_db():
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    # 進行中
    cursor.execute('''CREATE TABLE IF NOT EXISTS active_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, status TEXT, iqc_done INTEGER DEFAULT 0, oqc_done INTEGER DEFAULT 0, path TEXT)''')
    # 已歸檔歷史 (原本的)
    cursor.execute('''CREATE TABLE IF NOT EXISTS done_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, ship_date TEXT, path TEXT)''')
    # Database (SN 對照表)
    cursor.execute('''CREATE TABLE IF NOT EXISTS sn_lookup_db 
        (sn TEXT PRIMARY KEY, customer TEXT, cust_id TEXT, model TEXT)''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. 核心功能邏輯
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
            if date_keyword in folder and os.path.isdir(os.path.join(sn_path, folder)):
                has_files = any(os.path.isfile(os.path.join(sn_path, folder, f)) for f in os.listdir(os.path.join(sn_path, folder)))
                if "IQC" in folder and has_files: has_iqc = True
                if "OQC" in folder and has_files: has_oqc = True
    except: pass
    if has_iqc and has_oqc: return "✅ 已完成", "green"
    if has_iqc or has_oqc: return "⚠️ 部分上傳", "orange"
    return "❌ 尚未放照片", "red"

def auto_fill_by_sn(event=None):
    """從 Database 搜尋 SN 並自動帶入"""
    sn = entry_sn.get().strip()
    if not sn: return
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT customer, cust_id, model FROM sn_lookup_db WHERE sn = ?", (sn,))
    res = cursor.fetchone(); conn.close()
    if res:
        entry_customer.delete(0, tk.END); entry_customer.insert(0, res[0])
        entry_cust_id.delete(0, tk.END); entry_cust_id.insert(0, res[1])
        entry_model.delete(0, tk.END); entry_model.insert(0, res[2])
        lbl_status_msg.config(text=f"✅ 已從 Database 帶入 SN: {sn}", fg="blue")
    else:
        lbl_status_msg.config(text="🆕 這是新 SN，請輸入資訊", fg="orange")

def create_folders():
    customer, model, sn, cust_id, staff = entry_customer.get().strip(), entry_model.get().strip(), entry_sn.get().strip(), entry_cust_id.get().strip(), entry_staff.get().strip()
    if not all([customer, model, sn, cust_id, staff]):
        messagebox.showwarning("提示", "欄位未填齊"); return
    
    sn_path = os.path.join(DB_PATH, customer, cust_id, sn)
    today = datetime.now().strftime('%Y%m%d')
    iqc_path = os.path.join(sn_path, f"{today}_IQC")
    
    try:
        os.makedirs(iqc_path, exist_ok=True)
        os.makedirs(os.path.join(sn_path, f"{today}_OQC"), exist_ok=True)
        
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        cursor.execute("INSERT INTO active_records (time, customer, cust_id, model, sn, staff, status, path) VALUES (?,?,?,?,?,?,?,?)",
                       (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        # 建立資料夾時，同步更新 Database 對照表
        cursor.execute("INSERT OR REPLACE INTO sn_lookup_db VALUES (?,?,?,?)", (sn, customer, cust_id, model))
        conn.commit(); conn.close()
        
        messagebox.showinfo("成功", "資料夾已建立"); os.startfile(iqc_path)
        refresh_search(); refresh_database_tab()
    except Exception as e: messagebox.showerror("錯誤", str(e))

def archive_done_records():
    """【維持原狀】原本的歸檔 CSV 匯出邏輯"""
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records")
    active_rows = cursor.fetchall()
    
    to_archive = []
    for r in active_rows:
        date_keyword = r[1][:10].replace("-", "")
        status, _ = get_folder_status(r[10], date_keyword)
        if status == "✅ 已完成": to_archive.append(r)
    
    if not to_archive:
        messagebox.showinfo("提示", "無「已完成」資料可歸檔"); conn.close(); return

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
            messagebox.showinfo("成功", f"歸檔完成，已匯出至 CSV"); 
        except Exception as e: messagebox.showerror("錯誤", f"歸檔失敗：{e}")
    conn.close(); refresh_search(); refresh_done_tab()

def import_to_database_csv():
    """【新功能】這是用來匯入您的『舊資料 CSV』到 Database 的功能"""
    path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not path: return
    try:
        conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
        with open(path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # 需確保『舊資料 CSV』標頭有: SN, 客戶名稱, 客戶編號, 產品型號
                cursor.execute("INSERT OR REPLACE INTO sn_lookup_db VALUES (?,?,?,?)",
                               (row['SN'], row['客戶名稱'], row['客戶編號'], row['產品型號']))
        conn.commit(); conn.close()
        messagebox.showinfo("成功", "Database 已從外部 CSV 載入舊資料！")
        refresh_database_tab()
    except Exception as e: messagebox.showerror("錯誤", f"匯入失敗：{e}")

# ==========================================
# 3. 介面重新整理
# ==========================================

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records ORDER BY id DESC")
    for r in cursor.fetchall():
        if any(query in str(x).lower() for x in r[2:7]):
            date_keyword = r[1][:10].replace("-", "")
            status_text, color_tag = get_folder_status(r[10], date_keyword)
            tree.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], status_text, "📂 開啟 IQC", "📂 開啟 OQC", r[10], r[0]), tags=(color_tag,))
    conn.close()

def refresh_done_tab(event=None):
    query = entry_done_search.get().strip().lower()
    for item in tree_done.get_children(): tree_done.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM done_records ORDER BY id DESC")
    for r in cursor.fetchall():
        if any(query in str(x).lower() for x in r[2:8]):
            tree_done.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], r[7], "✅ 已歸檔", "📂 開啟 IQC", "📂 開啟 OQC", r[8]), tags=("gray",))
    conn.close()

def refresh_database_tab(event=None):
    query = entry_db_search.get().strip().lower()
    for item in tree_database.get_children(): tree_database.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM sn_lookup_db")
    for r in cursor.fetchall():
        if any(query in str(x).lower() for x in r):
            tree_database.insert("", "end", values=r)
    conn.close()

# ==========================================
# 4. UI 佈局 (順序調整：先宣告再整理)
# ==========================================
root = tk.Tk()
root.title(f"IQC/OQC 管理系統 v{CURRENT_VERSION}")
root.geometry("1400x850")

FONT_MAIN, FONT_BOLD = ("微軟正黑體", 12), ("微軟正黑體", 12, "bold")
style = ttk.Style(root); style.theme_use('clam')
style.configure("Treeview", font=FONT_MAIN, rowheight=35)
style.configure("Treeview.Heading", font=FONT_BOLD)

# 輸入區
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_BOLD, padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

tk.Label(frame_input, text="SN 編號:").grid(row=0, column=0)
entry_sn = tk.Entry(frame_input, width=20, font=FONT_MAIN, bg="#E8F8FF")
entry_sn.grid(row=0, column=1, padx=5, pady=5)
entry_sn.bind("<FocusOut>", auto_fill_by_sn)

tk.Label(frame_input, text="客戶名稱:").grid(row=0, column=2); entry_customer = tk.Entry(frame_input, width=15, font=FONT_MAIN); entry_customer.grid(row=0, column=3)
tk.Label(frame_input, text="產品型號:").grid(row=0, column=4); entry_model = tk.Entry(frame_input, width=15, font=FONT_MAIN); entry_model.grid(row=0, column=5)
tk.Label(frame_input, text="客戶編號:").grid(row=1, column=0); entry_cust_id = tk.Entry(frame_input, width=20, font=FONT_MAIN); entry_cust_id.grid(row=1, column=1)
tk.Label(frame_input, text="作業人員:").grid(row=1, column=2); entry_staff = tk.Entry(frame_input, width=15, font=FONT_MAIN); entry_staff.grid(row=1, column=3)

tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BOLD).grid(row=0, column=6, rowspan=2, padx=20)
lbl_status_msg = tk.Label(frame_input, text="提示：輸入 SN 後按 Tab 鍵自動檢索", font=("微軟正黑體", 10), fg="gray")
lbl_status_msg.grid(row=2, column=0, columnspan=6, sticky="w")

notebook = ttk.Notebook(root); notebook.pack(fill="both", expand=True, padx=20, pady=10)

# --- 1. 進行中資料 ---
tab_active = tk.Frame(notebook); notebook.add(tab_active, text=" 進行中資料 ")
frame_search = tk.Frame(tab_active); frame_search.pack(fill="x", pady=5)
entry_search = tk.Entry(frame_search, font=FONT_MAIN); entry_search.pack(side="left", fill="x", expand=True, padx=10)
entry_search.bind("<KeyRelease>", refresh_search)
tk.Button(frame_search, text="📦 產出CSV歸檔", command=archive_done_records, bg="#27AE60", fg="white", font=FONT_BOLD).pack(side="right", padx=5)

cols_active = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "目前狀態", "IQC 資料夾", "OQC 資料夾", "路徑", "ID")
tree = ttk.Treeview(tab_active, columns=cols_active, show="headings")
tree.pack(fill="both", expand=True)
for c in cols_active: tree.heading(c, text=c); tree.column(c, width=100, anchor="center")
tree.column("路徑", width=0, stretch=False); tree.column("ID", width=0, stretch=False)
tree.tag_configure("green", background="#DFF2BF"); tree.tag_configure("orange", background="#FEEFB3"); tree.tag_configure("red", background="#FFBABA")

# --- 2. 已歸檔歷史 ---
tab_done = tk.Frame(notebook); notebook.add(tab_done, text=" 已歸檔歷史 ")
frame_done_search = tk.Frame(tab_done); frame_done_search.pack(fill="x", pady=5)
entry_done_search = tk.Entry(frame_done_search, font=FONT_MAIN); entry_done_search.pack(side="left", fill="x", expand=True, padx=10)
entry_done_search.bind("<KeyRelease>", refresh_done_tab)

cols_done = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "出貨日期", "狀態", "IQC 資料夾", "OQC 資料夾", "路徑")
tree_done = ttk.Treeview(tab_done, columns=cols_done, show="headings")
tree_done.pack(fill="both", expand=True)
for c in cols_done: tree_done.heading(c, text=c); tree_done.column(c, width=100, anchor="center")
tree_done.column("路徑", width=0, stretch=False); tree_done.tag_configure("gray", background="#F2F2F2")

# --- 3. Database ---
tab_db = tk.Frame(notebook); notebook.add(tab_db, text=" Database ")
frame_db_ctrl = tk.Frame(tab_db); frame_db_ctrl.pack(fill="x", pady=10)
entry_db_search = tk.Entry(frame_db_ctrl, font=FONT_MAIN); entry_db_search.pack(side="left", padx=10, fill="x", expand=True)
entry_db_search.bind("<KeyRelease>", refresh_database_tab)
tk.Button(frame_db_ctrl, text="📥 匯入舊資訊 CSV", command=import_to_database_csv, bg="#D4AC0D", fg="white").pack(side="right", padx=10)

cols_db = ("SN 編號", "客戶名稱", "客戶編號", "產品型號")
tree_database = ttk.Treeview(tab_db, columns=cols_db, show="headings")
tree_database.pack(fill="both", expand=True, padx=10, pady=5)
for c in cols_db: tree_database.heading(c, text=c); tree_database.column(c, anchor="center")

# --- 初始化 ---
refresh_search(); refresh_done_tab(); refresh_database_tab()
root.mainloop()
