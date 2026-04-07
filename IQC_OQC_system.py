import os
import sys
import shutil
import sqlite3
import csv
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# ==========================================
# 0. 版本與路徑設定
# ==========================================
CURRENT_VERSION = "4.6.0" 
UPDATE_DIR = r"\\fs2\Dept(Q)\08_p03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
SQLITE_DB = os.path.join(HIST_DIR, "ioqc_management.db")

os.makedirs(HIST_DIR, exist_ok=True)

# ==========================================
# 1. 資料庫初始化
# ==========================================
def init_db():
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    # 進行中
    cursor.execute('''CREATE TABLE IF NOT EXISTS active_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, 
         time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, status TEXT, 
         iqc_done INTEGER DEFAULT 0, oqc_done INTEGER DEFAULT 0, 
         path TEXT)''')
    # 已歸檔歷史 (原本的，不做改變)
    cursor.execute('''CREATE TABLE IF NOT EXISTS done_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, 
         time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, ship_date TEXT, path TEXT)''')
    # 第三個 Sheet: Database (SN 對照表)
    cursor.execute('''CREATE TABLE IF NOT EXISTS sn_lookup_db 
        (sn TEXT PRIMARY KEY, customer TEXT, cust_id TEXT, model TEXT)''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. 核心功能邏輯
# ==========================================

def auto_fill_by_sn(event=None):
    """從 Database 搜尋 SN 並自動帶入"""
    sn = entry_sn.get().strip()
    if not sn: return
    
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT customer, cust_id, model FROM sn_lookup_db WHERE sn = ?", (sn,))
    res = cursor.fetchone()
    conn.close()
    
    if res:
        # 清除並填入資料
        entry_customer.delete(0, tk.END); entry_customer.insert(0, res[0])
        entry_cust_id.delete(0, tk.END); entry_cust_id.insert(0, res[1])
        entry_model.delete(0, tk.END); entry_model.insert(0, res[2])
        lbl_status_msg.config(text=f"✅ 已從 Database 帶入 SN: {sn}", fg="blue")
    else:
        lbl_status_msg.config(text="🆕 這是新 SN，請手動輸入完整資訊", fg="orange")

def create_folders():
    customer = entry_customer.get().strip()
    model = entry_model.get().strip()
    sn = entry_sn.get().strip()
    cust_id = entry_cust_id.get().strip()
    staff = entry_staff.get().strip()

    if not all([customer, model, sn, cust_id, staff]):
        messagebox.showwarning("提示", "請填寫所有欄位！")
        return
    
    sn_path = os.path.join(DB_PATH, customer, cust_id, sn)
    today = datetime.now().strftime('%Y%m%d')
    iqc_path = os.path.join(sn_path, f"{today}_IQC")
    oqc_path = os.path.join(sn_path, f"{today}_OQC")
    
    try:
        os.makedirs(iqc_path, exist_ok=True)
        os.makedirs(oqc_path, exist_ok=True)
        
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        # 紀錄於進行中
        cursor.execute("""INSERT INTO active_records 
            (time, customer, cust_id, model, sn, staff, status, path) 
            VALUES (?,?,?,?,?,?,?,?)""",
            (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        
        # 同步更新至 Database (第三個 Sheet)
        cursor.execute("INSERT OR REPLACE INTO sn_lookup_db (sn, customer, cust_id, model) VALUES (?,?,?,?)",
                       (sn, customer, cust_id, model))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("成功", f"資料夾已建立！")
        if os.path.exists(iqc_path): os.startfile(iqc_path)
        refresh_search(); refresh_database_tab()
    except Exception as e: 
        messagebox.showerror("錯誤", f"建立失敗：{e}")

def import_to_database_csv():
    """匯入 CSV 到 Database 分頁"""
    path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not path: return
    try:
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        with open(path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # 需確保 CSV 標頭有: SN, 客戶名稱, 客戶編號, 產品型號
                cursor.execute("INSERT OR REPLACE INTO sn_lookup_db VALUES (?,?,?,?)",
                               (row['SN'], row['客戶名稱'], row['客戶編號'], row['產品型號']))
        conn.commit(); conn.close()
        messagebox.showinfo("成功", "Database 已更新！")
        refresh_database_tab()
    except Exception as e:
        messagebox.showerror("錯誤", f"匯入失敗：{e}")

# (其他 refresh_search, refresh_done_tab 等維持您原本的 code)

def refresh_database_tab(event=None):
    query = entry_db_search.get().strip().lower()
    for item in tree_database.get_children(): tree_database.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM sn_lookup_db")
    rows = cursor.fetchall(); conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r):
            tree_database.insert("", "end", values=r)

# ==========================================
# 3. UI 介面佈局
# ==========================================
root = tk.Tk()
root.title(f"IQC/OQC 管理系統 v{CURRENT_VERSION}")
root.geometry("1400x850")

FONT_MAIN, FONT_BOLD = ("微軟正黑體", 12), ("微軟正黑體", 12, "bold")
style = ttk.Style()
style.theme_use('clam')
style.configure("Treeview", font=FONT_MAIN, rowheight=35)

# --- 輸入區 ---
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_BOLD, padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

tk.Label(frame_input, text="SN 編號:", font=FONT_BOLD).grid(row=0, column=0, sticky="w")
entry_sn = tk.Entry(frame_input, width=20, font=FONT_MAIN, bg="#E8F8FF") # 淺藍提示
entry_sn.grid(row=0, column=1, padx=5, pady=5)
entry_sn.bind("<FocusOut>", auto_fill_by_sn)
entry_sn.bind("<Return>", auto_fill_by_sn)

tk.Label(frame_input, text="客戶名稱:").grid(row=0, column=2, padx=5)
entry_customer = tk.Entry(frame_input, width=15, font=FONT_MAIN); entry_customer.grid(row=0, column=3)

tk.Label(frame_input, text="產品型號:").grid(row=0, column=4, padx=5)
entry_model = tk.Entry(frame_input, width=15, font=FONT_MAIN); entry_model.grid(row=0, column=5)

tk.Label(frame_input, text="客戶編號:").grid(row=1, column=0, sticky="w")
entry_cust_id = tk.Entry(frame_input, width=20, font=FONT_MAIN); entry_cust_id.grid(row=1, column=1)

tk.Label(frame_input, text="作業人員:").grid(row=1, column=2)
entry_staff = tk.Entry(frame_input, width=15, font=FONT_MAIN); entry_staff.grid(row=1, column=3)

btn_create = tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BOLD, width=15)
btn_create.grid(row=0, column=6, rowspan=2, padx=20)

lbl_status_msg = tk.Label(frame_input, text="提示：輸入 SN 後按 Tab 或 Enter 可帶入舊資料", font=("微軟正黑體", 10), fg="gray")
lbl_status_msg.grid(row=2, column=0, columnspan=6, sticky="w")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=20, pady=10)

# 1. 進行中資料 (同原 code)
tab_active = tk.Frame(notebook); notebook.add(tab_active, text=" 進行中資料 ")
# ... (Treeview tree 建立邏輯)

# 2. 已歸檔歷史 (原本的，不做改變)
tab_done = tk.Frame(notebook); notebook.add(tab_done, text=" 已歸檔歷史 ")
# ... (Treeview tree_done 建立邏輯)

# 3. Database 分頁 (新增加的)
tab_db = tk.Frame(notebook); notebook.add(tab_db, text=" Database ")
frame_db_ctrl = tk.Frame(tab_db); frame_db_ctrl.pack(fill="x", pady=10)
entry_db_search = tk.Entry(frame_db_ctrl, font=FONT_MAIN); entry_db_search.pack(side="left", padx=10, fill="x", expand=True)
entry_db_search.bind("<KeyRelease>", refresh_database_tab)
tk.Button(frame_db_ctrl, text="📥 匯入舊資訊 CSV", command=import_to_database_csv, bg="#D4AC0D", fg="white").pack(side="right", padx=10)

cols_db = ("SN 編號", "客戶名稱", "客戶編號", "產品型號")
tree_database = ttk.Treeview(tab_db, columns=cols_db, show="headings")
tree_database.pack(fill="both", expand=True, padx=10, pady=5)
for c in cols_db: 
    tree_database.heading(c, text=c)
    tree_database.column(c, anchor="center")

# 啟動與初始化
refresh_search(); refresh_database_tab()
root.mainloop()
