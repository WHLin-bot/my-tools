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
CURRENT_VERSION = "4.5.0" 
UPDATE_DIR = r"\\fs2\Dept(Q)\08_p03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
SQLITE_DB = os.path.join(HIST_DIR, "ioqc_management.db")

os.makedirs(HIST_DIR, exist_ok=True)

# ==========================================
# 1. 資料庫初始化 (新增 sn_history 表)
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
    # 已歸檔
    cursor.execute('''CREATE TABLE IF NOT EXISTS done_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, 
         time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, ship_date TEXT, path TEXT)''')
    # SN 歷史資訊 (用於自動帶入)
    cursor.execute('''CREATE TABLE IF NOT EXISTS sn_history 
        (sn TEXT PRIMARY KEY, customer TEXT, cust_id TEXT, model TEXT)''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. 核心功能邏輯
# ==========================================

def check_for_updates():
    if not getattr(sys, 'frozen', False): return
    try:
        if not os.path.exists(VERSION_FILE_PATH): return
        with open(VERSION_FILE_PATH, "r", encoding="utf-8") as f:
            server_version = f.read().strip()
        if server_version != CURRENT_VERSION:
            msg = f"偵測到新版本 v{server_version}！\n\n是否現在下載並關閉程式？"
            if messagebox.askyesno("更新提示", msg):
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                target = os.path.join(desktop, f"IQC_OQC_system_v{server_version}.zip")
                if os.path.exists(REMOTE_ZIP_PATH):
                    shutil.copy2(REMOTE_ZIP_PATH, target)
                    messagebox.showinfo("成功", "已下載至桌面，程式將關閉。")
                    os._exit(0)
    except: pass

def auto_fill_by_sn(event=None):
    """根據輸入的 SN 自動從資料庫抓取歷史資訊"""
    sn = entry_sn.get().strip()
    if not sn: return
    
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT customer, cust_id, model FROM sn_history WHERE sn = ?", (sn,))
    res = cursor.fetchone()
    conn.close()
    
    if res:
        entry_customer.delete(0, tk.END); entry_customer.insert(0, res[0])
        entry_cust_id.delete(0, tk.END); entry_cust_id.insert(0, res[1])
        entry_model.delete(0, tk.END); entry_model.insert(0, res[2])
        lbl_status_msg.config(text=f"✨ 已自動帶入 SN: {sn} 的歷史資料", fg="blue")
    else:
        lbl_status_msg.config(text="🆕 新 SN 編號，請手動輸入資訊", fg="orange")

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
        # 1. 寫入進行中資料表
        cursor.execute("""INSERT INTO active_records 
            (time, customer, cust_id, model, sn, staff, status, path) 
            VALUES (?,?,?,?,?,?,?,?)""",
            (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        # 2. 更新/存入 SN 歷史表，方便下次自動帶入
        cursor.execute("INSERT OR REPLACE INTO sn_history (sn, customer, cust_id, model) VALUES (?,?,?,?)",
                       (sn, customer, cust_id, model))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("成功", f"資料夾已建立！")
        if os.path.exists(iqc_path): os.startfile(iqc_path)
        refresh_search(); refresh_history_tab()
    except Exception as e: 
        messagebox.showerror("錯誤", f"建立失敗：{e}")

def import_history_csv():
    """匯入舊有的 CSV 資料至 sn_history"""
    path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not path: return
    try:
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        with open(path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            # 預期 CSV 標頭：SN, 客戶名稱, 客戶編號, 產品型號
            for row in reader:
                cursor.execute("INSERT OR REPLACE INTO sn_history VALUES (?,?,?,?)",
                               (row['SN'], row['客戶名稱'], row['客戶編號'], row['產品型號']))
        conn.commit()
        conn.close()
        messagebox.showinfo("成功", "歷史資訊匯入完成！")
        refresh_history_tab()
    except Exception as e:
        messagebox.showerror("錯誤", f"匯入失敗，請檢查 CSV 格式：\n{e}")

# (其餘 get_folder_status, get_oqc_ship_date, archive_done_records 等邏輯與你提供的相同，此處略作整合)
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

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records ORDER BY id DESC")
    rows = cursor.fetchall(); conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r[2:7]):
            date_keyword = r[1][:10].replace("-", "")
            status_text, color_tag = get_folder_status(r[10], date_keyword)
            tree.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], status_text, "📂 開啟 IQC", "📂 開啟 OQC", r[10], r[0]), tags=(color_tag,))

def refresh_history_tab(event=None):
    query = entry_hist_search.get().strip().lower()
    for item in tree_history.get_children(): tree_history.delete(item)
    conn = sqlite3.connect(SQLITE_DB); cursor = conn.cursor()
    cursor.execute("SELECT * FROM sn_history")
    rows = cursor.fetchall(); conn.close()
    for r in rows:
        if any(query in str(x).lower() for x in r):
            tree_history.insert("", "end", values=r)

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

# 輸入區
frame_input = tk.LabelFrame(root, text=" 建立IQC資訊 ", font=FONT_BOLD, padx=15, pady=15)
frame_input.pack(fill="x", padx=20, pady=10)

tk.Label(frame_input, text="SN 編號 (輸入完按 Tab):", font=FONT_BOLD).grid(row=0, column=0, sticky="w")
entry_sn = tk.Entry(frame_input, width=20, font=FONT_MAIN, bg="#FFF9C4")
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

lbl_status_msg = tk.Label(frame_input, text="請輸入 SN 開始作業", font=("微軟正黑體", 10))
lbl_status_msg.grid(row=2, column=0, columnspan=6, sticky="w")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=20, pady=10)

# --- 進行中分頁 (略，參考你原本的代碼) ---
tab_active = tk.Frame(notebook); notebook.add(tab_active, text=" 進行中資料 ")
frame_search = tk.Frame(tab_active); frame_search.pack(fill="x", pady=5)
entry_search = tk.Entry(frame_search, font=FONT_MAIN); entry_search.pack(side="left", fill="x", expand=True, padx=10)
entry_search.bind("<KeyRelease>", refresh_search)
# Treeview 定義... (略，保持你原本的 cols_active 與 tree 設定)
cols_active = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "目前狀態", "IQC 資料夾", "OQC 資料夾", "路徑", "ID")
tree = ttk.Treeview(tab_active, columns=cols_active, show="headings")
tree.pack(fill="both", expand=True)
for c in cols_active: tree.heading(c, text=c); tree.column(c, width=100, anchor="center")
tree.column("路徑", width=0, stretch=False); tree.column("ID", width=0, stretch=False)

# --- 歷史資訊管理分頁 (新增) ---
tab_hist = tk.Frame(notebook); notebook.add(tab_hist, text=" 歷史 SN 對照表 ")
frame_hist_ctrl = tk.Frame(tab_hist); frame_hist_ctrl.pack(fill="x", pady=10)
entry_hist_search = tk.Entry(frame_hist_ctrl, font=FONT_MAIN); entry_hist_search.pack(side="left", padx=10, fill="x", expand=True)
entry_hist_search.bind("<KeyRelease>", refresh_history_tab)
tk.Button(frame_hist_ctrl, text="📥 匯入舊資料 CSV", command=import_history_csv, bg="#8E44AD", fg="white").pack(side="right", padx=10)

cols_hist = ("SN 編號", "客戶名稱", "客戶編號", "產品型號")
tree_history = ttk.Treeview(tab_hist, columns=cols_hist, show="headings")
tree_history.pack(fill="both", expand=True, padx=10, pady=5)
for c in cols_hist: tree_history.heading(c, text=c); tree_history.column(c, anchor="center")

# 啟動與初始化
root.after(100, check_for_updates)
refresh_search(); refresh_history_tab()
root.mainloop()
