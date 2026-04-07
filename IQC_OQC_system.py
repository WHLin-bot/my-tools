import os
import sys
import shutil
import sqlite3
import csv
import zipfile
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# ==========================================
# 0. 版本號與自動更新邏輯 (核心功能)
# ==========================================
# 每次改完程式要打包前，請手動增加這個版本號
CURRENT_VERSION = "4.4.3" 

# 更新檔存放目錄
UPDATE_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

def check_for_updates():
    """檢查伺服器上的版本號，若不同則提示下載"""
    # 只有在打包成 EXE 後才執行更新檢查，避免開發時一直跳通知
    if not getattr(sys, 'frozen', False): 
        return

    try:
        if not os.path.exists(VERSION_FILE_PATH):
            return

        with open(VERSION_FILE_PATH, "r", encoding="utf-8") as f:
            server_version = f.read().strip()

        if server_version != CURRENT_VERSION:
            msg = f"偵測到新版本 v{server_version} (目前 v{CURRENT_VERSION})！\n\n系統將自動下載更新檔至桌面，請解壓縮後取代舊版程式。\n是否現在執行更新？"
            if messagebox.askyesno("系統更新提示", msg):
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                target_zip_path = os.path.join(desktop_path, f"IQC_OQC_system_v{server_version}.zip")
                
                if os.path.exists(REMOTE_ZIP_PATH):
                    shutil.copy2(REMOTE_ZIP_PATH, target_zip_path)
                    messagebox.showinfo("下載完成", f"更新檔已存至桌面：\n{os.path.basename(target_zip_path)}\n\n請關閉程式並更新。")
                    os._exit(0) # 強制關閉程式
                else:
                    messagebox.showerror("錯誤", "伺服器上找不到更新壓縮包 (IQC_OQC_system.zip)")
    except Exception as e:
        print(f"更新檢查失敗: {e}")

# ==========================================
# 1. 路徑與資料庫設定
# ==========================================
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = UPDATE_DIR # 統一放在歷史紀錄區
SQLITE_DB = os.path.join(HIST_DIR, "ioqc_management.db")

os.makedirs(HIST_DIR, exist_ok=True)

def init_db():
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    # 建立進行中資料表
    cursor.execute('''CREATE TABLE IF NOT EXISTS active_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, status TEXT, iqc_done INTEGER DEFAULT 0, oqc_done INTEGER DEFAULT 0, path TEXT)''')
    # 建立已完成資料表
    cursor.execute('''CREATE TABLE IF NOT EXISTS done_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, ship_date TEXT, path TEXT)''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. 核心功能邏輯 (路徑：客戶/型號/SN)
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
    
    # 修正後的路徑：客戶 / 型號 / SN
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
        
        messagebox.showinfo("成功", f"資料夾已建立！\n型號：{model}\nSN：{sn}")
        if os.path.exists(iqc_path): os.startfile(iqc_path)
        refresh_search()
    except Exception as e: 
        messagebox.showerror("錯誤", f"建立失敗：{e}")

# (中間 refresh_search, refresh_done_tab, archive_done_records 邏輯同前，確保調用最新 path)

# ... [省略重複的 Treeview 刷新邏輯以節省空間，請沿用上一版代碼] ...

# ==========================================
# 3. UI 介面
# ==========================================
# [此處沿用上一版的完整 UI 代碼]

# 關鍵在最後這一行：啟動後檢查更新
root.after(1000, check_for_updates) # 延遲 1 秒執行避免影響視窗開啟速度
root.mainloop()
