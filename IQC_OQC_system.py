import os
import sys
import shutil
from tkinter import messagebox

# ==========================================
# 0. 版本號與手動更新引導 (退而求其次方案)
# ==========================================
CURRENT_VERSION = "4.4.0" 

# 伺服器路徑設定
UPDATE_DIR = r"\\fs2\Dept(Q)\08_p03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
VERSION_FILE_PATH = os.path.join(UPDATE_DIR, "version.txt")
REMOTE_ZIP_PATH = os.path.join(UPDATE_DIR, "IQC_OQC_system.zip")

def check_for_updates():
    """偵測更新並將最新壓縮檔下載到桌面"""
    if not getattr(sys, 'frozen', False):
        return  # 如果是在編輯器執行就不跑更新邏輯

    try:
        if not os.path.exists(VERSION_FILE_PATH):
            return
            
        with open(VERSION_FILE_PATH, "r", encoding="utf-8") as f:
            server_version = f.read().strip()
            
        if server_version != CURRENT_VERSION:
            msg = (
                f"偵測到新版本 v{server_version}！\n\n"
                "【更新步驟說明】\n"
                "1. 點擊「是」後，系統會將最新版壓縮檔下載到您的『桌面』。\n"
                "2. 程式將會自動關閉。\n"
                "3. 請您刪除舊的資料夾，並將桌面上的新壓縮檔解壓縮即可。\n\n"
                "是否現在下載新版本並關閉程式？"
            )
            ans = messagebox.askyesno("系統更新提示", msg)
            
            if ans:
                # 取得使用者桌面路徑
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                local_zip_name = f"IQC_OQC_system_v{server_version}.zip"
                target_zip_path = os.path.join(desktop_path, local_zip_name)
                
                # 從伺服器複製到桌面
                if os.path.exists(REMOTE_ZIP_PATH):
                    shutil.copy2(REMOTE_ZIP_PATH, target_zip_path)
                    messagebox.showinfo("下載完成", f"最新版已下載至桌面：\n{local_zip_name}\n\n程式即將關閉，請手動更新。")
                    os._exit(0) # 關閉程式
                else:
                    messagebox.showerror("錯誤", "伺服器上找不到壓縮檔，請聯絡系統管理員。")
                    
    except Exception as e:
        # 更新失敗不要卡住主程式開啟
        print(f"檢查更新時發生錯誤: {e}")

# ==========================================
# 1. 路徑與資料庫設定
# ==========================================
DB_PATH = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區"
HIST_DIR = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history"
SQLITE_DB = os.path.join(HIST_DIR, "ioqc_management.db")

os.makedirs(HIST_DIR, exist_ok=True)

def init_db():
    """初始化 SQLite 資料庫與資料表 - 確保欄位完全對齊"""
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    # 進行中資料表 (active_records)
    cursor.execute('''CREATE TABLE IF NOT EXISTS active_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, 
         time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, status TEXT, 
         iqc_done INTEGER DEFAULT 0, oqc_done INTEGER DEFAULT 0, 
         path TEXT)''')
    # 已歸檔資料表 (done_records)
    cursor.execute('''CREATE TABLE IF NOT EXISTS done_records 
        (id INTEGER PRIMARY KEY AUTOINCREMENT, 
         time TEXT, customer TEXT, cust_id TEXT, model TEXT, 
         sn TEXT, staff TEXT, ship_date TEXT, path TEXT)''')
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
                full_folder_path = os.path.join(sn_path, folder)
                if os.path.isdir(full_folder_path):
                    files = [os.path.join(full_folder_path, f) for f in os.listdir(full_folder_path) if os.path.isfile(os.path.join(full_folder_path, f))]
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

def create_folders():
    # 取得輸入框內容
    customer = entry_customer.get().strip()
    model = entry_model.get().strip()
    sn = entry_sn.get().strip()
    cust_id = entry_cust_id.get().strip()
    staff = entry_staff.get().strip()

    if not all([customer, model, sn, cust_id, staff]):
        messagebox.showwarning("提示", "請填寫所有欄位！")
        return
    
    # 建立路徑邏輯：伺服器根目錄 \ 客戶 \ ID \ SN
    sn_path = os.path.join(DB_PATH, customer, cust_id, sn)
    today = datetime.now().strftime('%Y%m%d')
    iqc_path = os.path.join(sn_path, f"{today}_IQC")
    oqc_path = os.path.join(sn_path, f"{today}_OQC")
    
    try:
        os.makedirs(iqc_path, exist_ok=True)
        os.makedirs(oqc_path, exist_ok=True)
        
        conn = sqlite3.connect(SQLITE_DB)
        cursor = conn.cursor()
        # 寫入資料庫，status 預設「進行中」，iqc/oqc_done 預設 0
        cursor.execute("""INSERT INTO active_records 
            (time, customer, cust_id, model, sn, staff, status, iqc_done, oqc_done, path) 
            VALUES (?,?,?,?,?,?,?,0,0,?)""",
            (datetime.now().strftime('%Y-%m-%d %H:%M'), customer, cust_id, model, sn, staff, "進行中", sn_path))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("成功", f"資料夾已建立！\n路徑：{sn_path}")
        if os.path.exists(iqc_path): os.startfile(iqc_path)
        refresh_search()
    except Exception as e: 
        messagebox.showerror("錯誤", f"建立失敗：{e}")

def refresh_search(event=None):
    query = entry_search.get().strip().lower()
    for item in tree.get_children(): tree.delete(item)
    
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    
    for r in rows:
        # r[10] 為 path
        targets = [str(x).lower() for x in r[2:7]]
        if any(query in t for t in targets):
            date_keyword = r[1][:10].replace("-", "").replace("/", "")
            sn_path = r[10] if r[10] else ""
            status_text, color_tag = get_folder_status(sn_path, date_keyword)
            # 插入 Treeview (注意索引對齊)
            tree.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], status_text, "📂 開啟 IQC", "📂 開啟 OQC", sn_path, r[0]), tags=(color_tag,))

def refresh_done_tab(event=None):
    query = entry_done_search.get().strip().lower()
    for item in tree_done.get_children(): tree_done.delete(item)
    
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM done_records ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    
    for r in rows:
        # r[8] 為 path
        targets = [str(x).lower() for x in r[2:8]]
        if any(query in t for t in targets):
            sn_path = r[8] if r[8] else ""
            tree_done.insert("", "end", values=(r[1], r[2], r[3], r[4], r[5], r[6], r[7], "✅ 已歸檔", "📂 開啟 IQC", "📂 開啟 OQC", sn_path), tags=("gray",))

def archive_done_records():
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM active_records")
    active_rows = cursor.fetchall()
    
    to_archive = []
    for r in active_rows:
        date_keyword = r[1][:10].replace("-", "").replace("/", "")
        sn_path = r[10] if r[10] else ""
        status, _ = get_folder_status(sn_path, date_keyword)
        if status == "✅ 已完成":
            to_archive.append(r)
    
    if not to_archive:
        messagebox.showinfo("提示", "沒有符合「已完成」狀態的資料。")
        conn.close()
        return

    csv_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")],
                                            initialfile=f"IOQC_Archive_{datetime.now().strftime('%Y%m%d')}.csv")
    if csv_path:
        try:
            export_data = []
            for r in to_archive:
                sn_path = r[10] if r[10] else ""
                ship_date = get_oqc_ship_date(sn_path)
                cursor.execute("INSERT INTO done_records (time, customer, cust_id, model, sn, staff, ship_date, path) VALUES (?,?,?,?,?,?,?,?)",
                               (r[1], r[2], r[3], r[4], r[5], r[6], ship_date, sn_path))
                cursor.execute("DELETE FROM active_records WHERE id = ?", (r[0],))
                export_data.append({"客戶名稱": r[2], "產品型號": r[4], "SN編號": r[5], "客戶編號": r[3], "建立日期": r[1][:10], "作業人員": r[6], "出貨日期": ship_date})

            conn.commit()
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=["客戶名稱", "產品型號", "SN編號", "客戶編號", "建立日期", "作業人員", "出貨日期"])
                writer.writeheader()
                writer.writerows(export_data)
            messagebox.showinfo("成功", f"已成功歸檔 {len(to_archive)} 筆資料！")
        except Exception as e: messagebox.showerror("錯誤", f"歸檔失敗：{e}")
    conn.close()
    refresh_search(); refresh_done_tab()

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
    
    create_time_str = vals[0]
    date_keyword = create_time_str[:10].replace("-", "").replace("/", "")
    
    # 根據不同分頁判定索引
    if tv == tree:
        path = vals[9]   # 隱藏路徑在索引 9
        iqc_btn_id, oqc_btn_id = "#8", "#9"
    else:
        path = vals[10]  # 隱藏路徑在索引 10
        iqc_btn_id, oqc_btn_id = "#9", "#10"

    if not path or path == "None":
        messagebox.showerror("錯誤", "路徑遺失，無法開啟資料夾。")
        return

    if col_id == iqc_btn_id:
        target_path = os.path.join(path, f"{date_keyword}_IQC")
        if os.path.exists(target_path): os.startfile(target_path)
        else: messagebox.showwarning("提示", f"找不到 IQC 資料夾：\n{target_path}")
            
    elif col_id == oqc_btn_id:
        target_path = os.path.join(path, f"{date_keyword}_OQC")
        if os.path.exists(target_path): os.startfile(target_path)
        else: messagebox.showwarning("提示", f"找不到 OQC 資料夾：\n{target_path}")

# ==========================================
# 3. UI 介面佈局
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

# 定義輸入框
tk.Label(frame_input, text="客戶名稱:", font=FONT_MAIN).grid(row=0, column=0, sticky="w", padx=5, pady=8)
entry_customer = tk.Entry(frame_input, width=18, font=FONT_MAIN)
entry_customer.grid(row=0, column=1, padx=5)

tk.Label(frame_input, text="產品型號:", font=FONT_MAIN).grid(row=0, column=2, sticky="w", padx=5, pady=8)
entry_model = tk.Entry(frame_input, width=18, font=FONT_MAIN)
entry_model.grid(row=0, column=3, padx=5)

tk.Label(frame_input, text="SN 編號:", font=FONT_MAIN).grid(row=0, column=4, sticky="w", padx=5, pady=8)
entry_sn = tk.Entry(frame_input, width=18, font=FONT_MAIN)
entry_sn.grid(row=0, column=5, padx=5)

tk.Label(frame_input, text="客戶編號:", font=FONT_MAIN).grid(row=1, column=0, sticky="w", padx=5, pady=8)
entry_cust_id = tk.Entry(frame_input, width=18, font=FONT_MAIN)
entry_cust_id.grid(row=1, column=1, padx=5)

tk.Label(frame_input, text="作業人員:", font=FONT_MAIN).grid(row=1, column=2, sticky="w", padx=5, pady=8)
entry_staff = tk.Entry(frame_input, width=18, font=FONT_MAIN)
entry_staff.grid(row=1, column=3, padx=5)

tk.Button(frame_input, text="➕ 建立資料夾", command=create_folders, bg="#2E86C1", fg="white", font=FONT_BOLD).grid(row=0, column=6, rowspan=2, padx=20, sticky="nswe")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=20, pady=10)

# --- 進行中分頁 ---
tab_active = tk.Frame(notebook)
notebook.add(tab_active, text=" 進行中資料 ")
frame_search = tk.Frame(tab_active); frame_search.pack(fill="x", pady=10, padx=10)
entry_search = tk.Entry(frame_search, font=FONT_MAIN); entry_search.pack(side="left", fill="x", expand=True, padx=10)
entry_search.bind("<KeyRelease>", refresh_search)
tk.Button(frame_search, text="🔄 刷新列表", command=refresh_search, font=FONT_MAIN).pack(side="left", padx=5)
tk.Button(frame_search, text="📦 產出CSV歸檔", command=archive_done_records, bg="#27AE60", fg="white", font=FONT_BOLD).pack(side="right", padx=5)

frame_tree_act = tk.Frame(tab_active); frame_tree_act.pack(fill="both", expand=True, padx=10, pady=5)
cols_active = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "目前狀態", "IQC 資料夾", "OQC 資料夾", "路徑", "ID")
tree = ttk.Treeview(frame_tree_act, columns=cols_active, show="headings")
vsb_act = ttk.Scrollbar(frame_tree_act, orient="vertical", command=tree.yview); vsb_act.pack(side="right", fill="y")
tree.configure(yscrollcommand=vsb_act.set); tree.pack(side="left", fill="both", expand=True)

for c in cols_active: 
    tree.heading(c, text=c, command=lambda _c=c: treeview_sort_column(tree, _c, False))
    tree.column(c, width=110, anchor="center")
tree.column("路徑", width=0, stretch=False); tree.column("ID", width=0, stretch=False)

# --- 已歸檔分頁 ---
tab_done = tk.Frame(notebook)
notebook.add(tab_done, text=" 已歸檔歷史 ")
frame_done_search = tk.Frame(tab_done); frame_done_search.pack(fill="x", pady=10, padx=10)
entry_done_search = tk.Entry(frame_done_search, font=FONT_MAIN); entry_done_search.pack(side="left", fill="x", expand=True, padx=10)
entry_done_search.bind("<KeyRelease>", refresh_done_tab)
tk.Button(frame_done_search, text="🔄 刷新列表", command=refresh_done_tab, font=FONT_MAIN).pack(side="left", padx=5)

frame_tree_dn = tk.Frame(tab_done); frame_tree_dn.pack(fill="both", expand=True, padx=10, pady=10)
cols_done = ("建立時間", "客戶", "客戶編號", "型號", "SN", "作業人員", "出貨日期", "狀態", "IQC 資料夾", "OQC 資料夾", "路徑")
tree_done = ttk.Treeview(frame_tree_dn, columns=cols_done, show="headings")
vsb_dn = ttk.Scrollbar(frame_tree_dn, orient="vertical", command=tree_done.yview); vsb_dn.pack(side="right", fill="y")
tree_done.configure(yscrollcommand=vsb_dn.set); tree_done.pack(side="left", fill="both", expand=True)

for c in cols_done: 
    tree_done.heading(c, text=c, command=lambda _c=c: treeview_sort_column(tree_done, _c, False))
    tree_done.column(c, width=110, anchor="center")
tree_done.column("路徑", width=0, stretch=False)

# --- 事件綁定 ---
for t in [tree, tree_done]:
    t.tag_configure("green", background="#DFF2BF", foreground="#270")
    t.tag_configure("orange", background="#FEEFB3", foreground="#9F6000")
    t.tag_configure("red", background="#FFBABA", foreground="#D8000C")
    t.tag_configure("gray", background="#F2F2F2", foreground="#666")
    t.bind("<ButtonRelease-1>", open_selected)

footer_frame = tk.Frame(root, bg="#F0F0F0", height=30); footer_frame.pack(side="bottom", fill="x")
lbl_db_path = tk.Label(footer_frame, text=f"💡 資料來源：{SQLITE_DB}", fg="#555555", font=("微軟正黑體", 10), bg="#F0F0F0", padx=20)
lbl_db_path.pack(side="left")

# 呼叫更新檢查
root.after(100, check_for_updates) # 延遲 0.1 秒執行，確保主視窗已經跑出來

refresh_search(); refresh_done_tab()
root.after(100, check_for_updates)
root.mainloop()
