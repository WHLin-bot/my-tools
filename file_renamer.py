import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog  

class AdvancedFileRenamer:
    def __init__(self, root):
        self.root = root
        self.root.title("企業檔案機密等級批次標籤工具")
        self.root.geometry("950x650")  
        
        self.node_status = {} 
        self.selected_root_dir = ""

        self.UNCHECKED = "⬜ "
        self.CHECKED = "✅ "
        self.FOLDER_ICON = "📁 "
        self.FILE_ICON = "📄 "
        
        self.extensions = ('.docx', '.xlsx', '.xlsm', '.pptx', '.pdf')
        self.all_levels = ["【限閱】", "【一般】", "【機密】", "【極機密】"]

        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft JhengHei", 12)) # 這裡的 10 可以改成你想要的尺寸 (例如 11 或 12)
        
        top_frame = tk.LabelFrame(self.root, text=" 1. 設定變更參數 ", font=("Microsoft JhengHei", 14, "bold"), padx=10, pady=10)
        top_frame.pack(pady=10, padx=15, fill="x")

        lbl_desc = tk.Label(top_frame, text="此工具僅適用於'docx, xlsx, xlsm, pdf'副檔名，如有遺漏請自行手動修改", fg="#8B0000", font=("Microsoft JhengHei", 12, "bold"))
        lbl_desc.pack(anchor="w", pady=(0, 5))

        lbl_desc2 = tk.Label(top_frame, text="請選擇欲套用的機密等級標籤。可在左側樹狀圖中單獨勾選特定檔案或整個資料夾。", fg="#2B579A", font=("Microsoft JhengHei", 12, "bold"))
        lbl_desc2.pack(anchor="w", pady=(0, 5))

        ctrl_layout = tk.Frame(top_frame)
        ctrl_layout.pack(fill="x", expand=True)

        self.level_var = tk.StringVar(value="【一般】")
        levels = [("限閱", "【限閱】"), ("一般", "【一般】"), ("機密", "【機密】"), ("極機密", "【極機密】")]
        
        rb_frame = tk.Frame(ctrl_layout)
        rb_frame.pack(side="left", fill="y")
        for text, value in levels:
            rb = tk.Radiobutton(rb_frame, text=text, variable=self.level_var, value=value, font=("Microsoft JhengHei", 12, "bold"))
            rb.pack(side="left", padx=10)

        lbl_sep = tk.Label(ctrl_layout, text=" | ", fg="#A0A0A0", font=("Microsoft JhengHei", 12))
        lbl_sep.pack(side="left", padx=10)

        self.wave_force_top_secret = tk.BooleanVar(value=False)
        chk_wave = tk.Checkbutton(ctrl_layout, text="檔名含 '~~' 則一律設為【極機密】(忽略左側選擇)", 
                                  variable=self.wave_force_top_secret, fg="#D32F2F", font=("Microsoft JhengHei", 12, "bold"))
        chk_wave.pack(side="left", padx=10)

        path_frame = tk.Frame(self.root)
        path_frame.pack(pady=5, padx=15, fill="x")

        self.entry_path = tk.Entry(path_frame, font=("Microsoft JhengHei", 12))
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.entry_path.insert(0, "請點擊右側按鈕載入目標主資料夾...")
        self.entry_path.config(state="disabled")

        btn_browse = tk.Button(path_frame, text="載入主資料夾...", command=self.load_main_folder, bg="#008CBA", fg="white", font=("Microsoft JhengHei", 12, "bold"))
        btn_browse.pack(side="right")

        main_layout = tk.Frame(self.root)
        main_layout.pack(pady=5, padx=15, fill="both", expand=True)

        tree_panel = tk.LabelFrame(main_layout, text=" 2. 勾選要處理的資料夾或檔案 ", font=("Microsoft JhengHei", 14, "bold"))
        tree_panel.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.tree = ttk.Treeview(tree_panel, columns=("path", "type"), show="tree", selectmode="none")
        self.tree.column("#0", width=1200, minwidth=300, stretch=tk.YES) 
        
        self.tree.column("path", width=0, stretch=tk.NO)
        self.tree.column("type", width=0, stretch=tk.NO)

        tree_scroll_y = ttk.Scrollbar(tree_panel, orient="vertical", command=self.tree.yview)
        tree_scroll_y.pack(side="right", fill="y")
        
        tree_scroll_x = ttk.Scrollbar(tree_panel, orient="horizontal", command=self.tree.xview)
        tree_scroll_x.pack(side="bottom", fill="x")
        
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)

        log_panel = tk.LabelFrame(main_layout, text=" 3. 執行紀錄 ", font=("Microsoft JhengHei", 12, "bold"))
        log_panel.pack(side="right", fill="both", expand=True, padx=(5, 0))

        self.log_box = scrolledtext.ScrolledText(log_panel, font=("Consolas", 2), bg="#F9F9F9")
        self.log_box.pack(fill="both", expand=True)

        self.btn_start = tk.Button(self.root, text="開始批次歸類更名 (僅處理勾選項目)", command=self.start_rename, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 12, "bold"), height=2, state="disabled")
        self.btn_start.pack(pady=15, padx=15, fill="x")

    def load_main_folder(self):
        folder_selected = filedialog.askdirectory()
        if not folder_selected:
            return

        self.selected_root_dir = os.path.normpath(folder_selected)
        self.entry_path.config(state="normal")
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, self.selected_root_dir)
        self.entry_path.config(state="disabled")

        self.tree.delete(*self.tree.get_children())
        self.node_status.clear()
        self.log_box.delete(1.0, tk.END)

        self.log_box.insert(tk.END, f"正在掃描資料夾與檔案結構：{self.selected_root_dir} ...\n")
        
        root_name = os.path.basename(self.selected_root_dir) or self.selected_root_dir
        root_node = self.tree.insert("", "end", text=f"{self.UNCHECKED}{self.FOLDER_ICON}{root_name}", values=(self.selected_root_dir, "dir"), open=True)
        self.node_status[root_node] = False

        self.build_tree(self.selected_root_dir, root_node)
        
        self.log_box.insert(tk.END, "✅ 資料夾與檔案樹狀圖載入完成！您可以單獨勾選檔案或資料夾。\n")
        self.btn_start.config(state="normal")

    def build_tree(self, current_dir, parent_node):
        try:
            for entry in os.scandir(current_dir):
                path = os.path.normpath(entry.path)
                if entry.is_dir():
                    node = self.tree.insert(parent_node, "end", text=f"{self.UNCHECKED}{self.FOLDER_ICON}{entry.name}", values=(path, "dir"), open=False)
                    self.node_status[node] = False
                    self.build_tree(path, node)
                elif entry.is_file() and entry.name.lower().endswith(self.extensions):
                    node = self.tree.insert(parent_node, "end", text=f"{self.UNCHECKED}{self.FILE_ICON}{entry.name}", values=(path, "file"))
                    self.node_status[node] = False
        except Exception:
            pass

    def on_tree_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        element = self.tree.identify_element(event.x, event.y)
        if element != "text":
            return

        current_status = not self.node_status.get(item_id, False)
        self.set_node_status(item_id, current_status)
        self.set_child_nodes_status(item_id, current_status)

    def set_node_status(self, item_id, status):
        self.node_status[item_id] = status
        current_text = self.tree.item(item_id, "text")
        clean_text = current_text.replace(self.CHECKED, "").replace(self.UNCHECKED, "")
        icon = self.CHECKED if status else self.UNCHECKED
        self.tree.item(item_id, text=f"{icon}{clean_text}")

    def set_child_nodes_status(self, parent_id, status):
        for child_id in self.tree.get_children(parent_id):
            self.set_node_status(child_id, status)
            self.set_child_nodes_status(child_id, status)

    def get_selected_file_nodes(self):
        selected_nodes = []
        for node_id, is_checked in self.node_status.items():
            if is_checked:
                values = self.tree.item(node_id, "values")
                if values and values[1] == "file":  
                    selected_nodes.append((node_id, values[0])) 
        return selected_nodes

    def start_rename(self):
        selected_file_nodes = self.get_selected_file_nodes()
        if not selected_file_nodes:
            messagebox.showwarning("警告", "請至少勾選一個『檔案』進行處理！")
            return

        default_prefix = self.level_var.get()
        wave_rule_active = self.wave_force_top_secret.get()

        confirm_msg = f"確定要將已勾選的 {len(selected_file_nodes)} 個檔案進行標籤調整嗎？\n\n步驟一：一律調整為 ➔ {default_prefix}"
        if wave_rule_active:
            confirm_msg += "\n步驟二：若檔名內含 '~~' ➔ 強制升級為 【極機密】"

        if not messagebox.askyesno("確認執行", confirm_msg):
            return

        self.btn_start.config(state="disabled", text="處理中...")

        threading.Thread(
            target=self._execute_rename_thread, 
            args=(selected_file_nodes, default_prefix, wave_rule_active),
            daemon=True
        ).start()

    def _execute_rename_thread(self, selected_file_nodes, default_prefix, wave_rule_active):
        self.log_box.delete(1.0, tk.END)
        self.log_box.insert(tk.END, "🚀 開始執行批次檔案更名...\n\n")
        
        success_count = 0
        skip_count = 0
        updated_count = 0
        ui_updates = []

        for node_id, file_path in selected_file_nodes:
            if not os.path.exists(file_path):
                self.log_box.insert(tk.END, f"  ❌ 找不到檔案: {os.path.basename(file_path)}\n")
                continue
                
            folder = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            
            target_prefix = default_prefix
            is_wave_triggered = False

            clean_name = file_name
            existing_prefix = None
            for p in self.all_levels:
                if file_name.startswith(p):
                    existing_prefix = p
                    clean_name = file_name[len(p):]
                    break
            base_new_name = target_prefix + clean_name

            if wave_rule_active and "~~" in base_new_name:
                target_prefix = "【極機密】"
                base_new_name = target_prefix + clean_name
                is_wave_triggered = True

            if existing_prefix == target_prefix:
                if is_wave_triggered:
                    self.log_box.insert(tk.END, f"  ⚠️ 跳過 (含波浪號且原本已是極機密): {file_name}\n")
                else:
                    self.log_box.insert(tk.END, f"  ⚠️ 跳過 (等級已與預設相符): {file_name}\n")
                skip_count += 1
                continue

            new_path = os.path.normpath(os.path.join(folder, base_new_name))
            try:
                os.rename(file_path, new_path)
                
                if is_wave_triggered:
                    if existing_prefix:
                        log_msg = f"  🚨  {file_name} ➔ 先調為 {default_prefix} ➔ 終點強制升級: {base_new_name}\n"
                    else:
                        log_msg = f"  🚨  {file_name} ➔ 偵測波浪號直接強制歸類: {base_new_name}\n"
                    updated_count += 1
                else:
                    log_msg = f"  🔄 變更等級: {file_name} ➔ {base_new_name}\n" if existing_prefix else f"  ✅ 成功歸類: {file_name} ➔ {base_new_name}\n"
                    if existing_prefix: updated_count += 1
                    else: success_count += 1

                self.log_box.insert(tk.END, log_msg)
                ui_updates.append((node_id, base_new_name, new_path))
            except Exception as e:
                self.log_box.insert(tk.END, f"  ❌ 失敗: {file_name} (原因: {str(e)})\n")
            
            self.log_box.see(tk.END)

        if ui_updates:
            self.log_box.insert(tk.END, "\n🔄 正在批量刷新列表畫面...")
            for node_id, new_name, new_path in ui_updates:
                icon = self.CHECKED if self.node_status[node_id] else self.UNCHECKED
                self.tree.item(node_id, text=f"{icon}{self.FILE_ICON}{new_name}")
                self.tree.item(node_id, values=(new_path, "file"))
            self.log_box.insert(tk.END, " 完成！\n")

        self.btn_start.config(state="normal", text="開始批次歸類更名 (僅處理勾選項目)")

        total_done = success_count + updated_count
        summary = f"\n🎉 任務執行完畢！\n新歸類檔案: {success_count} 個\n變更等級檔案: {updated_count} 個\n跳過(無須變動): {skip_count} 個\n"
        self.log_box.insert(tk.END, summary)
        self.log_box.see(tk.END)
        
        messagebox.showinfo("完成", f"批次處理完成！\n總共變更: {total_done} 個檔案\n跳過不變: {skip_count} 個檔案\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedFileRenamer(root)
    root.mainloop()
