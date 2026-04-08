import sqlite3
import os

# 請確認你的資料庫路徑
SQLITE_DB = r"\\fs2\Dept(Q)\08_品保處\03_客戶服務部\99_Public\IQC_OQC_新竹_進出料檢照片區\IOQC_folder_history\ioqc_management.db"

def migrate_database():
    if not os.path.exists(SQLITE_DB):
        print("找不到資料庫檔案，請確認路徑！")
        return

    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()

    try:
        # 1. 將舊的 done_records 改名為 history_records
        # 檢查 history_records 是否已存在，避免重複執行報錯
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='history_records'")
        if not cursor.fetchone():
            cursor.execute("ALTER TABLE done_records RENAME TO history_records")
            print("成功：已將舊的 done_records 修改為 history_records")
        else:
            print("提示：history_records 已經存在，略過改名步驟")

        # 2. 建立新的 done_records 給 Tab3 使用
        # 結構與 active_records 相似，但不需要 status 欄位（因為到這裡一定都是已完成）
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS done_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                time TEXT,
                customer TEXT,
                cust_id TEXT,
                model TEXT,
                sn TEXT,
                staff TEXT,
                path TEXT
            )
        """)
        print("成功：已建立新的 done_records 資料表")

        conn.commit()
        print("\n--- 資料庫轉換完成！現在你可以執行 v4.6.0 的程式了 ---")
    except Exception as e:
        print(f"轉換失敗: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    migrate_database()
