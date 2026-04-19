import pandas as pd
import json

# 讀取 Excel 檔案
# 請確保 Excel 檔案在同一個資料夾，或使用完整路徑
df = pd.read_excel("HP93K.xlsx", sheet_name="名稱對應表")

# 假設：A欄是名稱 (A-01), B欄是POGO座標 (M12), C欄是正面座標 (C24)
data_dict = {}
for index, row in df.iterrows():
    name = str(row[0]).strip()
    pogo = str(row[1]).strip()
    coord = str(row[2]).strip()
    
    data_dict[name] = {"pogo": pogo, "coord": coord}

# 儲存成 JSON 檔案
with open('mapping.json', 'w', encoding='utf-8') as f:
    json.dump(data_dict, f, ensure_ascii=False, indent=4)

print("資料轉換完成！已生成 mapping.json")
