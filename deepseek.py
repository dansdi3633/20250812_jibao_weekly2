import pandas as pd
import openpyxl

def compare_holdings(file1, file2, output_file):
    # 讀取檔案
    df1 = pd.read_csv(file1, encoding='utf-8')
    df2 = pd.read_csv(file2, encoding='utf-8')
    
    # 過濾持股分級為15
    df1_level15 = df1[df1['持股分級'] == 15]
    df2_level15 = df2[df2['持股分級'] == 15]
    
    # 處理證券代號格式
    df1_level15['證券代號'] = df1_level15['證券代號'].astype(str).str.strip()
    df2_level15['證券代號'] = df2_level15['證券代號'].astype(str).str.strip()

    # 合併時使用outer，保留全部證券代號
    merged = pd.merge(
        df1_level15[['證券代號', '占集保庫存數比例%']],
        df2_level15[['證券代號', '占集保庫存數比例%']],
        on='證券代號',
        how='outer',
        suffixes=('_20250124', '_20250808')
    )
    
    # 計算變化值（NaN 以 0 處理或自訂）
    merged = merged.fillna(0)
    merged['變化值'] = merged['占集保庫存數比例%_20250808'] - merged['占集保庫存數比例%_20250124']
    
    # 排序、欄位名稱
    merged_sorted = merged.sort_values('變化值', ascending=False)
    result = merged_sorted[['證券代號', '占集保庫存數比例%_20250124', '占集保庫存數比例%_20250808', '變化值']]
    result.columns = ['證券代號', '2025/01/24比例', '2025/08/08比例', '變化值']
    
    # 輸出
    result.to_excel(output_file, index=False)
    print(f"結果已保存到 {output_file}")

if __name__ == "__main__":
    file1 = "20250124.csv"
    file2 = "20250808.csv"
    output_file = "持股分級15比例變化.xlsx"
    compare_holdings(file1, file2, output_file)