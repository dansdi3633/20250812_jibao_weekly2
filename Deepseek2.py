import pandas as pd
import openpyxl
import os

def compare_holdings(file1, file2, output_file):
    # 讀取檔案
    df1 = pd.read_csv(file1, encoding='utf-8')
    df2 = pd.read_csv(file2, encoding='utf-8')
    
    # 處理證券代號格式
    df1['證券代號'] = df1['證券代號'].astype(str).str.strip()
    df2['證券代號'] = df2['證券代號'].astype(str).str.strip()

    # 個別15級比例
    df1_level15 = df1[df1['持股分級'] == 15][['證券代號','占集保庫存數比例%']]
    df2_level15 = df2[df2['持股分級'] == 15][['證券代號','占集保庫存數比例%']]

    # 合併個別15級
    merged = pd.merge(
        df1_level15,
        df2_level15,
        on='證券代號',
        how='outer',
        suffixes=('_20250124', '_20250808')
    ).fillna(0)
    merged['變化值'] = merged['占集保庫存數比例%_20250808'] - merged['占集保庫存數比例%_20250124']
    
    # 累計12~15級——20250124
    df1_lv12_15 = df1[df1['持股分級'].between(12, 15)].groupby('證券代號')['占集保庫存數比例%'].sum().reset_index()
    df1_lv12_15.columns = ['證券代號', '累積12~15_20250124']
    # 累計12~15級——20250808
    df2_lv12_15 = df2[df2['持股分級'].between(12, 15)].groupby('證券代號')['占集保庫存數比例%'].sum().reset_index()
    df2_lv12_15.columns = ['證券代號', '累積12~15_20250808']

    # 合併到主表
    merged = pd.merge(merged, df1_lv12_15, on='證券代號', how='left')
    merged = pd.merge(merged, df2_lv12_15, on='證券代號', how='left')
    merged = merged.fillna(0)

    # 新增累積變化值欄位
    merged['累積12~15變化值'] = merged['累積12~15_20250808'] - merged['累積12~15_20250124']

    # 排序及挑選欄位
    merged_sorted = merged.sort_values('變化值', ascending=False)
    result = merged_sorted[['證券代號',
                            '占集保庫存數比例%_20250124',
                            '占集保庫存數比例%_20250808',
                            '變化值',
                            '累積12~15變化值']]
    result.columns = ['證券代號', '2025/01/24比例', '2025/08/08比例', '變化值', '累積12~15變化值']
    
    # 輸出
    result.to_excel(output_file, index=False)
    print(f"結果已保存到 {output_file}")

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    file1 = os.path.join(base_dir, "20250124.csv")
    file2 = os.path.join(base_dir, "20250808.csv")
    output_file = os.path.join(base_dir, "持股分級15比例變化.xlsx")
    compare_holdings(file1, file2, output_file)