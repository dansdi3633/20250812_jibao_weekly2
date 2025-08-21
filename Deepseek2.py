import pandas as pd
import openpyxl
import os
from tkinter import Tk, filedialog

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
    suffix1 = '_' + os.path.splitext(os.path.basename(file1))[0]
    suffix2 = '_' + os.path.splitext(os.path.basename(file2))[0]
    merged = pd.merge(
        df1_level15,
        df2_level15,
        on='證券代號',
        how='outer',
        suffixes=(suffix1, suffix2)
    ).fillna(0)
    merged['變化值'] = merged[f'占集保庫存數比例%{suffix2}'] - merged[f'占集保庫存數比例%{suffix1}']
    # 累計12~15級
    df1_lv12_15 = df1[df1['持股分級'].between(12, 15)].groupby('證券代號')['占集保庫存數比例%'].sum().reset_index()
    df1_lv12_15.columns = ['證券代號', f'累積12~15{suffix1}']
    df2_lv12_15 = df2[df2['持股分級'].between(12, 15)].groupby('證券代號')['占集保庫存數比例%'].sum().reset_index()
    df2_lv12_15.columns = ['證券代號', f'累積12~15{suffix2}']
    # 合併到主表
    merged = pd.merge(merged, df1_lv12_15, on='證券代號', how='left')
    merged = pd.merge(merged, df2_lv12_15, on='證券代號', how='left')
    merged = merged.fillna(0)
    # 新增累積變化值欄位
    merged['累積12~15變化值'] = merged[f'累積12~15{suffix2}'] - merged[f'累積12~15{suffix1}']
    # 排序及挑選欄位
    merged_sorted = merged.sort_values('變化值', ascending=False)
    result = merged_sorted[['證券代號',
                            f'占集保庫存數比例%{suffix1}',
                            f'占集保庫存數比例%{suffix2}',
                            '變化值',
                            '累積12~15變化值']]
    result.columns = ['證券代號', f'{os.path.splitext(os.path.basename(file1))[0]}比例',
                      f'{os.path.splitext(os.path.basename(file2))[0]}比例',
                      '變化值', '累積12~15變化值']
    # 輸出
    result.to_excel(output_file, index=False)
    print(f"結果已保存到 {output_file}")

if __name__ == "__main__":
    # 防止 tkinter 彈出主視窗
    root = Tk()
    root.withdraw()
    print("請選擇第一個基準CSV檔案")
    file1 = filedialog.askopenfilename(title="選擇第一個基準CSV檔案", filetypes=[("CSV files", "*.csv")])
    print("請選擇要比較的第二個CSV檔案")
    file2 = filedialog.askopenfilename(title="選擇第二個要比較的CSV檔案", filetypes=[("CSV files", "*.csv")])
    if file1 and file2:
        # 結果自動以檔名日期標示
        base_path = os.path.dirname(file1)
        date1 = os.path.splitext(os.path.basename(file1))[0]
        date2 = os.path.splitext(os.path.basename(file2))[0]
        output_file = os.path.join(base_path, f"持股分級15比例變化_{date1}_to_{date2}.xlsx")
        compare_holdings(file1, file2, output_file)
    else:
        print("未選擇兩個檔案，比對作業取消。")