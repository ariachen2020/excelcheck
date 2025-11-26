#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
創建示例 Excel 檔案用於測試
"""

import pandas as pd
import os


def create_sample_files():
    """創建示例 Excel 檔案"""

    # 創建正確的檔案
    print("正在創建 correct_sample.xlsx...")
    data_correct = {
        '姓名': ['張三', '李四', '王五'],
        '年齡': [25.0, 30.0, 28.0],
        '薪資': [50000.50, 60000.75, 55000.25],
        '部門': ['銷售', '技術', '人資']
    }
    df_correct = pd.DataFrame(data_correct)
    df_correct.to_excel('correct_sample.xlsx', index=False)
    print("✓ 已創建 correct_sample.xlsx")

    # 創建待檢查的檔案（有格式差異）
    print("\n正在創建 test_sample.xlsx...")
    data_test = {
        '姓名': ['張三', '李四', '王五'],
        '年齡': [25, 30, 28],  # 整數而非浮點數
        '部門': ['銷售', '技術', '人資'],
        '薪資': [50000.50, 60000.75, 55000.25]  # 欄位順序不同
    }
    df_test = pd.DataFrame(data_test)
    df_test.to_excel('test_sample.xlsx', index=False)
    print("✓ 已創建 test_sample.xlsx")

    print("\n" + "=" * 60)
    print("示例檔案創建完成！")
    print("=" * 60)
    print("\n你現在可以：")
    print("1. 運行 Streamlit 應用:")
    print("   streamlit run app.py")
    print("\n2. 上傳這兩個示例檔案進行測試:")
    print("   - correct_sample.xlsx（正確的檔案）")
    print("   - test_sample.xlsx（待檢查的檔案）")
    print("\n應該會發現以下差異：")
    print("- 年齡欄位的資料類型不同（float64 vs int64）")
    print("- 欄位順序不同（部門和薪資位置交換）")
    print("=" * 60)


if __name__ == "__main__":
    create_sample_files()
