import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-强力兼容版", layout="wide")

st.title("📋 拣货单提取 (报错自动修复版)")

# --- 1. 基础资料智能加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            # 自动处理可能存在的表头空行
            df = pd.read_excel(name)
            if df.columns[0].startswith('Unnamed'):
                df.columns = df.iloc[0]
                df = df[1:]
            return df
        except:
            return None
    return None

df_info = load_data("product_info.xlsx")
df_name = load_data("name_map.xlsx")

# --- 2. 核心匹配逻辑 (带容错) ---
if df_info is not None and df_name is not None:
    # 智能定位列名的函数
    def find_col(df, keywords):
        for col in df.columns:
            if any(k in str(col) for k in keywords):
                return col
        return df.columns[0]

    # 预处理对照表
    name_key = find_col(df_name, ['编码', '货号', 'SKU'])
    df_name[name_key] = df_name[name_key].astype(str).str.strip()
    name_dict = df_name.drop_duplicates(name_key).set_index(name_key).iloc[:, 0].to_dict()

    # 预处理基础信息表
    info_key = find_col(df_info, ['SKU货号', '识别码', '编码', '货号'])
    df_info[info_key] = df_info[info_key].astype(str).str.strip()
    # 确保提取店铺和标签时不会因为列名不存在而崩溃
    def get_val(row, keys, default="-"):
        for k in keys:
            if k in row: return row[k]
        return default

    info_dict = df_info.drop_duplicates(info_key).set_index(info_key).to_dict('index')

    # ... (PDF处理部分逻辑保持不变)
