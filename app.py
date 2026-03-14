import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强-路径修复版", layout="wide")

st.title("📋 拣货单自动提取 (路径修复版)")

# --- 1. 深度搜索加载函数 ---
def smart_load(base_name):
    # 在当前目录及子目录搜索匹配的文件
    possible_names = [f"{base_name}.xlsx", f"{base_name}.XLSX"]
    
    for root, dirs, files in os.walk("."):
        for file in files:
            if file in possible_names:
                file_path = os.path.join(root, file)
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    # 清理列名中的空格和换行
                    df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
                    return df
                except Exception as e:
                    st.sidebar.error(f"读取 {file} 失败: {e}")
    return None

# 执行搜索加载
df_info = smart_load("product_info")
df_name = smart_load("name_map")

# --- 2. 侧边栏实时监测 ---
with st.sidebar:
    st.header("⚙️ 仓库文件监测")
    if df_info is not None:
        st.success("✅ 基础信息表已连接")
    else:
        st.error("❌ 未找到 product_info.xlsx")
        st.info("请检查文件名是否准确，且已上传至 GitHub")

    if df_name is not None:
        st.success("✅ 名称对照表已连接")
    else:
        st.error("❌ 未找到 name_map.xlsx")

# --- 3. 匹配提取逻辑 (仅在文件都找到时运行) ---
if df_info is not None and df_name is not None:
    # (此部分包含之前优化的模糊匹配逻辑，确保 3.12 的表能对上)
    def find_col(df, keywords):
        for col in df.columns:
            if any(k in str(col) for k in keywords):
                return col
        return None

    n_col = find_col(df_name, ['编码', '货号', 'SKU'])
    v_col = find_col(df_name, ['名称', '标题'])
    name_dict = df_name.set_index(n_col)[v_col].to_dict() if n_col and v_col else {}

    i_col = find_col(df_info, ['SKU货号', '识别码', '编码'])
    info_dict = df_info.set_index(i_col).to_dict('index') if i_col else {}

    uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

    if uploaded_file:
        # ... (PDF 处理逻辑)
        st.info("PDF 处理中...")
        # (此处省略部分重复的 PDF 解析逻辑，请直接沿用上一版的解析部分)
