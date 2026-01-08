import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-最终调试版", layout="wide")

st.title("📋 拣货单自动提取")

# --- 调试：列出仓库里所有的文件 ---
with st.sidebar:
    st.header("📂 仓库文件检查")
    all_files = os.listdir(".")
    st.write("当前仓库内的文件：", all_files)

# --- 智能加载函数 ---
def load_excel_smart(possible_names):
    for name in possible_names:
        if os.path.exists(name):
            try:
                return pd.read_excel(name)
            except:
                continue
    return None

# 自动匹配可能的文件名（防止大小写或多空格问题）
df_prod = load_excel_smart(["product_info.xlsx", "PRODUCT_INFO.xlsx", "product_info.XLSX"])
df_label = load_excel_smart(["label_info.xlsx", "LABEL_INFO.xlsx", "label_info.XLSX"])

with st.sidebar:
    st.divider()
    if df_prod is not None: st.success("✅ 商品信息：已连接")
    else: st.error("❌ 缺失 product_info.xlsx")
    
    if df_label is not None: st.success("✅ 标签信息：已连接")
    else: st.error("❌ 缺失 label_info.xlsx")

# --- 主程序 ---
uploaded_pdf = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_pdf and df_prod is not None and df_label is not None:
    results = []
    # 预处理基础表：将匹配列全部转为字符串，去掉空格，防止匹配不上
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()
    
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if table:
                headers = table[0]
                try:
                    skc_text_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
                    sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                    
                    for row in table[1:]:
                        if not row[sku_idx] or "合计" in str(row): continue
                        
                        sku = str(row[sku_idx]).strip().replace('\n', '')
                        qty = str(row[qty_idx]).strip()
                        
                        skc_id = ""
                        info_cell = str(row[skc_text_idx])
                        skc_match = re.search(r"SKC:\s*(\d+)", info_cell)
                        if skc_match: skc_id = skc_match.group(1)

                        # 匹配
                        prod_name = "-"
                        p_match = df_prod[df_prod['商品编码'] == sku]
                        if not p_match.empty: prod_name = p_match.iloc[0]['商品名称']

                        label_type = "-"
                        if skc_id:
                            l_match = df_label[df_label['SKC ID'] == skc_id]
                            if not l_match.empty: label_type = l_match.iloc[0]['回收标签']

                        results.append({
                            "发货仓库": current_wh,
                            "SKC ID": skc_id,
                            "回收标签类别": label_type,
                            "货品编码": sku,
                            "商品名称": prod_name,
                            "发货数量": qty
                        })
                except: continue

    if results:
        st.success("处理完成！")
        st.dataframe(pd.DataFrame(results), use_container_width=True)
