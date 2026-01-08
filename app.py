import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-全精准修复版", layout="wide")

# --- 1. 基础资料加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            # 兼容读取 xlsx
            return pd.read_excel(name)
        except:
            return None
    return None

df_prod = load_data("product_info.xlsx")
df_label = load_data("label_info.xlsx")

st.title("📋 拣货单自动提取 (仓库+SKC 修正版)")

# --- 2. 处理 PDF ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_prod is not None and df_label is not None:
    results = []
    
    # 预处理基础表（去空格，转字符串）
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            # --- 每一页都重新提取本页信息 ---
            text = page.extract_text() or ""
            
            # 1. 提取发货仓库（针对每一页单独搜索）
            # 查找“收货仓:”或“仓库:”后面的文字，直到空格或换行
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            # 2. 提取表格
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
            except:
                continue

            active_skc = "" # 本页内的 SKC 指针
            
            for row in table[1:]:
                # 过滤空行或合计行
                if not row[sku_idx] or "合计" in str(row):
                    continue
                
                # --- 向下关联 SKC 逻辑 ---
                # 在“商品信息”单元格中找 SKC
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                
                if skc_match:
                    active_skc = skc_match.group(1)
                
                # 如果这行没写 SKC，会沿用同页上方最近的一个 SKC
                
                sku = str(row[sku_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()

                # 匹配 Excel
                p_name = "-"
                p_match = df_prod[df_prod['商品编码'] == sku]
                if not p_match.empty:
                    p_name = p_match.iloc[0]['商品名称']

                l_type = "-"
                if active_skc:
                    l_match = df_label[df_label['SKC ID'] == active_skc]
                    if not l_match.empty:
                        l_type = l_match.iloc[0]['回收标签']

                results.append({
                    "发货仓库": current_wh, # 确保每一行都对应本页抓到的仓库
                    "SKC ID": active_skc,
                    "回收标签类别": l_type,
                    "货品编码": sku,
                    "商品名称": p_name,
                    "发货数量": qty
                })

    if results:
        df_res = pd.DataFrame(results)
        st.success("仓库与 SKC 匹配完成！")
        st.dataframe(df_res, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载 Excel 结果", output.getvalue(), "修正结果.xlsx")
