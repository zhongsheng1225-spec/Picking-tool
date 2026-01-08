import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-自动补全版", layout="wide")

st.title("📋 拣货单自动提取 (SKC 自动补全版)")

# --- 1. 基础资料加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

df_prod = load_data("product_info.xlsx")
df_label = load_data("label_info.xlsx")

# --- 2. 处理 PDF ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_prod is not None and df_label is not None:
    results = []
    
    # 数据清洗：确保基础表 ID 是字符串格式
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()

    with pdfplumber.open(uploaded_file) as pdf:
        # 初始化一个变量，用于记住上一个有效的 SKC ID
        last_valid_skc = ""
        
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # 提取仓库
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            # 深度抓取本页所有 SKC (关键词模式 + 纯数字模式)
            found_skcs = re.findall(r"SKC[:：\s]+(\d+)", text)
            if not found_skcs:
                found_skcs = re.findall(r"\b(\d{9,15})\b", text)

            table = page.extract_table()
            if table:
                headers = table[0]
                try:
                    sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                    
                    row_count = 0
                    for row in table[1:]:
                        if not row[sku_idx] or "合计" in str(row): continue
                        
                        sku = str(row[sku_idx]).strip().replace('\n', '')
                        qty = str(row[qty_idx]).strip()
                        
                        # --- 【核心修复：向下填充逻辑】 ---
                        # 如果当前行在列表中有对应的 SKC，就更新 last_valid_skc
                        if row_count < len(found_skcs):
                            last_valid_skc = found_skcs[row_count]
                        
                        # 如果列表用完了，它会自动沿用上一个 last_valid_skc (即实现补全)
                        skc_id = last_valid_skc
                        
                        # VLOOKUP 匹配商品名称
                        p_name = "-"
                        p_match = df_prod[df_prod['商品编码'] == sku]
                        if not p_match.empty: p_name = p_match.iloc[0]['商品名称']

                        # VLOOKUP 匹配标签
                        l_type = "-"
                        if skc_id:
                            l_match = df_label[df_label['SKC ID'] == skc_id]
                            if not l_match.empty: l_type = l_match.iloc[0]['回收标签']

                        results.append({
                            "发货仓库": current_wh,
                            "SKC ID": skc_id,
                            "回收标签类别": l_type,
                            "货品编码": sku,
                            "商品名称": p_name,
                            "发货数量": qty
                        })
                        row_count += 1
                except: continue

    if results:
        df_res = pd.DataFrame(results)
        st.success("处理完成！SKC 已根据排版自动向下补全。")
        st.dataframe(df_res, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载完整结果", output.getvalue(), "提取结果.xlsx")
