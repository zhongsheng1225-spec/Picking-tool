import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-SKC全能版", layout="wide")

st.title("📋 拣货单自动提取 (SKC 深度识别版)")

# --- 1. 基础资料加载 (Excel) ---
def load_excel_smart(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

df_prod = load_excel_smart("product_info.xlsx")
df_label = load_excel_smart("label_info.xlsx")

# --- 2. 处理 PDF ---
uploaded_pdf = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_pdf and df_prod is not None and df_label is not None:
    results = []
    
    # 预处理基础表
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()
    
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            # 提取页面所有文字对象，包含坐标信息
            words = page.extract_words()
            full_text = page.extract_text() or ""
            
            # A. 提取仓库
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", full_text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            # B. 深度扫描 SKC (支持 换行/空格/中英文冒号)
            # 逻辑：先找关键词 "SKC"，再找它附近的数字
            found_skcs = []
            # 匹配模式：SKC 后面跟着冒号或空格，再跟着 5-15 位数字
            skc_pattern = re.compile(r"SKC[:：\s]+(\d{5,15})")
            for match in skc_pattern.finditer(full_text):
                found_skcs.append(match.group(1))
            
            # C. 提取表格
            table_obj = page.find_table()
            if table_obj:
                table_data = table_obj.extract()
                headers = table_data[0]
                
                try:
                    sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                    
                    row_count = 0
                    for row in table_data[1:]:
                        if not row[sku_idx] or "合计" in str(row):
                            continue
                        
                        sku = str(row[sku_idx]).strip().replace('\n', '')
                        qty = str(row[qty_idx]).strip()
                        
                        # 尝试精准分配 SKC
                        # 逻辑：拣货单中 SKC 出现的个数通常与商品行数一致
                        skc_id = found_skcs[row_count] if row_count < len(found_skcs) else ""
                        
                        # 3. 关联 Excel 数据
                        prod_name = "-"
                        p_match = df_prod[df_prod['商品编码'] == sku]
                        if not p_match.empty:
                            prod_name = p_match.iloc[0]['商品名称']

                        label_type = "-"
                        if skc_id:
                            l_match = df_label[df_label['SKC ID'] == skc_id]
                            if not l_match.empty:
                                label_type = l_match.iloc[0]['回收标签']

                        results.append({
                            "发货仓库": current_wh,
                            "SKC ID": skc_id,
                            "回收标签类别": label_type,
