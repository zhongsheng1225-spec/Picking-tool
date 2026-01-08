import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

# ... (加载 Excel 的函数保持不变) ...

def process_pdf_precision(pdf_file, df_prod, df_label):
    results = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            # 提取仓库
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            # 动态寻找列索引
            try:
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
                sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
            except: continue

            active_skc = "" # 记录当前正在处理的 SKC 组

            for row in table[1:]:
                # 1. 检查当前行的“商品信息”列是否包含新的 SKC
                cell_text = str(row[info_idx])
                skc_match = re.search(r"SKC[:：]\s*(\d+)", cell_text)
                
                if skc_match:
                    active_skc = skc_match.group(1) # 发现新组，更新 active_skc
                
                # 2. 如果当前行没有 SKC，但有 SKU（说明是同组下的另一个规格）
                sku = str(row[sku_idx]).strip().replace('\n', '')
                if sku and sku != "None" and "合计" not in str(row):
                    qty = str(row[qty_idx]).strip()
                    
                    # 匹配商品名称
                    prod_name = "-"
                    if df_prod is not None:
                        p_m = df_prod[df_prod['商品编码'].astype(str) == sku]
                        if not p_m.empty: prod_name = p_m.iloc[0]['商品名称']

                    # 匹配回收标签 (使用当前生效的 active_skc)
                    label_type = "-"
                    if active_skc and df_label is not None:
                        l_m = df_label[df_label['SKC ID'].astype(str) == active_skc]
                        if not l_m.empty: label_type = l_m.iloc[0]['回收标签']

                    results.append({
                        "发货仓库": current_wh,
                        "SKC ID": active_skc,
                        "回收标签类别": label_type,
                        "货品编码": sku,
                        "商品名称": prod_name,
                        "发货数量": qty
                    })
    return results
