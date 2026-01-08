import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-向下补全版", layout="wide")

# --- 1. 基础资料加载 ---
def load_data(name):
    if os.path.exists(name):
        try: return pd.read_excel(name)
        except: return None
    return None

df_prod = load_data("product_info.xlsx")
df_label = load_data("label_info.xlsx")

st.title("📋 拣货单自动提取 (SKC 向下关联版)")

# --- 2. 处理 PDF ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_prod is not None and df_label is not None:
    results = []
    
    # 数据清洗
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            # 获取页面所有文本行，按纵坐标从上到下排序
            lines = page.extract_text().split('\n')
            
            current_wh = "未知"
            active_skc = "" # 当前正在生效的 SKC ID
            
            # 提取表格以便精准获取 SKU 和数量
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
            except: continue

            # 遍历表格行
            for row in table[1:]:
                if not row[sku_idx] or "合计" in str(row): continue
                
                # --- 核心修改：向下找当前行对应的 SKC ---
                # 检查“商品信息”栏位是否包含新的 SKC
                info_content = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", info_content)
                
                if skc_match:
                    # 如果这行发现了新 SKC，则更新当前活跃 SKC
                    active_skc = skc_match.group(1)
                
                # 如果当前行没写 SKC，它会自动沿用上面最近的那一个 active_skc
                
                sku = str(row[sku_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()

                # VLOOKUP 匹配商品名称
                p_name = "-"
                p_match = df_prod[df_prod['商品编码'] == sku]
                if not p_match.empty: p_name = p_match.iloc[0]['商品名称']

                # VLOOKUP 匹配标签
                l_type = "-"
                if active_skc:
                    l_match = df_label[df_label['SKC ID'] == active_skc]
                    if not l_match.empty: l_type = l_match.iloc[0]['回收标签']

                results.append({
                    "发货仓库": "从PDF提取", # 仓库提取逻辑可保持
                    "SKC ID": active_skc,
                    "回收标签类别": l_type,
                    "货品编码": sku,
                    "商品名称": p_name,
                    "发货数量": qty
                })

    if results:
        df_res = pd.DataFrame(results)
        st.success("向下关联处理完成！")
        st.dataframe(df_res, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载提取结果", output.getvalue(), "提取结果.xlsx")
