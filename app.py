import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="调试版-拣货工具", layout="wide")

# --- 智能加载函数 ---
def load_local_data(file_name):
    if not os.path.exists(file_name):
        return None
    try:
        # 尝试多种读取方式
        try: return pd.read_excel(file_name)
        except: pass
        try: return pd.read_csv(file_name, encoding='utf-8-sig')
        except: return pd.read_csv(file_name, encoding='gbk')
    except Exception as e:
        st.error(f"读取 {file_name} 失败: {e}")
        return None

df_prod = load_local_data("product_info.csv")
df_label = load_local_data("label_info.csv")

st.title("📋 拣货单自动提取 (数据匹配预览)")

# 侧边栏实时诊断
with st.sidebar:
    st.header("🔍 数据诊断")
    if df_prod is not None:
        st.success(f"商品表: 已加载 {len(df_prod)} 行")
        st.write("列名:", list(df_prod.columns)) # 显示表头，方便对齐
    else:
        st.error("未找到 product_info.csv")
        
    if df_label is not None:
        st.success(f"标签表: 已加载 {len(df_label)} 行")
        st.write("列名:", list(df_label.columns))
    else:
        st.error("未找到 label_info.csv")

uploaded_pdf = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_pdf is not None:
    results = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if table:
                headers = table[0]
                try:
                    # 动态匹配 PDF 的列
                    skc_text_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
                    sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                    
                    for row in table[1:]:
                        if not row[sku_idx] or "合计" in str(row): continue
                        
                        sku = str(row[sku_idx]).strip().replace('\n', '')
                        qty = str(row[qty_idx]).strip()
                        
                        # 提取 SKC
                        skc_id = ""
                        skc_match = re.search(r"SKC:\s*(\d+)", str(row[skc_text_idx]))
                        if skc_match: skc_id = skc_match.group(1)

                        # 关联名称
                        p_name = "-"
                        if df_prod is not None:
                            # 强制转为字符串匹配
                            m = df_prod[df_prod['商品编码'].astype(str).str.strip() == sku]
                            if not m.empty: p_name = m.iloc[0]['商品名称']

                        # 关联标签
                        l_type = "-"
                        if df_label is not None and skc_id:
                            m = df_label[df_label['SKC ID'].astype(str).str.strip() == skc_id]
                            if not m.empty: l_type = m.iloc[0]['回收标签']

                        results.append({
                            "发货仓库": current_wh,
                            "SKC ID": skc_id,
                            "回收标签类别": l_type,
                            "货品编码": sku,
                            "商品名称": p_name,
                            "发货数量": qty
                        })
                except Exception: continue

    if results:
        df_res = pd.DataFrame(results)
        st.dataframe(df_res, use_container_width=True)
        
        # 导出结果
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载 Excel", output.getvalue(), "result.xlsx")
