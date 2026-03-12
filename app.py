import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-强力纠错版", layout="wide")

st.title("📋 拣货单自动提取 (双表强力匹配版)")

# --- 1. 基础资料加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            # 尝试读取 Excel，如果失败尝试读取 CSV
            if name.endswith('.csv'):
                return pd.read_csv(name)
            return pd.read_excel(name)
        except:
            return None
    return None

# 尝试加载两张表
df_info = load_data("product_info.xlsx")
df_name = load_data("name_map.xlsx")

with st.sidebar:
    st.header("⚙️ 资料状态检查")
    if df_info is not None: st.success("✅ 基础信息已载入")
    else: st.error("❌ 缺失 product_info.xlsx")
    
    if df_name is not None: st.success("✅ 名称对照已载入")
    else: st.error("❌ 缺失 name_map.xlsx")

# --- 2. 处理 PDF 主逻辑 ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_info is not None and df_name is not None:
    results = []
    
    # --- 智能寻找匹配列 ---
    def get_match_col(df, keywords):
        for col in df.columns:
            if any(key in str(col) for key in keywords):
                return col
        return df.columns[0] # 找不到就默认第一列

    # 1. 处理名称对照表 (匹配编码)
    name_key = get_match_col(df_name, ['编码', 'SKU', '货号'])
    df_name[name_key] = df_name[name_key].astype(str).str.strip()
    name_dict = df_name.drop_duplicates(name_key).set_index(name_key).iloc[:, 0].to_dict()

    # 2. 处理基础信息表 (匹配货号)
    info_key = get_match_col(df_info, ['SKU货号', '商品识别码', '编码'])
    df_info[info_key] = df_info[info_key].astype(str).str.strip()
    info_dict = df_info.drop_duplicates(info_key).set_index(info_key).to_dict('index')

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                # 寻找 PDF 表格列
                sku_idx = next(i for i, h in enumerate(headers) if h and ('货号' in str(h) or '编码' in str(h)))
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in str(h))
                qty_idx = next(i for i, h in enumerate(headers) if h and '发货数' in str(h))
            except: continue

            active_skc = "" 
            for row in table[1:]:
                if not row[sku_idx] or "合计" in str(row): continue
                
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                if skc_match: active_skc = skc_match.group(1)
                
                sku = str(row[sku_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()

                # 双表关联
                res_prod_name = name_dict.get(sku, "-")
                res_shop_name = "-"
                res_label = "-"
                if sku in info_dict:
                    res_shop_name = info_dict[sku].get('店铺名称', '-')
                    res_label = info_dict[sku].get('回收标签', '-')

                results.append({
                    "发货仓库": current_wh,
                    "店铺名称": res_shop_name,
                    "SKC ID": active_skc,
                    "回收标签类别": res_label,
                    "货品编码": sku,
                    "商品名称": res_prod_name,
                    "发货数量": qty
                })

    if results:
        df_res = pd.DataFrame(results)
        st.subheader("🔍 结果预览")
        st.dataframe(df_res, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载 Excel 结果", output.getvalue(), "拣货结果.xlsx")
