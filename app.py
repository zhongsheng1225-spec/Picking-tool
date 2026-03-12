import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-双表关联版", layout="wide")

st.title("📋 拣货单自动提取 (多表关联版)")

# --- 1. 基础资料加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

# 加载两张核心表
df_info = load_data("product_info.xlsx") # 基础信息表：取店铺、回收标签
df_name = load_data("name_map.xlsx")     # 对照表：取商品名称

with st.sidebar:
    st.header("⚙️ 基础资料状态")
    status_info = "✅ 基础信息已就绪" if df_info is not None else "❌ 缺失 product_info.xlsx"
    status_name = "✅ 名称对照已就绪" if df_name is not None else "❌ 缺失 name_map.xlsx"
    
    st.write(status_info)
    st.write(status_name)
    
    if df_info is not None and df_name is not None:
        st.divider()
        st.info("💡 逻辑：从对照表取[商品名称]，从基础表取[店铺]和[标签]")

# --- 2. 处理 PDF 主逻辑 ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_info is not None and df_name is not None:
    results = []
    
    # --- 数据预处理：统一格式并建立快速索引 ---
    # 1. 处理名称对照表 (name_map)
    df_name['商品编码'] = df_name['商品编码'].astype(str).str.strip()
    name_dict = df_name.drop_duplicates('商品编码').set_index('商品编码')['商品名称'].to_dict()

    # 2. 处理基础信息表 (product_info) - 优先使用 [商品识别码] 或 [SKU货号] 匹配，取决于您的表结构
    # 这里根据您上传的 CSV 预览，匹配列应为 '商品识别码' 或 'SKU货号'，我们统一按您要求的“商品编码”逻辑处理
    # 注意：如果基础表里叫“商品识别码”，请确保代码中的 key 一致
    search_col = '商品识别码' if '商品识别码' in df_info.columns else '商品编码'
    df_info[search_col] = df_info[search_col].astype(str).str.strip()
    info_dict = df_info.drop_duplicates(search_col).set_index(search_col).to_dict('index')

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # A. 提取仓库
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            # B. 提取表格
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
                qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
            except: continue

            active_skc = "" # SKC指针
            
            for row in table[1:]:
                if not row[sku_idx] or "合计" in str(row): continue
                
                # C. SKC 向下补全
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                if skc_match:
                    active_skc = skc_match.group(1)
                
                sku = str(row[sku_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()

                # D. 双表关联逻辑
                # 1. 从对照表取名称
                res_prod_name = name_dict.get(sku, "-")
                
                # 2. 从基础信息表取店铺和标签
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

    # --- 3. 结果展示 ---
    if results:
        df_res = pd.DataFrame(results)
        
        st.subheader("🔍 提取质量看板")
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("处理总数", len(df_res))
        with m2:
            fail_name = len(df_res[df_res['商品名称'] == '-'])
            st.metric("名称匹配失败", fail_name, delta_color="inverse")
        with m3:
            fail_info = len(df_res[df_res['店铺名称'] == '-'])
            st.metric("基础信息缺失", fail_info, delta_color="inverse")

        st.dataframe(df_res, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载关联结果", output.getvalue(), "拣货单增强结果.xlsx")
