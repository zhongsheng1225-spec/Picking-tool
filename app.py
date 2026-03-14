import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-终极兼容版", layout="wide")

st.title("📋 拣货单自动提取 (终极修复版)")

# --- 1. 强力加载函数 ---
def load_data(name):
    full_path = name + ".xlsx"
    if os.path.exists(full_path):
        try:
            # 加上 engine='openpyxl' 强制读取，并跳过可能的空行
            df = pd.read_excel(full_path, engine='openpyxl').dropna(how='all', axis=0)
            # 统一清理列名中的空格和换行
            df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
            return df
        except Exception as e:
            st.error(f"读取 {full_path} 出错: {e}")
            return None
    return None

df_info = load_data("product_info")
df_name = load_data("name_map")

# --- 2. 侧边栏监控 ---
with st.sidebar:
    st.header("⚙️ 资料状态")
    if df_info is not None: st.success("✅ 基础信息已载入")
    else: st.error("❌ 未找到 product_info.xlsx")
    
    if df_name is not None: st.success("✅ 名称对照已载入")
    else: st.error("❌ 未找到 name_map.xlsx")

# --- 3. 核心匹配逻辑 ---
if df_info is not None and df_name is not None:
    # 智能寻找匹配列的索引
    def get_col_idx(df, keywords):
        for i, col in enumerate(df.columns):
            if any(k in str(col) for k in keywords):
                return i
        return 0

    # 预处理对照表 (取商品名称)
    n_idx = get_col_idx(df_name, ['编码', '货号', 'SKU'])
    df_name.iloc[:, n_idx] = df_name.iloc[:, n_idx].astype(str).str.strip()
    name_dict = df_name.drop_duplicates(df_name.columns[n_idx]).set_index(df_name.columns[n_idx]).iloc[:, 0].to_dict()

    # 预处理基础表 (取店铺和标签)
    i_idx = get_col_idx(df_info, ['SKU货号', '识别码', '编码'])
    df_info.iloc[:, i_idx] = df_info.iloc[:, i_idx].astype(str).str.strip()
    
    # 找“店铺名称”和“回收标签”的索引，找不到就给个负数
    shop_idx = get_col_idx(df_info, ['店铺名称'])
    label_idx = get_col_idx(df_info, ['回收标签'])
    
    # 转为字典提速
    info_dict = df_info.set_index(df_info.columns[i_idx]).to_dict('index')

    uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

    if uploaded_file:
        results = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
                current_wh = wh_match.group(1) if wh_match else "未知"
                
                table = page.extract_table()
                if not table or len(table) < 2: continue
                
                headers = [str(h).strip().replace('\n', '') for h in table[0] if h]
                try:
                    sku_idx = next(i for i, h in enumerate(headers) if '货号' in h or '编码' in h)
                    info_idx = next(i for i, h in enumerate(headers) if '商品信息' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if '实际' in h or '数量' in h or '发货数' in h)
                except: continue

                active_skc = "" 
                for row in table[1:]:
                    if not row[sku_idx] or "合计" in str(row): continue
                    
                    skc_match = re.search(r"SKC[:：\s]+(\d+)", str(row[info_idx]))
                    if skc_match: active_skc = skc_match.group(1)
                    
                    sku = str(row[sku_idx]).strip().replace('\n', '')
                    
                    # 执行双表匹配
                    res_prod_name = str(name_dict.get(sku, "-"))
                    res_shop_name, res_label = "-", "-"
                    
                    if sku in info_dict:
                        target_row = info_dict[sku]
                        # 用索引取值，避免列名带空格导致的 KeyError
                        res_shop_name = target_row.get(df_info.columns[shop_idx], "-")
                        res_label = target_row.get(df_info.columns[label_idx], "-")

                    results.append({
                        "发货仓库": current_wh,
                        "店铺名称": res_shop_name,
                        "SKC ID": active_skc,
                        "回收标签类别": res_label,
                        "货品编码": sku,
                        "商品名称": res_prod_name,
                        "发货数量": row[qty_idx]
                    })

        if results:
            df_res = pd.DataFrame(results)
            st.subheader("🔍 自动体检结果")
            c1, c2, c3 = st.columns(3)
            c1.metric("总行数", len(df_res))
            c2.metric("名称未匹配", len(df_res[df_res['商品名称'] == '-']))
            c3.metric("店铺未匹配", len(df_res[df_res['店铺名称'] == '-']))
            
            st.dataframe(df_res, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False)
            st.download_button("📥 下载 Excel 结果", output.getvalue(), "提取结果.xlsx")
