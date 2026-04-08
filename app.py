import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-完整版", layout="wide")
st.title("📋 拣货单自动提取 (双表关联+自动体检)")

# --- 1. 基础资料加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

df_info = load_data("product_info.xlsx")
df_name = load_data("name_map.xlsx")

with st.sidebar:
    st.header("⚙️ 资料体检状态")
    if df_info is not None: st.success("✅ product_info.xlsx 已就绪")
    else: st.error("❌ 缺失 product_info.xlsx")
    if df_name is not None: st.success("✅ name_map.xlsx 已就绪")
    else: st.error("❌ 缺失 name_map.xlsx")
    st.divider()
    st.info("💡 校验逻辑：\n1. [名称对照表] 找具体商品名\n2. [基础信息表] 找店铺和回收标签")

# --- 2. 处理 PDF 主逻辑 ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_info is not None and df_name is not None:
    results = []

    # 智能匹配列
    def get_match_col(df, keywords):
        for col in df.columns:
            if any(key in str(col).lower() for key in keywords):
                return col
        return df.columns[0]

    # ====================== 名称对照表 ======================
    name_key = get_match_col(df_name, ['编码', 'sku', '货号'])
    df_name[name_key] = df_name[name_key].astype(str).str.strip().str.upper()
    name_dict = df_name.set_index(name_key).iloc[:, 0].to_dict()

    # ====================== 产品信息表（核心：绝对精准匹配） ======================
    info_key = get_match_col(df_info, ['skuid', 'sku货号', '商品识别码', '编码'])
    df_info[info_key] = df_info[info_key].astype(str).str.strip().str.upper()
    # 【重要】不做任何去重！不做任何排序！完全保留你表格原始数据
    info_dict = df_info.set_index(info_key).to_dict('index')

    # 开始解析PDF
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                sku_idx = next(i for i, h in enumerate(headers) if h and any(k in str(h).lower() for k in ['货号', '编码', 'sku']))
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in str(h))
                qty_idx = next(i for i, h in enumerate(headers) if h and '发货数' in str(h))
            except:
                continue

            active_skc = ""
            for row in table[1:]:
                if not row[sku_idx] or "合计" in str(row):
                    continue
                
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                if skc_match:
                    active_skc = skc_match.group(1)
                
                # 清理SKU，保持唯一匹配
                sku = str(row[sku_idx]).strip().replace('\n', '').replace(' ', '').upper()
                qty = str(row[qty_idx]).strip()

                # 【绝对精准匹配】表格填什么，就输出什么
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

    # --- 3. 结果展示 ---
    if results:
        df_res = pd.DataFrame(results)
        st.subheader("🔍 自动体检看板")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("处理总行数", len(df_res))
        with col2:
            shops = df_res[df_res['店铺名称'] != '-']['店铺名称'].nunique()
            st.metric("涉及店铺数", shops)
        with col3:
            missing_name = len(df_res[df_res['商品名称'] == '-'])
            st.metric("名称未匹配", missing_name, delta_color="inverse")
        with col4:
            missing_info = len(df_res[df_res['店铺名称'] == '-'])
            st.metric("基础信息未匹配", missing_info, delta_color="inverse")

        if missing_name > 0 or missing_info > 0:
            st.warning("🚨 提示：部分货品未能在 Excel 中找到。请检查 `name_map.xlsx` 和 `product_info.xlsx` 是否已包含最新货号。")

        st.dataframe(df_res, use_container_width=True)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载校验后的 Excel", output.getvalue(), "拣货单结果.xlsx")
