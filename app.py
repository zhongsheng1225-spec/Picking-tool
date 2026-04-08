import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-完整版", layout="wide")
st.title("📋 拣货单自动提取 (按 SKU ID 唯一匹配)")

# --- 1. 基础资料智能加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

df_info = load_data("product_info.xlsx")   # 基础信息表
df_name = load_data("name_map.xlsx")       # 名称对照表

with st.sidebar:
    st.header("⚙️ 资料体检状态")
    if df_info is not None:
        st.success("✅ product_info.xlsx 已就绪")
    else:
        st.error("❌ 缺失 product_info.xlsx")

    if df_name is not None:
        st.success("✅ name_map.xlsx 已就绪")
    else:
        st.error("❌ 缺失 name_map.xlsx")

    st.divider()
    st.info("💡 校验逻辑：\n1. PDF 提取 SKU ID\n2. product_info.xlsx 用 SKU ID 匹配店铺和回收标签\n3. name_map.xlsx 用 SKU货号匹配商品名称")

# --- 2. 工具函数 ---
def get_match_col(df, keywords):
    for col in df.columns:
        col_str = str(col).strip().lower()
        if any(str(key).strip().lower() in col_str for key in keywords):
            return col
    return None

uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_info is not None and df_name is not None:
    results = []

    # --- 3. 名称表：SKU货号 -> 商品名称 ---
    name_key = get_match_col(df_name, ['编码', 'sku', '货号'])
    if name_key is None:
        st.error("❌ name_map.xlsx 未找到可用的 SKU货号/编码列")
        st.stop()

    df_name[name_key] = df_name[name_key].astype(str).str.strip()
    name_value_col = [c for c in df_name.columns if c != name_key][0]
    name_dict = df_name.drop_duplicates(name_key).set_index(name_key)[name_value_col].to_dict()

    # --- 4. 基础信息表：SKU ID -> 店铺名称 / 回收标签 ---
    info_key = get_match_col(df_info, ['sku id', 'skuid'])
    if info_key is None:
        st.error("❌ product_info.xlsx 未找到 SKU ID 列")
        st.stop()

    shop_col = get_match_col(df_info, ['店铺名称', '店铺'])
    label_col = get_match_col(df_info, ['回收标签'])

    if shop_col is None or label_col is None:
        st.error("❌ product_info.xlsx 缺少 店铺名称 或 回收标签 列")
        st.stop()

    df_info[info_key] = df_info[info_key].astype(str).str.strip()
    info_dict = df_info.drop_duplicates(info_key).set_index(info_key).to_dict('index')

    # --- 5. PDF 解析 ---
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [line.strip() for line in text.split('\n') if line.strip()]

            # 提取仓库
            wh_match = re.search(r"收货仓[:：]\s*([^\n]+)", text)
            current_wh = wh_match.group(1).strip() if wh_match else "未知"

            active_skc = ""

            for line in lines:
                # 提取 SKC
                skc_match = re.match(r"SKC[:：]\s*(\d+)", line)
                if skc_match:
                    active_skc = skc_match.group(1)
                    continue

                # 跳过无效行
                if (
                    "备货母单号" in line or
                    "备货单号" in line or
                    "创建时间" in line or
                    "要求发货时间" in line or
                    "收货仓" in line or
                    "打印时间" in line or
                    "序号" in line or
                    "商品信息" in line or
                    "合计" in line or
                    "【VMI】" in line or
                    "数量：" in line
                ):
                    continue

                # -----------------------------
                # 明细行格式：
                # 属性集   SKU ID   SKU货号   实际发货数
                # 例如：
                # 2.0m*25cm 3619002957 Y160224081 10
                # One Chair 7413329453 Y060325014 0
                # Chair with Accessories 17674900294 202503222401 12
                # -----------------------------
                detail_match = re.match(r"(.+?)\s+(\d{7,})\s+([A-Za-z0-9]+)\s+(\d+)$", line)
                if detail_match:
                    attr = detail_match.group(1).strip()
                    sku_id = detail_match.group(2).strip()
                    sku_code = detail_match.group(3).strip()
                    qty = detail_match.group(4).strip()

                    # 名称仍然按 SKU货号 去查
                    res_prod_name = name_dict.get(sku_code, "-")

                    # 店铺 / 标签 改成按 SKU ID 查
                    res_shop_name = "-"
                    res_label = "-"
                    if sku_id in info_dict:
                        res_shop_name = info_dict[sku_id].get(shop_col, "-")
                        res_label = info_dict[sku_id].get(label_col, "-")

                    results.append({
                        "发货仓库": current_wh,
                        "店铺名称": res_shop_name,
                        "SKC ID": active_skc,
                        "属性集": attr,
                        "SKU ID": sku_id,
                        "回收标签类别": res_label,
                        "货品编码": sku_code,
                        "商品名称": res_prod_name,
                        "发货数量": qty
                    })

    # --- 6. 输出结果 ---
    if results:
        df_res = pd.DataFrame(results)

        st.subheader("🔍 自动体检看板")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("处理总行数", len(df_res))
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
            st.warning("🚨 提示：部分货品未能在 Excel 中找到。请检查 name_map.xlsx 和 product_info.xlsx 是否包含最新数据。")

        st.dataframe(df_res, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)

        st.download_button("📥 下载校验后的 Excel", output.getvalue(), "拣货单结果.xlsx")
    else:
        st.warning("⚠️ 未提取到任何有效数据，请检查 PDF 格式。")
