import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-完整版", layout="wide")
st.title("📋 拣货单自动提取（按 SKU ID 唯一匹配）")

# =========================
# 1. 基础资料加载
# =========================
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except Exception as e:
            st.error(f"读取 {name} 失败：{e}")
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
    st.info(
        "💡 匹配逻辑：\n"
        "1. PDF 提取 SKU ID + SKU货号\n"
        "2. product_info.xlsx 用 SKU ID 匹配店铺和回收标签\n"
        "3. name_map.xlsx 用 SKU货号匹配商品名称"
    )

# =========================
# 2. 工具函数
# =========================
def get_match_col(df, keywords):
    """
    在表头里找包含关键词的列名
    """
    for col in df.columns:
        col_str = str(col).strip().lower()
        if any(str(key).strip().lower() in col_str for key in keywords):
            return col
    return None

def clean_numeric_id(val):
    """
    清洗 SKU ID 之类的字段：
    - 去空格
    - 去掉 .0
    - 只保留数字
    """
    if pd.isna(val):
        return ""
    val = str(val).strip()
    # 提取连续数字
    match = re.search(r"\d+", val)
    return match.group(0) if match else ""

def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

# =========================
# 3. 上传 PDF
# =========================
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_info is not None and df_name is not None:
    results = []

    # =========================
    # 4. 处理 name_map.xlsx
    #    SKU货号 -> 商品名称
    # =========================
    name_key = get_match_col(df_name, ['sku货号', '货号', '编码', 'sku'])
    if name_key is None:
        st.error("❌ name_map.xlsx 未找到 SKU货号 / 编码 列")
        st.stop()

    # 尝试找“商品名称”列；找不到就默认取除 key 外的第一列
    name_value_col = get_match_col(df_name, ['商品名称', '名称', '品名'])
    if name_value_col is None:
        other_cols = [c for c in df_name.columns if c != name_key]
        if not other_cols:
            st.error("❌ name_map.xlsx 未找到商品名称列")
            st.stop()
        name_value_col = other_cols[0]

    df_name[name_key] = df_name[name_key].apply(clean_text)
    df_name[name_value_col] = df_name[name_value_col].apply(clean_text)

    name_dict = (
        df_name
        .drop_duplicates(subset=[name_key])
        .set_index(name_key)[name_value_col]
        .to_dict()
    )

    # =========================
    # 5. 处理 product_info.xlsx
    #    SKU ID -> 店铺名称 / 回收标签
    # =========================
    # 这里优先直接写死列名逻辑，避免“智能匹配”误伤
    info_key = get_match_col(df_info, ['sku id', 'skuid'])
    if info_key is None:
        st.error("❌ product_info.xlsx 未找到 SKU ID 列")
        st.stop()

    shop_col = get_match_col(df_info, ['店铺名称', '店铺'])
    label_col = get_match_col(df_info, ['回收标签'])
    sku_code_col = get_match_col(df_info, ['sku货号', '货号'])

    if shop_col is None:
        st.error("❌ product_info.xlsx 未找到 店铺名称 列")
        st.stop()

    if label_col is None:
        st.error("❌ product_info.xlsx 未找到 回收标签 列")
        st.stop()

    df_info[info_key] = df_info[info_key].apply(clean_numeric_id)

    if sku_code_col is not None:
        df_info[sku_code_col] = df_info[sku_code_col].apply(clean_text)

    # 去掉空 SKU ID
    df_info = df_info[df_info[info_key] != ""].copy()

    # 如果 SKU ID 真的是唯一的，这里就不会有问题
    # 就算有重复，也保留第一条；后面可以用调试列排查
    info_dict = (
        df_info
        .drop_duplicates(subset=[info_key])
        .set_index(info_key)
        .to_dict('index')
    )

    # =========================
    # 6. 解析 PDF（按文本行）
    # =========================
    with pdfplumber.open(uploaded_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            lines = [line.strip() for line in text.split('\n') if line.strip()]

            # 提取仓库
            wh_match = re.search(r"收货仓[:：]\s*([^\n]+)", text)
            current_wh = wh_match.group(1).strip() if wh_match else "未知"

            active_skc = ""

            for line in lines:
                # 1) 抓 SKC
                skc_match = re.match(r"SKC[:：]\s*(\d+)", line)
                if skc_match:
                    active_skc = skc_match.group(1)
                    continue

                # 2) 跳过无效行
                if any(key in line for key in [
                    "备货母单号", "备货单号", "创建时间", "要求发货时间",
                    "收货仓", "打印时间", "序号", "商品信息",
                    "SKU ID", "SKU货号", "实际发货数", "拣货数",
                    "合计", "【VMI】", "数量：", "SKC货号："
                ]):
                    continue

                # 3) 解析明细行
                # 格式大致是：
                # 属性集   SKU ID   SKU货号   数量
                # 例如：
                # 2.0m*25cm 3619002957 Y160224081 10
                # One Chair 7413329453 Y060325014 0
                # 250*26cm 33010749823 202412022526 5
                detail_match = re.match(r"(.+?)\s+(\d{7,})\s+([A-Za-z0-9]+)\s+(\d+)$", line)
                if detail_match:
                    attr = clean_text(detail_match.group(1))
                    sku_id = clean_numeric_id(detail_match.group(2))
                    sku_code = clean_text(detail_match.group(3))
                    qty = clean_text(detail_match.group(4))

                    # 商品名称：仍然按 SKU货号 去匹配
                    res_prod_name = name_dict.get(sku_code, "-")

                    # 店铺名称 / 回收标签：必须按 SKU ID 去匹配
                    matched_row = info_dict.get(sku_id, None)

                    if matched_row is not None:
                        res_shop_name = clean_text(matched_row.get(shop_col, "-")) or "-"
                        res_label = clean_text(matched_row.get(label_col, "-")) or "-"
                        matched_sku_id = sku_id
                        match_status = "命中"
                    else:
                        res_shop_name = "-"
                        res_label = "-"
                        matched_sku_id = "未找到"
                        match_status = "未命中"

                    results.append({
                        "页码": page_num,
                        "发货仓库": current_wh,
                        "SKC ID": active_skc,
                        "属性集": attr,
                        "SKU ID": sku_id,
                        "货品编码": sku_code,
                        "商品名称": res_prod_name,
                        "店铺名称": res_shop_name,
                        "回收标签类别": res_label,
                        "发货数量": qty,
                        "匹配状态": match_status,
                        "匹配SKU ID": matched_sku_id
                    })

    # =========================
    # 7. 输出结果
    # =========================
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
            missing_info = len(df_res[df_res['匹配状态'] == '未命中'])
            st.metric("基础信息未匹配", missing_info, delta_color="inverse")

        if missing_name > 0 or missing_info > 0:
            st.warning("🚨 提示：部分货品未能在 Excel 中找到，请检查表中数据是否完整。")

        st.subheader("📄 提取结果")
        st.dataframe(df_res, use_container_width=True)

        # 只看未命中的
        with st.expander("🔎 查看未命中的记录"):
            df_unmatched = df_res[df_res["匹配状态"] == "未命中"]
            if len(df_unmatched) > 0:
                st.dataframe(df_unmatched, use_container_width=True)
            else:
                st.success("✅ 没有未命中的记录")

        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name="结果")

            # 单独导出未命中 sheet
            df_unmatched = df_res[df_res["匹配状态"] == "未命中"]
            if len(df_unmatched) > 0:
                df_unmatched.to_excel(writer, index=False, sheet_name="未命中记录")

        st.download_button(
            "📥 下载校验后的 Excel",
            output.getvalue(),
            file_name="拣货单结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ 未提取到任何有效数据，请检查 PDF 格式。")
