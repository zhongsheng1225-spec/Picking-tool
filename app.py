import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单自动提取", layout="wide")
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
    st.header("⚙️ 资料状态")
    if df_info is not None:
        st.success("✅ product_info.xlsx 已就绪")
    else:
        st.error("❌ 缺失 product_info.xlsx")

    if df_name is not None:
        st.success("✅ name_map.xlsx 已就绪")
    else:
        st.error("❌ 缺失 name_map.xlsx")


# =========================
# 2. 工具函数
# =========================
def get_match_col(df, keywords):
    for col in df.columns:
        col_str = str(col).strip().lower()
        if any(str(key).strip().lower() in col_str for key in keywords):
            return col
    return None

def clean_numeric_id(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    match = re.search(r"\d+", val)
    return match.group(0) if match else ""

def clean_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def parse_detail_line(line):
    """
    解析一条完整明细：
    属性集 + SKU ID + SKU货号 + 数量

    示例：
    2.0m*25cm 3619002957 Y160224081 10
    Chair with Accessories 17674900294 202503222401 12
    BLACK-125cm 2-layer 38866424354 Y010425027 13
    """
    line = re.sub(r"\s+", " ", line).strip()

    # 最后三段固定为：SKU ID、SKU货号、数量
    m = re.match(r"(.+?)\s+(\d{7,})\s+([A-Za-z0-9]+)\s+(\d+)$", line)
    if not m:
        return None

    attr = clean_text(m.group(1))
    sku_id = clean_numeric_id(m.group(2))
    sku_code = clean_text(m.group(3))
    qty = clean_text(m.group(4))

    if not sku_id or not sku_code or not qty:
        return None

    return {
        "属性集": attr,
        "SKU ID": sku_id,
        "SKU货号": sku_code,
        "数量": qty
    }

def is_control_line(line):
    """
    是否是非明细控制行
    """
    keywords = [
        "备货母单号", "备货单号", "创建时间", "要求发货时间",
        "收货仓", "打印时间", "序号", "商品信息",
        "SKU ID", "SKU货号", "实际发货数", "拣货数",
        "数量：", "SKC货号："
    ]
    return any(key in line for key in keywords)

def is_complete_detail_candidate(text):
    """
    判断当前拼接内容是否已经像一条完整明细
    """
    text = re.sub(r"\s+", " ", text).strip()
    return re.match(r".+?\s+\d{7,}\s+[A-Za-z0-9]+\s+\d+$", text) is not None


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
    name_key = get_match_col(df_name, ["sku货号", "货号", "编码", "sku"])
    if name_key is None:
        st.error("❌ name_map.xlsx 未找到 SKU货号 / 编码 列")
        st.stop()

    name_value_col = get_match_col(df_name, ["商品名称", "名称", "品名"])
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
    info_key = get_match_col(df_info, ["sku id", "skuid"])
    if info_key is None:
        st.error("❌ product_info.xlsx 未找到 SKU ID 列")
        st.stop()

    shop_col = get_match_col(df_info, ["店铺名称", "店铺"])
    label_col = get_match_col(df_info, ["回收标签"])

    if shop_col is None:
        st.error("❌ product_info.xlsx 未找到 店铺名称 列")
        st.stop()

    if label_col is None:
        st.error("❌ product_info.xlsx 未找到 回收标签 列")
        st.stop()

    df_info[info_key] = df_info[info_key].apply(clean_numeric_id)
    df_info = df_info[df_info[info_key] != ""].copy()

    info_dict = (
        df_info
        .drop_duplicates(subset=[info_key])
        .set_index(info_key)
        .to_dict("index")
    )

    # =========================
    # 6. 解析 PDF（支持跨行明细）
    # =========================
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [line.strip() for line in text.split("\n") if line.strip()]

            active_skc = ""
            in_vmi_block = False
            detail_buffer = ""

            def flush_buffer():
                nonlocal detail_buffer, active_skc, results
                if not detail_buffer:
                    return

                parsed = parse_detail_line(detail_buffer)
                if parsed:
                    sku_id = parsed["SKU ID"]
                    sku_code = parsed["SKU货号"]
                    qty = parsed["数量"]

                    res_prod_name = name_dict.get(sku_code, "-")

                    matched_row = info_dict.get(sku_id, None)
                    if matched_row is not None:
                        res_shop_name = clean_text(matched_row.get(shop_col, "-")) or "-"
                        res_label = clean_text(matched_row.get(label_col, "-")) or "-"
                    else:
                        res_shop_name = "-"
                        res_label = "-"

                    results.append({
                        "店铺名称": res_shop_name,
                        "SKC ID": active_skc,
                        "SKU ID": sku_id,
                        "SKU货号": sku_code,
                        "商品名称": res_prod_name,
                        "回收标签": res_label,
                        "数量": qty
                    })

                detail_buffer = ""

            for line in lines:
                # 抓 SKC
                skc_match = re.match(r"SKC[:：]\s*(\d+)", line)
                if skc_match:
                    # 遇到新 SKC 前，先尝试冲掉上一个 buffer
                    flush_buffer()
                    active_skc = skc_match.group(1)
                    in_vmi_block = False
                    continue

                # 进入明细区
                if "【VMI】" in line:
                    in_vmi_block = True
                    flush_buffer()
                    continue

                # 合计说明当前这个 SKC 的明细结束
                if line.startswith("合计"):
                    flush_buffer()
                    in_vmi_block = False
                    continue

                # 不是明细区，跳过
                if not in_vmi_block:
                    continue

                # 跳过控制行
                if is_control_line(line):
                    continue

                # 明细拼接逻辑
                if not detail_buffer:
                    detail_buffer = line
                else:
                    detail_buffer = f"{detail_buffer} {line}"

                # 如果已经拼成完整明细，就落表
                if is_complete_detail_candidate(detail_buffer):
                    flush_buffer()

            # 一页结束，冲掉残留
            flush_buffer()

    # =========================
    # 7. 输出结果
    # =========================
    if results:
        df_res = pd.DataFrame(results)

        # 固定列顺序
        df_show = df_res[
            ["店铺名称", "SKC ID", "SKU ID", "SKU货号", "商品名称", "回收标签", "数量"]
        ].copy()

        st.subheader("📄 提取结果")
        st.dataframe(df_show, use_container_width=True)

        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_show.to_excel(writer, index=False, sheet_name="结果")

        st.download_button(
            "📥 下载 Excel",
            output.getvalue(),
            file_name="拣货单结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ 未提取到任何有效数据，请检查 PDF 格式。")
