import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(
    page_title="拣货单增强工具-完整版",
    layout="wide",
)

st.title("📋 拣货单自动提取（路标35暂停测试版）")

# =========================================================
# 1. 加载基础 Excel
# =========================================================
st.write("🟢 路标 ①：程序已启动，准备读取基础 Excel")


def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except Exception as exc:
            st.error(f"读取 {name} 失败：{exc}")
            return None

    return None


df_info = load_data("product_info.xlsx")
st.write("🟢 路标 ②：product_info.xlsx 读取步骤已结束")

df_name = load_data("name_map.xlsx")
st.write("🟢 路标 ③：name_map.xlsx 读取步骤已结束")


with st.sidebar:
    st.header("⚙️ 资料体检状态")

    if df_info is not None:
        st.success("✅ product_info.xlsx 已就绪")
    else:
        st.error("❌ 缺失或无法读取 product_info.xlsx")

    if df_name is not None:
        st.success("✅ name_map.xlsx 已就绪")
    else:
        st.error("❌ 缺失或无法读取 name_map.xlsx")

    st.divider()

    st.info(
        "💡 校验逻辑：\n"
        "1. [名称对照表] 找具体商品名\n"
        "2. [基础信息表] 找店铺和回收标签\n"
        "3. 优先使用 SKU ID 匹配"
    )


# =========================================================
# 2. 辅助函数
# =========================================================
def get_match_col(df, keywords):
    for col in df.columns:
        if any(key in str(col) for key in keywords):
            return col

    return df.columns[0]


st.write("🟢 路标 ④：页面和辅助函数加载完成")


# =========================================================
# 3. 上传 PDF
# =========================================================
uploaded_file = st.file_uploader(
    "上传 PDF 拣货单",
    type=["pdf"],
)

if uploaded_file is not None:
    st.success(f"已收到文件：{uploaded_file.name}")
    st.write("🟢 路标 ⑤：PDF 上传成功")


# =========================================================
# 4. 正式处理
# =========================================================
if uploaded_file is not None and df_info is not None and df_name is not None:

    try:
        results = []

        # -------------------------------------------------
        # 4.1 准备基础信息表
        # -------------------------------------------------
        st.write("🟡 路标 ⑥：开始查找 SKU ID 列")

        sku_id_col = None

        for col in df_info.columns:
            if "SKU ID" in str(col):
                sku_id_col = col
                break

        if sku_id_col is None:
            st.error(
                "❌ product_info.xlsx 中缺少 SKU ID 列，"
                "请检查文件。"
            )
            st.stop()

        st.write(f"🟢 路标 ⑦：已找到 SKU ID 列：{sku_id_col}")

        df_info[sku_id_col] = (
            df_info[sku_id_col]
            .astype(str)
            .str.strip()
        )

        st.write("🟡 路标 ⑧：准备生成 SKU ID 信息字典")

        info_dict = (
            df_info
            .drop_duplicates(sku_id_col)
            .set_index(sku_id_col)
            .to_dict("index")
        )

        st.write(
            f"🟢 路标 ⑨：SKU ID 信息字典生成完成，"
            f"共 {len(info_dict)} 条"
        )

        # -------------------------------------------------
        # 4.2 SKU 货号降级匹配
        # -------------------------------------------------
        st.write("🟡 路标 ⑩：开始查找 SKU 货号列")

        sku_code_col = None

        for col in df_info.columns:
            if "SKU货号" in str(col):
                sku_code_col = col
                break

        if sku_code_col is not None:
            st.write(
                f"🟢 路标 ⑪：已找到 SKU 货号列："
                f"{sku_code_col}"
            )

            df_info[sku_code_col] = (
                df_info[sku_code_col]
                .astype(str)
                .str.strip()
            )

            st.write("🟡 路标 ⑫：准备生成 SKU 货号字典")

            sku_code_dict = (
                df_info
                .drop_duplicates(sku_code_col)
                .set_index(sku_code_col)
                .to_dict("index")
            )

            st.write(
                f"🟢 路标 ⑬：SKU 货号字典生成完成，"
                f"共 {len(sku_code_dict)} 条"
            )

        else:
            sku_code_dict = {}

            st.warning(
                "🟠 路标 ⑬：未找到 SKU货号 列，"
                "已跳过降级匹配字典"
            )

        # -------------------------------------------------
        # 4.3 准备名称对照表
        # -------------------------------------------------
        st.write("🟡 路标 ⑭：开始查找名称对照表匹配列")

        name_key = get_match_col(
            df_name,
            ["编码", "SKU", "货号"],
        )

        st.write(
            f"🟢 路标 ⑮：名称对照表匹配列：{name_key}"
        )

        df_name[name_key] = (
            df_name[name_key]
            .astype(str)
            .str.strip()
        )

        st.write("🟡 路标 ⑯：准备生成商品名称字典")

        name_dict = (
            df_name
            .drop_duplicates(name_key)
            .set_index(name_key)
            .iloc[:, 0]
            .to_dict()
        )

        st.write(
            f"🟢 路标 ⑰：商品名称字典生成完成，"
            f"共 {len(name_dict)} 条"
        )

        # -------------------------------------------------
        # 4.4 读取 PDF
        # -------------------------------------------------
        st.write("🟡 路标 ⑱：准备读取 PDF 二进制内容")

        pdf_bytes = uploaded_file.getvalue()

        st.write(
            f"🟢 路标 ⑲：PDF 内容读取成功，"
            f"共 {len(pdf_bytes)} 字节"
        )

        st.write("🟡 路标 ⑳：准备打开 PDF")

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:

            st.write(
                f"🟢 路标 ㉑：PDF 打开成功，"
                f"共 {len(pdf.pages)} 页"
            )

            for page_number, page in enumerate(
                pdf.pages,
                start=1,
            ):

                st.write(
                    f"🟡 路标 ㉒：开始处理第 "
                    f"{page_number} 页"
                )

                text = page.extract_text() or ""

                st.write(
                    f"🟢 路标 ㉓：第 {page_number} 页"
                    f"文字提取完成，共 {len(text)} 个字符"
                )

                wh_match = re.search(
                    r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)",
                    text,
                )

                current_wh = (
                    wh_match.group(1)
                    if wh_match
                    else "未知"
                )

                st.write(
                    f"🟢 路标 ㉔：第 {page_number} 页"
                    f"仓库：{current_wh}"
                )

                st.write(
                    f"🟡 路标 ㉕：准备提取第 "
                    f"{page_number} 页表格"
                )

                table = page.extract_table()

                st.write(
                    f"🟢 路标 ㉖：第 {page_number} 页"
                    f"表格提取结束"
                )

                if not table:
                    st.warning(
                        f"第 {page_number} 页没有识别到表格，"
                        "已跳过"
                    )
                    continue

                st.write(
                    f"🟢 路标 ㉗：第 {page_number} 页"
                    f"表格共有 {len(table)} 行"
                )

                headers = table[0]

                st.write(
                    f"🟢 路标 ㉘：第 {page_number} 页表头"
                )

                st.write(headers)

                st.write(
                    f"🟡 路标 ㉙：开始查找第 "
                    f"{page_number} 页各列位置"
                )

                try:
                    sku_code_idx = next(
                        i
                        for i, h in enumerate(headers)
                        if h
                        and (
                            "货号" in str(h)
                            or "编码" in str(h)
                        )
                    )

                    info_idx = next(
                        i
                        for i, h in enumerate(headers)
                        if h
                        and "商品信息" in str(h)
                    )

                    qty_idx = next(
                        i
                        for i, h in enumerate(headers)
                        if h
                        and "发货数" in str(h)
                    )

                    sku_id_idx = None

                    for i, h in enumerate(headers):
                        if h and (
                            "SKU ID" in str(h)
                            or "SKUID"
                            in str(h).replace(" ", "")
                        ):
                            sku_id_idx = i
                            break

                except StopIteration:
                    st.warning(
                        f"第 {page_number} 页缺少必要列，"
                        "已跳过"
                    )
                    continue

                st.write(
                    f"🟢 路标 ㉚：第 {page_number} 页"
                    f"列位置查找完成"
                )

                st.write(
                    {
                        "货号列位置": sku_code_idx,
                        "商品信息列位置": info_idx,
                        "发货数量列位置": qty_idx,
                        "SKU ID列位置": sku_id_idx,
                    }
                )

                active_skc = ""

                st.write(
                    f"🟡 路标 ㉛：开始遍历第 "
                    f"{page_number} 页数据行"
                )

                for row_number, row in enumerate(
                    table[1:],
                    start=2,
                ):
                    required_indexes = [
                        sku_code_idx,
                        info_idx,
                        qty_idx,
                    ]

                    if any(
                        index >= len(row)
                        for index in required_indexes
                    ):
                        st.warning(
                            f"第 {page_number} 页第 "
                            f"{row_number} 行列数不足，已跳过"
                        )
                        continue

                    if (
                        not row[sku_code_idx]
                        or "合计" in str(row)
                    ):
                        continue

                    cell_info = str(row[info_idx])

                    skc_match = re.search(
                        r"SKC[:：\s]+(\d+)",
                        cell_info,
                    )

                    if skc_match:
                        active_skc = skc_match.group(1)

                    sku_code = (
                        str(row[sku_code_idx])
                        .strip()
                        .replace("\n", "")
                    )

                    qty = str(row[qty_idx]).strip()

                    prod_name = name_dict.get(
                        sku_code,
                        "-",
                    )

                    res_shop_name = "-"
                    res_label = "-"
                    matched = False

                    if (
                        sku_id_idx is not None
                        and sku_id_idx < len(row)
                    ):
                        sku_id = str(
                            row[sku_id_idx]
                        ).strip()

                        if sku_id and sku_id in info_dict:
                            res_shop_name = (
                                info_dict[sku_id]
                                .get("店铺名称", "-")
                            )

                            res_label = (
                                info_dict[sku_id]
                                .get("回收标签", "-")
                            )

                            matched = True

                    if (
                        not matched
                        and sku_code
                        and sku_code in sku_code_dict
                    ):
                        res_shop_name = (
                            sku_code_dict[sku_code]
                            .get("店铺名称", "-")
                        )

                        res_label = (
                            sku_code_dict[sku_code]
                            .get("回收标签", "-")
                        )

                        matched = True

                    if (
                        not matched
                        and sku_code
                        and sku_code in info_dict
                    ):
                        res_shop_name = (
                            info_dict[sku_code]
                            .get("店铺名称", "-")
                        )

                        res_label = (
                            info_dict[sku_code]
                            .get("回收标签", "-")
                        )

                    results.append(
                        {
                            "发货仓库": current_wh,
                            "店铺名称": res_shop_name,
                            "SKC ID": active_skc,
                            "回收标签类别": res_label,
                            "货品编码": sku_code,
                            "商品名称": prod_name,
                            "发货数量": qty,
                        }
                    )

                st.write(
                    f"🟢 路标 ㉜：第 {page_number} 页"
                    f"遍历完成，当前累计结果 "
                    f"{len(results)} 行"
                )

        # -------------------------------------------------
        # 4.5 生成结果表
        # -------------------------------------------------
        st.write(
            "🟡 路标 ㉝：PDF 全部处理完成，"
            "准备生成结果表"
        )

        if results:
            df_res = pd.DataFrame(results)

            st.write(
                f"🟢 路标 ㉞：结果表生成完成，"
                f"共 {len(df_res)} 行"
            )

            st.write("🟡 路标 ㉟：准备计算体检数据")

            # 暂停测试：确认路标35以前不会崩
# ==========================
# 路标36：计算统计
# ==========================
shops = (
    df_res[df_res["店铺名称"] != "-"]["店铺名称"]
    .nunique()
)

missing_name = len(
    df_res[df_res["商品名称"] == "-"]
)

missing_info = len(
    df_res[df_res["店铺名称"] == "-"]
)

st.write("🟢 路标 ㊱：体检数据计算完成")

st.subheader("🔍 自动体检看板")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("处理总行数", len(df_res))

with col2:
    st.metric("涉及店铺数", shops)

with col3:
    st.metric("名称未匹配", missing_name)

with col4:
    st.metric("基础信息未匹配", missing_info)

st.write("🟢 路标 ㊲：Metric 显示完成")

# ==========================
# 测试 dataframe
# ==========================
st.write("🟡 路标 ㊳：准备显示 DataFrame")

st.dataframe(
    df_res,
    width="stretch"
)

st.write("🟢 路标 ㊴：DataFrame 显示完成")

# 到这里停止，不测试下载
st.stop()
