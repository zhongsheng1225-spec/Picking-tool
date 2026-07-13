import io
import os
import re
from typing import Any

import pandas as pd
import streamlit as st
from pypdf import PdfReader


APP_TITLE = "拣货单校验工具｜20260713"
RESULT_COLUMNS = [
    "发货仓库",
    "店铺名称",
    "SKC ID",
    "回收标签类别",
    "货品编码",
    "商品名称",
    "发货数量",
]


st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📋",
    layout="wide",
)

st.title(f"📋 {APP_TITLE}")
st.caption(
    "使用纯文本解析 PDF，不调用表格识别；保留 SKU ID 优先匹配、"
    "SKU 货号降级匹配、名称校验和 Excel 导出。"
)


def normalize_key(value: Any) -> str:
    """统一 Excel / PDF 中的编号格式，避免 12345.0、空格、换行导致匹配失败。"""
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip().replace("\n", "").replace("\r", "")
    if text.lower() in {"nan", "none"}:
        return ""

    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    return text


def find_column(
    dataframe: pd.DataFrame,
    candidates: list[str],
    required: bool = True,
) -> str | None:
    """按候选关键词寻找列名，优先完全匹配，其次包含匹配。"""
    columns = [str(col).strip() for col in dataframe.columns]

    for candidate in candidates:
        for original, normalized in zip(dataframe.columns, columns):
            if normalized == candidate:
                return original

    for candidate in candidates:
        for original, normalized in zip(dataframe.columns, columns):
            if candidate in normalized:
                return original

    if required:
        raise ValueError(
            f"找不到需要的列：{' / '.join(candidates)}。"
            f"当前文件列名：{', '.join(columns)}"
        )

    return None


@st.cache_data(show_spinner=False)
def load_excel_file(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(f"仓库中缺少 {path}")

    return pd.read_excel(path, dtype=object)


def build_reference_data(
    df_info: pd.DataFrame,
    df_name: pd.DataFrame,
) -> tuple[dict[str, dict[str, str]], dict[str, dict[str, str]], dict[str, str], dict[str, str]]:
    """生成 SKU ID、SKU 货号及商品名称映射，并返回识别到的列名。"""
    sku_id_col = find_column(df_info, ["SKU ID", "SKUID"])
    sku_code_col = find_column(
        df_info,
        ["SKU货号", "SKU 货号", "货品编码", "货号"],
        required=False,
    )
    shop_col = find_column(df_info, ["店铺名称", "店铺"])
    label_col = find_column(
        df_info,
        ["回收标签类别", "回收标签", "标签类别"],
    )

    info_by_sku_id: dict[str, dict[str, str]] = {}
    info_by_sku_code: dict[str, dict[str, str]] = {}

    for _, row in df_info.iterrows():
        shop_name = normalize_key(row.get(shop_col))
        recycle_label = normalize_key(row.get(label_col))

        sku_id = normalize_key(row.get(sku_id_col))
        if sku_id and sku_id not in info_by_sku_id:
            info_by_sku_id[sku_id] = {
                "店铺名称": shop_name or "-",
                "回收标签类别": recycle_label or "-",
            }

        if sku_code_col is not None:
            sku_code = normalize_key(row.get(sku_code_col))
            if sku_code and sku_code not in info_by_sku_code:
                info_by_sku_code[sku_code] = {
                    "店铺名称": shop_name or "-",
                    "回收标签类别": recycle_label or "-",
                }

    name_key_col = find_column(
        df_name,
        ["货品编码", "SKU货号", "SKU 货号", "编码", "货号", "SKU"],
    )
    product_name_col = find_column(
        df_name,
        ["商品名称", "产品名称", "名称", "品名"],
    )

    name_by_sku_code: dict[str, str] = {}

    for _, row in df_name.iterrows():
        sku_code = normalize_key(row.get(name_key_col))
        product_name = normalize_key(row.get(product_name_col))

        if sku_code and sku_code not in name_by_sku_code:
            name_by_sku_code[sku_code] = product_name or "-"

    detected_columns = {
        "基础信息 SKU ID 列": str(sku_id_col),
        "基础信息 SKU 货号列": str(sku_code_col or "未找到"),
        "基础信息 店铺列": str(shop_col),
        "基础信息 标签列": str(label_col),
        "名称表 货号列": str(name_key_col),
        "名称表 商品名称列": str(product_name_col),
    }

    return (
        info_by_sku_id,
        info_by_sku_code,
        name_by_sku_code,
        detected_columns,
    )


def extract_warehouse(text: str) -> str:
    match = re.search(
        r"收货仓[:：]\s*(.+?)(?:\s+第\d+页/共\d+页|\s+拣货单|$)",
        text,
    )
    return match.group(1).strip() if match else "未知"


def parse_pdf_text(pdf_bytes: bytes) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """
    从 PDF 的文本层提取：
    发货仓库、SKC ID、SKU ID、SKU 货号、实际发货数量。

    数据行格式示例：
    黑色 74650202615 Y010724008 20
    6.17 FT 91090497298 202601717020 10
    """
    records: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []

    reader = PdfReader(io.BytesIO(pdf_bytes))

    detail_pattern = re.compile(
        r"^(?P<attribute>.+?)\s+"
        r"(?P<sku_id>\d{8,})\s+"
        r"(?P<sku_code>[A-Za-z0-9_-]{5,})\s+"
        r"(?P<qty>\d+)\s*$"
    )

    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        warehouse = extract_warehouse(text)
        active_skc = ""
        parsed_on_page = 0

        lines = [
            line.strip()
            for line in text.splitlines()
            if line and line.strip()
        ]

        for line_number, line in enumerate(lines, start=1):
            skc_match = re.search(r"SKC[:：]\s*(\d+)", line)
            if skc_match:
                active_skc = skc_match.group(1)
                continue

            detail_match = detail_pattern.match(line)
            if not detail_match:
                continue

            sku_id = normalize_key(detail_match.group("sku_id"))
            sku_code = normalize_key(detail_match.group("sku_code"))
            qty_text = normalize_key(detail_match.group("qty"))

            if not active_skc:
                issues.append(
                    {
                        "页码": page_number,
                        "行号": line_number,
                        "问题类型": "缺少 SKC ID",
                        "原始内容": line,
                    }
                )

            if not sku_id or not sku_code or not qty_text.isdigit():
                issues.append(
                    {
                        "页码": page_number,
                        "行号": line_number,
                        "问题类型": "明细字段不完整",
                        "原始内容": line,
                    }
                )
                continue

            records.append(
                {
                    "页码": page_number,
                    "发货仓库": warehouse,
                    "SKC ID": active_skc or "-",
                    "SKU ID": sku_id,
                    "货品编码": sku_code,
                    "发货数量": int(qty_text),
                    "属性集": detail_match.group("attribute").strip(),
                    "原始内容": line,
                }
            )
            parsed_on_page += 1

        if warehouse == "未知":
            issues.append(
                {
                    "页码": page_number,
                    "行号": "-",
                    "问题类型": "未识别到发货仓库",
                    "原始内容": "",
                }
            )

        if parsed_on_page == 0:
            issues.append(
                {
                    "页码": page_number,
                    "行号": "-",
                    "问题类型": "本页未提取到明细",
                    "原始内容": "",
                }
            )

    return records, issues


def enrich_and_validate(
    raw_records: list[dict[str, Any]],
    info_by_sku_id: dict[str, dict[str, str]],
    info_by_sku_code: dict[str, dict[str, str]],
    name_by_sku_code: dict[str, str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    result_rows: list[dict[str, Any]] = []
    validation_rows: list[dict[str, Any]] = []

    for row_number, record in enumerate(raw_records, start=1):
        sku_id = normalize_key(record["SKU ID"])
        sku_code = normalize_key(record["货品编码"])

        matched_info = info_by_sku_id.get(sku_id)
        match_method = "SKU ID"

        if matched_info is None:
            matched_info = info_by_sku_code.get(sku_code)
            match_method = "SKU 货号降级匹配"

        if matched_info is None:
            matched_info = {
                "店铺名称": "-",
                "回收标签类别": "-",
            }
            match_method = "未匹配"

        product_name = name_by_sku_code.get(sku_code, "-")

        result_rows.append(
            {
                "发货仓库": record["发货仓库"],
                "店铺名称": matched_info["店铺名称"],
                "SKC ID": record["SKC ID"],
                "回收标签类别": matched_info["回收标签类别"],
                "货品编码": sku_code,
                "商品名称": product_name,
                "发货数量": record["发货数量"],
            }
        )

        problems: list[str] = []

        if record["发货仓库"] == "未知":
            problems.append("发货仓库未识别")
        if record["SKC ID"] in {"", "-"}:
            problems.append("SKC ID 缺失")
        if matched_info["店铺名称"] == "-":
            problems.append("店铺名称未匹配")
        if matched_info["回收标签类别"] == "-":
            problems.append("回收标签类别未匹配")
        if product_name == "-":
            problems.append("商品名称未匹配")

        validation_rows.append(
            {
                "结果行号": row_number,
                "页码": record["页码"],
                "SKU ID": sku_id,
                "货品编码": sku_code,
                "匹配方式": match_method,
                "校验结果": "通过" if not problems else "；".join(problems),
                "原始内容": record["原始内容"],
            }
        )

    result_df = pd.DataFrame(result_rows, columns=RESULT_COLUMNS)
    validation_df = pd.DataFrame(validation_rows)

    return result_df, validation_df


def build_excel(
    result_df: pd.DataFrame,
    validation_df: pd.DataFrame,
    parse_issues_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(
            writer,
            index=False,
            sheet_name="提取结果",
        )
        validation_df.to_excel(
            writer,
            index=False,
            sheet_name="校验报告",
        )
        parse_issues_df.to_excel(
            writer,
            index=False,
            sheet_name="解析异常",
        )

        workbook = writer.book

        header_format = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        )
        warning_format = workbook.add_format(
            {
                "bg_color": "#FFF2CC",
                "font_color": "#9C6500",
            }
        )

        for sheet_name, dataframe in {
            "提取结果": result_df,
            "校验报告": validation_df,
            "解析异常": parse_issues_df,
        }.items():
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(
                0,
                0,
                max(len(dataframe), 1),
                max(len(dataframe.columns) - 1, 0),
            )

            for col_index, column_name in enumerate(dataframe.columns):
                worksheet.write(
                    0,
                    col_index,
                    column_name,
                    header_format,
                )

                values = dataframe[column_name].astype(str).tolist()
                max_width = max(
                    [len(str(column_name))]
                    + [len(value) for value in values[:1000]]
                )
                worksheet.set_column(
                    col_index,
                    col_index,
                    min(max_width + 2, 42),
                )

        if not validation_df.empty:
            result_col = validation_df.columns.get_loc("校验结果")
            worksheet = writer.sheets["校验报告"]
            worksheet.conditional_format(
                1,
                result_col,
                len(validation_df),
                result_col,
                {
                    "type": "text",
                    "criteria": "not containing",
                    "value": "通过",
                    "format": warning_format,
                },
            )

    return output.getvalue()


# =========================================================
# 加载两张基础表
# =========================================================
load_errors: list[str] = []

try:
    df_info = load_excel_file("product_info.xlsx")
except Exception as exc:
    df_info = None
    load_errors.append(str(exc))

try:
    df_name = load_excel_file("name_map.xlsx")
except Exception as exc:
    df_name = None
    load_errors.append(str(exc))


with st.sidebar:
    st.header("⚙️ 基础资料状态")

    if df_info is not None:
        st.success(f"product_info.xlsx：{len(df_info)} 行")
    else:
        st.error("product_info.xlsx 未就绪")

    if df_name is not None:
        st.success(f"name_map.xlsx：{len(df_name)} 行")
    else:
        st.error("name_map.xlsx 未就绪")

    if load_errors:
        with st.expander("查看读取错误"):
            for error in load_errors:
                st.error(error)


uploaded_file = st.file_uploader(
    "上传 PDF 拣货单",
    type=["pdf"],
)


if uploaded_file is not None:
    if df_info is None or df_name is None:
        st.error(
            "PDF 已上传，但基础 Excel 未全部读取成功。"
            "请先确认 product_info.xlsx 和 name_map.xlsx "
            "已放在 GitHub 仓库根目录。"
        )
        st.stop()

    try:
        (
            info_by_sku_id,
            info_by_sku_code,
            name_by_sku_code,
            detected_columns,
        ) = build_reference_data(df_info, df_name)
    except Exception as exc:
        st.error("基础 Excel 的列名识别失败：")
        st.exception(exc)
        st.stop()

    with st.expander("查看自动识别到的基础表列名"):
        st.json(detected_columns)

    with st.spinner("正在提取并校验拣货单……"):
        try:
            raw_records, parse_issues = parse_pdf_text(
                uploaded_file.getvalue()
            )
        except Exception as exc:
            st.error("PDF 文本读取失败：")
            st.exception(exc)
            st.stop()

        if not raw_records:
            st.error(
                "没有提取到任何 SKU 明细。"
                "请查看“解析异常”，或确认 PDF 是否仍是同类拣货单格式。"
            )

            issues_df = pd.DataFrame(parse_issues)
            if not issues_df.empty:
                st.dataframe(
                    issues_df,
                    use_container_width=True,
                )
            st.stop()

        result_df, validation_df = enrich_and_validate(
            raw_records,
            info_by_sku_id,
            info_by_sku_code,
            name_by_sku_code,
        )

        parse_issues_df = pd.DataFrame(
            parse_issues,
            columns=["页码", "行号", "问题类型", "原始内容"],
        )

    total_rows = len(result_df)
    shop_count = (
        result_df.loc[
            result_df["店铺名称"] != "-",
            "店铺名称",
        ].nunique()
    )
    missing_names = int((result_df["商品名称"] == "-").sum())
    missing_info = int(
        (
            (result_df["店铺名称"] == "-")
            | (result_df["回收标签类别"] == "-")
        ).sum()
    )
    validation_problem_count = int(
        (validation_df["校验结果"] != "通过").sum()
    )

    st.subheader("🔍 自动体检看板")

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric("处理总行数", total_rows)
    with col2:
        st.metric("涉及店铺数", shop_count)
    with col3:
        st.metric("名称未匹配", missing_names)
    with col4:
        st.metric("基础信息未匹配", missing_info)
    with col5:
        st.metric("需人工复核", validation_problem_count)

    if validation_problem_count > 0 or not parse_issues_df.empty:
        st.warning(
            "存在未匹配或解析异常，请先查看校验报告再下载使用。"
        )
    else:
        st.success("所有提取行均通过当前校验。")

    st.subheader("📄 提取结果")
    st.dataframe(
        result_df,
        use_container_width=True,
    )

    with st.expander("查看逐行校验报告"):
        st.dataframe(
            validation_df,
            use_container_width=True,
        )

    if not parse_issues_df.empty:
        with st.expander("查看解析异常"):
            st.dataframe(
                parse_issues_df,
                use_container_width=True,
            )

    excel_bytes = build_excel(
        result_df,
        validation_df,
        parse_issues_df,
    )

    st.download_button(
        label="📥 下载提取及校验结果",
        data=excel_bytes,
        file_name="拣货单提取结果_稳定文本解析版.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
