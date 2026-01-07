import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 页面配置
st.set_page_config(page_title="拣货单自动化提取", layout="wide")

st.title("📋 拣货单数据自动提取工具")
st.info("💡 提示：上传 PDF 后可自动识别仓库、货品编码和发货数量，并支持下载 Excel。")

# 文件上传组件
uploaded_file = st.file_uploader("请上传 PDF 拣货单文件", type="pdf")

if uploaded_file is not None:
    results = []
    # 使用 pdfplumber 打开上传的文件
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # 1. [span_0](start_span)[span_1](start_span)[span_2](start_span)提取收货仓名称[span_0](end_span)[span_1](end_span)[span_2](end_span)
            warehouse = "未知仓库"
            wh_match = re.search(r"收货仓:\s*([^\s\n]+)", text)
            if wh_match:
                warehouse = wh_match.group(1)
            
            # 2. [span_3](start_span)[span_4](start_span)[span_5](start_span)[span_6](start_span)提取表格数据[span_3](end_span)[span_4](end_span)[span_5](end_span)[span_6](end_span)
            table = page.extract_table()
            if table:
                headers = table[0]
                try:
                    # 动态定位列索引，防止格式偏移
                    sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                    
                    for row in table[1:]:
                        # [span_7](start_span)过滤无效行：空行、只有序号的行、或包含“合计”的行[span_7](end_span)
                        if not row[sku_idx] or "合计" in str(row):
                            continue
                        
                        sku = str(row[sku_idx]).strip().replace('\n', '')
                        qty = str(row[qty_idx]).strip()
                        
                        if sku and sku != "None":
                            results.append({
                                "发货仓库": warehouse,
                                "货品编码": sku,
                                "发货数量": qty
                            })
                except (StopIteration, IndexError):
                    continue

    if results:
        df = pd.DataFrame(results)
        st.success(f"成功提取 {len(df)} 条数据！")
        
        # 展示预览表格
        st.dataframe(df, use_container_width=True)

        # 转换为 Excel 供下载
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='提取结果')
        
        st.download_button(
            label="点击下载提取后的 Excel 文件",
            data=output.getvalue(),
            file_name="拣货单数据提取.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("未能识别出有效数据，请确保 PDF 格式正确。")
