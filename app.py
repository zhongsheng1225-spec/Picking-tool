import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="拣货单数据提取-人工校验版", layout="wide")

st.title("📋 拣货单自动提取 (含校验标识)")
st.info("💡 逻辑：若页面无法识别到'收货仓'，该页所有货品将标记为'未知'，方便您人工核对。")

uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file is not None:
    results = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # 1. 严格提取当前页面的仓库名称
            # 匹配“收货仓:”后面紧跟的非空文字
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
            current_page_warehouse = wh_match.group(1) if wh_match else "未知"
            
            # 2. 提取表格数据
            table = page.extract_table()
            if table:
                headers = table[0]
                try:
                    # 定位关键列索引
                    sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                    
                    for row in table[1:]:
                        # 排除合计行或空行
                        if not row[sku_idx] or "合计" in str(row):
                            continue
                        
                        sku = str(row[sku_idx]).strip().replace('\n', '')
                        qty = str(row[qty_idx]).strip()
                        
                        if sku and sku != "None":
                            results.append({
                                "发货仓库": current_page_warehouse,
                                "货品编码": sku,
                                "发货数量": qty
                            })
                except (StopIteration, IndexError):
                    continue

    if results:
        df = pd.DataFrame(results)
        st.success(f"处理完成，共计 {len(df)} 行数据。")
        
        # 统计有多少行是“未知”，提醒用户
        unknown_count = len(df[df["发货仓库"] == "未知"])
        if unknown_count > 0:
            st.warning(f"注意：共有 {unknown_count} 行数据的仓库显示为'未知'，请重点人工校验。")
        
        st.dataframe(df, use_container_width=True)

        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='待校验结果')
        
        st.download_button(
            label="下载结果进行人工校验",
            data=output.getvalue(),
            file_name="拣货单待校验明细.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
