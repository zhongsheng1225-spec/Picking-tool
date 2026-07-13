import streamlit as st
import pdfplumber
import io

st.title("PDF 表格测试")

uploaded_file = st.file_uploader(
    "上传 PDF",
    type=["pdf"],
)

if uploaded_file is not None:
    pdf_bytes = uploaded_file.getvalue()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]

        st.write("开始提取表格...")

        table = page.extract_table()

        st.write("提取完成！")

        if table:
            st.write("表格行数：", len(table))
            st.write(table[:3])
        else:
            st.warning("没有找到表格")
