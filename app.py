import streamlit as st
import pdfplumber
import io

st.title("PDF 测试")

uploaded_file = st.file_uploader(
    "上传 PDF",
    type=["pdf"],
)

if uploaded_file is not None:
    st.success(f"已收到：{uploaded_file.name}")

    pdf_bytes = uploaded_file.getvalue()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        st.write("PDF 页数：", len(pdf.pages))

        text = pdf.pages[0].extract_text()

        if text:
            st.subheader("第一页前200个字符")
            st.text(text[:200])
        else:
            st.warning("第一页没有提取到文字")
