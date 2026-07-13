import streamlit as st

st.title("上传测试")

uploaded_file = st.file_uploader(
    "上传文件",
    type=["pdf", "xlsx", "xls"],
)

if uploaded_file is not None:
    st.success(f"已收到：{uploaded_file.name}")
    st.write("文件大小：", uploaded_file.size)
