import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-Excel版", layout="wide")

st.title("📋 拣货单自动提取 (Excel基础资料版)")
st.info("💡 提示：系统将直接读取仓库内的 .xlsx 文件，解决中文乱码问题。")

# --- 核心：直接读取 Excel 的函数 ---
def load_excel_data(file_name):
    if os.path.exists(file_name):
        try:
            # 直接使用 pandas 读取 excel，不会有乱码问题
            return pd.read_excel(file_name)
        except Exception as e:
            st.error(f"加载 {file_name} 失败: {e}")
    return None

# 预加载基础数据 (直接读 xlsx)
df_prod = load_excel_data("product_info.xlsx")
df_label = load_excel_data("label_info.xlsx")

# 侧边栏状态检查
with st.sidebar:
    st.header("⚙️ 资料状态 (Excel)")
    if df_prod is not None:
        st.success(f"✅ 商品信息已就绪")
    else:
        st.error("❌ 缺失 product_info.xlsx")
        
    if df_label is not None:
        st.success(f"✅ 标签信息已就绪")
    else:
        st.error("❌ 缺失 label_info.xlsx")

# --- 主程序：处理 PDF ---
uploaded_pdf = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_pdf is not None:
    if df_prod is None or df_label is None:
        st.warning("基础资料未加载，请检查 GitHub 中是否存在 product_info.xlsx 和 label_info.xlsx")
    else:
        results = []
        with pdfplumber.open(uploaded_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                # 提取仓库
                wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
                current_wh = wh_match.group(1) if wh_match else "未知"
                
                table = page.extract_table()
                if table:
                    headers = table[0]
                    try:
                        # 查找关键列索引
                        skc_text_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
                        sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                        qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                        
                        for row in table[1:]:
                            if not row[sku_idx] or "合计" in str(row):
                                continue
                            
                            sku = str(row[sku_idx]).strip().replace('\n', '')
                            qty = str(row[qty_idx]).strip()
                            
                            # 解析 SKC ID
                            skc_id = ""
                            info_cell = str(row[skc_text_idx])
                            skc_match = re.search(r"SKC:\s*(\d+)", info_cell)
                            if skc_match:
                                skc_id = skc_match.group(1)

                            # 1. 关联商品名称
                            prod_name = "-"
                            m_prod = df_prod[df_prod['商品编码'].astype(str) == sku]
                            if not m_prod.empty:
                                prod_name = m_prod.iloc[0]['商品名称']

                            # 2. 关联回收标签
                            label_type = "-"
                            if skc_id:
                                m_lab = df_label[df_label['SKC ID'].astype(str) == skc_id]
                                if not m_lab.empty:
                                    label_type = m_lab.iloc[0]['回收标签']

                            results.append({
                                "发货仓库": current_wh,
                                "SKC ID": skc_id,
                                "回收标签类别": label_type,
                                "货品编码": sku,
                                "商品名称": prod_name,
                                "发货数量": qty
                            })
                    except:
                        continue

        if results:
            df_final = pd.DataFrame(results)
            st.success("处理成功！")
            st.dataframe(df_final, use_container_width=True)

            # 导出 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='结果')
            
            st.download_button("📥 下载完整结果 (Excel)", output.getvalue(), "拣货单最终结果.xlsx")
