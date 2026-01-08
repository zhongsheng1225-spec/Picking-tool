import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单自动提取", layout="wide")

st.title("📋 拣货单自动提取工具")

# --- 加载基础资料 ---
def load_excel_smart(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except Exception as e:
            st.sidebar.error(f"读取 {name} 失败: {e}")
            return None
    return None

df_prod = load_excel_smart("product_info.xlsx")
df_label = load_excel_smart("label_info.xlsx")

# 侧边栏状态检查
with st.sidebar:
    st.header("⚙️ 资料状态")
    if df_prod is not None: st.success("✅ 商品信息已就绪")
    else: st.error("❌ 缺失 product_info.xlsx")
    if df_label is not None: st.success("✅ 标签信息已就绪")
    else: st.error("❌ 缺失 label_info.xlsx")

# --- 主程序 ---
uploaded_pdf = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_pdf:
    if df_prod is None or df_label is None:
        st.warning("请确保 GitHub 中已上传 product_info.xlsx 和 label_info.xlsx")
    else:
        results = []
        # 预处理基础表
        df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
        df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()
        
        try:
            with pdfplumber.open(uploaded_pdf) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    
                    # 1. 提取仓库
                    wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
                    current_wh = wh_match.group(1) if wh_match else "未知"
                    
                    # 2. 提取本页所有 SKC
                    all_skcs = re.findall(r"SKC[:：]\s*(\d+)", text)
                    
                    # 3. 提取表格
                    table = page.extract_table()
                    if table:
                        headers = table[0]
                        try:
                            sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                            qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                            
                            row_count = 0
                            for row in table[1:]:
                                if not row[sku_idx] or "合计" in str(row): continue
                                
                                sku = str(row[sku_idx]).strip().replace('\n', '')
                                qty = str(row[qty_idx]).strip()
                                
                                # 分配 SKC
                                skc_id = all_skcs[row_count] if row_count < len(all_skcs) else ""
                                
                                # 匹配商品信息
                                prod_name = "-"
                                p_match = df_prod[df_prod['商品编码'] == sku]
                                if not p_match.empty: prod_name = p_match.iloc[0]['商品名称']

                                # 匹配标签
                                label_type = "-"
                                if skc_id:
                                    l_match = df_label[df_label['SKC ID'] == skc_id]
                                    if not l_match.empty: label_type = l_match.iloc[0]['回收标签']

                                results.append({
                                    "发货仓库": current_wh,
                                    "SKC ID": skc_id,
                                    "回收标签类别": label_type,
                                    "货品编码": sku,
                                    "商品名称": prod_name,
                                    "发货数量": qty
                                })
                                row_count += 1
                        except: continue
            
            if results:
                st.success("处理完成！")
                df_res = pd.DataFrame(results)
                st.dataframe(df_res, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_res.to_excel(writer, index=False)
                st.download_button("📥 下载 Excel 结果", output.getvalue(), "拣货单结果.xlsx")
        except Exception as e:
            st.error(f"解析过程中发生错误: {e}")
