import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单提取工具", layout="wide")

# --- 1. 基础资料智能加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except Exception as e:
            st.error(f"读取文件 {name} 出错: {e}")
    return None

df_prod = load_data("product_info.xlsx")
df_label = load_data("label_info.xlsx")

st.title("📋 拣货单自动提取")

# 侧边栏状态栏
with st.sidebar:
    st.header("系统检查")
    if df_prod is not None: st.success("商品信息：已就绪")
    else: st.error("缺失 product_info.xlsx")
    if df_label is not None: st.success("标签信息：已就绪")
    else: st.error("缺失 label_info.xlsx")

# --- 2. PDF 处理主逻辑 ---
uploaded_file = st.file_uploader("请上传 PDF 拣货单", type="pdf")

if uploaded_file and df_prod is not None and df_label is not None:
    results = []
    
    # 数据清洗：统一转为字符串并去空格
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()

    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                
                # 提取仓库
                wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
                current_wh = wh_match.group(1) if wh_match else "未知"
                
                # 【全方位抓取 SKC】：匹配 SKC 字样后的 5 位以上数字
                found_skcs = re.findall(r"SKC[:：\s]+(\d{5,})", text)
                
                table = page.extract_table()
                if table:
                    headers = table[0]
                    # 动态寻找列名位置
                    try:
                        sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                        qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                        
                        row_count = 0
                        for row in table[1:]:
                            if not row[sku_idx] or "合计" in str(row): continue
                            
                            sku = str(row[sku_idx]).strip().replace('\n', '')
                            qty = str(row[qty_idx]).strip()
                            
                            # 按顺序分配 SKC ID
                            skc_id = found_skcs[row_count] if row_count < len(found_skcs) else ""
                            
                            # 关联 VLOOKUP
                            p_name = "-"
                            p_match = df_prod[df_prod['商品编码'] == sku]
                            if not p_match.empty: p_name = p_match.iloc[0]['商品名称']

                            l_type = "-"
                            if skc_id:
                                l_match = df_label[df_label['SKC ID'] == skc_id]
                                if not l_match.empty: l_type = l_match.iloc[0]['回收标签']

                            results.append({
                                "发货仓库": current_wh,
                                "SKC ID": skc_id,
                                "回收标签类别": l_type,
                                "货品编码": sku,
                                "商品名称": p_name,
                                "发货数量": qty
                            })
                            row_count += 1
                    except Exception as e:
                        st.warning(f"页面列识别跳过: {e}")

        if results:
            df_res = pd.DataFrame(results)
            st.success("数据处理完毕！")
            st.dataframe(df_res, use_container_width=True)
            
            # 生成下载
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False)
            st.download_button("📥 下载提取结果 (Excel)", output.getvalue(), "提取结果.xlsx")

    except Exception as e:
        st.error(f"解析 PDF 时发生严重错误: {e}")
