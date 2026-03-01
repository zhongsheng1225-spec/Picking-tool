import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-店铺合并版", layout="wide")

st.title("📋 拣货单自动化提取 (新增店铺名称)")

# --- 1. 基础资料智能加载 (单表版) ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

# 现在只需要加载一张合并后的表
df_base = load_data("product_info.xlsx")

with st.sidebar:
    st.header("⚙️ 基础资料状态")
    if df_base is not None:
        st.success("✅ 商品基础信息表已就绪")
        # 预检查必要的列
        cols = df_base.columns.tolist()
        st.write("已检测到列：", ", ".join(cols[:5]) + "...")
    else:
        st.error("❌ 缺失 product_info.xlsx")

# --- 2. 处理 PDF 主逻辑 ---
uploaded_file = st.file_uploader("上传 PDF 拣货单文件", type="pdf")

if uploaded_file and df_base is not None:
    results = []
    
    # 数据预清洗：确保关键匹配列为字符串
    df_base['商品编码'] = df_base['商品编码'].astype(str).str.strip()
    # 建立一个以商品编码为索引的字典，提速查询
    base_dict = df_base.drop_duplicates('商品编码').set_index('商品编码').to_dict('index')

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # A. 提取本页仓库
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            # B. 提取表格
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
                qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
            except: continue

            active_skc = "" # SKC指针
            
            for row in table[1:]:
                if not row[sku_idx] or "合计" in str(row): continue
                
                # C. SKC 向下补全
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                if skc_match:
                    active_skc = skc_match.group(1)
                
                sku = str(row[sku_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()

                # D. 关联匹配 (从单表获取所有信息)
                shop_name = "-"
                prod_name = "-"
                label_type = "-"
                
                if sku in base_dict:
                    item = base_dict[sku]
                    shop_name = item.get('店铺名称', '-')
                    prod_name = item.get('商品名称', '-')
                    label_type = item.get('回收标签', '-')

                results.append({
                    "发货仓库": current_wh,
                    "店铺名称": shop_name,
                    "SKC ID": active_skc,
                    "回收标签类别": label_type,
                    "货品编码": sku,
                    "商品名称": prod_name,
                    "发货数量": qty
                })

    # --- 3. 结果展示与自动校验 ---
    if results:
        df_res = pd.DataFrame(results)
        
        st.subheader("🔍 自动体检看板")
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("处理总行数", len(df_res))
        with m2:
            shops = [s for s in df_res['店铺名称'].unique() if s != "-"]
            st.metric("涉及店铺数", len(shops), help=f"涉及店铺：{', '.join(shops)}")
        with m3:
            fail_match = len(df_res[df_res['商品名称'] == '-'])
            st.metric("匹配失败数", fail_match, delta_color="inverse")

        st.dataframe(df_res, use_container_width=True)
        
        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='拣货提取结果')
        
        st.download_button("📥 下载完整 Excel 结果", output.getvalue(), "拣货单结果_含店铺.xlsx")
