import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-全表扫描版", layout="wide")

st.title("📋 拣货单自动提取 (3.12 兼容版)")

# --- 1. 强力加载与列名清洗 ---
def load_and_clean(name):
    full_path = name + ".xlsx"
    if os.path.exists(full_path):
        try:
            df = pd.read_excel(full_path, engine='openpyxl')
            # 清理列名：转字符串、去空格、去换行
            df.columns = [str(c).strip().replace('\n', '') for c in df.columns]
            return df
        except Exception as e:
            st.sidebar.error(f"读取 {name} 失败: {e}")
            return None
    return None

df_info = load_and_clean("product_info")
df_name = load_and_clean("name_map")

# --- 2. 状态检查 ---
with st.sidebar:
    st.header("⚙️ 资料状态")
    if df_info is not None: st.success("✅ 基础信息表已载入")
    else: st.error("❌ 缺失 product_info.xlsx")
    
    if df_name is not None: st.success("✅ 名称对照表已载入")
    else: st.error("❌ 缺失 name_map.xlsx")

# --- 3. 核心提取逻辑 ---
if df_info is not None and df_name is not None:
    # 模糊寻找列名函数
    def find_best_col(df, targets):
        for col in df.columns:
            if any(t in col for t in targets):
                return col
        return None

    # A. 预处理名称对照表
    n_col = find_best_col(df_name, ['商品编码', '货号', 'SKU'])
    val_col = find_best_col(df_name, ['商品名称', '标题'])
    if n_col and val_col:
        df_name[n_col] = df_name[n_col].astype(str).str.strip()
        name_dict = df_name.drop_duplicates(n_col).set_index(n_col)[val_col].to_dict()
    else:
        name_dict = {}

    # B. 预处理基础信息表
    i_col = find_best_col(df_info, ['SKU货号', '商品识别码', '编码'])
    s_col = find_best_col(df_info, ['店铺名称'])
    l_col = find_best_col(df_info, ['回收标签'])
    
    if i_col:
        df_info[i_col] = df_info[i_col].astype(str).str.strip()
        # 转为字典，方便快速查询
        info_dict = df_info.drop_duplicates(i_col).set_index(i_col).to_dict('index')
    else:
        info_dict = {}

    uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

    if uploaded_file:
        results = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                # 提取仓库
                wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
                current_wh = wh_match.group(1) if wh_match else "未知"
                
                table = page.extract_table()
                if not table or len(table) < 2: continue
                
                # 清洗 PDF 表头
                headers = [str(h).strip().replace('\n', '') for h in table[0] if h]
                try:
                    sku_idx = next(i for i, h in enumerate(headers) if '货号' in h or '编码' in h)
                    info_idx = next(i for i, h in enumerate(headers) if '商品信息' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if '实际' in h or '发货数' in h)
                except: continue

                active_skc = "" 
                for row in table[1:]:
                    if not row[sku_idx] or "合计" in str(row): continue
                    
                    # 提取 SKC
                    skc_match = re.search(r"SKC[:：\s]+(\d+)", str(row[info_idx]))
                    if skc_match: active_skc = skc_match.group(1)
                    
                    sku = str(row[sku_idx]).strip().replace('\n', '')
                    
                    # 执行关联
                    res_name = name_dict.get(sku, "-")
                    res_shop = "-"
                    res_label = "-"
                    
                    if sku in info_dict:
                        res_shop = info_dict[sku].get(s_col, "-") if s_col else "-"
                        res_label = info_dict[sku].get(l_col, "-") if l_col else "-"

                    results.append({
                        "发货仓库": current_wh,
                        "店铺名称": res_shop,
                        "SKC ID": active_skc,
                        "回收标签类别": res_label,
                        "货品编码": sku,
                        "商品名称": res_name,
                        "发货数量": row[qty_idx]
                    })

        if results:
            df_res = pd.DataFrame(results)
            st.subheader("🔍 处理结果")
            st.dataframe(df_res, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False)
            st.download_button("📥 下载 Excel 结果", output.getvalue(), "提取结果.xlsx")
