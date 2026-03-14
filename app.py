import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单-雷达探测版", layout="wide")

st.title("📋 拣货单自动提取 (雷达排错版)")

# --- 0. 雷达探测：看看仓库里到底有什么 ---
with st.sidebar:
    st.header("📡 仓库文件雷达")
    st.write("当前 GitHub 里的真实文件列表：")
    current_files = os.listdir('.')
    for f in current_files:
        if not f.startswith('.'): # 隐藏系统自带文件
            st.code(f)
    
    st.divider()
    st.header("⚙️ 资料连接状态")

# --- 1. 智能抓取加载 ---
def load_data(keyword):
    # 扫描文件夹，只要文件名包含关键字就尝试读取
    for f in os.listdir('.'):
        if keyword in f and (f.endswith('.xlsx') or f.endswith('.csv')):
            try:
                if f.endswith('.csv'):
                    df = pd.read_csv(f, encoding='utf-8-sig')
                else:
                    df = pd.read_excel(f, engine='openpyxl')
                df.columns = [str(c).strip() for c in df.columns]
                return df
            except Exception as e:
                st.sidebar.error(f"读表报错: {e}")
                return None
    return None

df_info = load_data("product_info")
df_name = load_data("name_map")

# --- 2. 状态更新 ---
with st.sidebar:
    if df_info is not None: st.success("✅ 基础信息已找到")
    else: st.error("❌ 没找到 product_info 表格")
    
    if df_name is not None: st.success("✅ 名称对照已找到")
    else: st.error("❌ 没找到 name_map 表格")

# --- 3. 核心提取逻辑 ---
if df_info is not None and df_name is not None:
    # A. 对照表 (第一列是编码，第二列是名称)
    name_col = df_name.columns[0]
    val_col = df_name.columns[1] if len(df_name.columns) > 1 else df_name.columns[0]
    df_name[name_col] = df_name[name_col].astype(str).str.strip()
    name_dict = df_name.set_index(name_col)[val_col].to_dict()

    # B. 基础表 (寻找 SKU货号)
    info_dict = {}
    target_key = 'SKU货号'
    # 模糊寻找列名防报错
    if target_key not in df_info.columns:
        for c in df_info.columns:
            if '货号' in c or '识别码' in c or '编码' in c:
                target_key = c
                break
                
    if target_key in df_info.columns:
        df_info[target_key] = df_info[target_key].astype(str).str.strip()
        info_dict = df_info.drop_duplicates(target_key).set_index(target_key).to_dict('index')

    uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

    if uploaded_file:
        results = []
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
                current_wh = wh_match.group(1) if wh_match else "未知"
                
                table = page.extract_table()
                if not table or len(table) < 2: continue
                
                headers = [str(h).strip().replace('\n', '') for h in table[0] if h]
                try:
                    sku_idx = next(i for i, h in enumerate(headers) if '货号' in h or '编码' in h)
                    info_idx = next(i for i, h in enumerate(headers) if '商品信息' in h)
                    qty_idx = next(i for i, h in enumerate(headers) if '实际' in h or '发货数' in h)
                except: continue

                active_skc = "" 
                for row in table[1:]:
                    if not row[sku_idx] or "合计" in str(row): continue
                    
                    skc_match = re.search(r"SKC[:：\s]+(\d+)", str(row[info_idx]))
                    if skc_match: active_skc = skc_match.group(1)
                    
                    sku = str(row[sku_idx]).strip().replace('\n', '')
                    
                    # 关联匹配
                    res_name = name_dict.get(sku, "-")
                    res_shop = "-"
                    res_label = "-"
                    
                    if sku in info_dict:
                        res_shop = info_dict[sku].get('店铺名称', '-')
                        res_label = info_dict[sku].get('回收标签', '-')

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
            st.subheader("🔍 结果看板")
            st.dataframe(df_res, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_res.to_excel(writer, index=False)
            st.download_button("📥 下载 Excel 结果", output.getvalue(), "拣货单结果.xlsx")
