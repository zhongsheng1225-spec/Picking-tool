import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-完整版", layout="wide")

st.title("📋 拣货单自动提取 (双表关联+自动体检)")

# --- 1. 基础资料智能加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

df_info = load_data("product_info.xlsx") # 基础信息表
df_name = load_data("name_map.xlsx")     # 名称对照表

with st.sidebar:
    st.header("⚙️ 资料体检状态")
    if df_info is not None: st.success("✅ product_info.xlsx 已就绪")
    else: st.error("❌ 缺失 product_info.xlsx")
    
    if df_name is not None: st.success("✅ name_map.xlsx 已就绪")
    else: st.error("❌ 缺失 name_map.xlsx")
    
    st.divider()
    st.info("💡 校验逻辑：\n1. [名称对照表] 找具体商品名\n2. [基础信息表] 找店铺和回收标签（优先使用 SKU ID 匹配）")

# --- 辅助函数：智能查找匹配列 ---
def get_match_col(df, keywords):
    for col in df.columns:
        if any(key in str(col) for key in keywords):
            return col
    return df.columns[0]

# --- 2. 处理 PDF 主逻辑 ---
uploaded_file = st.file_uploader("上传 PDF 拣货单", type="pdf")

if uploaded_file and df_info is not None and df_name is not None:
    results = []
    
    # --------------------------
    # 1. 准备基础信息表：以 SKU ID 为唯一键
    # --------------------------
    # 查找 SKU ID 列
    sku_id_col = None
    for col in df_info.columns:
        if 'SKU ID' in str(col):
            sku_id_col = col
            break
    if sku_id_col is None:
        st.error("❌ product_info.xlsx 中缺少 'SKU ID' 列，无法进行唯一匹配。请检查文件。")
        st.stop()
    
    # 统一转为字符串并去空格
    df_info[sku_id_col] = df_info[sku_id_col].astype(str).str.strip()
    # 使用 SKU ID 作为索引，理论上唯一，但为防止重复保留第一个
    info_dict = df_info.drop_duplicates(sku_id_col).set_index(sku_id_col).to_dict('index')
    
    # 额外保留 SKU 货号 -> 店铺信息的映射（用于降级匹配）
    # 查找 SKU 货号列（用于降级）
    sku_code_col = None
    for col in df_info.columns:
        if 'SKU货号' in str(col):
            sku_code_col = col
            break
    if sku_code_col is not None:
        df_info[sku_code_col] = df_info[sku_code_col].astype(str).str.strip()
        sku_code_dict = df_info.drop_duplicates(sku_code_col).set_index(sku_code_col).to_dict('index')
    else:
        sku_code_dict = {}
    
    # --------------------------
    # 2. 准备名称对照表（货号 -> 商品名）
    # --------------------------
    name_key = get_match_col(df_name, ['编码', 'SKU', '货号'])
    df_name[name_key] = df_name[name_key].astype(str).str.strip()
    name_dict = df_name.drop_duplicates(name_key).set_index(name_key).iloc[:, 0].to_dict()
    
    # --------------------------
    # 3. 解析 PDF
    # --------------------------
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            # 提取仓库
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            # 查找各列索引
            try:
                # 货号列（用于名称对照和降级匹配）
                sku_code_idx = next(i for i, h in enumerate(headers) if h and ('货号' in str(h) or '编码' in str(h)))
                # 商品信息列（含 SKC ID）
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in str(h))
                # 发货数量列
                qty_idx = next(i for i, h in enumerate(headers) if h and '发货数' in str(h))
                # 尝试查找 SKU ID 列
                sku_id_idx = None
                for i, h in enumerate(headers):
                    if h and ('SKU ID' in str(h) or 'SKUID' in str(h).replace(' ', '')):
                        sku_id_idx = i
                        break
            except StopIteration:
                continue
            
            active_skc = "" 
            for row in table[1:]:
                if not row[sku_code_idx] or "合计" in str(row):
                    continue
                
                # 提取 SKC ID（从商品信息单元格）
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                if skc_match:
                    active_skc = skc_match.group(1)
                
                # 获取 SKU 货号（用于名称对照和降级匹配）
                sku_code = str(row[sku_code_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()
                
                # 商品名称：从 name_dict 中获取
                prod_name = name_dict.get(sku_code, "-")
                
                # ---- 店铺和回收标签匹配：优先使用 SKU ID ----
                res_shop_name = "-"
                res_label = "-"
                matched = False
                
                if sku_id_idx is not None:
                    sku_id = str(row[sku_id_idx]).strip()
                    if sku_id and sku_id in info_dict:
                        res_shop_name = info_dict[sku_id].get('店铺名称', '-')
                        res_label = info_dict[sku_id].get('回收标签', '-')
                        matched = True
                
                # 降级：如果 PDF 中没有 SKU ID 或匹配失败，则使用 SKU 货号匹配
                if not matched and sku_code and sku_code in sku_code_dict:
                    res_shop_name = sku_code_dict[sku_code].get('店铺名称', '-')
                    res_label = sku_code_dict[sku_code].get('回收标签', '-')
                    matched = True
                
                # 如果还没有匹配到，尝试用 sku_code 在 info_dict 中查找（可能信息表里也有货号作为索引）
                if not matched and sku_code and sku_code in info_dict:
                    res_shop_name = info_dict[sku_code].get('店铺名称', '-')
                    res_label = info_dict[sku_code].get('回收标签', '-')
                
                results.append({
                    "发货仓库": current_wh,
                    "店铺名称": res_shop_name,
                    "SKC ID": active_skc,
                    "回收标签类别": res_label,
                    "货品编码": sku_code,
                    "商品名称": prod_name,
                    "发货数量": qty
                })
    
    # --- 3. 校验看板 ---
    if results:
        df_res = pd.DataFrame(results)
        
        st.subheader("🔍 自动体检看板")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("处理总行数", len(df_res))
        with col2:
            shops = df_res[df_res['店铺名称'] != '-']['店铺名称'].nunique()
            st.metric("涉及店铺数", shops)
        with col3:
            missing_name = len(df_res[df_res['商品名称'] == '-'])
            st.metric("名称未匹配", missing_name, delta_color="inverse")
        with col4:
            missing_info = len(df_res[df_res['店铺名称'] == '-'])
            st.metric("基础信息未匹配", missing_info, delta_color="inverse")

        if missing_name > 0 or missing_info > 0:
            st.warning("🚨 提示：部分货品未能在 Excel 中找到。请检查 `name_map.xlsx` 和 `product_info.xlsx` 是否已包含最新数据。")
        
        st.dataframe(df_res, use_container_width=True)
        
        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False)
        st.download_button("📥 下载校验后的 Excel", output.getvalue(), "拣货单结果.xlsx")
    else:
        st.warning("未从 PDF 中提取到有效数据，请检查文件格式。")
