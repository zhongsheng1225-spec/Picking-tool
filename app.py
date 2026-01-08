import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-最终整合版", layout="wide")

st.title("📋 拣货单自动化提取与校验——福星高照")

# --- 1. 基础资料智能加载 ---
def load_data(name):
    if os.path.exists(name):
        try:
            return pd.read_excel(name)
        except:
            return None
    return None

df_prod = load_data("product_info.xlsx")
df_label = load_data("label_info.xlsx")

# 侧边栏状态监测
with st.sidebar:
    st.header("⚙️ 基础资料状态")
    if df_prod is not None: st.success("✅ 商品信息已就绪")
    else: st.error("❌ 缺失 product_info.xlsx")
    
    if df_label is not None: st.success("✅ 标签信息已就绪")
    else: st.error("❌ 缺失 label_info.xlsx")
    
    st.divider()
    st.info("💡 校验逻辑说明：\n1. 仓库：每页重新抓取表头仓库。\n2. SKC：向下关联，直到遇到新SKC。")

# --- 2. 处理 PDF 主逻辑 ---
uploaded_file = st.file_uploader("上传 PDF 拣货单文件", type="pdf")

if uploaded_file and df_prod is not None and df_label is not None:
    results = []
    
    # 数据清洗：统一转为字符串并去空格，防止匹配失败
    df_prod['商品编码'] = df_prod['商品编码'].astype(str).str.strip()
    df_label['SKC ID'] = df_label['SKC ID'].astype(str).str.strip()

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # A. 提取本页仓库 (精准定位表头)
            wh_match = re.search(r"(?:收货仓|仓库)[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            # B. 提取表格
            table = page.extract_table()
            if not table: continue
            
            headers = table[0]
            try:
                sku_idx = next(i for i, h in enumerate(headers) if h and 'SKU货号' in h)
                qty_idx = next(i for i, h in enumerate(headers) if h and '实际发货数' in h)
                info_idx = next(i for i, h in enumerate(headers) if h and '商品信息' in h)
            except: continue

            active_skc = "" # 本页当前的活跃 SKC 指针
            
            for row in table[1:]:
                if not row[sku_idx] or "合计" in str(row): continue
                
                # C. SKC 向下补全逻辑
                cell_info = str(row[info_idx])
                skc_match = re.search(r"SKC[:：\s]+(\d+)", cell_info)
                if skc_match:
                    active_skc = skc_match.group(1)
                
                sku = str(row[sku_idx]).strip().replace('\n', '')
                qty = str(row[qty_idx]).strip()

                # D. 关联匹配 (VLOOKUP)
                p_name = "-"
                p_match = df_prod[df_prod['商品编码'] == sku]
                if not p_match.empty:
                    p_name = p_match.iloc[0]['商品名称']

                l_type = "-"
                if active_skc:
                    l_match = df_label[df_label['SKC ID'] == active_skc]
                    if not l_match.empty:
                        l_type = l_match.iloc[0]['回收标签']

                results.append({
                    "发货仓库": current_wh,
                    "SKC ID": active_skc,
                    "回收标签类别": l_type,
                    "货品编码": sku,
                    "商品名称": p_name,
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
            wh_list = df_res['发货仓库'].unique()
            st.metric("识别仓库数", len(wh_list), help=f"识别到的仓库：{', '.join(wh_list)}")
        with m3:
            # 统计未匹配到结果的比例
            fail_match = len(df_res[df_res['商品名称'] == '-'])
            st.metric("匹配失败数", fail_match, delta_color="inverse")

        # 异常提醒
        if fail_match > 0:
            st.warning(f"🚨 提示：有 {fail_match} 行货品未能匹配到商品名称，请检查基础资料表是否完整。")
        if "未知" in wh_list:
            st.error("🚨 警告：部分页面的仓库未能正确识别，请手动核对“未知”行。")

        st.dataframe(df_res, use_container_width=True)
        
        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_res.to_excel(writer, index=False, sheet_name='提取结果')
        
        st.download_button(
            label="📥 下载 Excel 结果报表",
            data=output.getvalue(),
            file_name="拣货单增强结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
