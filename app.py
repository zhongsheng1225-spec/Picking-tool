import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="拣货单增强工具-全兼容版", layout="wide")

st.title("📋 拣货单自动提取 (基础资料固定版)")
st.info("💡 提示：系统会自动识别仓库、关联商品名称与回收标签。找不到仓库时将显示‘未知’。")

# --- 核心：全兼容加载函数 ---
def load_local_data(file_name):
    if not os.path.exists(file_name):
        return None
    
    # 方案1：尝试作为 Excel 读取 (兼容伪装成 csv 的 xlsx)
    try:
        return pd.read_excel(file_name)
    except:
        pass
    
    # 方案2：尝试作为 CSV 读取 (先尝试 UTF-8 编码)
    try:
        return pd.read_csv(file_name, encoding='utf-8-sig')
    except:
        pass
    
    # 方案3：尝试 GBK 编码 (处理部分 Excel 直接导出的 CSV)
    try:
        return pd.read_csv(file_name, encoding='gbk')
    except Exception as e:
        st.error(f"无法读取文件 {file_name}: {e}")
        return None

# 预加载基础数据
df_prod = load_local_data("product_info.csv")
df_label = load_local_data("label_info.csv")

# 侧边栏状态检查
with st.sidebar:
    st.header("⚙️ 基础资料状态")
    if df_prod is not None:
        st.success(f"✅ 商品信息已加载 ({len(df_prod)}条)")
    else:
        st.error("❌ 缺失 product_info.csv")
        
    if df_label is not None:
        st.success(f"✅ 标签信息已加载 ({len(df_label)}条)")
    else:
        st.error("❌ 缺失 label_info.csv")
    
    if st.button("刷新数据"):
        st.rerun()

# --- 主程序：处理 PDF ---
uploaded_pdf = st.file_uploader("点击或拖拽上传 PDF 拣货单", type="pdf")

if uploaded_pdf is not None:
    results = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # 1. 提取当前页仓库 (找不到即为未知，不沿用上一页)
            wh_match = re.search(r"收货仓[:：]\s*([^\s\n]+)", text)
            current_wh = wh_match.group(1) if wh_match else "未知"
            
            table = page.extract_table()
            if table:
                headers = table[0]
                try:
                    # 获取关键列的动态索引
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

                        # 2. 关联商品名称 (来自 product_info.csv)
                        prod_name = "-"
                        if df_prod is not None:
                            # 确保商品编码作为字符串进行匹配
                            m = df_prod[df_prod['商品编码'].astype(str) == sku]
                            if not m.empty:
                                prod_name = m.iloc[0]['商品名称']

                        # 3. 关联回收标签 (来自 label_info.csv)
                        label_type = "-"
                        if df_label is not None and skc_id:
                            # 确保 SKC ID 作为字符串进行匹配
                            m = df_label[df_label['SKC ID'].astype(str) == skc_id]
                            if not m.empty:
                                label_type = m.iloc[0]['回收标签']

                        results.append({
                            "发货仓库": current_wh,
                            "SKC ID": skc_id,
                            "回收标签类别": label_type,
                            "货品编码": sku,
                            "商品名称": prod_name,
                            "发货数量": qty
                        })
                except Exception:
                    continue

    if results:
        df_final = pd.DataFrame(results)
        st.success(f"处理成功！提取到 {len(df_final)} 条数据。")
        
        # 结果预览
        st.dataframe(df_final, use_container_width=True)

        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='拣货导出明细')
        
        st.download_button(
            label="📥 下载结果 (Excel格式)",
            data=output.getvalue(),
            file_name="拣货单增强处理结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
