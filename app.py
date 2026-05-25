"""
📊 Excel智能处理工具 - 加强版
author: AI Assistant
version: 2.0
"""
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os
import streamlit as st
import pandas as pd
from datetime import datetime
import os
# 导入功能模块
from pages.math_ops import render_math_ops
from pages.text_ops import render_text_ops
from pages.date_ops import render_date_ops
from pages.lookup_ops import render_lookup_ops
from pages.stats_ops import render_stats_ops
from pages.data_clean import render_data_clean
from pages.formula_engine import render_formula_engine
from utils.helpers import load_excel, df_to_excel, get_col_types, generate_filename


# # ======================== 页面配置 ========================
# st.set_page_config(
#     page_title="Excel智能处理工具 Pro",
#     page_icon="📊",
#     layout="wide",
#     initial_sidebar_state="expanded"
# )

# # 加载自定义样式
# css_path = os.path.join(os.path.dirname(__file__), "assets", "style.css")
# if os.path.exists(css_path):
#     with open(css_path) as f:
#         st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


# ======================== 页面配置 ========================
st.set_page_config(
    page_title="Excel智能处理工具 Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://github.com',
        'Report a bug': 'https://github.com',
        'About': "### Excel智能处理工具 Pro\n版本 2.0"
    }
)

# 加载自定义CSS
css_path = os.path.join(os.path.dirname(__file__), "assets", "style.css")
print(css_path)
if os.path.exists(css_path):
    with open(css_path, encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title("📊 Excel智能处理工具 Pro")
st.caption("强大、易用、无需安装的在线 Excel 函数处理平台")
st.markdown("---")



# ======================== 初始化Session State ========================
if 'df' not in st.session_state:
    st.session_state.df = None
if 'original_df' not in st.session_state:
    st.session_state.original_df = None
if 'operation_log' not in st.session_state:
    st.session_state.operation_log = []
if 'current_sheet' not in st.session_state:
    st.session_state.current_sheet = None
if 'sheets' not in st.session_state:
    st.session_state.sheets = {}


# ======================== 侧边栏 ========================
with st.sidebar:
    st.markdown("# 📊 Excel处理工具")
    st.markdown("### Pro 加强版 v2.0")
    st.markdown("---")
    
    # 文件上传
    st.markdown("### 📁 数据源")
    uploaded_file = st.file_uploader(
        "上传Excel文件",
        type=["xlsx", "xls", "csv"],
        help="支持 .xlsx / .xls / .csv 格式"
    )
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                encoding = st.selectbox("CSV编码", ["utf-8", "gbk", "gb2312", "utf-8-sig"],
                                        key="csv_encoding")
                st.session_state.sheets = {"Sheet1": pd.read_csv(uploaded_file, encoding=encoding)}
            else:
                st.session_state.sheets = load_excel(uploaded_file)
            
            # 选择sheet
            sheet_names = list(st.session_state.sheets.keys())
            selected_sheet = st.selectbox("选择工作表", sheet_names, key="sheet_select")
            
            if selected_sheet != st.session_state.current_sheet:
                st.session_state.current_sheet = selected_sheet
                st.session_state.df = st.session_state.sheets[selected_sheet].copy()
                st.session_state.original_df = st.session_state.sheets[selected_sheet].copy()
            
            # 数据概况
            if st.session_state.df is not None:
                df = st.session_state.df
                st.markdown("---")
                st.markdown("### 📋 数据概况")
                st.metric("行数", f"{len(df):,}")
                st.metric("列数", f"{len(df.columns)}")
                
                col_types = get_col_types(df)
                st.markdown(f"""
                - 🔢 数值列: **{len(col_types['numeric'])}**
                - 📝 文本列: **{len(col_types['text'])}**
                - 📅 日期列: **{len(col_types['datetime'])}**
                """)
                
                missing_total = df.isna().sum().sum()
                if missing_total > 0:
                    st.warning(f"⚠️ 缺失值: {missing_total} 个")
        
        except Exception as e:
            st.error(f"文件读取失败: {e}")
    
    st.markdown("---")
    
    # 功能导航
    st.markdown("### 🧭 功能导航")
    menu_options = [
        "🏠 数据总览",
        "🧹 数据清洗",
        "➕ 数学运算",
        "📝 文本处理",
        "📅 日期处理",
        "🔎 查找匹配",
        "📊 统计分析",
        "🧮 自定义公式",
    ]
    
    selected_menu = st.radio("选择功能", menu_options, key="menu", label_visibility="collapsed")
    
    st.markdown("---")
    
    # 撤销功能
    if st.button("↩️ 撤销所有修改（恢复原始数据）", use_container_width=True):
        if st.session_state.original_df is not None:
            st.session_state.df = st.session_state.original_df.copy()
            st.success("✅ 已恢复原始数据")
            st.rerun()
    
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; color:#666; font-size:12px;">
    Made with ❤️ by Excel Tool Pro<br>
    Powered by Streamlit + Pandas
    </div>
    """, unsafe_allow_html=True)


# ======================== 主区域 ========================
if st.session_state.df is None:
    # 欢迎页面
    st.markdown("""
    <div style="text-align:center; padding:60px 0;">
        <h1 style="font-size:3rem;">📊 Excel 智能处理工具 Pro</h1>
        <p style="font-size:1.2rem; color:#666; margin:20px 0;">
            告别手写Excel函数，所有操作点点就能完成
        </p>
        <div style="background:linear-gradient(135deg,#667eea,#764ba2); 
                    color:white; padding:30px; border-radius:16px; 
                    max-width:600px; margin:30px auto;">
            <h3>👈 请在左侧上传Excel文件开始</h3>
            <p>支持 .xlsx / .xls / .csv 格式</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 功能一览
    st.markdown("---")
    st.markdown("### ✨ 功能一览（20+常用Excel函数，点点即用）")
    
    func_list = {
        "➕ 数学运算": ["SUM求和", "AVERAGE平均", "MAX/MIN", "四则运算", "ROUND四舍五入", 
                       "ABS绝对值", "CUMSUM累计", "RANK排名", "百分比占比"],
        "📝 文本处理": ["CONCATENATE拼接", "LEFT/RIGHT/MID截取", "SUBSTITUTE替换",
                       "UPPER/LOWER大小写", "TRIM去空格", "LEN长度", "正则提取", "文本分列"],
        "📅 日期处理": ["YEAR/MONTH/DAY提取", "DATEDIF日期差", "日期偏移", "工作日判断", "文本转日期"],
        "🔎 查找匹配": ["IF条件判断", "IFS多条件", "VLOOKUP跨表查找", "COUNTIF计数", "去重标记"],
        "📊 统计分析": ["数据透视表", "SUMIF条件求和", "描述性统计", "数值分箱"],
        "🧹 数据清洗": ["缺失值处理", "类型转换", "列管理", "数据排序", "数据筛选"],
    }
    
    cols = st.columns(3)
    for idx, (category, funcs) in enumerate(func_list.items()):
        with cols[idx % 3]:
            st.markdown(f"**{category}**")
            for func in funcs:
                st.markdown(f"- ✅ {func}")
    
    st.stop()


# ======================== 功能路由 ========================
df = st.session_state.df

# 顶部信息栏
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📊 数据行数", f"{len(df):,}")
with col2:
    st.metric("📋 数据列数", f"{len(df.columns)}")
with col3:
    st.metric("🕳️ 缺失值", f"{df.isna().sum().sum():,}")
with col4:
    memory_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
    st.metric("💾 内存占用", f"{memory_mb:.1f} MB")

st.markdown("---")

# 根据菜单选择渲染不同功能
if selected_menu == "🏠 数据总览":
    st.subheader("📋 数据预览")
    
    # 可编辑数据表
    edited_df = st.data_editor(
        df, 
        use_container_width=True, 
        num_rows="dynamic",
        hide_index=False,
        height=500
    )
    st.session_state.df = edited_df
    df = edited_df
    
    # 列信息
    with st.expander("📊 列详细信息"):
        col_info = pd.DataFrame({
            '列名': df.columns,
            '数据类型': df.dtypes.astype(str).values,
            '非空数': df.notna().sum().values,
            '空值数': df.isna().sum().values,
            '唯一值数': df.nunique().values,
            '示例值': [str(df[col].dropna().iloc[0]) if df[col].notna().any() else 'N/A' 
                      for col in df.columns]
        })
        st.dataframe(col_info, use_container_width=True, hide_index=True)

elif selected_menu == "🧹 数据清洗":
    st.session_state.df = render_data_clean(df)

elif selected_menu == "➕ 数学运算":
    st.session_state.df = render_math_ops(df)

elif selected_menu == "📝 文本处理":
    st.session_state.df = render_text_ops(df)

elif selected_menu == "📅 日期处理":
    st.session_state.df = render_date_ops(df)

elif selected_menu == "🔎 查找匹配":
    st.session_state.df = render_lookup_ops(df)

elif selected_menu == "📊 统计分析":
    st.session_state.df = render_stats_ops(df)

elif selected_menu == "🧮 自定义公式":
    st.session_state.df = render_formula_engine(df)


# ======================== 底部：导出区域 ========================
st.markdown("---")
st.markdown("### 💾 导出数据")

c1, c2, c3 = st.columns([3, 2, 2])

with c1:
    output_name = st.text_input("文件名", value=generate_filename(), key="output_name")

with c2:
    export_format = st.selectbox("格式", ["Excel (.xlsx)", "CSV (.csv)"], key="export_fmt")

with c3:
    include_index = st.checkbox("包含行号", value=False, key="include_idx")

# 导出按钮
df_export = st.session_state.df

if export_format == "Excel (.xlsx)":
    # 检查是否有汇总表
    export_sheets = {"处理结果": df_export}
    if 'pivot_table' in st.session_state:
        export_sheets["汇总表"] = st.session_state.pivot_table
    if 'stats_table' in st.session_state:
        export_sheets["统计概览"] = st.session_state.stats_table
    
    excel_data = df_to_excel(export_sheets, index=include_index)
    
    st.download_button(
        label="⬇️ 下载 Excel 文件",
        data=excel_data,
        file_name=f"{output_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
else:
    csv_data = df_export.to_csv(index=include_index).encode('utf-8-sig')
    st.download_button(
        label="⬇️ 下载 CSV 文件",
        data=csv_data,
        file_name=f"{output_name}.csv",
        mime="text/csv",
        type="primary",
        use_container_width=True
    )

# 预览导出数据
with st.expander("👁️ 预览将要导出的数据"):
    st.dataframe(df_export, use_container_width=True, height=300)
    st.info(f"共 {len(df_export)} 行 × {len(df_export.columns)} 列")