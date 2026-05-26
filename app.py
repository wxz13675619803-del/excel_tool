"""
📊 Excel智能处理工具 Pro - 终极优化版
单页面架构 + 缓存优化 + 分页加载
"""
import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO
from datetime import datetime

from utils.helpers import (
    load_excel_cached, optimize_dtypes, df_to_excel_optimized,
    get_col_types, paginate_dataframe, generate_filename
)

# ======================== 页面配置 ========================
st.set_page_config(
    page_title="Excel智能处理工具 Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 加载CSS
css_path = os.path.join(os.path.dirname(__file__), "assets", "style.css")
if os.path.exists(css_path):
    with open(css_path, encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# ======================== Session State ========================
defaults = {
    'df': None, 'original_df': None, 'sheets': {},
    'current_sheet': None, 'history': [],
    'page_num': 1, 'page_size': 500,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def save_snapshot():
    """保存操作快照（用于撤销）"""
    if st.session_state.df is not None:
        st.session_state.history.append(st.session_state.df.copy())
        if len(st.session_state.history) > 20:
            st.session_state.history.pop(0)


def undo():
    """撤销上一步"""
    if st.session_state.history:
        st.session_state.df = st.session_state.history.pop()
        st.rerun()


# ======================== 侧边栏 ========================
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:10px 0 20px 0;">
        <div style="font-size:2.5rem;">📊</div>
        <div style="font-size:1.1rem; font-weight:700; color:#a78bfa;">Excel 处理工具</div>
        <div style="font-size:0.75rem; color:#6b7280; margin-top:4px;">Pro v3.0 · 终极版</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 文件上传
    uploaded_file = st.file_uploader(
        "📁 上传数据文件",
        type=["xlsx", "xls", "csv"],
        help="支持 Excel(.xlsx/.xls) 和 CSV 文件，最大500MB"
    )
    
    if uploaded_file:
        file_bytes = uploaded_file.getvalue()
        
        with st.spinner("📖 正在读取文件..."):
            try:
                st.session_state.sheets = load_excel_cached(file_bytes, uploaded_file.name)
                
                sheet_names = list(st.session_state.sheets.keys())
                selected_sheet = st.selectbox("📄 选择工作表", sheet_names)
                
                if selected_sheet != st.session_state.current_sheet:
                    st.session_state.current_sheet = selected_sheet
                    raw_df = st.session_state.sheets[selected_sheet].copy()
                    
                    # 大数据自动优化内存
                    if len(raw_df) > 5000:
                        raw_df = optimize_dtypes(raw_df)
                    
                    st.session_state.df = raw_df
                    st.session_state.original_df = raw_df.copy()
                    st.session_state.history = []
            except Exception as e:
                st.error(f"❌ 文件读取失败: {e}")
    
    if st.session_state.df is not None:
        df = st.session_state.df
        col_types = get_col_types(df)
        
        st.markdown("---")
        st.markdown("##### 📋 数据概况")
        
        c1, c2 = st.columns(2)
        with c1:
            st.metric("行", f"{len(df):,}")
        with c2:
            st.metric("列", f"{len(df.columns)}")
        
        mem_mb = df.memory_usage(deep=True).sum() / 1024**2
        st.caption(f"💾 内存: {mem_mb:.1f}MB | 🔢 数值列: {len(col_types['numeric'])} | 📝 文本列: {len(col_types['text'])}")
        
        missing = df.isna().sum().sum()
        if missing > 0:
            st.warning(f"⚠️ {missing} 个缺失值")
    
    st.markdown("---")
    
    # 功能菜单
    st.markdown("##### 🧭 功能菜单")
    menu = st.radio(
        "nav", 
        ["🏠 数据总览", "🧹 数据清洗", "➕ 数学运算", "📝 文本处理",
         "📅 日期处理", "🔎 条件与查找", "📊 统计汇总", "🧮 公式引擎"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    # 操作按钮
    c1, c2 = st.columns(2)
    with c1:
        if st.button("↩️ 撤销", use_container_width=True, disabled=len(st.session_state.history)==0):
            undo()
    with c2:
        if st.button("🔄 重置", use_container_width=True):
            if st.session_state.original_df is not None:
                st.session_state.df = st.session_state.original_df.copy()
                st.session_state.history = []
                st.rerun()


# ======================== 欢迎页 ========================
if st.session_state.df is None:
    st.markdown("""
    <div class="welcome-card" style="text-align:center; padding:40px 0;">
        <h1 style="font-size:2.8rem; margin-bottom:8px;">
            <span class="gradient-text">Excel 智能处理工具</span>
        </h1>
        <p style="font-size:1.15rem; color:#8e8ea0; max-width:600px; margin:0 auto 30px auto;">
            告别手写函数的痛苦，30+ 常用功能点点即用<br>
            支持大文件处理，无需安装任何软件
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    features = [
        ("➕", "数学运算", "SUM / AVERAGE / MAX / MIN<br>排名 / 占比 / 累计 / 四则运算"),
        ("📝", "文本处理", "拼接 / 截取 / 替换 / 分列<br>提取数字/手机号 / 大小写"),
        ("📅", "日期处理", "提取年月日 / 日期差<br>偏移 / 工作日判断"),
        ("🔎", "条件查找", "IF判断 / 多条件IFS<br>VLOOKUP / COUNTIF / 去重"),
        ("📊", "统计汇总", "数据透视表 / SUMIF<br>分箱 / 描述统计"),
        ("🧮", "公式引擎", "自定义公式 / 模板库<br>条件计算器"),
    ]
    
    cols = st.columns(3)
    for i, (icon, title, desc) in enumerate(features):
        with cols[i % 3]:
            st.markdown(f"""
            <div class="feature-card">
                <div class="feature-icon">{icon}</div>
                <div class="feature-title">{title}</div>
                <div class="feature-desc">{desc}</div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("")
    
    st.markdown("---")
    st.info("👈 请在左侧上传 Excel / CSV 文件开始使用")
    st.stop()


# ======================== 主区域（数据已加载）========================
df = st.session_state.df

# 顶部指标
m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("📊 总行数", f"{len(df):,}")
with m2:
    st.metric("📋 总列数", f"{len(df.columns)}")
with m3:
    st.metric("🔢 数值列", f"{len(get_col_types(df)['numeric'])}")
with m4:
    st.metric("📝 文本列", f"{len(get_col_types(df)['text'])}")
with m5:
    st.metric("🕳️ 缺失值", f"{df.isna().sum().sum():,}")

st.markdown("---")

numeric_cols = get_col_types(df)['numeric']
text_cols = get_col_types(df)['text']
all_cols = df.columns.tolist()


# ================================================================
#                        🏠 数据总览
# ================================================================
if menu == "🏠 数据总览":
    st.subheader("🏠 数据总览与编辑")
    
    # 分页控制
    total_rows = len(df)
    page_size = st.select_slider("每页显示行数", [100, 200, 500, 1000, 2000, 5000], value=500)
    
    page_df, total_pages = paginate_dataframe(df, page_size, st.session_state.page_num)
    
    if total_pages > 1:
        col1, col2, col3 = st.columns([1, 3, 1])
        with col2:
            st.session_state.page_num = st.number_input(
                f"页码（共{total_pages}页，{total_rows:,}行）",
                min_value=1, max_value=total_pages,
                value=st.session_state.page_num
            )
    
    # 可编辑表格
    edited = st.data_editor(page_df, use_container_width=True, num_rows="dynamic", height=500)
    
    # 将编辑结果写回
    if not edited.equals(page_df):
        save_snapshot()
        start = (st.session_state.page_num - 1) * page_size
        for col in edited.columns:
            df.loc[edited.index, col] = edited[col]
        st.session_state.df = df
    
    # 列信息
    with st.expander("📊 列详细信息"):
        info_df = pd.DataFrame({
            '列名': df.columns,
            '类型': df.dtypes.astype(str).values,
            '非空': df.notna().sum().values,
            '缺失': df.isna().sum().values,
            '唯一值': df.nunique().values,
        })
        st.dataframe(info_df, use_container_width=True, hide_index=True)


# ================================================================
#                        🧹 数据清洗
# ================================================================
elif menu == "🧹 数据清洗":
    st.subheader("🧹 数据清洗")
    
    t1, t2, t3, t4, t5 = st.tabs(["缺失值处理", "类型转换", "列管理", "排序筛选", "异常值处理"])
    
    # ---------- 缺失值 ----------
    with t1:
        missing = df.isna().sum()
        miss_df = pd.DataFrame({'列名': missing.index, '缺失数': missing.values, 
                                '缺失率%': (missing/len(df)*100).round(1).values})
        miss_df = miss_df[miss_df['缺失数'] > 0]
        
        if len(miss_df) > 0:
            st.dataframe(miss_df, use_container_width=True, hide_index=True)
            
            c1, c2 = st.columns(2)
            with c1:
                fill_col = st.selectbox("选择列", miss_df['列名'].tolist(), key="fc")
            with c2:
                fill_method = st.selectbox("处理方式", [
                    "删除缺失行", "填充 0", "填充均值", "填充中位数", "填充众数",
                    "向下填充", "向上填充", "填充固定值", "填充空字符串"
                ], key="fm")
            
            fill_val = ""
            if fill_method == "填充固定值":
                fill_val = st.text_input("固定值", key="fv")
            
            if st.button("✅ 执行", key="bf", type="primary"):
                save_snapshot()
                if fill_method == "删除缺失行":
                    before = len(df)
                    df = df.dropna(subset=[fill_col])
                    st.success(f"✅ 删除 {before-len(df)} 行")
                elif fill_method == "填充 0":
                    df[fill_col] = df[fill_col].fillna(0)
                elif fill_method == "填充均值":
                    df[fill_col] = df[fill_col].fillna(df[fill_col].mean())
                elif fill_method == "填充中位数":
                    df[fill_col] = df[fill_col].fillna(df[fill_col].median())
                elif fill_method == "填充众数":
                    m = df[fill_col].mode()
                    if len(m): df[fill_col] = df[fill_col].fillna(m[0])
                elif fill_method == "向下填充":
                    df[fill_col] = df[fill_col].ffill()
                elif fill_method == "向上填充":
                    df[fill_col] = df[fill_col].bfill()
                elif fill_method == "填充固定值":
                    df[fill_col] = df[fill_col].fillna(fill_val)
                else:
                    df[fill_col] = df[fill_col].fillna('')
                st.session_state.df = df
                st.success("✅ 处理完成")
        else:
            st.success("🎉 无缺失值")
    
    # ---------- 类型转换 ----------
    with t2:
        c1, c2 = st.columns(2)
        with c1:
            tc = st.selectbox("选择列", all_cols, key="tc")
            st.info(f"当前类型: `{df[tc].dtype}`")
        with c2:
            tt = st.selectbox("转换为", ["文本str", "整数int", "浮点数float", "日期datetime", "分类category"], key="tt")
        
        if st.button("✅ 转换", key="bt", type="primary"):
            save_snapshot()
            try:
                m = {"文本str": lambda: df[tc].astype(str),
                     "整数int": lambda: pd.to_numeric(df[tc], errors='coerce').astype('Int64'),
                     "浮点数float": lambda: pd.to_numeric(df[tc], errors='coerce'),
                     "日期datetime": lambda: pd.to_datetime(df[tc], errors='coerce'),
                     "分类category": lambda: df[tc].astype('category')}
                df[tc] = m[tt]()
                st.session_state.df = df
                st.success(f"✅ 已转换为 {tt}")
            except Exception as e:
                st.error(f"转换失败: {e}")
    
    # ---------- 列管理 ----------
    with t3:
        op = st.radio("操作", ["重命名列", "删除列", "调整列顺序", "新增空列"], horizontal=True, key="cm")
        
        if op == "重命名列":
            rc = st.selectbox("选择列", df.columns.tolist(), key="rc")
            nn = st.text_input("新名称", key="nn")
            if st.button("✅ 重命名", key="brn") and nn:
                save_snapshot()
                df = df.rename(columns={rc: nn})
                st.session_state.df = df
                st.success(f"✅ {rc} → {nn}")
                st.rerun()
        
        elif op == "删除列":
            dc = st.multiselect("选择要删除的列", df.columns.tolist(), key="dc")
            if st.button("✅ 删除", key="bdc", type="primary") and dc:
                save_snapshot()
                df = df.drop(columns=dc)
                st.session_state.df = df
                st.success(f"✅ 已删除 {len(dc)} 列")
                st.rerun()
        
        elif op == "调整列顺序":
            new_order = st.multiselect("按顺序选择所有列", df.columns.tolist(), default=df.columns.tolist(), key="co")
            if st.button("✅ 应用", key="bco") and len(new_order) == len(df.columns):
                save_snapshot()
                df = df[new_order]
                st.session_state.df = df
                st.success("✅ 列顺序已调整")
        
        else:
            nc = st.text_input("新列名", key="nc")
            nv = st.text_input("默认值（留空为空值）", key="nv")
            if st.button("✅ 添加", key="bnc") and nc:
                save_snapshot()
                df[nc] = nv if nv else np.nan
                st.session_state.df = df
                st.success(f"✅ 已添加列 [{nc}]")
    
    # ---------- 排序筛选 ----------
    with t4:
        st.markdown("**排序**")
        sc = st.multiselect("排序列", df.columns.tolist(), key="sc")
        if sc:
            orders = []
            cols_row = st.columns(len(sc))
            for i, col in enumerate(sc):
                with cols_row[i]:
                    o = st.radio(f"{col}", ["升序↑", "降序↓"], horizontal=True, key=f"so_{col}")
                    orders.append(o == "升序↑")
            
            if st.button("✅ 排序", key="bs", type="primary"):
                save_snapshot()
                df = df.sort_values(by=sc, ascending=orders).reset_index(drop=True)
                st.session_state.df = df
                st.success("✅ 排序完成")
        
        st.markdown("---")
        st.markdown("**筛选**")
        fc2 = st.selectbox("筛选列", df.columns.tolist(), key="fc2")
        
        if df[fc2].dtype in ['object', 'category']:
            uv = df[fc2].dropna().unique().tolist()
            sv = st.multiselect("保留的值", uv, default=uv, key="sv")
            if st.button("✅ 筛选", key="bf2", type="primary"):
                save_snapshot()
                df = df[df[fc2].isin(sv)].reset_index(drop=True)
                st.session_state.df = df
                st.success(f"✅ 保留 {len(df)} 行")
        else:
            c1, c2 = st.columns(2)
            with c1:
                mn = st.number_input("最小值", value=float(df[fc2].min()), key="mn")
            with c2:
                mx = st.number_input("最大值", value=float(df[fc2].max()), key="mx")
            if st.button("✅ 筛选", key="bf3", type="primary"):
                save_snapshot()
                df = df[(df[fc2]>=mn)&(df[fc2]<=mx)].reset_index(drop=True)
                st.session_state.df = df
                st.success(f"✅ 保留 {len(df)} 行")
    
    # ---------- 异常值处理 ----------
    with t5:
        if numeric_cols:
            oc = st.selectbox("选择数值列", numeric_cols, key="oc")
            om = st.radio("检测方法", ["IQR四分位法", "Z-Score标准差法", "固定范围"], horizontal=True, key="om")
            
            if om == "IQR四分位法":
                q1, q3 = df[oc].quantile(0.25), df[oc].quantile(0.75)
                iqr = q3 - q1
                lower, upper = q1 - 1.5*iqr, q3 + 1.5*iqr
                st.info(f"正常范围: [{lower:.2f}, {upper:.2f}]")
                outliers = df[(df[oc]<lower)|(df[oc]>upper)]
            elif om == "Z-Score标准差法":
                threshold = st.slider("Z-Score阈值", 1.0, 5.0, 3.0, key="zs")
                mean, std = df[oc].mean(), df[oc].std()
                z_scores = ((df[oc]-mean)/std).abs()
                outliers = df[z_scores > threshold]
                lower, upper = mean - threshold*std, mean + threshold*std
                st.info(f"正常范围: [{lower:.2f}, {upper:.2f}]")
            else:
                c1, c2 = st.columns(2)
                with c1: lower = st.number_input("下限", value=float(df[oc].min()), key="ol")
                with c2: upper = st.number_input("上限", value=float(df[oc].max()), key="ou")
                outliers = df[(df[oc]<lower)|(df[oc]>upper)]
            
            st.warning(f"发现 {len(outliers)} 个异常值")
            
            oa = st.radio("处理方式", ["标记异常值", "删除异常行", "替换为边界值(Winsorize)", "替换为空值"], 
                         horizontal=True, key="oa")
            
            if st.button("✅ 处理异常值", key="boa", type="primary"):
                save_snapshot()
                mask = (df[oc]<lower)|(df[oc]>upper)
                if oa == "标记异常值":
                    df[f"{oc}_异常"] = mask.map({True: "异常", False: "正常"})
                elif oa == "删除异常行":
                    df = df[~mask].reset_index(drop=True)
                elif oa == "替换为边界值(Winsorize)":
                    df[oc] = df[oc].clip(lower=lower, upper=upper)
                else:
                    df.loc[mask, oc] = np.nan
                st.session_state.df = df
                st.success("✅ 处理完成")
        else:
            st.warning("无数值列")


# ================================================================
#                        ➕ 数学运算
# ================================================================
elif menu == "➕ 数学运算":
    st.subheader("➕ 数学运算")
    
    if not numeric_cols:
        st.warning("⚠️ 当前数据没有数值列")
        st.stop()
    
    t1, t2, t3, t4 = st.tabs(["基础运算", "高级运算", "行列统计", "数值转换"])
    
    # ---------- 基础运算 ----------
    with t1:
        with st.expander("➕ 多列求和（SUM）", expanded=True):
            cs = st.multiselect("选择列", numeric_cols, key="sum_c")
            sn = st.text_input("新列名", value="合计", key="sum_n")
            if st.button("✅ 求和", key="b_sum", type="primary") and cs:
                save_snapshot()
                df[sn] = df[cs].sum(axis=1)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{sn}]")
        
        with st.expander("📊 多列平均值（AVERAGE）"):
            ca = st.multiselect("选择列", numeric_cols, key="avg_c")
            an = st.text_input("新列名", value="平均值", key="avg_n")
            if st.button("✅ 平均", key="b_avg", type="primary") and ca:
                save_snapshot()
                df[an] = df[ca].mean(axis=1)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{an}]")
        
        with st.expander("⬆️⬇️ 最大/最小值（MAX / MIN）"):
            cm = st.multiselect("选择列", numeric_cols, key="mm_c")
            mm_type = st.radio("类型", ["最大值MAX", "最小值MIN"], horizontal=True, key="mm_t")
            mn2 = st.text_input("新列名", value="最大值" if "MAX" in mm_type else "最小值", key="mm_n")
            if st.button("✅ 执行", key="b_mm", type="primary") and cm:
                save_snapshot()
                df[mn2] = df[cm].max(axis=1) if "MAX" in mm_type else df[cm].min(axis=1)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{mn2}]")
        
        with st.expander("🔢 两列四则运算"):
            c1, c2, c3 = st.columns(3)
            with c1: col_a = st.selectbox("列A", numeric_cols, key="a_a")
            with c2: op = st.selectbox("运算符", ["+", "-", "×", "÷", "取余%", "幂^"], key="a_op")
            with c3: col_b = st.selectbox("列B", numeric_cols, key="a_b")
            ar_n = st.text_input("新列名", value="运算结果", key="a_n")
            
            if st.button("✅ 运算", key="b_ar", type="primary"):
                save_snapshot()
                ops = {"+": df[col_a]+df[col_b], "-": df[col_a]-df[col_b],
                       "×": df[col_a]*df[col_b], "÷": df[col_a]/df[col_b].replace(0,np.nan),
                       "取余%": df[col_a]%df[col_b].replace(0,np.nan), "幂^": df[col_a]**df[col_b]}
                df[ar_n] = ops[op]
                st.session_state.df = df
                st.success(f"✅ {col_a} {op} {col_b} → [{ar_n}]")
    
    # ---------- 高级运算 ----------
    with t2:
        with st.expander("🔄 四舍五入（ROUND / FLOOR / CEIL）"):
            cr = st.selectbox("选择列", numeric_cols, key="r_c")
            rt = st.radio("类型", ["四舍五入ROUND", "向下取整FLOOR", "向上取整CEIL", "截断TRUNC"], horizontal=True, key="r_t")
            if "ROUND" in rt:
                dec = st.number_input("小数位数", 0, 10, 2, key="r_d")
            mode = st.radio("模式", ["覆盖原列", "生成新列"], horizontal=True, key="r_m")
            
            if st.button("✅ 执行", key="b_r", type="primary"):
                save_snapshot()
                if "ROUND" in rt: result = df[cr].round(dec)
                elif "FLOOR" in rt: result = np.floor(df[cr])
                elif "CEIL" in rt: result = np.ceil(df[cr])
                else: result = np.trunc(df[cr])
                
                target = cr if mode=="覆盖原列" else f"{cr}_{rt.split('(')[0].strip()}"
                df[target] = result
                st.session_state.df = df
                st.success(f"✅ 完成")
        
        with st.expander("| | 绝对值（ABS）"):
            cab = st.selectbox("选择列", numeric_cols, key="ab_c")
            if st.button("✅ 绝对值", key="b_ab", type="primary"):
                save_snapshot()
                df[f"{cab}_abs"] = df[cab].abs()
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("📐 数学函数（SQRT / LOG / EXP / POWER）"):
            cmf = st.selectbox("选择列", numeric_cols, key="mf_c")
            mf = st.selectbox("函数", ["平方根SQRT", "自然对数LN", "常用对数LOG10", 
                                        "指数EXP", "平方", "立方"], key="mf_f")
            if st.button("✅ 执行", key="b_mf", type="primary"):
                save_snapshot()
                fm = {"平方根SQRT": np.sqrt(df[cmf].abs()), "自然对数LN": np.log(df[cmf].clip(lower=0.001)),
                      "常用对数LOG10": np.log10(df[cmf].clip(lower=0.001)), "指数EXP": np.exp(df[cmf].clip(upper=500)),
                      "平方": df[cmf]**2, "立方": df[cmf]**3}
                df[f"{cmf}_{mf}"] = fm[mf]
                st.session_state.df = df
                st.success("✅ 完成")
    
    # ---------- 行列统计 ----------
    with t3:
        with st.expander("📈 累计求和（CUMSUM / Running Total）"):
            ccs = st.selectbox("选择列", numeric_cols, key="cs_c")
            if st.button("✅ 累计", key="b_cs", type="primary"):
                save_snapshot()
                df[f"{ccs}_累计"] = df[ccs].cumsum()
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("🏆 排名（RANK）"):
            crk = st.selectbox("选择列", numeric_cols, key="rk_c")
            rko = st.radio("排序", ["降序（最大=第1）", "升序（最小=第1）"], horizontal=True, key="rk_o")
            if st.button("✅ 排名", key="b_rk", type="primary"):
                save_snapshot()
                asc = "升序" in rko
                df[f"{crk}_排名"] = df[crk].rank(ascending=asc, method='min').astype(int)
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("📐 占比 / 百分比"):
            cp = st.selectbox("选择列", numeric_cols, key="p_c")
            pf = st.radio("格式", ["小数(0.25)", "百分比(25.00%)"], horizontal=True, key="p_f")
            if st.button("✅ 占比", key="b_p", type="primary"):
                save_snapshot()
                total = df[cp].sum()
                if total != 0:
                    if "小数" in pf:
                        df[f"{cp}_占比"] = (df[cp]/total).round(4)
                    else:
                        df[f"{cp}_占比%"] = ((df[cp]/total)*100).round(2).astype(str) + '%'
                    st.session_state.df = df
                    st.success("✅ 完成")
                else:
                    st.error("总和为0")
        
        with st.expander("📉 同比/环比增长率"):
            cg = st.selectbox("选择列", numeric_cols, key="g_c")
            gp = st.number_input("对比间隔行数", 1, 365, 1, key="g_p", help="1=环比，12=同比(月数据)")
            if st.button("✅ 计算增长率", key="b_g", type="primary"):
                save_snapshot()
                df[f"{cg}_增长率%"] = (df[cg].pct_change(periods=gp) * 100).round(2)
                st.session_state.df = df
                st.success("✅ 完成")
    
    # ---------- 数值转换 ----------
    with t4:
        with st.expander("💰 金额大写转换"):
            ck = st.selectbox("选择金额列", numeric_cols, key="k_c")
            if st.button("✅ 转换", key="b_k", type="primary"):
                save_snapshot()
                def to_rmb(n):
                    try:
                        n = round(float(n), 2)
                        units = ['', '拾', '佰', '仟', '万', '拾', '佰', '仟', '亿']
                        digits = '零壹贰叁肆伍陆柒捌玖'
                        s = str(int(abs(n)*100))
                        result = ''
                        for i, d in enumerate(reversed(s)):
                            if i == 0: result = digits[int(d)] + '分' + result if int(d) else result
                            elif i == 1: result = digits[int(d)] + '角' + result if int(d) else result
                            else:
                                idx = i - 2
                                if idx < len(units):
                                    result = digits[int(d)] + units[idx] + result
                        return ('负' if n < 0 else '') + (result or '零') + '整'
                    except:
                        return str(n)
                df[f"{ck}_大写"] = df[ck].apply(to_rmb)
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("🔀 单位换算"):
            cu = st.selectbox("选择列", numeric_cols, key="u_c")
            ut = st.selectbox("换算类型", [
                "元→万元(÷10000)", "万元→元(×10000)", "元→亿元(÷100000000)",
                "kg→吨(÷1000)", "吨→kg(×1000)",
                "米→千米(÷1000)", "千米→米(×1000)",
                "摄氏度→华氏度", "华氏度→摄氏度"
            ], key="u_t")
            
            if st.button("✅ 换算", key="b_u", type="primary"):
                save_snapshot()
                um = {
                    "元→万元(÷10000)": df[cu]/10000, "万元→元(×10000)": df[cu]*10000,
                    "元→亿元(÷100000000)": df[cu]/100000000,
                    "kg→吨(÷1000)": df[cu]/1000, "吨→kg(×1000)": df[cu]*1000,
                    "米→千米(÷1000)": df[cu]/1000, "千米→米(×1000)": df[cu]*1000,
                    "摄氏度→华氏度": df[cu]*9/5+32, "华氏度→摄氏度": (df[cu]-32)*5/9,
                }
                df[f"{cu}_{ut.split('(')[0]}"] = um[ut].round(4)
                st.session_state.df = df
                st.success("✅ 完成")


# ================================================================
#                        📝 文本处理
# ================================================================
elif menu == "📝 文本处理":
    st.subheader("📝 文本处理")
    
    t1, t2, t3, t4 = st.tabs(["拼接截取", "替换清理", "提取转换", "分列编码"])
    
    with t1:
        with st.expander("🔗 文本拼接（CONCATENATE）", expanded=True):
            cc = st.multiselect("选择要拼接的列", all_cols, key="ct_c")
            sep = st.text_input("分隔符（留空=直接拼接）", key="ct_s")
            cn = st.text_input("新列名", value="拼接结果", key="ct_n")
            if st.button("✅ 拼接", key="b_ct", type="primary") and cc:
                save_snapshot()
                df[cn] = df[cc].astype(str).agg(sep.join, axis=1)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{cn}]")
        
        with st.expander("✂️ 截取文本（LEFT / RIGHT / MID）"):
            cl = st.selectbox("选择列", all_cols, key="lr_c")
            lm = st.radio("方式", ["从左(LEFT)", "从右(RIGHT)", "中间(MID)"], horizontal=True, key="lr_m")
            if "中间" in lm:
                c1, c2 = st.columns(2)
                with c1: sp = st.number_input("起始位置", 0, 1000, 0, key="lr_s")
                with c2: cnt = st.number_input("字符数", 1, 1000, 3, key="lr_cnt")
            else:
                cnt = st.number_input("字符数", 1, 1000, 3, key="lr_cnt2")
            
            if st.button("✅ 截取", key="b_lr", type="primary"):
                save_snapshot()
                s = df[cl].astype(str)
                if "左" in lm: df[f"{cl}_left{cnt}"] = s.str[:cnt]
                elif "右" in lm: df[f"{cl}_right{cnt}"] = s.str[-cnt:]
                else: df[f"{cl}_mid"] = s.str[sp:sp+cnt]
                st.session_state.df = df
                st.success("✅ 完成")
    
    with t2:
        with st.expander("🔍 查找替换（SUBSTITUTE）"):
            cr2 = st.selectbox("选择列", all_cols, key="sub_c")
            ft = st.text_input("查找内容", key="sub_f")
            rt2 = st.text_input("替换为", key="sub_r")
            regex = st.checkbox("使用正则", key="sub_re")
            if st.button("✅ 替换", key="b_sub", type="primary") and ft:
                save_snapshot()
                df[cr2] = df[cr2].astype(str).str.replace(ft, rt2, regex=regex)
                st.session_state.df = df
                st.success("✅ 替换完成")
        
        with st.expander("🧹 去除空格（TRIM）"):
            ct2 = st.selectbox("选择列", all_cols, key="trm_c")
            tm = st.radio("模式", ["首尾空格", "所有空格", "多空格变一个"], horizontal=True, key="trm_m")
            if st.button("✅ 清理", key="b_trm", type="primary"):
                save_snapshot()
                s = df[ct2].astype(str)
                if "首尾" in tm: df[ct2] = s.str.strip()
                elif "所有" in tm: df[ct2] = s.str.replace(r'\s+', '', regex=True)
                else: df[ct2] = s.str.replace(r'\s+', ' ', regex=True).str.strip()
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("🔠 大小写转换"):
            cc2 = st.selectbox("选择列", all_cols, key="case_c")
            cm2 = st.radio("方式", ["全部大写UPPER", "全部小写LOWER", "首字母大写PROPER"], horizontal=True, key="case_m")
            if st.button("✅ 转换", key="b_case", type="primary"):
                save_snapshot()
                s = df[cc2].astype(str)
                if "大写" in cm2 and "首" not in cm2: df[cc2] = s.str.upper()
                elif "小写" in cm2: df[cc2] = s.str.lower()
                else: df[cc2] = s.str.title()
                st.session_state.df = df
                st.success("✅ 完成")
    
    with t3:
        with st.expander("🎯 提取指定内容（正则提取）"):
            ce = st.selectbox("选择列", all_cols, key="ext_c")
            em = st.selectbox("提取类型", ["数字", "中文", "英文字母", "邮箱", "手机号", "身份证号", "自定义正则"], key="ext_m")
            
            pm = {"数字": r'(\d+\.?\d*)', "中文": r'([\u4e00-\u9fa5]+)', "英文字母": r'([a-zA-Z]+)',
                  "邮箱": r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
                  "手机号": r'(1[3-9]\d{9})', "身份证号": r'(\d{17}[\dXx])'}
            
            if em == "自定义正则":
                pattern = st.text_input("正则表达式", key="ext_p")
            else:
                pattern = pm[em]
                st.code(f"正则: {pattern}")
            
            if st.button("✅ 提取", key="b_ext", type="primary") and pattern:
                save_snapshot()
                df[f"{ce}_提取"] = df[ce].astype(str).str.extract(pattern, expand=False)
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("📏 文本长度（LEN）"):
            cln = st.selectbox("选择列", all_cols, key="len_c")
            if st.button("✅ 长度", key="b_len", type="primary"):
                save_snapshot()
                df[f"{cln}_长度"] = df[cln].astype(str).str.len()
                st.session_state.df = df
                st.success("✅ 完成")
    
    with t4:
        with st.expander("📋 文本分列"):
            csp = st.selectbox("选择列", all_cols, key="spl_c")
            ssp = st.text_input("分隔符", value=",", key="spl_s")
            msp = st.number_input("最多拆分列数", 2, 20, 3, key="spl_m")
            if st.button("✅ 分列", key="b_spl", type="primary") and ssp:
                save_snapshot()
                result = df[csp].astype(str).str.split(ssp, n=msp-1, expand=True)
                for i in range(result.shape[1]):
                    df[f"{csp}_part{i+1}"] = result[i]
                st.session_state.df = df
                st.success(f"✅ 已拆分为 {result.shape[1]} 列")
        
        with st.expander("🔢 添加前缀/后缀/编号"):
            cpf = st.selectbox("选择列", all_cols, key="pf_c")
            pft = st.radio("类型", ["添加前缀", "添加后缀", "生成行编号"], horizontal=True, key="pf_t")
            
            if "编号" in pft:
                c1, c2 = st.columns(2)
                with c1: start_num = st.number_input("起始编号", 1, 999999, 1, key="pf_sn")
                with c2: prefix_str = st.text_input("前缀", value="NO.", key="pf_ps")
                pad = st.number_input("编号位数(补零)", 1, 10, 4, key="pf_pad")
            else:
                affix = st.text_input("前缀/后缀文本", key="pf_af")
            
            if st.button("✅ 执行", key="b_pf", type="primary"):
                save_snapshot()
                if "前缀" in pft:
                    df[f"{cpf}_加前缀"] = affix + df[cpf].astype(str)
                elif "后缀" in pft:
                    df[f"{cpf}_加后缀"] = df[cpf].astype(str) + affix
                else:
                    nums = range(start_num, start_num + len(df))
                    df["编号"] = [f"{prefix_str}{str(n).zfill(pad)}" for n in nums]
                st.session_state.df = df
                st.success("✅ 完成")


# ================================================================
#                        📅 日期处理
# ================================================================
elif menu == "📅 日期处理":
    st.subheader("📅 日期处理")
    
    t1, t2, t3 = st.tabs(["日期转换", "日期提取与计算", "日期生成"])
    
    with t1:
        with st.expander("📅 文本转日期", expanded=True):
            ctd = st.selectbox("选择列", all_cols, key="td_c")
            df_fmt = st.selectbox("格式", ["自动识别", "%Y-%m-%d", "%Y/%m/%d", "%Y%m%d",
                                            "%d/%m/%Y", "%m/%d/%Y", "%Y年%m月%d日"], key="td_f")
            if st.button("✅ 转换", key="b_td", type="primary"):
                save_snapshot()
                try:
                    if df_fmt == "自动识别":
                        df[ctd] = pd.to_datetime(df[ctd], errors='coerce')
                    else:
                        df[ctd] = pd.to_datetime(df[ctd], format=df_fmt, errors='coerce')
                    st.session_state.df = df
                    st.success("✅ 转换完成")
                except Exception as e:
                    st.error(f"失败: {e}")
    
    date_cols = df.select_dtypes(include='datetime').columns.tolist()
    
    with t2:
        if not date_cols:
            st.info("💡 请先在上方将文本列转换为日期格式")
        else:
            with st.expander("📆 提取年/月/日/星期/季度"):
                ced = st.selectbox("日期列", date_cols, key="de_c")
                parts = st.multiselect("提取", ["年份", "月份", "日", "星期", "季度", "周数", "年月", "是否月末"],
                                        default=["年份", "月份"], key="de_p")
                if st.button("✅ 提取", key="b_de", type="primary") and parts:
                    save_snapshot()
                    dt = df[ced].dt
                    pm2 = {"年份": dt.year, "月份": dt.month, "日": dt.day, "星期": dt.day_name(),
                           "季度": dt.quarter, "周数": dt.isocalendar().week.astype(int),
                           "年月": dt.strftime('%Y-%m'),
                           "是否月末": (dt.day == dt.days_in_month).map({True:"是",False:"否"})}
                    for p in parts:
                        df[f"{ced}_{p}"] = pm2[p]
                    st.session_state.df = df
                    st.success("✅ 完成")
            
            with st.expander("⏱️ 日期差（DATEDIF）"):
                c1, c2 = st.columns(2)
                with c1: ds = st.selectbox("开始日期", date_cols, key="dd_s")
                with c2: de = st.selectbox("结束日期", date_cols, key="dd_e")
                du = st.radio("单位", ["天", "小时", "月(近似)", "年(近似)"], horizontal=True, key="dd_u")
                if st.button("✅ 计算", key="b_dd", type="primary"):
                    save_snapshot()
                    delta = df[de] - df[ds]
                    um2 = {"天": delta.dt.days, "小时": (delta.dt.total_seconds()/3600).round(1),
                           "月(近似)": (delta.dt.days/30.44).round(1), "年(近似)": (delta.dt.days/365.25).round(2)}
                    df[f"日期差_{du}"] = um2[du]
                    st.session_state.df = df
                    st.success("✅ 完成")
            
            with st.expander("➡️ 日期偏移"):
                co = st.selectbox("日期列", date_cols, key="do_c")
                c1, c2 = st.columns(2)
                with c1: ov = st.number_input("偏移量", value=7, key="do_v")
                with c2: ou = st.selectbox("单位", ["天", "周", "月", "年"], key="do_u")
                if st.button("✅ 偏移", key="b_do", type="primary"):
                    save_snapshot()
                    om2 = {"天": pd.Timedelta(days=ov), "周": pd.Timedelta(weeks=ov),
                           "月": pd.DateOffset(months=int(ov)), "年": pd.DateOffset(years=int(ov))}
                    df[f"{co}_偏移"] = df[co] + om2[ou]
                    st.session_state.df = df
                    st.success("✅ 完成")
    
    with t3:
        with st.expander("📅 生成日期序列"):
            c1, c2, c3 = st.columns(3)
            with c1: sd = st.date_input("开始日期", key="gs_s")
            with c2: ed = st.date_input("结束日期", key="gs_e")
            with c3: freq = st.selectbox("频率", ["每天", "每周", "每月", "每季度", "每年"], key="gs_f")
            
            if st.button("✅ 生成", key="b_gs", type="primary"):
                fm = {"每天": "D", "每周": "W", "每月": "MS", "每季度": "QS", "每年": "YS"}
                dates = pd.date_range(start=sd, end=ed, freq=fm[freq])
                new_df = pd.DataFrame({"日期": dates})
                st.session_state.df = new_df
                df = new_df
                st.success(f"✅ 生成 {len(dates)} 条日期")
                st.dataframe(new_df.head(20))


# ================================================================
#                        🔎 条件与查找
# ================================================================
elif menu == "🔎 条件与查找":
    st.subheader("🔎 条件判断与查找匹配")
    
    t1, t2, t3, t4 = st.tabs(["IF条件判断", "多条件IFS", "VLOOKUP查找", "去重与计数"])
    
    with t1:
        st.markdown("##### ❓ IF 条件判断")
        c1, c2, c3 = st.columns(3)
        with c1: ic = st.selectbox("判断列", all_cols, key="if_c")
        with c2: it = st.selectbox("条件", [">", ">=", "<", "<=", "==", "!=", "包含", "不包含", "为空", "不为空"], key="if_t")
        with c3:
            if it not in ["为空", "不为空"]:
                iv = st.text_input("条件值", key="if_v")
            else:
                iv = ""
        
        c1, c2 = st.columns(2)
        with c1: tv = st.text_input("满足时 =", value="是", key="if_tv")
        with c2: fv = st.text_input("不满足时 =", value="否", key="if_fv")
        
        ifn = st.text_input("新列名", value="判断结果", key="if_n")
        
        if st.button("✅ 执行IF", key="b_if", type="primary"):
            save_snapshot()
            col_data = df[ic]
            
            try:
                threshold = float(iv) if iv else 0
                col_num = pd.to_numeric(col_data, errors='coerce')
                use_num = True
            except:
                threshold = iv
                col_num = col_data
                use_num = False
            
            cd = col_num if use_num else col_data
            cond_map = {
                ">": cd > threshold, ">=": cd >= threshold,
                "<": cd < threshold, "<=": cd <= threshold,
                "==": cd == threshold, "!=": cd != threshold,
                "包含": col_data.astype(str).str.contains(str(iv), na=False),
                "不包含": ~col_data.astype(str).str.contains(str(iv), na=False),
                "为空": col_data.isna() | (col_data.astype(str).str.strip()==''),
                "不为空": col_data.notna() & (col_data.astype(str).str.strip()!=''),
            }
            df[ifn] = np.where(cond_map[it], tv, fv)
            st.session_state.df = df
            st.success(f"✅ 已生成 [{ifn}]")
    
    with t2:
        st.markdown("##### 🔀 多条件分类（IFS）")
        ifs_col = st.selectbox("判断列", all_cols, key="ifs_c")
        num_cond = st.number_input("条件数量", 2, 15, 3, key="ifs_n")
        
        conditions = []
        for i in range(int(num_cond)):
            c1, c2, c3 = st.columns([1,1,1])
            with c1: op2 = st.selectbox(f"条件{i+1}运算符", [">=",">","<=","<","==","包含"], key=f"ifs_o_{i}")
            with c2: val = st.text_input(f"条件{i+1}值", key=f"ifs_v_{i}")
            with c3: res = st.text_input(f"条件{i+1}结果", key=f"ifs_r_{i}")
            conditions.append((op2, val, res))
        
        default = st.text_input("默认值", value="其他", key="ifs_d")
        ifs_name = st.text_input("新列名", value="分类结果", key="ifs_nm")
        
        if st.button("✅ 执行分类", key="b_ifs", type="primary"):
            save_snapshot()
            result = pd.Series(default, index=df.index)
            for op2, val, res in reversed(conditions):
                if not val or not res: continue
                try:
                    vn = float(val)
                    cd2 = pd.to_numeric(df[ifs_col], errors='coerce')
                except:
                    vn = val
                    cd2 = df[ifs_col].astype(str)
                
                om3 = {">=": cd2>=vn, ">": cd2>vn, "<=": cd2<=vn, "<": cd2<vn,
                       "==": cd2==vn, "包含": df[ifs_col].astype(str).str.contains(str(val), na=False)}
                result[om3[op2]] = res
            
            df[ifs_name] = result
            st.session_state.df = df
            st.success(f"✅ 已生成 [{ifs_name}]")
    
    with t3:
        st.markdown("##### 🔎 VLOOKUP 跨表查找")
        lf = st.file_uploader("上传查找表", type=["xlsx","xls","csv"], key="vl_f")
        
        if lf:
            if lf.name.endswith('.csv'):
                df_lk = pd.read_csv(lf)
            else:
                df_lk = pd.read_excel(lf)
            
            st.write("查找表预览：", df_lk.head())
            
            c1, c2 = st.columns(2)
            with c1: mk = st.selectbox("主表匹配列", all_cols, key="vl_mk")
            with c2: lk = st.selectbox("查找表匹配列", df_lk.columns.tolist(), key="vl_lk")
            
            rc = st.multiselect("要匹配回来的列", [c for c in df_lk.columns if c != lk], key="vl_rc")
            
            if st.button("✅ VLOOKUP", key="b_vl", type="primary") and rc:
                save_snapshot()
                lk_sub = df_lk[[lk]+rc].drop_duplicates(subset=lk)
                df = df.merge(lk_sub, left_on=mk, right_on=lk, how='left', suffixes=('','_查找'))
                if lk != mk and lk in df.columns:
                    df = df.drop(columns=[lk])
                st.session_state.df = df
                st.success(f"✅ VLOOKUP完成")
    
    with t4:
        with st.expander("🔢 计数（COUNTIF）"):
            ccf = st.selectbox("选择列", all_cols, key="cf_c")
            if st.button("✅ 计数", key="b_cf", type="primary"):
                save_snapshot()
                df[f"{ccf}_出现次数"] = df[ccf].map(df[ccf].value_counts())
                st.session_state.df = df
                st.success("✅ 完成")
        
        with st.expander("🏷️ 去重"):
            dpc = st.multiselect("判重列", all_cols, key="dp_c")
            dpa = st.radio("操作", ["标记重复", "删除重复(保留首条)", "删除重复(保留末条)"], horizontal=True, key="dp_a")
            if st.button("✅ 执行", key="b_dp", type="primary") and dpc:
                save_snapshot()
                if "标记" in dpa:
                    df["是否重复"] = df.duplicated(subset=dpc, keep=False).map({True:"重复",False:"唯一"})
                elif "首条" in dpa:
                    b = len(df)
                    df = df.drop_duplicates(subset=dpc, keep='first').reset_index(drop=True)
                    st.success(f"✅ {b}→{len(df)}行")
                else:
                    b = len(df)
                    df = df.drop_duplicates(subset=dpc, keep='last').reset_index(drop=True)
                    st.success(f"✅ {b}→{len(df)}行")
                st.session_state.df = df


# ================================================================
#                        📊 统计汇总
# ================================================================
elif menu == "📊 统计汇总":
    st.subheader("📊 统计汇总与分析")
    
    t1, t2, t3, t4 = st.tabs(["分组汇总", "条件汇总SUMIF", "描述统计", "数值分箱"])
    
    with t1:
        st.markdown("##### 📊 分组汇总（数据透视表）")
        gc = st.multiselect("分组列", all_cols, key="pv_g")
        ac = st.multiselect("汇总列", numeric_cols, key="pv_a")
        af = st.multiselect("汇总方式", ["求和sum", "平均值mean", "计数count", "最大值max", "最小值min", "中位数median"],
                             default=["求和sum"], key="pv_f")
        
        if st.button("✅ 汇总", key="b_pv", type="primary") and gc and ac and af:
            fm2 = {"求和sum":"sum","平均值mean":"mean","计数count":"count",
                   "最大值max":"max","最小值min":"min","中位数median":"median"}
            funcs = [fm2[f] for f in af]
            pivot = df.groupby(gc)[ac].agg(funcs).round(2)
            if isinstance(pivot.columns, pd.MultiIndex):
                pivot.columns = ['_'.join(c).strip() for c in pivot.columns]
            pivot = pivot.reset_index()
            
            st.dataframe(pivot, use_container_width=True)
            st.session_state['pivot_table'] = pivot
            
            if st.checkbox("合并回主表", key="pv_merge"):
                save_snapshot()
                df = df.merge(pivot, on=gc, how='left', suffixes=('','_汇总'))
                st.session_state.df = df
                st.success("✅ 已合并")
    
    with t2:
        st.markdown("##### 🎯 条件汇总（SUMIF / AVERAGEIF）")
        sg = st.selectbox("分组列", all_cols, key="si_g")
        sv2 = st.selectbox("汇总列", numeric_cols, key="si_v")
        sf = st.selectbox("方式", ["SUMIF求和", "AVERAGEIF平均", "COUNTIF计数", "MAXIF最大", "MINIF最小"], key="si_f")
        
        if st.button("✅ 执行", key="b_si", type="primary"):
            save_snapshot()
            fm3 = {"SUMIF求和":"sum","AVERAGEIF平均":"mean","COUNTIF计数":"count",
                   "MAXIF最大":"max","MINIF最小":"min"}
            f = fm3[sf]
            df[f"{sv2}_{f}_by_{sg}"] = df.groupby(sg)[sv2].transform(f).round(2)
            st.session_state.df = df
            st.success("✅ 完成")
    
    with t3:
        if numeric_cols:
            stats = df[numeric_cols].describe().round(2).T
            stats.columns = ['计数','均值','标准差','最小','25%','中位数','75%','最大']
            stats['总和'] = df[numeric_cols].sum().round(2)
            stats['缺失'] = df[numeric_cols].isna().sum()
            st.dataframe(stats, use_container_width=True)
            st.session_state['stats_table'] = stats.reset_index().rename(columns={'index':'列名'})
    
    with t4:
        if numeric_cols:
            bc = st.selectbox("选择列", numeric_cols, key="bn_c")
            bm = st.radio("分箱方式", ["等距分箱", "自定义分界点", "等频分箱"], horizontal=True, key="bn_m")
            
            if bm == "等距分箱":
                nb = st.number_input("分段数", 2, 20, 5, key="bn_n")
            elif bm == "自定义分界点":
                be = st.text_input("分界点(逗号分隔)", value="0,60,80,100", key="bn_e")
            else:
                nb = st.number_input("分组数", 2, 20, 5, key="bn_q")
            
            bl = st.text_input("自定义标签(可选,逗号分隔)", key="bn_l")
            
            if st.button("✅ 分箱", key="b_bn", type="primary"):
                save_snapshot()
                try:
                    labels = bl.split(",") if bl else None
                    if bm == "等距分箱":
                        df[f"{bc}_分段"] = pd.cut(df[bc], bins=nb, labels=labels)
                    elif bm == "自定义分界点":
                        edges = [float(x.strip()) for x in be.split(",")]
                        df[f"{bc}_分段"] = pd.cut(df[bc], bins=edges, labels=labels, include_lowest=True)
                    else:
                        df[f"{bc}_分段"] = pd.qcut(df[bc], q=nb, labels=labels, duplicates='drop')
                    st.session_state.df = df
                    st.success("✅ 完成")
                except Exception as e:
                    st.error(f"失败: {e}")


# ================================================================
#                        🧮 公式引擎（小白友好版）
# ================================================================
elif menu == "🧮 公式引擎":
    st.subheader("🧮 智能公式生成器")
    st.caption("✨ 全程点选，无需写代码！选列名 + 选运算 + 点生成 = 完成")
    
    t1, t2, t3, t4 = st.tabs(["🎯 快捷公式（推荐）", "🔧 公式构建器", "📚 常用模板", "💻 高级模式"])
    
    # ============================================================
    # Tab 1: 快捷公式 - 一键完成最常见需求
    # ============================================================
    with t1:
        st.markdown("##### 👇 选择您要做的事情")
        
        scene = st.selectbox(
            "我想要...",
            [
                "💰 计算含税金额（金额 × 1.13）",
                "💰 计算税额（金额 × 13%）",
                "💰 计算折扣价（原价 × 折扣率）",
                "💰 计算利润（收入 - 成本）",
                "💰 计算利润率%（(收入-成本)/收入×100）",
                "📊 两列相加",
                "📊 两列相减",
                "📊 两列相乘",
                "📊 两列相除",
                "📊 多列求和",
                "📊 多列求平均",
                "🔢 某列 × 固定数字",
                "🔢 某列 ÷ 固定数字",
                "🔢 某列 + 固定数字",
                "🔢 某列 - 固定数字",
                "🔢 某列保留N位小数",
                "🔢 某列百分比（×100加%）",
                "📝 两列文本拼接",
                "📝 添加前缀",
                "📝 添加后缀",
            ],
            key="quick_scene"
        )
        
        st.markdown("---")
        
        # ============ 财务场景 ============
        if "含税金额" in scene:
            col = st.selectbox("📌 选择「金额」列", numeric_cols, key="q1_c")
            new_name = st.text_input("✏️ 新列名", value="含税金额", key="q1_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col} × 1.13</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q1_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = (df[col] * 1.13).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成新列 [{new_name}]"); st.rerun()
        
        elif "税额" in scene:
            col = st.selectbox("📌 选择「金额」列", numeric_cols, key="q2_c")
            rate = st.number_input("💯 税率(%)", value=13.0, key="q2_r")
            new_name = st.text_input("✏️ 新列名", value="税额", key="q2_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col} × {rate}%</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q2_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = (df[col] * rate / 100).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "折扣价" in scene:
            c1, c2 = st.columns(2)
            with c1: col_a = st.selectbox("📌 「原价」列", numeric_cols, key="q3_a")
            with c2: col_b = st.selectbox("📌 「折扣率」列", numeric_cols, key="q3_b")
            new_name = st.text_input("✏️ 新列名", value="折扣价", key="q3_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col_a} × {col_b}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q3_b2", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = (df[col_a] * df[col_b]).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif scene.startswith("💰 计算利润（"):
            c1, c2 = st.columns(2)
            with c1: col_a = st.selectbox("📌 「收入」列", numeric_cols, key="q4_a")
            with c2: col_b = st.selectbox("📌 「成本」列", numeric_cols, key="q4_b")
            new_name = st.text_input("✏️ 新列名", value="利润", key="q4_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col_a} - {col_b}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q4_b2", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = (df[col_a] - df[col_b]).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "利润率" in scene:
            c1, c2 = st.columns(2)
            with c1: col_a = st.selectbox("📌 「收入」列", numeric_cols, key="q5_a")
            with c2: col_b = st.selectbox("📌 「成本」列", numeric_cols, key="q5_b")
            new_name = st.text_input("✏️ 新列名", value="利润率%", key="q5_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>({col_a} - {col_b}) / {col_a} × 100</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q5_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = ((df[col_a] - df[col_b]) / df[col_a].replace(0, np.nan) * 100).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        # ============ 两列运算 ============
        elif scene in ["📊 两列相加", "📊 两列相减", "📊 两列相乘", "📊 两列相除"]:
            op_map = {
                "相加": ("+", lambda a, b: a + b),
                "相减": ("-", lambda a, b: a - b),
                "相乘": ("×", lambda a, b: a * b),
                "相除": ("÷", lambda a, b: a / b.replace(0, np.nan)),
            }
            op_key = [k for k in op_map if k in scene][0]
            symbol, func = op_map[op_key]
            
            # 使用 op_key（中文）作为唯一key前缀，避免符号冲突
            c1, c2 = st.columns(2)
            with c1: 
                col_a = st.selectbox("📌 列A", numeric_cols, key=f"q6_{op_key}_colA")
            with c2: 
                col_b = st.selectbox("📌 列B", numeric_cols, key=f"q6_{op_key}_colB")
            new_name = st.text_input("✏️ 新列名", value=f"{col_a}{symbol}{col_b}", key=f"q6_{op_key}_name")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col_a} {symbol} {col_b}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key=f"q6_{op_key}_btn", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = func(df[col_a], df[col_b]).round(4)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        # ============ 多列汇总 ============
        elif "多列求和" in scene:
            cols = st.multiselect("📌 选择要求和的多列", numeric_cols, key="q7_c")
            new_name = st.text_input("✏️ 新列名", value="合计", key="q7_n")
            if cols:
                st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{" + ".join(cols)}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q7_b", type="primary", use_container_width=True) and cols:
                save_snapshot()
                df[new_name] = df[cols].sum(axis=1).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "多列求平均" in scene:
            cols = st.multiselect("📌 选择要计算平均值的多列", numeric_cols, key="q8_c")
            new_name = st.text_input("✏️ 新列名", value="平均值", key="q8_n")
            if cols:
                st.markdown(f'<div class="formula-preview">📐 计算公式: <b>({" + ".join(cols)}) ÷ {len(cols)}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q8_b", type="primary", use_container_width=True) and cols:
                save_snapshot()
                df[new_name] = df[cols].mean(axis=1).round(2)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        # ============ 列与固定数字运算 ============
        elif scene in ["🔢 某列 × 固定数字", "🔢 某列 ÷ 固定数字", 
                       "🔢 某列 + 固定数字", "🔢 某列 - 固定数字"]:
            op_map2 = {
                "乘": ("×", lambda a, n: a * n),
                "除": ("÷", lambda a, n: a / n if n != 0 else np.nan),
                "加": ("+", lambda a, n: a + n),
                "减": ("-", lambda a, n: a - n),
            }
            # 用中文动词作为唯一key
            op_key2 = None
            for k in op_map2:
                if k in scene:
                    op_key2 = k
                    break
            symbol, func2 = op_map2[op_key2]
            
            c1, c2 = st.columns(2)
            with c1: 
                col = st.selectbox("📌 选择列", numeric_cols, key=f"q9_{op_key2}_col")
            with c2: 
                num = st.number_input("🔢 输入数字", value=1.0, key=f"q9_{op_key2}_num")
            new_name = st.text_input("✏️ 新列名", value=f"{col}{symbol}{num}", key=f"q9_{op_key2}_name")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col} {symbol} {num}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key=f"q9_{op_key2}_btn", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = func2(df[col], num).round(4)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "保留N位小数" in scene:
            col = st.selectbox("📌 选择列", numeric_cols, key="q10_c")
            digits = st.number_input("🔢 保留小数位数", 0, 10, 2, key="q10_d")
            new_name = st.text_input("✏️ 新列名", value=f"{col}_四舍五入", key="q10_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>四舍五入({col}, {digits}位小数)</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q10_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = df[col].round(digits)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "百分比" in scene:
            col = st.selectbox("📌 选择列", numeric_cols, key="q11_c")
            new_name = st.text_input("✏️ 新列名", value=f"{col}_百分比", key="q11_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col} × 100 + "%"</b>（例：0.25 → 25.00%）</div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q11_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = (df[col] * 100).round(2).astype(str) + '%'
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        # ============ 文本场景 ============
        elif "两列文本拼接" in scene:
            c1, c2, c3 = st.columns(3)
            with c1: col_a = st.selectbox("📌 列A", all_cols, key="q12_a")
            with c2: sep = st.text_input("🔗 中间分隔符", value="-", key="q12_s")
            with c3: col_b = st.selectbox("📌 列B", all_cols, key="q12_b")
            new_name = st.text_input("✏️ 新列名", value="拼接结果", key="q12_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col_a} + "{sep}" + {col_b}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q12_b2", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = df[col_a].astype(str) + sep + df[col_b].astype(str)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "添加前缀" in scene:
            col = st.selectbox("📌 选择列", all_cols, key="q13_c")
            prefix = st.text_input("✏️ 前缀文本", value="PRE_", key="q13_p")
            new_name = st.text_input("✏️ 新列名", value=f"{col}_加前缀", key="q13_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>"{prefix}" + {col}</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q13_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = prefix + df[col].astype(str)
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
        
        elif "添加后缀" in scene:
            col = st.selectbox("📌 选择列", all_cols, key="q14_c")
            suffix = st.text_input("✏️ 后缀文本", value="_END", key="q14_s")
            new_name = st.text_input("✏️ 新列名", value=f"{col}_加后缀", key="q14_n")
            st.markdown(f'<div class="formula-preview">📐 计算公式: <b>{col} + "{suffix}"</b></div>', unsafe_allow_html=True)
            if st.button("✅ 立即生成", key="q14_b", type="primary", use_container_width=True):
                save_snapshot()
                df[new_name] = df[col].astype(str) + suffix
                st.session_state.df = df
                st.success(f"✅ 已生成 [{new_name}]"); st.rerun()
    
    # ============================================================
    # Tab 2: 公式构建器 - 可视化拼接公式
    # ============================================================
    with t2:
        st.markdown("##### 🔧 自由组合：列 + 运算符 + 列/数字")
        st.caption("最多支持 5 个操作数组合（A 运算符 B 运算符 C...）")
        
        num_operands = st.slider("参与运算的项数", 2, 5, 2, key="fb_n")
        
        # 初始化构建器状态
        if 'builder_parts' not in st.session_state:
            st.session_state.builder_parts = []
        
        operands = []  # 操作数
        operators = []  # 运算符
        
        # 第一个操作数
        st.markdown("**操作数 1**")
        c1, c2 = st.columns([1, 3])
        with c1:
            t = st.radio("类型", ["列", "数字"], key="fb_t_0", horizontal=True)
        with c2:
            if t == "列":
                v = st.selectbox("选择列", numeric_cols, key="fb_v_0", label_visibility="collapsed")
                operands.append(("col", v))
            else:
                v = st.number_input("输入数字", value=1.0, key="fb_v_0", label_visibility="collapsed")
                operands.append(("num", v))
        
        # 后续操作数
        for i in range(1, num_operands):
            c1, c2, c3 = st.columns([1, 1, 3])
            with c1:
                op = st.selectbox(f"运算符", ["+", "-", "×", "÷"], key=f"fb_op_{i}")
                operators.append(op)
            with c2:
                t = st.radio(f"类型", ["列", "数字"], key=f"fb_t_{i}", horizontal=True)
            with c3:
                if t == "列":
                    v = st.selectbox("选择列", numeric_cols, key=f"fb_v_{i}", label_visibility="collapsed")
                    operands.append(("col", v))
                else:
                    v = st.number_input("输入数字", value=1.0, key=f"fb_v_{i}", label_visibility="collapsed")
                    operands.append(("num", v))
        
        # 构建可读公式
        readable = str(operands[0][1])
        for i, op in enumerate(operators):
            readable += f" {op} {operands[i+1][1]}"
        
        st.markdown(f'<div class="formula-preview">📐 公式预览: <b>{readable}</b></div>', unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            round_digits = st.number_input("🔢 结果保留小数位", 0, 10, 2, key="fb_rd")
        with c2:
            fb_name = st.text_input("✏️ 新列名", value="构建结果", key="fb_name")
        
        if st.button("✅ 计算并生成新列", type="primary", use_container_width=True, key="fb_btn"):
            save_snapshot()
            try:
                # 计算
                def get_val(operand):
                    typ, val = operand
                    return df[val] if typ == "col" else val
                
                result = get_val(operands[0])
                for i, op in enumerate(operators):
                    next_val = get_val(operands[i+1])
                    if op == "+": result = result + next_val
                    elif op == "-": result = result - next_val
                    elif op == "×": result = result * next_val
                    elif op == "÷":
                        if isinstance(next_val, pd.Series):
                            result = result / next_val.replace(0, np.nan)
                        else:
                            result = result / next_val if next_val != 0 else np.nan
                
                if isinstance(result, pd.Series):
                    df[fb_name] = result.round(round_digits)
                else:
                    df[fb_name] = round(result, round_digits)
                
                st.session_state.df = df
                st.success(f"✅ 已生成 [{fb_name}]")
                st.rerun()
            except Exception as e:
                st.error(f"计算失败: {e}")
    
    # ============================================================
    # Tab 3: 常用模板（保留可视化、隐藏代码细节）
    # ============================================================
    with t3:
        st.markdown("##### 📚 业务场景模板")
        st.caption("点击应用按钮，按提示选择对应列即可")
        
        templates = {
            "💼 销售数据": [
                {"name": "销售提成（销售额×提成比例）", "type": "two_col_mul", "fields": ["销售额列", "提成比例列"], "result": "提成金额"},
                {"name": "客单价（销售额÷订单数）", "type": "two_col_div", "fields": ["销售额列", "订单数列"], "result": "客单价"},
                {"name": "环比增长率%", "type": "growth_rate", "fields": ["数值列"], "result": "环比增长率%"},
            ],
            "📦 库存管理": [
                {"name": "库存金额（数量×单价）", "type": "two_col_mul", "fields": ["数量列", "单价列"], "result": "库存金额"},
                {"name": "库存周转天数（库存÷日均销量）", "type": "two_col_div", "fields": ["库存列", "日均销量列"], "result": "周转天数"},
            ],
            "👤 员工管理": [
                {"name": "实发工资（应发-五险一金-个税）", "type": "subtract_3", "fields": ["应发工资", "五险一金", "个税"], "result": "实发工资"},
                {"name": "年薪（月薪×12）", "type": "col_mul_num", "fields": ["月薪列"], "num": 12, "result": "年薪"},
            ],
            "📈 数据分析": [
                {"name": "Z-Score标准化", "type": "zscore", "fields": ["数值列"], "result": "标准化值"},
                {"name": "Min-Max归一化到0-1", "type": "minmax", "fields": ["数值列"], "result": "归一化值"},
                {"name": "占比%（每行÷总和）", "type": "pct", "fields": ["数值列"], "result": "占比%"},
            ],
        }
        
        for category, items in templates.items():
            with st.expander(category, expanded=False):
                for idx, tpl in enumerate(items):
                    st.markdown(f"**📌 {tpl['name']}**")
                    
                    fields_data = {}
                    if len(tpl['fields']) == 1:
                        fields_data[tpl['fields'][0]] = st.selectbox(
                            tpl['fields'][0], numeric_cols, key=f"tpl_{category}_{idx}_0")
                    else:
                        cols_ui = st.columns(len(tpl['fields']))
                        for fi, f in enumerate(tpl['fields']):
                            with cols_ui[fi]:
                                fields_data[f] = st.selectbox(f, numeric_cols, key=f"tpl_{category}_{idx}_{fi}")
                    
                    result_name = st.text_input("结果列名", value=tpl['result'], key=f"tpl_{category}_{idx}_rn")
                    
                    if st.button(f"✅ 应用", key=f"tpl_{category}_{idx}_btn", type="primary"):
                        save_snapshot()
                        try:
                            vals = list(fields_data.values())
                            if tpl['type'] == "two_col_mul":
                                df[result_name] = (df[vals[0]] * df[vals[1]]).round(2)
                            elif tpl['type'] == "two_col_div":
                                df[result_name] = (df[vals[0]] / df[vals[1]].replace(0, np.nan)).round(2)
                            elif tpl['type'] == "subtract_3":
                                df[result_name] = (df[vals[0]] - df[vals[1]] - df[vals[2]]).round(2)
                            elif tpl['type'] == "col_mul_num":
                                df[result_name] = (df[vals[0]] * tpl['num']).round(2)
                            elif tpl['type'] == "growth_rate":
                                df[result_name] = (df[vals[0]].pct_change() * 100).round(2)
                            elif tpl['type'] == "zscore":
                                df[result_name] = ((df[vals[0]] - df[vals[0]].mean()) / df[vals[0]].std()).round(4)
                            elif tpl['type'] == "minmax":
                                mn, mx = df[vals[0]].min(), df[vals[0]].max()
                                df[result_name] = ((df[vals[0]] - mn) / (mx - mn)).round(4)
                            elif tpl['type'] == "pct":
                                total = df[vals[0]].sum()
                                df[result_name] = (df[vals[0]] / total * 100).round(2).astype(str) + '%'
                            
                            st.session_state.df = df
                            st.success(f"✅ 已生成 [{result_name}]")
                            st.rerun()
                        except Exception as e:
                            st.error(f"失败: {e}")
                    
                    st.markdown("---")
    
    # ============================================================
    # Tab 4: 高级模式（保留给会写代码的人）
    # ============================================================
    with t4:
        st.markdown("##### 💻 高级模式：直接写 Python/Pandas 公式")
        st.caption("适合熟悉编程的用户，支持完整的 pandas 语法")
        
        with st.expander("📋 列名速查表"):
            ref = pd.DataFrame({
                '列名': df.columns,
                '类型': df.dtypes.astype(str).values,
                '公式中引用方式': [f"df['{c}']" for c in df.columns]
            })
            st.dataframe(ref, use_container_width=True, hide_index=True)
        
        formula = st.text_area(
            "输入公式",
            value="",
            height=100,
            placeholder="例如: df['金额'] * 1.13\n或: np.where(df['销售额']>1000, '高', '低')",
            key="adv_formula"
        )
        
        new_col = st.text_input("新列名", value="计算结果", key="adv_name")
        
        c1, c2 = st.columns(2)
        with c1:
            if st.button("👁️ 预览前10行", key="adv_preview", use_container_width=True) and formula:
                try:
                    result = eval(formula, {"__builtins__": {}, "np": np, "pd": pd}, {"df": df})
                    st.write(result.head(10) if isinstance(result, pd.Series) else result)
                except Exception as e:
                    st.error(f"❌ {e}")
        with c2:
            if st.button("✅ 生成新列", key="adv_exec", type="primary", use_container_width=True) and formula and new_col:
                save_snapshot()
                try:
                    result = eval(formula, {"__builtins__": {}, "np": np, "pd": pd}, {"df": df})
                    df[new_col] = result
                    st.session_state.df = df
                    st.success(f"✅ 已生成 [{new_col}]")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")


# ================================================================
#                        💾 导出区域（始终显示）
# ================================================================
st.markdown("---")
st.markdown("### 💾 导出数据")

df = st.session_state.df

c1, c2, c3 = st.columns([3, 2, 2])
with c1:
    out_name = st.text_input("📄 文件名", value=generate_filename(), key="out_n")
with c2:
    out_fmt = st.selectbox("📦 格式", ["Excel (.xlsx)", "CSV (.csv)"], key="out_f")
with c3:
    inc_idx = st.checkbox("包含行号", False, key="out_i")

# 构建导出数据
export_sheets = {"处理结果": df}
if 'pivot_table' in st.session_state:
    export_sheets["汇总表"] = st.session_state.pivot_table
if 'stats_table' in st.session_state:
    export_sheets["统计概览"] = st.session_state.stats_table

if out_fmt == "Excel (.xlsx)":
    excel_data = df_to_excel_optimized(export_sheets, index=inc_idx)
    st.download_button("⬇️ 下载 Excel 文件", data=excel_data,
                       file_name=f"{out_name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)
else:
    csv = df.to_csv(index=inc_idx).encode('utf-8-sig')
    st.download_button("⬇️ 下载 CSV 文件", data=csv,
                       file_name=f"{out_name}.csv", mime="text/csv",
                       type="primary", use_container_width=True)

with st.expander(f"👁️ 预览导出数据（{len(df):,}行 × {len(df.columns)}列）"):
    st.dataframe(df.head(200), use_container_width=True, height=300)
