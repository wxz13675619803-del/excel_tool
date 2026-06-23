"""数据清洗模块"""
import streamlit as st
import pandas as pd
import numpy as np


def render_data_clean(df: pd.DataFrame) -> pd.DataFrame:
    """渲染数据清洗功能面板"""
    
    all_cols = df.columns.tolist()
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    # ===================== 1. 缺失值处理 =====================
    with st.expander("🕳️ 缺失值处理", expanded=False):
        # 显示缺失值概况
        missing = df.isna().sum()
        missing_pct = (missing / len(df) * 100).round(1)
        missing_df = pd.DataFrame({
            '列名': missing.index,
            '缺失数': missing.values,
            '缺失率(%)': missing_pct.values
        })
        missing_df = missing_df[missing_df['缺失数'] > 0]
        
        if len(missing_df) > 0:
            st.dataframe(missing_df, use_container_width=True, hide_index=True)
            
            col_fill = st.selectbox("选择要处理的列", 
                                     missing_df['列名'].tolist(), key="fill_col")
            fill_method = st.selectbox("填充方式", [
                "删除缺失行",
                "填充固定值",
                "填充均值",
                "填充中位数",
                "填充众数",
                "前值填充（向下填充）",
                "后值填充（向上填充）",
                "填充为 0",
                "填充为 空字符串",
            ], key="fill_method")
            
            if fill_method == "填充固定值":
                fill_value = st.text_input("输入填充值", key="fill_value")
            
            if st.button("✅ 执行处理", key="btn_fill", type="primary"):
                if fill_method == "删除缺失行":
                    before = len(df)
                    df = df.dropna(subset=[col_fill])
                    st.success(f"✅ 删除 {before - len(df)} 行")
                elif fill_method == "填充固定值":
                    df[col_fill] = df[col_fill].fillna(fill_value)
                    st.success("✅ 填充完成")
                elif fill_method == "填充均值":
                    df[col_fill] = df[col_fill].fillna(df[col_fill].mean())
                    st.success("✅ 填充完成")
                elif fill_method == "填充中位数":
                    df[col_fill] = df[col_fill].fillna(df[col_fill].median())
                    st.success("✅ 填充完成")
                elif fill_method == "填充众数":
                    mode_val = df[col_fill].mode()
                    if len(mode_val) > 0:
                        df[col_fill] = df[col_fill].fillna(mode_val[0])
                    st.success("✅ 填充完成")
                elif fill_method == "前值填充（向下填充）":
                    df[col_fill] = df[col_fill].ffill()
                    st.success("✅ 填充完成")
                elif fill_method == "后值填充（向上填充）":
                    df[col_fill] = df[col_fill].bfill()
                    st.success("✅ 填充完成")
                elif fill_method == "填充为 0":
                    df[col_fill] = df[col_fill].fillna(0)
                    st.success("✅ 填充完成")
                else:
                    df[col_fill] = df[col_fill].fillna('')
                    st.success("✅ 填充完成")
        else:
            st.success("🎉 数据没有缺失值！")
    
    # ===================== 2. 数据类型转换 =====================
    with st.expander("🔄 数据类型转换"):
        col_type = st.selectbox("选择列", all_cols, key="type_col")
        current_type = str(df[col_type].dtype)
        st.info(f"当前类型: **{current_type}**")
        
        target_type = st.selectbox("转换为", [
            "文本(str)", "整数(int)", "浮点数(float)", 
            "日期(datetime)", "布尔(bool)", "分类(category)"
        ], key="target_type")
        
        if st.button("✅ 执行转换", key="btn_type", type="primary"):
            try:
                type_map = {
                    "文本(str)": lambda: df[col_type].astype(str),
                    "整数(int)": lambda: pd.to_numeric(df[col_type], errors='coerce').astype('Int64'),
                    "浮点数(float)": lambda: pd.to_numeric(df[col_type], errors='coerce'),
                    "日期(datetime)": lambda: pd.to_datetime(df[col_type], errors='coerce'),
                    "布尔(bool)": lambda: df[col_type].astype(bool),
                    "分类(category)": lambda: df[col_type].astype('category'),
                }
                df[col_type] = type_map[target_type]()
                st.success(f"✅ [{col_type}] 已转换为 {target_type}")
            except Exception as e:
                st.error(f"转换失败: {e}")
    
    # ===================== 3. 列操作 =====================
    with st.expander("📝 列管理（重命名/删除/排序）"):
        tab1, tab2, tab3 = st.tabs(["重命名", "删除列", "调整顺序"])
        
        with tab1:
            col_rename = st.selectbox("选择列", df.columns.tolist(), key="rename_col")
            new_name = st.text_input("新名称", key="rename_new")
            if st.button("✅ 重命名", key="btn_rename"):
                if new_name:
                    df = df.rename(columns={col_rename: new_name})
                    st.success(f"✅ {col_rename} → {new_name}")
        
        with tab2:
            cols_drop = st.multiselect("选择要删除的列", df.columns.tolist(), key="drop_cols")
            if st.button("✅ 删除选中列", key="btn_drop", type="primary"):
                if cols_drop:
                    df = df.drop(columns=cols_drop)
                    st.success(f"✅ 已删除 {len(cols_drop)} 列")
        
        with tab3:
            new_order = st.multiselect("拖拽调整列顺序（按选择顺序排列）", 
                                        df.columns.tolist(),
                                        default=df.columns.tolist(),
                                        key="col_order")
            if st.button("✅ 应用新顺序", key="btn_order"):
                if len(new_order) == len(df.columns):
                    df = df[new_order]
                    st.success("✅ 列顺序已调整")
                else:
                    st.error("请选择所有列")
    
    # ===================== 4. 排序 =====================
    with st.expander("🔃 数据排序"):
        sort_cols = st.multiselect("排序列", df.columns.tolist(), key="sort_cols")
        if sort_cols:
            sort_orders = []
            for col in sort_cols:
                order = st.radio(f"{col} 排序方式", ["升序 ↑", "降序 ↓"], 
                                horizontal=True, key=f"sort_{col}")
                sort_orders.append(order == "升序 ↑")
            
            if st.button("✅ 执行排序", key="btn_sort", type="primary"):
                df = df.sort_values(by=sort_cols, ascending=sort_orders).reset_index(drop=True)
                st.success("✅ 排序完成")
    
    # ===================== 5. 筛选数据 =====================
    with st.expander("🔍 数据筛选"):
        filter_col = st.selectbox("筛选列", df.columns.tolist(), key="filter_col")
        
        if df[filter_col].dtype in ['object', 'category']:
            unique_vals = df[filter_col].dropna().unique().tolist()
            selected_vals = st.multiselect("选择保留的值", unique_vals, 
                                            default=unique_vals, key="filter_vals")
            if st.button("✅ 应用筛选", key="btn_filter", type="primary"):
                before = len(df)
                df = df[df[filter_col].isin(selected_vals)]
                st.success(f"✅ 筛选完成：{before} → {len(df)} 行")
        else:
            c1, c2 = st.columns(2)
            with c1:
                min_val = st.number_input("最小值", value=float(df[filter_col].min()),
                                          key="filter_min")
            with c2:
                max_val = st.number_input("最大值", value=float(df[filter_col].max()),
                                          key="filter_max")
            if st.button("✅ 应用筛选", key="btn_filter_num", type="primary"):
                before = len(df)
                df = df[(df[filter_col] >= min_val) & (df[filter_col] <= max_val)]
                st.success(f"✅ 筛选完成：{before} → {len(df)} 行")
    
    return df