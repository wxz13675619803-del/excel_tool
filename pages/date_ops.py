"""日期处理模块 - 对应Excel的日期函数"""
import streamlit as st
import pandas as pd
import numpy as np


def render_date_ops(df: pd.DataFrame) -> pd.DataFrame:
    """渲染日期处理功能面板"""
    
    all_cols = df.columns.tolist()
    
    # ===================== 1. 文本转日期 =====================
    with st.expander("📅 文本转日期格式"):
        st.caption("将文本列转换为标准日期格式")
        col_todate = st.selectbox("选择列", all_cols, key="todate_col")
        date_format = st.selectbox("原始格式", [
            "自动识别",
            "%Y-%m-%d (2024-01-15)",
            "%Y/%m/%d (2024/01/15)",
            "%Y%m%d (20240115)",
            "%d/%m/%Y (15/01/2024)",
            "%m/%d/%Y (01/15/2024)",
            "%Y年%m月%d日",
        ], key="date_fmt")
        
        if st.button("✅ 转换", key="btn_todate", type="primary"):
            try:
                if date_format == "自动识别":
                    df[col_todate] = pd.to_datetime(df[col_todate], errors='coerce')
                else:
                    fmt = date_format.split(" ")[0]
                    df[col_todate] = pd.to_datetime(df[col_todate], format=fmt, errors='coerce')
                st.success("✅ 转换完成")
            except Exception as e:
                st.error(f"转换失败: {e}")
    
    # 尝试自动检测日期列
    date_cols = df.select_dtypes(include='datetime').columns.tolist()
    
    if not date_cols:
        st.info("💡 提示：请先使用上方功能将文本列转换为日期格式")
        return df
    
    # ===================== 2. 提取年/月/日/星期 =====================
    with st.expander("📆 提取年/月/日/星期/季度（YEAR/MONTH/DAY）"):
        col_extract = st.selectbox("选择日期列", date_cols, key="date_extract_col")
        extract_parts = st.multiselect("提取内容", 
                                        ["年份", "月份", "日", "星期几", "季度", "周数", "年月"],
                                        default=["年份", "月份"], key="date_parts")
        
        if st.button("✅ 执行提取", key="btn_date_extract", type="primary"):
            dt = df[col_extract].dt
            part_map = {
                "年份": ("年份", dt.year),
                "月份": ("月份", dt.month),
                "日": ("日", dt.day),
                "星期几": ("星期", dt.day_name()),
                "季度": ("季度", dt.quarter),
                "周数": ("周数", dt.isocalendar().week.astype(int)),
                "年月": ("年月", dt.strftime('%Y-%m')),
            }
            for part in extract_parts:
                name, values = part_map[part]
                df[f"{col_extract}_{name}"] = values
            st.success("✅ 提取完成")
    
    # ===================== 3. 日期差 DATEDIF =====================
    with st.expander("⏱️ 日期差（DATEDIF）"):
        st.caption("计算两个日期之间的差值")
        c1, c2 = st.columns(2)
        with c1:
            date_start = st.selectbox("开始日期列", date_cols, key="datedif_start")
        with c2:
            date_end = st.selectbox("结束日期列", date_cols, key="datedif_end")
        
        diff_unit = st.radio("差值单位", ["天", "小时", "月（近似）", "年（近似）"],
                             horizontal=True, key="diff_unit")
        
        if st.button("✅ 计算差值", key="btn_datedif", type="primary"):
            delta = df[date_end] - df[date_start]
            if diff_unit == "天":
                df["日期差_天"] = delta.dt.days
            elif diff_unit == "小时":
                df["日期差_小时"] = (delta.dt.total_seconds() / 3600).round(1)
            elif diff_unit == "月（近似）":
                df["日期差_月"] = (delta.dt.days / 30.44).round(1)
            else:
                df["日期差_年"] = (delta.dt.days / 365.25).round(2)
            st.success("✅ 计算完成")
    
    # ===================== 4. 日期偏移 =====================
    with st.expander("➡️ 日期偏移（DATE加减天数）"):
        col_offset = st.selectbox("选择日期列", date_cols, key="offset_col")
        c1, c2 = st.columns(2)
        with c1:
            offset_val = st.number_input("偏移量", value=7, key="offset_val")
        with c2:
            offset_unit = st.selectbox("单位", ["天", "周", "月", "年"], key="offset_unit")
        
        if st.button("✅ 执行偏移", key="btn_offset", type="primary"):
            unit_map = {"天": "D", "周": "W", "月": "ME", "年": "YE"}
            try:
                if offset_unit == "天":
                    df[f"{col_offset}_偏移"] = df[col_offset] + pd.Timedelta(days=offset_val)
                elif offset_unit == "周":
                    df[f"{col_offset}_偏移"] = df[col_offset] + pd.Timedelta(weeks=offset_val)
                elif offset_unit == "月":
                    df[f"{col_offset}_偏移"] = df[col_offset] + pd.DateOffset(months=int(offset_val))
                else:
                    df[f"{col_offset}_偏移"] = df[col_offset] + pd.DateOffset(years=int(offset_val))
                st.success("✅ 偏移完成")
            except Exception as e:
                st.error(f"偏移失败: {e}")
    
    # ===================== 5. 工作日计算 =====================
    with st.expander("🏢 是否工作日 / 月末"):
        col_wd = st.selectbox("选择日期列", date_cols, key="workday_col")
        wd_ops = st.multiselect("生成内容", 
                                 ["是否工作日", "是否月末", "是否月初", "是否季末"],
                                 key="wd_ops")
        
        if st.button("✅ 执行", key="btn_wd", type="primary"):
            dt = df[col_wd].dt
            if "是否工作日" in wd_ops:
                df[f"{col_wd}_工作日"] = dt.dayofweek.apply(lambda x: "是" if x < 5 else "否")
            if "是否月末" in wd_ops:
                df[f"{col_wd}_月末"] = (dt.day == dt.days_in_month).map({True: "是", False: "否"})
            if "是否月初" in wd_ops:
                df[f"{col_wd}_月初"] = (dt.day == 1).map({True: "是", False: "否"})
            if "是否季末" in wd_ops:
                df[f"{col_wd}_季末"] = ((dt.month % 3 == 0) & (dt.day == dt.days_in_month)).map({True: "是", False: "否"})
            st.success("✅ 生成完成")
    
    return df