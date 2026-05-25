"""文本处理模块 - 对应Excel的文本函数"""
import streamlit as st
import pandas as pd
import re


def render_text_ops(df: pd.DataFrame) -> pd.DataFrame:
    """渲染文本处理功能面板"""
    
    text_cols = df.select_dtypes(include='object').columns.tolist()
    all_cols = df.columns.tolist()
    
    if not text_cols:
        st.warning("⚠️ 当前数据没有文本列")
        # 允许将数值列转为文本操作
        text_cols = all_cols
    
    # ===================== 1. 文本拼接 CONCATENATE =====================
    with st.expander("🔗 文本拼接（CONCATENATE）"):
        st.caption("将多列文本合并为一列")
        cols_concat = st.multiselect("选择要拼接的列", all_cols, key="concat_cols")
        separator = st.text_input("分隔符", value="", key="concat_sep", 
                                   help="留空则直接拼接，常用：-、_、空格")
        concat_name = st.text_input("新列名称", value="拼接结果", key="concat_name")
        
        if st.button("✅ 执行拼接", key="btn_concat", type="primary"):
            if cols_concat:
                df[concat_name] = df[cols_concat].astype(str).agg(separator.join, axis=1)
                st.success(f"✅ 已生成 [{concat_name}]")
    
    # ===================== 2. 截取文本 LEFT/RIGHT/MID =====================
    with st.expander("✂️ 截取文本（LEFT / RIGHT / MID）"):
        col_cut = st.selectbox("选择列", all_cols, key="cut_col")
        cut_mode = st.radio("截取方式", ["从左截取(LEFT)", "从右截取(RIGHT)", "中间截取(MID)"],
                            horizontal=True, key="cut_mode")
        
        if cut_mode == "中间截取(MID)":
            c1, c2 = st.columns(2)
            with c1:
                start_pos = st.number_input("起始位置（从0开始）", min_value=0, value=0, key="mid_start")
            with c2:
                char_count = st.number_input("截取字符数", min_value=1, value=3, key="mid_count")
        else:
            char_count = st.number_input("截取字符数", min_value=1, value=3, key="cut_count")
        
        if st.button("✅ 执行截取", key="btn_cut", type="primary"):
            col_str = df[col_cut].astype(str)
            if cut_mode == "从左截取(LEFT)":
                df[f"{col_cut}_left{char_count}"] = col_str.str[:char_count]
            elif cut_mode == "从右截取(RIGHT)":
                df[f"{col_cut}_right{char_count}"] = col_str.str[-char_count:]
            else:
                df[f"{col_cut}_mid"] = col_str.str[start_pos:start_pos+char_count]
            st.success("✅ 截取完成")
    
    # ===================== 3. 查找替换 SUBSTITUTE =====================
    with st.expander("🔍 查找替换（SUBSTITUTE）"):
        col_replace = st.selectbox("选择列", all_cols, key="replace_col")
        find_text = st.text_input("查找内容", key="find_text")
        replace_text = st.text_input("替换为", key="replace_text")
        use_regex = st.checkbox("使用正则表达式", key="use_regex")
        
        if st.button("✅ 执行替换", key="btn_replace", type="primary"):
            if find_text:
                df[col_replace] = df[col_replace].astype(str).str.replace(
                    find_text, replace_text, regex=use_regex
                )
                st.success("✅ 替换完成")
    
    # ===================== 4. 大小写转换 UPPER/LOWER/PROPER =====================
    with st.expander("🔠 大小写转换（UPPER/LOWER/PROPER）"):
        col_case = st.selectbox("选择列", all_cols, key="case_col")
        case_mode = st.radio("转换方式", 
                             ["全部大写(UPPER)", "全部小写(LOWER)", "首字母大写(PROPER)"],
                             horizontal=True, key="case_mode")
        
        if st.button("✅ 执行", key="btn_case", type="primary"):
            col_str = df[col_case].astype(str)
            if case_mode == "全部大写(UPPER)":
                df[col_case] = col_str.str.upper()
            elif case_mode == "全部小写(LOWER)":
                df[col_case] = col_str.str.lower()
            else:
                df[col_case] = col_str.str.title()
            st.success("✅ 转换完成")
    
    # ===================== 5. 去除空格 TRIM =====================
    with st.expander("🧹 去除空格（TRIM）"):
        col_trim = st.selectbox("选择列", all_cols, key="trim_col")
        trim_mode = st.radio("模式", 
                             ["去除首尾空格", "去除所有空格", "多个空格变一个"],
                             horizontal=True, key="trim_mode")
        
        if st.button("✅ 执行", key="btn_trim", type="primary"):
            col_str = df[col_trim].astype(str)
            if trim_mode == "去除首尾空格":
                df[col_trim] = col_str.str.strip()
            elif trim_mode == "去除所有空格":
                df[col_trim] = col_str.str.replace(r'\s+', '', regex=True)
            else:
                df[col_trim] = col_str.str.replace(r'\s+', ' ', regex=True).str.strip()
            st.success("✅ 处理完成")
    
    # ===================== 6. 文本长度 LEN =====================
    with st.expander("📏 文本长度（LEN）"):
        col_len = st.selectbox("选择列", all_cols, key="len_col")
        if st.button("✅ 执行", key="btn_len", type="primary"):
            df[f"{col_len}_长度"] = df[col_len].astype(str).str.len()
            st.success(f"✅ 已生成 [{col_len}_长度]")
    
    # ===================== 7. 提取数字/中文/英文 =====================
    with st.expander("🎯 提取指定内容（正则提取）"):
        col_extract = st.selectbox("选择列", all_cols, key="extract_col")
        extract_mode = st.radio("提取内容", 
                                ["提取数字", "提取中文", "提取英文字母", "提取邮箱", "提取手机号", "自定义正则"],
                                horizontal=True, key="extract_mode")
        
        pattern_map = {
            "提取数字": r'(\d+\.?\d*)',
            "提取中文": r'([\u4e00-\u9fa5]+)',
            "提取英文字母": r'([a-zA-Z]+)',
            "提取邮箱": r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            "提取手机号": r'(1[3-9]\d{9})',
        }
        
        if extract_mode == "自定义正则":
            pattern = st.text_input("输入正则表达式", key="custom_regex")
        else:
            pattern = pattern_map[extract_mode]
            st.code(f"使用正则: {pattern}")
        
        if st.button("✅ 执行提取", key="btn_extract", type="primary"):
            if pattern:
                df[f"{col_extract}_提取"] = df[col_extract].astype(str).str.extract(pattern, expand=False)
                st.success("✅ 提取完成")
    
    # ===================== 8. 文本分列 =====================
    with st.expander("📋 文本分列（按分隔符拆分）"):
        col_split = st.selectbox("选择列", all_cols, key="split_col")
        split_sep = st.text_input("分隔符", value=",", key="split_sep",
                                   help="常用：逗号,  空格  横杠-  下划线_  斜杠/")
        max_split = st.number_input("最多拆分为几列", min_value=2, max_value=20, value=3, key="max_split")
        
        if st.button("✅ 执行分列", key="btn_split", type="primary"):
            if split_sep:
                split_result = df[col_split].astype(str).str.split(split_sep, n=max_split-1, expand=True)
                for i in range(split_result.shape[1]):
                    df[f"{col_split}_part{i+1}"] = split_result[i]
                st.success(f"✅ 已拆分为 {split_result.shape[1]} 列")
    
    return df