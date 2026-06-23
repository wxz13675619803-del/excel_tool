"""查找匹配模块 - 对应Excel的VLOOKUP/IF等"""
import streamlit as st
import pandas as pd
import numpy as np


def render_lookup_ops(df: pd.DataFrame) -> pd.DataFrame:
    """渲染查找匹配功能面板"""
    
    all_cols = df.columns.tolist()
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    # ===================== 1. IF条件判断 =====================
    with st.expander("❓ IF条件判断", expanded=False):
        st.caption("根据条件生成新列，类似 Excel 的 IF 函数")
        col_if = st.selectbox("判断列", all_cols, key="if_col")
        
        condition_type = st.selectbox("条件类型", [
            "大于", "小于", "等于", "大于等于", "小于等于", 
            "不等于", "包含文本", "不包含文本", "为空", "不为空"
        ], key="if_cond")
        
        if condition_type not in ["为空", "不为空"]:
            threshold = st.text_input("条件值", key="if_threshold",
                                       help="数值直接输入数字，文本直接输入文本")
        
        c1, c2 = st.columns(2)
        with c1:
            true_val = st.text_input("满足条件时的值", value="是", key="if_true")
        with c2:
            false_val = st.text_input("不满足条件时的值", value="否", key="if_false")
        
        if_col_name = st.text_input("新列名称", value="判断结果", key="if_name")
        
        if st.button("✅ 执行判断", key="btn_if", type="primary"):
            try:
                col_data = df[col_if]
                
                # 尝试数值比较
                if condition_type in ["大于", "小于", "等于", "大于等于", "小于等于", "不等于"]:
                    try:
                        threshold_val = float(threshold)
                        col_data_num = pd.to_numeric(col_data, errors='coerce')
                    except:
                        threshold_val = threshold
                        col_data_num = col_data
                    
                    cond_map = {
                        "大于": col_data_num > threshold_val,
                        "小于": col_data_num < threshold_val,
                        "等于": col_data_num == threshold_val,
                        "大于等于": col_data_num >= threshold_val,
                        "小于等于": col_data_num <= threshold_val,
                        "不等于": col_data_num != threshold_val,
                    }
                    mask = cond_map[condition_type]
                elif condition_type == "包含文本":
                    mask = col_data.astype(str).str.contains(threshold, na=False)
                elif condition_type == "不包含文本":
                    mask = ~col_data.astype(str).str.contains(threshold, na=False)
                elif condition_type == "为空":
                    mask = col_data.isna() | (col_data.astype(str).str.strip() == '')
                else:  # 不为空
                    mask = col_data.notna() & (col_data.astype(str).str.strip() != '')
                
                df[if_col_name] = np.where(mask, true_val, false_val)
                st.success(f"✅ 已生成 [{if_col_name}]")
            except Exception as e:
                st.error(f"执行失败: {e}")
    
    # ===================== 2. 多条件IF（IFS / 嵌套IF）=====================
    with st.expander("🔀 多条件分类（IFS / 嵌套IF）"):
        st.caption("按多个区间或条件分类，从上到下匹配，命中第一个就停止")
        col_ifs = st.selectbox("判断列", all_cols, key="ifs_col")
        
        num_conditions = st.number_input("条件数量", min_value=2, max_value=10, value=3, key="ifs_num")
        
        conditions = []
        for i in range(int(num_conditions)):
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1:
                op = st.selectbox(f"条件{i+1} 运算符", 
                                  [">=", ">", "<=", "<", "==", "包含"],
                                  key=f"ifs_op_{i}")
            with c2:
                val = st.text_input(f"条件{i+1} 值", key=f"ifs_val_{i}")
            with c3:
                result = st.text_input(f"条件{i+1} 结果", key=f"ifs_result_{i}")
            conditions.append((op, val, result))
        
        default_val = st.text_input("都不满足时的默认值", value="其他", key="ifs_default")
        ifs_name = st.text_input("新列名称", value="分类结果", key="ifs_name")
        
        if st.button("✅ 执行分类", key="btn_ifs", type="primary"):
            try:
                result_series = pd.Series(default_val, index=df.index)
                
                # 从后往前处理（后面的条件先赋值，前面的覆盖）
                for op, val, result in reversed(conditions):
                    if not val or not result:
                        continue
                    
                    try:
                        val_num = float(val)
                        col_data = pd.to_numeric(df[col_ifs], errors='coerce')
                    except:
                        val_num = val
                        col_data = df[col_ifs].astype(str)
                    
                    if op == ">=":
                        mask = col_data >= val_num
                    elif op == ">":
                        mask = col_data > val_num
                    elif op == "<=":
                        mask = col_data <= val_num
                    elif op == "<":
                        mask = col_data < val_num
                    elif op == "==":
                        mask = col_data == val_num
                    elif op == "包含":
                        mask = df[col_ifs].astype(str).str.contains(val, na=False)
                    
                    result_series[mask] = result
                
                df[ifs_name] = result_series
                st.success(f"✅ 已生成 [{ifs_name}]")
            except Exception as e:
                st.error(f"执行失败: {e}")
    
    # ===================== 3. VLOOKUP =====================
    with st.expander("🔎 VLOOKUP（跨表查找匹配）"):
        st.caption("从另一个Excel文件中查找匹配数据，类似Excel的VLOOKUP")
        
        lookup_file = st.file_uploader("上传查找表（参照表）", type=["xlsx", "xls"],
                                        key="vlookup_file")
        
        if lookup_file is not None:
            df_lookup = pd.read_excel(lookup_file)
            st.write("查找表预览：", df_lookup.head())
            
            c1, c2 = st.columns(2)
            with c1:
                main_key = st.selectbox("主表匹配列", all_cols, key="vlookup_main_key")
            with c2:
                lookup_key = st.selectbox("查找表匹配列", df_lookup.columns.tolist(),
                                           key="vlookup_lookup_key")
            
            return_cols = st.multiselect("要匹配回来的列",
                                          [c for c in df_lookup.columns if c != lookup_key],
                                          key="vlookup_return")
            
            if st.button("✅ 执行VLOOKUP", key="btn_vlookup", type="primary"):
                if return_cols:
                    lookup_subset = df_lookup[[lookup_key] + return_cols].drop_duplicates(subset=lookup_key)
                    df = df.merge(lookup_subset, left_on=main_key, right_on=lookup_key,
                                  how='left', suffixes=('', '_查找'))
                    if lookup_key != main_key and lookup_key in df.columns:
                        df = df.drop(columns=[lookup_key])
                    st.success(f"✅ VLOOKUP 完成！匹配了 {return_cols} 列")
                else:
                    st.error("请选择要匹配的列")
    
    # ===================== 4. COUNTIF =====================
    with st.expander("🔢 计数（COUNTIF）"):
        st.caption("统计某个值在该列中出现的次数")
        col_countif = st.selectbox("选择列", all_cols, key="countif_col")
        
        if st.button("✅ 生成计数列", key="btn_countif", type="primary"):
            value_counts = df[col_countif].map(df[col_countif].value_counts())
            df[f"{col_countif}_出现次数"] = value_counts
            st.success("✅ 已生成计数列")
    
    # ===================== 5. 去重标记 =====================
    with st.expander("🏷️ 重复值标记 / 去重"):
        col_dup = st.multiselect("选择判重列", all_cols, key="dup_cols")
        dup_action = st.radio("操作", ["标记重复项", "删除重复项（保留第一条）", 
                                       "删除重复项（保留最后一条）"],
                              horizontal=True, key="dup_action")
        
        if st.button("✅ 执行", key="btn_dup", type="primary"):
            if col_dup:
                if dup_action == "标记重复项":
                    df["是否重复"] = df.duplicated(subset=col_dup, keep=False).map(
                        {True: "重复", False: "唯一"})
                    dup_count = (df["是否重复"] == "重复").sum()
                    st.success(f"✅ 发现 {dup_count} 条重复记录")
                elif "第一条" in dup_action:
                    before = len(df)
                    df = df.drop_duplicates(subset=col_dup, keep='first')
                    st.success(f"✅ 去重完成：{before} → {len(df)} 条（删除了 {before-len(df)} 条）")
                else:
                    before = len(df)
                    df = df.drop_duplicates(subset=col_dup, keep='last')
                    st.success(f"✅ 去重完成：{before} → {len(df)} 条")
    
    return df