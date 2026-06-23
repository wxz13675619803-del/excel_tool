"""统计分析模块 - 对应Excel的数据透视表等"""
import streamlit as st
import pandas as pd
import numpy as np


def render_stats_ops(df: pd.DataFrame) -> pd.DataFrame:
    """渲染统计分析功能面板"""
    
    all_cols = df.columns.tolist()
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    # ===================== 1. 分组汇总（数据透视表）=====================
    with st.expander("📊 分组汇总（数据透视表 / SUMIF）", expanded=False):
        st.caption("按某列分组，对数值列进行汇总统计")
        
        group_cols = st.multiselect("分组列（行标签）", all_cols, key="group_cols")
        agg_cols = st.multiselect("汇总列（值）", numeric_cols, key="agg_cols")
        agg_funcs = st.multiselect("汇总方式", 
                                    ["求和(sum)", "平均值(mean)", "计数(count)", 
                                     "最大值(max)", "最小值(min)", "中位数(median)",
                                     "标准差(std)"],
                                    default=["求和(sum)"], key="agg_funcs")
        
        if st.button("✅ 生成汇总表", key="btn_pivot", type="primary"):
            if group_cols and agg_cols and agg_funcs:
                func_map = {
                    "求和(sum)": "sum", "平均值(mean)": "mean",
                    "计数(count)": "count", "最大值(max)": "max",
                    "最小值(min)": "min", "中位数(median)": "median",
                    "标准差(std)": "std"
                }
                funcs = [func_map[f] for f in agg_funcs]
                
                pivot = df.groupby(group_cols)[agg_cols].agg(funcs).round(2)
                
                # 扁平化多级列名
                if isinstance(pivot.columns, pd.MultiIndex):
                    pivot.columns = ['_'.join(col).strip() for col in pivot.columns]
                
                pivot = pivot.reset_index()
                st.write("📋 汇总结果：")
                st.dataframe(pivot, use_container_width=True)
                
                # 提供将汇总结果合并回主表的选项
                if st.checkbox("将汇总结果合并回主表", key="merge_pivot"):
                    df = df.merge(pivot, on=group_cols, how='left', suffixes=('', '_汇总'))
                    st.success("✅ 已合并回主表")
                
                # 单独保存汇总表
                st.session_state['pivot_table'] = pivot
    
    # ===================== 2. SUMIF / AVERAGEIF =====================
    with st.expander("🎯 条件汇总（SUMIF / AVERAGEIF）"):
        st.caption("按条件对数据进行汇总计算，结果写回每一行")
        
        group_col_sumif = st.selectbox("分组列", all_cols, key="sumif_group")
        val_col_sumif = st.selectbox("汇总列", numeric_cols, key="sumif_val")
        sumif_func = st.selectbox("汇总方式", 
                                   ["求和(SUMIF)", "平均值(AVERAGEIF)", "计数(COUNTIF)",
                                    "最大值(MAXIF)", "最小值(MINIF)"],
                                   key="sumif_func")
        
        if st.button("✅ 执行", key="btn_sumif", type="primary"):
            func_map = {
                "求和(SUMIF)": "sum", "平均值(AVERAGEIF)": "mean",
                "计数(COUNTIF)": "count", "最大值(MAXIF)": "max",
                "最小值(MINIF)": "min"
            }
            func = func_map[sumif_func]
            agg_result = df.groupby(group_col_sumif)[val_col_sumif].transform(func)
            col_name = f"{val_col_sumif}_{func}_{group_col_sumif}"
            df[col_name] = agg_result.round(2)
            st.success(f"✅ 已生成 [{col_name}]")
    
    # ===================== 3. 描述性统计 =====================
    with st.expander("📈 描述性统计一览"):
        if numeric_cols:
            stats_df = df[numeric_cols].describe().round(2).T
            stats_df.columns = ['计数', '均值', '标准差', '最小值', '25%分位', 
                               '中位数', '75%分位', '最大值']
            stats_df['总和'] = df[numeric_cols].sum().round(2)
            stats_df['缺失数'] = df[numeric_cols].isna().sum()
            
            st.dataframe(stats_df, use_container_width=True)
            st.session_state['stats_table'] = stats_df.reset_index().rename(columns={'index': '列名'})
        else:
            st.warning("没有数值列")
    
    # ===================== 4. 频率分布 / 分箱 =====================
    with st.expander("📦 数值分箱（分段统计）"):
        st.caption("将连续数值划分为区间，类似Excel的FREQUENCY")
        col_bin = st.selectbox("选择列", numeric_cols, key="bin_col")
        
        bin_method = st.radio("分箱方式", ["等距分箱", "自定义分界点", "等频分箱"],
                              horizontal=True, key="bin_method")
        
        if bin_method == "等距分箱":
            num_bins = st.number_input("分几段", min_value=2, max_value=20, value=5, key="num_bins")
        elif bin_method == "自定义分界点":
            bin_edges = st.text_input("输入分界点（逗号分隔）", 
                                      value="0,60,80,100", key="bin_edges",
                                      help="例如成绩分段：0,60,70,80,90,100")
        else:
            num_bins = st.number_input("分几组", min_value=2, max_value=20, value=5, key="qnum_bins")
        
        bin_labels_input = st.text_input("自定义标签（可选，逗号分隔）", key="bin_labels",
                                          help="例如：不及格,及格,良好,优秀")
        
        if st.button("✅ 执行分箱", key="btn_bin", type="primary"):
            try:
                if bin_method == "等距分箱":
                    labels = bin_labels_input.split(",") if bin_labels_input else None
                    if labels and len(labels) != num_bins:
                        st.error(f"标签数量({len(labels)})需要等于分段数({num_bins})")
                    else:
                        df[f"{col_bin}_分段"] = pd.cut(df[col_bin], bins=num_bins, labels=labels)
                        st.success("✅ 分箱完成")
                
                elif bin_method == "自定义分界点":
                    edges = [float(x.strip()) for x in bin_edges.split(",")]
                    labels = bin_labels_input.split(",") if bin_labels_input else None
                    if labels and len(labels) != len(edges) - 1:
                        st.error(f"标签数量应为 {len(edges)-1}")
                    else:
                        df[f"{col_bin}_分段"] = pd.cut(df[col_bin], bins=edges, labels=labels, 
                                                      include_lowest=True)
                        st.success("✅ 分箱完成")
                
                else:  # 等频分箱
                    labels = bin_labels_input.split(",") if bin_labels_input else None
                    if labels and len(labels) != num_bins:
                        st.error(f"标签数量({len(labels)})需要等于分组数({num_bins})")
                    else:
                        df[f"{col_bin}_分段"] = pd.qcut(df[col_bin], q=num_bins, labels=labels,
                                                       duplicates='drop')
                        st.success("✅ 分箱完成")
            except Exception as e:
                st.error(f"分箱失败: {e}")
    
    return df