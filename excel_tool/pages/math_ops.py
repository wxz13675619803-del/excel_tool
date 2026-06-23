"""数学运算模块 - 对应Excel中的数学函数"""
import streamlit as st
import pandas as pd
import numpy as np


def render_math_ops(df: pd.DataFrame) -> pd.DataFrame:
    """渲染数学运算功能面板"""
    
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    if not numeric_cols:
        st.warning("⚠️ 当前数据没有数值列，无法进行数学运算")
        return df
    
    # ===================== 1. 行求和 SUM =====================
    with st.expander("➕ 行求和（SUM）", expanded=False):
        st.caption("对选中的多列进行每行求和，生成新列")
        cols_sum = st.multiselect("选择要求和的列", numeric_cols, key="sum_cols")
        sum_col_name = st.text_input("新列名称", value="合计", key="sum_name")
        if st.button("✅ 执行求和", key="btn_sum", type="primary"):
            if cols_sum:
                df[sum_col_name] = df[cols_sum].sum(axis=1)
                st.success(f"✅ 已生成 [{sum_col_name}] 列")
            else:
                st.error("请至少选择一列")
    
    # ===================== 2. 行平均 AVERAGE =====================
    with st.expander("📊 行平均值（AVERAGE）"):
        st.caption("计算选中列的每行平均值")
        cols_avg = st.multiselect("选择列", numeric_cols, key="avg_cols")
        avg_col_name = st.text_input("新列名称", value="平均值", key="avg_name")
        if st.button("✅ 执行", key="btn_avg", type="primary"):
            if cols_avg:
                df[avg_col_name] = df[cols_avg].mean(axis=1)
                st.success(f"✅ 已生成 [{avg_col_name}] 列")
    
    # ===================== 3. 最大值 MAX =====================
    with st.expander("⬆️ 行最大值（MAX）"):
        st.caption("取选中列每行的最大值")
        cols_max = st.multiselect("选择列", numeric_cols, key="max_cols")
        max_col_name = st.text_input("新列名称", value="最大值", key="max_name")
        if st.button("✅ 执行", key="btn_max", type="primary"):
            if cols_max:
                df[max_col_name] = df[cols_max].max(axis=1)
                st.success(f"✅ 已生成 [{max_col_name}] 列")
    
    # ===================== 4. 最小值 MIN =====================
    with st.expander("⬇️ 行最小值（MIN）"):
        cols_min = st.multiselect("选择列", numeric_cols, key="min_cols")
        min_col_name = st.text_input("新列名称", value="最小值", key="min_name")
        if st.button("✅ 执行", key="btn_min", type="primary"):
            if cols_min:
                df[min_col_name] = df[cols_min].min(axis=1)
                st.success(f"✅ 已生成 [{min_col_name}] 列")
    
    # ===================== 5. 四则运算 =====================
    with st.expander("🔢 两列四则运算"):
        st.caption("选择两列进行加减乘除运算")
        c1, c2, c3 = st.columns(3)
        with c1:
            col_a = st.selectbox("列A", numeric_cols, key="arith_a")
        with c2:
            operator = st.selectbox("运算符", ["+", "-", "×", "÷", "取余%", "幂^"], key="arith_op")
        with c3:
            col_b = st.selectbox("列B", numeric_cols, key="arith_b")
        
        arith_name = st.text_input("新列名称", value="运算结果", key="arith_name")
        
        if st.button("✅ 执行运算", key="btn_arith", type="primary"):
            op_map = {
                "+": df[col_a] + df[col_b],
                "-": df[col_a] - df[col_b],
                "×": df[col_a] * df[col_b],
                "÷": df[col_a] / df[col_b].replace(0, np.nan),
                "取余%": df[col_a] % df[col_b].replace(0, np.nan),
                "幂^": df[col_a] ** df[col_b],
            }
            df[arith_name] = op_map[operator]
            st.success(f"✅ {col_a} {operator} {col_b} → [{arith_name}]")
    
    # ===================== 6. 四舍五入 ROUND =====================
    with st.expander("🔄 四舍五入（ROUND）"):
        col_round = st.selectbox("选择列", numeric_cols, key="round_col")
        decimals = st.number_input("保留小数位数", min_value=0, max_value=10, value=2, key="round_dec")
        round_mode = st.radio("模式", ["覆盖原列", "生成新列"], horizontal=True, key="round_mode")
        
        if st.button("✅ 执行", key="btn_round", type="primary"):
            if round_mode == "覆盖原列":
                df[col_round] = df[col_round].round(decimals)
                st.success(f"✅ [{col_round}] 已四舍五入到 {decimals} 位")
            else:
                new_name = f"{col_round}_round{decimals}"
                df[new_name] = df[col_round].round(decimals)
                st.success(f"✅ 已生成 [{new_name}]")
    
    # ===================== 7. 绝对值 ABS =====================
    with st.expander("| | 绝对值（ABS）"):
        col_abs = st.selectbox("选择列", numeric_cols, key="abs_col")
        if st.button("✅ 执行", key="btn_abs", type="primary"):
            df[f"{col_abs}_abs"] = df[col_abs].abs()
            st.success(f"✅ 已生成 [{col_abs}_abs]")
    
    # ===================== 8. 累计求和 CUMSUM =====================
    with st.expander("📈 累计求和（CUMSUM / Running Total）"):
        st.caption("从第一行开始逐行累加，常用于计算累计销售额等")
        col_cumsum = st.selectbox("选择列", numeric_cols, key="cumsum_col")
        if st.button("✅ 执行", key="btn_cumsum", type="primary"):
            df[f"{col_cumsum}_累计"] = df[col_cumsum].cumsum()
            st.success(f"✅ 已生成 [{col_cumsum}_累计]")
    
    # ===================== 9. 排名 RANK =====================
    with st.expander("🏆 排名（RANK）"):
        col_rank = st.selectbox("选择列", numeric_cols, key="rank_col")
        rank_order = st.radio("排序方式", ["降序（最大=第1名）", "升序（最小=第1名）"], 
                              horizontal=True, key="rank_order")
        ascending = rank_order == "升序（最小=第1名）"
        
        if st.button("✅ 执行", key="btn_rank", type="primary"):
            df[f"{col_rank}_排名"] = df[col_rank].rank(ascending=ascending, method='min').astype(int)
            st.success(f"✅ 已生成 [{col_rank}_排名]")
    
    # ===================== 10. 百分比/占比 =====================
    with st.expander("📐 占比/百分比"):
        st.caption("计算每行在该列总和中的占比")
        col_pct = st.selectbox("选择列", numeric_cols, key="pct_col")
        pct_format = st.radio("格式", ["小数（0.25）", "百分比（25%）"], horizontal=True, key="pct_fmt")
        
        if st.button("✅ 执行", key="btn_pct", type="primary"):
            total = df[col_pct].sum()
            if total != 0:
                if pct_format == "小数（0.25）":
                    df[f"{col_pct}_占比"] = (df[col_pct] / total).round(4)
                else:
                    df[f"{col_pct}_占比%"] = ((df[col_pct] / total) * 100).round(2).astype(str) + '%'
                st.success("✅ 已生成占比列")
            else:
                st.error("该列总和为0，无法计算占比")
    
    return df