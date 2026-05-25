"""自定义公式引擎 - 支持类Excel公式输入"""
import streamlit as st
import pandas as pd
import numpy as np


# 公式模板库
FORMULA_TEMPLATES = {
    "基础运算": {
        "两列相加": {"formula": "df['列A'] + df['列B']", "desc": "列A + 列B"},
        "两列相乘": {"formula": "df['列A'] * df['列B']", "desc": "列A × 列B"},
        "含税价格": {"formula": "df['单价'] * 1.13", "desc": "单价 × 1.13（13%增值税）"},
        "折扣价": {"formula": "df['原价'] * df['折扣率']", "desc": "原价 × 折扣率"},
        "利润率": {"formula": "(df['收入'] - df['成本']) / df['收入'] * 100", "desc": "(收入-成本)/收入×100"},
    },
    "条件计算": {
        "正数变0": {"formula": "df['列A'].clip(upper=0)", "desc": "大于0的变为0"},
        "负数变0": {"formula": "df['列A'].clip(lower=0)", "desc": "小于0的变为0"},
        "限定范围": {"formula": "df['列A'].clip(lower=0, upper=100)", "desc": "限定在0-100之间"},
        "空值变0": {"formula": "df['列A'].fillna(0)", "desc": "空值填充为0"},
    },
    "文本处理": {
        "姓+名拼接": {"formula": "df['姓'] + df['名']", "desc": "姓 + 名"},
        "添加前缀": {"formula": "'PRE_' + df['列A'].astype(str)", "desc": "添加前缀PRE_"},
        "添加后缀": {"formula": "df['列A'].astype(str) + '_后缀'", "desc": "添加后缀"},
        "取前3个字符": {"formula": "df['列A'].astype(str).str[:3]", "desc": "LEFT(列A, 3)"},
    },
    "数学函数": {
        "平方": {"formula": "df['列A'] ** 2", "desc": "列A的平方"},
        "平方根": {"formula": "np.sqrt(df['列A'].abs())", "desc": "列A的平方根"},
        "对数(ln)": {"formula": "np.log(df['列A'].clip(lower=0.001))", "desc": "自然对数"},
        "取整(向下)": {"formula": "np.floor(df['列A'])", "desc": "向下取整"},
        "取整(向上)": {"formula": "np.ceil(df['列A'])", "desc": "向上取整"},
    },
}


def render_formula_engine(df: pd.DataFrame) -> pd.DataFrame:
    """渲染自定义公式引擎"""
    
    st.subheader("🧮 自定义公式引擎")
    st.caption("像Excel一样写公式，但更强大！支持Python/Pandas语法")
    
    # 显示当前列名（方便用户参考）
    with st.expander("📋 查看当前所有列名（复制使用）"):
        col_info = pd.DataFrame({
            '列名': df.columns,
            '类型': df.dtypes.astype(str).values,
            '示例值': [str(df[col].iloc[0]) if len(df) > 0 else '' for col in df.columns],
            '公式引用': [f"df['{col}']" for col in df.columns]
        })
        st.dataframe(col_info, use_container_width=True, hide_index=True)
    
    # 公式模板
    st.markdown("### 📚 公式模板（点击快速填入）")
    
    selected_formula = ""
    
    tabs = st.tabs(list(FORMULA_TEMPLATES.keys()))
    for tab, (category, templates) in zip(tabs, FORMULA_TEMPLATES.items()):
        with tab:
            for name, info in templates.items():
                c1, c2, c3 = st.columns([2, 4, 1])
                with c1:
                    st.markdown(f"**{name}**")
                with c2:
                    st.code(info['formula'], language='python')
                with c3:
                    if st.button("使用", key=f"tmpl_{category}_{name}"):
                        selected_formula = info['formula']
                        st.session_state['formula_input'] = selected_formula
    
    st.markdown("---")
    st.markdown("### ✏️ 编写公式")
    
    # 公式输入
    formula = st.text_area(
        "输入公式",
        value=st.session_state.get('formula_input', ''),
        height=80,
        key="formula_text",
        help="""
        💡 语法说明：
        - 引用列：df['列名'] 
        - 加减乘除：+ - * /
        - 条件：np.where(条件, 真值, 假值)
        - 数学：np.sqrt(), np.log(), np.abs()
        - 文本：df['列名'].str.upper(), .str.contains()
        """
    )
    
    new_col_name = st.text_input("新列名称", value="计算结果", key="formula_col_name")
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁️ 预览结果（前10行）", key="btn_preview", use_container_width=True):
            if formula:
                try:
                    result = eval(formula, {"__builtins__": {}, "np": np, "pd": pd}, {"df": df})
                    if isinstance(result, pd.Series):
                        st.write(result.head(10))
                    else:
                        st.write(result)
                except Exception as e:
                    st.error(f"❌ 公式错误: {e}")
    
    with c2:
        if st.button("✅ 生成新列", key="btn_formula_exec", type="primary", use_container_width=True):
            if formula and new_col_name:
                try:
                    result = eval(formula, {"__builtins__": {}, "np": np, "pd": pd}, {"df": df})
                    if isinstance(result, pd.Series):
                        df[new_col_name] = result
                    else:
                        df[new_col_name] = result
                    st.success(f"✅ 已生成 [{new_col_name}]")
                except Exception as e:
                    st.error(f"❌ 公式错误: {e}")
                    st.info("💡 请检查列名是否正确，格式应为 df['列名']")
    
    # 常用快捷公式（np.where 条件计算器）
    st.markdown("---")
    st.markdown("### 🎛️ 条件计算器（可视化IF公式）")
    st.caption("不用写代码，点点就能完成条件计算")
    
    all_cols = df.columns.tolist()
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    c1, c2, c3 = st.columns(3)
    with c1:
        cond_col = st.selectbox("条件列", all_cols, key="cond_calc_col")
    with c2:
        cond_op = st.selectbox("条件", [">", ">=", "<", "<=", "==", "!=", "包含"], key="cond_calc_op")
    with c3:
        cond_val = st.text_input("条件值", key="cond_calc_val")
    
    c1, c2 = st.columns(2)
    with c1:
        true_formula = st.text_input("条件为真时 = ", value="df['列A'] * 1.1", key="true_formula",
                                      help="可以是固定值如 100，也可以是公式如 df['价格']*0.9")
    with c2:
        false_formula = st.text_input("条件为假时 = ", value="df['列A'] * 0.9", key="false_formula")
    
    result_name = st.text_input("结果列名", value="条件计算结果", key="cond_result_name")
    
    if st.button("✅ 执行条件计算", key="btn_cond_calc", type="primary", use_container_width=True):
        if cond_val:
            try:
                # 构建条件
                col_data = df[cond_col]
                try:
                    cond_val_parsed = float(cond_val)
                    col_data = pd.to_numeric(col_data, errors='coerce')
                except ValueError:
                    cond_val_parsed = cond_val
                
                op_map = {
                    ">": col_data > cond_val_parsed,
                    ">=": col_data >= cond_val_parsed,
                    "<": col_data < cond_val_parsed,
                    "<=": col_data <= cond_val_parsed,
                    "==": col_data == cond_val_parsed,
                    "!=": col_data != cond_val_parsed,
                    "包含": col_data.astype(str).str.contains(str(cond_val_parsed), na=False),
                }
                condition = op_map[cond_op]
                
                # 解析真/假值
                ctx = {"__builtins__": {}, "np": np, "pd": pd, "df": df}
                
                try:
                    true_result = eval(true_formula, ctx)
                except:
                    true_result = true_formula
                
                try:
                    false_result = eval(false_formula, ctx)
                except:
                    false_result = false_formula
                
                df[result_name] = np.where(condition, true_result, false_result)
                st.success(f"✅ 已生成 [{result_name}]")
            except Exception as e:
                st.error(f"执行失败: {e}")
    
    return df