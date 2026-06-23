"""
三个超 Excel 杀手级功能模块
功能一：智能异常值检测（多算法融合 + 一键修复）
功能二：自然语言透视表（"按地区看销售额" → 自动建图表）
功能三：数据血缘追踪（每一列从哪来、经过了什么操作）
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime


# ═══════════════════════════════════════════════════════
# 功能一：智能多算法异常值检测
# Excel 只能手动设颜色，这里自动用 3 种算法投票，可一键修复
# ═══════════════════════════════════════════════════════

def detect_anomalies_multi(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """
    三算法投票异常值检测：
    - IQR（四分位距）：经典、对偏态分布鲁棒
    - Z-Score（3σ）：适合正态分布
    - 孤立森林（Isolation Forest）：适合高维非线性
    返回带 is_anomaly 列的 DataFrame
    """
    result = df[[col]].copy()
    series = pd.to_numeric(df[col], errors="coerce")
    votes = pd.Series(0, index=df.index)

    # ── 算法 1: IQR ──
    q1, q3 = series.quantile(0.25), series.quantile(0.75)
    iqr = q3 - q1
    iqr_mask = (series < q1 - 1.5 * iqr) | (series > q3 + 1.5 * iqr)
    votes += iqr_mask.astype(int)

    # ── 算法 2: Z-Score ──
    mean, std = series.mean(), series.std()
    if std > 0:
        zscore_mask = ((series - mean).abs() / std) > 3
        votes += zscore_mask.astype(int)

    # ── 算法 3: Isolation Forest（仅数据量 > 100 时启用）──
    if len(series.dropna()) > 100:
        try:
            from sklearn.ensemble import IsolationForest
            clean = series.dropna()
            iso = IsolationForest(contamination=0.05, random_state=42, n_jobs=-1)
            preds = iso.fit_predict(clean.values.reshape(-1, 1))
            iso_idx = clean.index[preds == -1]
            votes.loc[iso_idx] += 1
        except ImportError:
            pass  # 没有 sklearn 也能工作

    # ── 投票：至少 2/3 算法认为异常 ──
    result["异常票数"] = votes
    result["是否异常"] = votes >= 2
    result["异常原因"] = ""
    result.loc[iqr_mask, "异常原因"] += "IQR超界 "
    if std > 0:
        result.loc[zscore_mask, "异常原因"] += "Z-Score>3 "
    result["异常原因"] = result["异常原因"].str.strip()
    result["原始值"] = series
    result["IQR下界"] = q1 - 1.5 * iqr
    result["IQR上界"] = q3 + 1.5 * iqr

    return result[result["是否异常"]].copy()


def render_anomaly_detector(df: pd.DataFrame) -> pd.DataFrame | None:
    """
    异常值检测 UI
    返回修复后的 DataFrame（如果用户执行了修复），否则返回 None
    """
    st.subheader("🔬 智能异常值检测")
    st.caption("三算法投票：IQR + Z-Score + 孤立森林，比 Excel 条件格式精准 10 倍")

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if not numeric_cols:
        st.warning("没有数值列")
        return None

    c1, c2 = st.columns([2, 2])
    with c1:
        col = st.selectbox("选择检测列", numeric_cols, key="ano_col")
    with c2:
        fix_method = st.selectbox(
            "异常值处理方式",
            ["仅标记（不修改）", "替换为均值", "替换为中位数", "替换为上下界", "删除异常行"],
            key="ano_fix",
        )

    if st.button("🚀 开始检测", type="primary", key="ano_run"):
        with st.spinner(f"正在用 3 种算法分析「{col}」列..."):
            anomalies = detect_anomalies_multi(df, col)

        if anomalies.empty:
            st.success(f"✅ 「{col}」列未发现异常值，数据质量良好！")
            return None

        n = len(anomalies)
        rate = n / len(df) * 100
        st.error(f"🚨 发现 **{n:,}** 个异常值（占 {rate:.1f}%）")

        # 展示异常值详情
        st.dataframe(
            anomalies[["原始值", "异常原因", "异常票数", "IQR下界", "IQR上界"]].reset_index(),
            use_container_width=True, height=250
        )

        # 分布直方图
        import plotly.graph_objects as go
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        fig = go.Figure()
        fig.add_trace(go.Histogram(x=series, name="全部数据", nbinsx=50,
                                   marker_color="#6C63FF", opacity=0.7))
        fig.add_trace(go.Scatter(x=anomalies["原始值"], y=[0] * len(anomalies),
                                 mode="markers", name="异常点",
                                 marker=dict(color="red", size=8, symbol="x")))
        fig.update_layout(title=f"「{col}」分布图（红 × = 异常值）",
                          height=300, showlegend=True)
        st.plotly_chart(fig, use_container_width=True)

        # 执行修复
        if fix_method != "仅标记（不修改）":
            if st.button(f"✅ 执行「{fix_method}」", type="primary", key="ano_apply"):
                df_fixed = df.copy()
                mask = df.index.isin(anomalies.index)
                series_num = pd.to_numeric(df_fixed[col], errors="coerce")

                if fix_method == "替换为均值":
                    fill_val = series_num[~mask].mean()
                    df_fixed.loc[mask, col] = round(fill_val, 4)
                elif fix_method == "替换为中位数":
                    fill_val = series_num[~mask].median()
                    df_fixed.loc[mask, col] = round(fill_val, 4)
                elif fix_method == "替换为上下界":
                    q1, q3 = series_num.quantile(0.25), series_num.quantile(0.75)
                    iqr = q3 - q1
                    lb, ub = q1 - 1.5 * iqr, q3 + 1.5 * iqr
                    df_fixed.loc[series_num < lb, col] = round(lb, 4)
                    df_fixed.loc[series_num > ub, col] = round(ub, 4)
                elif fix_method == "删除异常行":
                    df_fixed = df_fixed.drop(index=anomalies.index).reset_index(drop=True)

                st.toast(f"✅ 已处理 {n} 个异常值")
                return df_fixed
        else:
            # 仅标记：新增标记列
            df_marked = df.copy()
            df_marked[f"{col}_异常标记"] = "正常"
            df_marked.loc[anomalies.index, f"{col}_异常标记"] = "⚠️异常"
            st.info("已在原数据中添加「异常标记」列，请在数据预览中查看")
            return df_marked

    return None


# ═══════════════════════════════════════════════════════
# 功能二：自然语言透视表
# 输入："各地区销售额对比" → 自动识别列 → 建透视 + 图表
# ═══════════════════════════════════════════════════════

def nl_pivot(query: str, df: pd.DataFrame) -> dict:
    """
    解析自然语言查询，返回透视配置
    不依赖 AI，本地规则匹配（零延迟、离线可用）
    """
    query_lower = query.lower()
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    text_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

    # 识别聚合方式
    agg = "sum"
    agg_label = "求和"
    for kw, fn, label in [
        ("平均", "mean", "平均值"), ("均值", "mean", "平均值"),
        ("计数", "count", "计数"), ("个数", "count", "计数"),
        ("最大", "max", "最大值"), ("最高", "max", "最大值"),
        ("最小", "min", "最小值"), ("最低", "min", "最小值"),
        ("中位", "median", "中位数"),
    ]:
        if kw in query:
            agg, agg_label = fn, label
            break

    # 识别列名：优先匹配用户 query 中出现的列名
    matched_group = None
    matched_value = None

    for col in text_cols:
        if col in query or any(c in query for c in col):
            matched_group = col
            break

    for col in numeric_cols:
        if col in query or any(c in query for c in col):
            matched_value = col
            break

    # fallback：用第一个文本列 / 数值列
    if matched_group is None and text_cols:
        matched_group = text_cols[0]
    if matched_value is None and numeric_cols:
        matched_value = numeric_cols[0]

    return {
        "group_col": matched_group,
        "value_col": matched_value,
        "agg": agg,
        "agg_label": agg_label,
    }


def render_nl_pivot(df: pd.DataFrame):
    """自然语言透视表 UI"""
    st.subheader("🗣️ 自然语言透视表")
    st.caption("输入一句话，秒出透视表 + 图表（不需要 AI，本地瞬间完成）")

    examples = [
        "各地区销售额汇总", "每个销售员的平均订单金额",
        "按月份统计收入", "各部门员工计数",
    ]
    st.markdown("**💬 试试：**")
    cols = st.columns(len(examples))
    for i, ex in enumerate(examples):
        with cols[i]:
            if st.button(ex, key=f"nlpiv_{i}", use_container_width=True):
                st.session_state["_nlpiv_query"] = ex

    query = st.text_input(
        "输入查询",
        value=st.session_state.get("_nlpiv_query", ""),
        placeholder="例：各地区销售额汇总",
        key="nlpiv_q",
        label_visibility="collapsed",
    )

    if not query:
        return

    cfg = nl_pivot(query, df)
    if not cfg["group_col"] or not cfg["value_col"]:
        st.warning("未能识别分组列或数值列，请手动在「汇总」模块操作")
        return

    st.info(f"🤖 理解为：按 **{cfg['group_col']}** 分组，对 **{cfg['value_col']}** 做 **{cfg['agg_label']}**")

    try:
        pivot = (
            df.groupby(cfg["group_col"])[cfg["value_col"]]
            .agg(cfg["agg"])
            .round(2)
            .reset_index()
            .sort_values(cfg["value_col"], ascending=False)
        )

        col_left, col_right = st.columns([1, 2])
        with col_left:
            st.dataframe(pivot, use_container_width=True, height=300)
            csv = pivot.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ 下载", csv,
                               file_name=f"透视_{datetime.now().strftime('%H%M%S')}.csv",
                               mime="text/csv", key="nlpiv_dl")
        with col_right:
            import plotly.express as px
            top_n = min(20, len(pivot))
            fig = px.bar(
                pivot.head(top_n),
                x=cfg["group_col"], y=cfg["value_col"],
                title=f"{cfg['group_col']} × {cfg['value_col']} {cfg['agg_label']}（Top {top_n}）",
                color=cfg["value_col"], color_continuous_scale="Purples",
            )
            fig.update_layout(height=300)
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"生成失败：{e}")


# ═══════════════════════════════════════════════════════
# 功能三：数据血缘追踪
# 记录每一列的来源、每一步操作的 before/after，Excel 根本没有
# ═══════════════════════════════════════════════════════

def init_lineage():
    if "data_lineage" not in st.session_state:
        st.session_state["data_lineage"] = {}  # col_name → {source, ops, created_at}


def record_column_lineage(col_name: str, source: str, operation: str,
                           input_cols: list[str] | None = None):
    """记录一列的来源"""
    init_lineage()
    lineage = st.session_state["data_lineage"]
    if col_name not in lineage:
        lineage[col_name] = {
            "source": source,
            "input_cols": input_cols or [],
            "ops": [],
            "created_at": datetime.now().strftime("%H:%M:%S"),
        }
    lineage[col_name]["ops"].append({
        "time": datetime.now().strftime("%H:%M:%S"),
        "op": operation,
    })


def render_lineage_explorer(df: pd.DataFrame):
    """数据血缘浏览器"""
    st.subheader("🔗 数据血缘追踪")
    st.caption("每一列从哪来、经过了什么操作 — Excel 根本没有这个功能")

    init_lineage()
    lineage = st.session_state.get("data_lineage", {})

    if not lineage:
        st.info("📝 暂无血缘记录。当你在各功能页面生成新列时，血缘会自动记录。")
        return

    for col in df.columns:
        if col in lineage:
            info = lineage[col]
            with st.expander(f"📌 {col}（{info['created_at']} 生成）"):
                st.markdown(f"**来源操作：** {info['source']}")
                if info["input_cols"]:
                    st.markdown(f"**依赖列：** {', '.join(info['input_cols'])}")
                if info["ops"]:
                    st.markdown("**操作历史：**")
                    for op in info["ops"]:
                        st.markdown(f"- `{op['time']}` {op['op']}")
        else:
            with st.expander(f"📄 {col}（原始列）"):
                st.markdown("**来源：** 原始文件上传")

    # 血缘关系图（Mermaid）
    if lineage:
        mermaid_lines = ["graph LR"]
        for col, info in lineage.items():
            for src_col in info.get("input_cols", []):
                safe_src = src_col.replace(" ", "_").replace("（", "").replace("）", "")
                safe_col = col.replace(" ", "_").replace("（", "").replace("）", "")
                mermaid_lines.append(f"    {safe_src}[{src_col}] --> {safe_col}[{col}]")

        if len(mermaid_lines) > 1:
            mermaid_code = "\n".join(mermaid_lines)
            st.markdown("**📊 列依赖关系图：**")
            st.markdown(f"```mermaid\n{mermaid_code}\n```")
