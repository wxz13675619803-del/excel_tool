"""
📊 Excel智能处理工具 Pro v5.0 - 极简实用版
- 操作即可见：每个页面底部固定显示数据
- 公式极简：一个面板搞定所有运算
- 新列高亮：生成后立刻看到效果
"""
import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from io import BytesIO
from datetime import datetime

from utils.helpers import (
    load_excel_cached, optimize_dtypes, df_to_excel_optimized,
    get_col_types, generate_filename, detect_data_quality, get_column_stats,
    lazy_load_large_file
)
from utils.ai_helper import (
    ai_insight, ai_chat_to_code, ai_explain_result,
    safe_exec_pandas_code, get_ai_client
)
from utils.state_manager import (
    init_session, save_snapshot, undo, redo,
    is_large_data, estimate_memory_mb, on_page_change,
    LARGE_DATA_THRESHOLD
)

# ======================== 页面配置 ========================
st.set_page_config(
    page_title="Excel智能处理工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

css_path = os.path.join(os.path.dirname(__file__), "assets", "style.css")
if os.path.exists(css_path):
    with open(css_path, encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# ======================== Session State ========================
init_session()


# ======================== 工具函数 ========================
def save_snapshot(desc=""):
    from utils.state_manager import save_snapshot as _save
    _save(desc)


def undo():
    from utils.state_manager import undo as _undo
    _undo()


def redo():
    from utils.state_manager import redo as _redo
    _redo()


def mark_new_col(col_name):
    """标记新列用于高亮显示"""
    if col_name not in st.session_state.new_cols:
        st.session_state.new_cols.append(col_name)
    if len(st.session_state.new_cols) > 5:
        st.session_state.new_cols = st.session_state.new_cols[-5:]


def safe_numeric_cols(df):
    return df.select_dtypes(include='number').columns.tolist()


def safe_text_cols(df):
    return df.select_dtypes(include=['object', 'category']).columns.tolist()


def friendly_error(e):
    msg = str(e)
    if "KeyError" in msg or "not in index" in msg:
        return "❌ 找不到指定的列"
    if "could not convert" in msg or "invalid literal" in msg:
        return "❌ 数据类型不匹配"
    if "division by zero" in msg.lower():
        return "❌ 除数不能为0"
    return f"❌ {msg[:80]}"


def show_data_preview(df, title="📋 实时数据预览"):
    """底部固定数据预览区（核心改进：操作后立刻看到结果）"""
    st.markdown("---")
    
    # 高亮显示新列
    new_cols_set = set(st.session_state.new_cols) & set(df.columns)
    
    col_t, col_n = st.columns([3, 1])
    with col_t:
        if new_cols_set:
            highlight_html = " ".join([f'<span class="new-col-badge">✨ {c}</span>' for c in new_cols_set])
            st.markdown(f"#### {title} {highlight_html}", unsafe_allow_html=True)
        else:
            st.markdown(f"#### {title}")
    with col_n:
        n_show = st.selectbox("显示行数", [10, 20, 50, 100, 500], index=1, 
                              key=f"prev_n_{st.session_state.get('menu_key','')}", 
                              label_visibility="collapsed")
    
    # 快捷工具栏
    st.markdown("**⚡ 快捷操作：**")
    quick_btns = st.columns([2, 2, 2, 2, 2, 2])
    with quick_btns[0]:
        if st.button("🗑️ 删除空行", use_container_width=True):
            save_snapshot("删除空行")
            before = len(df)
            df = df.dropna(how='all').reset_index(drop=True)
            st.session_state.df = df
            st.toast(f"✅ 删除 {before - len(df)} 行")
            st.rerun()
    with quick_btns[1]:
        if st.button("🔁 去除重复", use_container_width=True):
            save_snapshot("去除重复")
            before = len(df)
            df = df.drop_duplicates().reset_index(drop=True)
            st.session_state.df = df
            st.toast(f"✅ 删除 {before - len(df)} 条重复")
            st.rerun()
    with quick_btns[2]:
        if st.button("✂️ 去除空格", use_container_width=True):
            save_snapshot("去空格")
            text_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
            for c in text_cols:
                df[c] = df[c].astype(str).str.strip()
            st.session_state.df = df
            st.toast("✅ 已清理")
            st.rerun()
    with quick_btns[3]:
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 CSV", csv,
                          file_name=f"数据_{datetime.now().strftime('%H%M%S')}.csv",
                          mime="text/csv", use_container_width=True, key=f"dl_csv")
    with quick_btns[4]:
        output = df_to_excel_optimized({"Sheet1": df})
        st.download_button("📥 Excel", output,
                          file_name=f"数据_{datetime.now().strftime('%H%M%S')}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          use_container_width=True, key=f"dl_xlsx")
    with quick_btns[5]:
        if st.button("🧹 清除高亮", use_container_width=True):
            st.session_state.new_cols = []
            st.rerun()
    
    # 显示数据，新列加样式
    if new_cols_set:
        def highlight_new(s):
            return ['background-color: rgba(108,99,255,0.15); font-weight:600' 
                    if s.name in new_cols_set else '' for _ in s]
        try:
            styled = df.head(n_show).style.apply(highlight_new, axis=0)
            st.dataframe(styled, use_container_width=True, height=350)
        except:
            st.dataframe(df.head(n_show), use_container_width=True, height=350)
    else:
        st.dataframe(df.head(n_show), use_container_width=True, height=350)


# ======================== 侧边栏 ========================
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:8px 0 16px 0;">
        <div style="font-size:2rem;">📊</div>
        <div style="font-size:1rem; font-weight:700; color:#a78bfa;">Excel 处理工具</div>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("📁 上传文件", type=["xlsx", "xls", "csv"])
    
    if uploaded_file:
        file_bytes = uploaded_file.getvalue()
        file_hash = hash(file_bytes)
        
        if file_hash != st.session_state.last_file_hash:
            st.session_state.last_file_hash = file_hash
            st.session_state.current_sheet = None
            st.session_state.history = []
            st.session_state.redo_stack = []
            st.session_state.op_log = []
            st.session_state.new_cols = []
            st.session_state.file_complete = True
        
        with st.spinner("读取中..."):
            try:
                # 大数据量提示
                file_size_mb = len(file_bytes) / (1024 * 1024)
                if file_size_mb > 10:
                    st.warning(f"⚠️ 文件较大 ({file_size_mb:.1f} MB)，正在进行懒加载...")
                    st.session_state.sheets, st.session_state.file_complete = lazy_load_large_file(
                        file_bytes, uploaded_file.name, max_rows=20000
                    )
                else:
                    st.session_state.sheets = load_excel_cached(file_bytes, uploaded_file.name)
                    st.session_state.file_complete = True
                
                sheet_names = list(st.session_state.sheets.keys())
                
                if len(sheet_names) > 1:
                    selected_sheet = st.selectbox("工作表", sheet_names)
                else:
                    selected_sheet = sheet_names[0]
                
                if selected_sheet != st.session_state.current_sheet:
                    st.session_state.current_sheet = selected_sheet
                    raw_df = st.session_state.sheets[selected_sheet].copy()
                    if len(raw_df) > 5000:
                        raw_df = optimize_dtypes(raw_df)
                    st.session_state.df = raw_df
                    st.session_state.original_df = raw_df.copy()
                    st.session_state.history = []
                    st.session_state.new_cols = []
                
                # 文件未完整加载提示
                if not st.session_state.file_complete:
                    st.info("💡 文件过大，仅加载了前20000行进行预览。如需处理全部数据，请使用较小的文件或在本地运行。")
            except Exception as e:
                st.error(friendly_error(e))
    
    if st.session_state.df is not None:
        df_sb = st.session_state.df
        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("行", f"{len(df_sb):,}")
        with c2:
            st.metric("列", f"{len(df_sb.columns)}")
        # 大数据模式提示
        if len(df_sb) > LARGE_DATA_THRESHOLD:
            st.warning(f"⚡ 大数据模式（>{LARGE_DATA_THRESHOLD//10000}万行）\n撤销历史已禁用以节省内存")
    
    st.markdown("---")
    
    # menu = st.radio(
    #     "nav",
    #     ["🏠 数据", "🧹 清洗", "🧮 计算", "📝 文本",
    #      "📅 日期", "🔎 查找", "📊 汇总", "📜 历史"],
    #     label_visibility="collapsed"
    # )
    menu = st.radio(
    "nav",
    ["🤖 AI 助手", "🏠 数据", "🧹 清洗", "🧮 计算", "📝 文本",
     "📅 日期", "🔎 查找", "📊 汇总", "📈 图表", "📜 历史"],
    label_visibility="collapsed"
    )
    st.session_state['menu_key'] = menu

    # 切页时清理临时缓存，释放内存
    prev_menu = st.session_state.get("_prev_menu")
    if prev_menu and prev_menu != menu:
        on_page_change()
    st.session_state["_prev_menu"] = menu

    # 内存用量显示（大数据模式下）
    mem_mb = estimate_memory_mb()
    if mem_mb > 200:
        st.caption(f"💾 当前内存占用约 {mem_mb:.0f} MB")
    
    st.markdown("---")
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("↩️ 撤销", use_container_width=True,
                     disabled=len(st.session_state.history) == 0):
            undo()
    with c2:
        if st.button("↪️ 重做", use_container_width=True,
                     disabled=len(st.session_state.redo_stack) == 0):
            redo()
    
    if st.button("🔄 重置", use_container_width=True):
        if st.session_state.original_df is not None:
            st.session_state.df = st.session_state.original_df.copy()
            st.session_state.history = []
            st.session_state.new_cols = []
            st.rerun()


# ======================== 欢迎页（极简）========================
if st.session_state.df is None:
    st.markdown("""
    <div style="text-align:center; padding:80px 0;">
        <div style="font-size:5rem;">📊</div>
        <h1 style="font-size:2.5rem; margin:10px 0;">
            <span class="gradient-text">Excel 智能处理工具</span>
        </h1>
        <p style="font-size:1.1rem; color:#8e8ea0; margin:15px 0 40px 0;">
            上传文件 → 选择功能 → 点击执行 → 下载结果
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("👈 请在左侧上传 Excel 或 CSV 文件")
    st.stop()


# ======================== 主区域 ========================
df = st.session_state.df
numeric_cols = safe_numeric_cols(df)
text_cols = safe_text_cols(df)
all_cols = df.columns.tolist()



# ================================================================
#                        🤖 AI 助手（核心新功能）
# ================================================================
if menu == "🤖 AI 助手":
    st.subheader("🤖 AI 智能助手")
    
    client, _ = get_ai_client()
    if not client:
        st.warning("""
        ⚠️ **AI功能未启用**
        
        请在 Streamlit Cloud 的 **App Settings → Secrets** 中添加：
        ```
        DEEPSEEK_API_KEY = "你的密钥"
        ```
        
        👉 [免费获取 DeepSeek API Key](https://platform.deepseek.com)（新用户送¥10额度）
        """)
        st.stop()
    
    ai_tab1, ai_tab2, ai_tab3 = st.tabs(["💬 对话操作", "🔮 数据洞察", "🎬 操作录制"])
    
    # ===== Tab 1: 对话操作 =====
    with ai_tab1:
        st.caption("💡 用大白话告诉 AI 您想做什么，它会自动帮您完成")
        
        # 示例
        st.markdown("**🎯 试试这些（点击直接使用）：**")
        ex_col1, ex_col2 = st.columns(2)
        examples = [
            "算每个销售员的总销售额，按降序排",
            "找出销售额最高的前10条记录",
            "把每个地区的销售额做成汇总",
            "计算每个产品类别的平均单价",
            "找出销售额大于平均值的记录",
            "按月份统计销售额",
        ]
        for i, ex in enumerate(examples):
            with (ex_col1 if i % 2 == 0 else ex_col2):
                if st.button(f"💬 {ex}", key=f"ai_ex_{i}", use_container_width=True):
                    st.session_state['ai_query_input'] = ex
        
        st.markdown("---")
        
        # 输入框
        user_query = st.text_area(
            "🗣️ 告诉 AI 您想做什么",
            value=st.session_state.get('ai_query_input', ''),
            height=80,
            placeholder="例如：算每个销售员的销售额总和，并按降序排列",
            key="ai_query_text"
        )
        
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            run_ai = st.button("🚀 让 AI 执行", type="primary", use_container_width=True)
        with col_btn2:
            if st.button("🗑️ 清空对话历史", use_container_width=True):
                st.session_state.ai_chat_history = []
                st.rerun()
        
        if run_ai and user_query:
            with st.spinner("🤖 AI 思考中..."):
                ai_result = ai_chat_to_code(user_query, df)
            
            if "error" in ai_result:
                st.error(f"❌ {ai_result['error']}")
            else:
                # 显示AI的理解
                st.markdown(f"""
                <div style="background:rgba(108,99,255,0.08); border-left:4px solid #6C63FF;
                            padding:14px; border-radius:8px; margin:10px 0;">
                    🤖 <b>AI 理解：</b>{ai_result.get('explanation', '')}
                </div>
                """, unsafe_allow_html=True)
                
                # 显示代码（折叠）
                with st.expander("👀 查看 AI 生成的代码"):
                    st.code(ai_result.get('code', ''), language='python')
                
                # 执行代码
                code = ai_result.get('code', '')
                result_type = ai_result.get('result_type', 'modify_df')
                
                if code:
                    new_df, result, error = safe_exec_pandas_code(code, df)
                    
                    if error:
                        st.error(error)
                        st.info("💡 试着换个说法，或者用更具体的列名")
                    else:
                        if result_type == "new_table" and result is not None:
                            st.success("✅ 已生成结果表：")
                            st.dataframe(result, use_container_width=True, height=350)
                            
                            # 提供下载
                            csv = result.to_csv(index=False).encode('utf-8-sig')
                            st.download_button("⬇️ 下载结果", csv,
                                              file_name=f"AI结果_{datetime.now().strftime('%H%M%S')}.csv",
                                              mime="text/csv")
                            
                            # 保存到 session 供导出
                            st.session_state['ai_result_table'] = result
                            
                            # 添加到聊天历史
                            st.session_state.ai_chat_history.append({
                                "query": user_query,
                                "explanation": ai_result.get('explanation', ''),
                                "result_shape": f"{len(result)}行 × {len(result.columns)}列"
                            })
                        
                        elif new_df is not None and not new_df.equals(df):
                            save_snapshot(f"AI: {user_query[:30]}")
                            # 标记新列
                            new_cols_set = set(new_df.columns) - set(df.columns)
                            for nc in new_cols_set:
                                mark_new_col(nc)
                            st.session_state.df = new_df
                            st.success("✅ 已应用到主数据")
                            st.toast("✅ 操作完成！查看下方数据预览")
                            st.balloons()
                            st.rerun()
                        else:
                            st.info("ℹ️ AI 执行完成，但数据无变化")
        
        # 显示对话历史
        if st.session_state.ai_chat_history:
            st.markdown("---")
            st.markdown("##### 📜 对话历史")
            for i, chat in enumerate(reversed(st.session_state.ai_chat_history[-5:])):
                st.markdown(f"""
                <div style="background:rgba(128,128,128,0.05); padding:10px 14px;
                            border-radius:8px; margin:6px 0; font-size:0.9rem;">
                    <b>🙋 您：</b> {chat['query']}<br>
                    <b>🤖 AI：</b> {chat['explanation']}<br>
                    <small style="color:#94a3b8;">结果：{chat.get('result_shape', '已应用')}</small>
                </div>
                """, unsafe_allow_html=True)
    
    # ===== Tab 2: 数据洞察 =====
    with ai_tab2:
        st.caption("🔮 让 AI 自动分析您的数据，发现亮点和问题")
        
        # 缓存机制：同一份数据只调用一次AI
        cache_key = f"{len(df)}_{len(df.columns)}_{','.join(df.columns[:5])}"
        
        if st.button("🚀 生成 AI 洞察报告", type="primary"):
            with st.spinner("🤖 AI 正在分析您的数据（约10秒）..."):
                insight = ai_insight(df)
                st.session_state.ai_insight_cache[cache_key] = insight
        
        if cache_key in st.session_state.ai_insight_cache:
            st.markdown(f"""
            <div style="background:linear-gradient(135deg, rgba(108,99,255,0.05), rgba(168,85,247,0.05));
                        border:1px solid rgba(108,99,255,0.2); border-radius:14px;
                        padding:24px; margin:14px 0; line-height:1.8;">
                {st.session_state.ai_insight_cache[cache_key].replace(chr(10), '<br>')}
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("🔄 重新分析"):
                del st.session_state.ai_insight_cache[cache_key]
                st.rerun()
        else:
            st.info("👆 点击上方按钮，AI 会用 10 秒分析您的数据，告诉您：\n"
                   "- 数据是什么内容\n"
                   "- 有什么关键发现\n"
                   "- 推荐做哪些操作\n"
                   "- 数据质量问题")
    
    # ===== Tab 3: 操作录制 =====
    with ai_tab3:
        st.caption("🎬 录制您的操作，保存为「配方」，下次相同数据一键应用")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            if not st.session_state.recording:
                if st.button("🔴 开始录制", type="primary", use_container_width=True):
                    st.session_state.recording = True
                    st.session_state.recipe = []
                    st.toast("🔴 录制中... 现在去操作数据吧")
                    st.rerun()
            else:
                if st.button("⏹️ 停止录制", type="primary", use_container_width=True):
                    st.session_state.recording = False
                    st.toast(f"✅ 录制完成，共 {len(st.session_state.recipe)} 步")
                    st.rerun()
        
        with c2:
            if st.session_state.recipe and st.button("💾 保存配方", use_container_width=True):
                recipe_json = json.dumps(st.session_state.recipe, ensure_ascii=False, indent=2)
                st.download_button("⬇️ 下载配方文件",
                                  recipe_json.encode('utf-8'),
                                  file_name=f"操作配方_{datetime.now().strftime('%Y%m%d')}.json",
                                  mime="application/json")
        
        with c3:
            uploaded_recipe = st.file_uploader("📂 载入配方", type=["json"], 
                                              label_visibility="collapsed")
            if uploaded_recipe:
                try:
                    loaded = json.loads(uploaded_recipe.read().decode('utf-8'))
                    st.session_state.recipe = loaded
                    st.toast(f"✅ 已载入 {len(loaded)} 步操作")
                except Exception as e:
                    st.error(f"载入失败: {e}")
        
        if st.session_state.recording:
            st.warning("🔴 正在录制中... 请在其他页面进行操作，操作会自动记录")
        
        if st.session_state.recipe:
            st.markdown(f"##### 📋 当前配方（{len(st.session_state.recipe)} 步）")
            for i, step in enumerate(st.session_state.recipe):
                st.markdown(f"""
                <div style="background:rgba(108,99,255,0.05); border-left:3px solid #6C63FF;
                            padding:8px 14px; margin:4px 0; border-radius:6px;">
                    <small style="color:#a78bfa;">步骤 {i+1}</small><br>
                    {step.get('desc', '未知操作')}
                </div>
                """, unsafe_allow_html=True)


# ================================================================
#                        📈 图表（新功能）
# ================================================================
elif menu == "📈 图表":
    st.subheader("📈 智能图表")
    
    import plotly.express as px
    
    chart_type = st.radio(
        "类型",
        ["📊 柱状图", "📈 折线图", "🥧 饼图", "🔵 散点图", "📦 箱线图", "🌡️ 热力图"],
        horizontal=True, label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    try:
        if chart_type == "📊 柱状图":
            c1, c2, c3 = st.columns(3)
            with c1:
                x_col = st.selectbox("X 轴（分类）", all_cols, key="bar_x")
            with c2:
                y_col = st.selectbox("Y 轴（数值）", numeric_cols, key="bar_y")
            with c3:
                color_col = st.selectbox("颜色（可选）", ["无"] + all_cols, key="bar_c")
            
            agg = st.selectbox("汇总方式", ["求和", "平均", "计数", "最大", "最小"], key="bar_a")
            agg_func = {"求和": "sum", "平均": "mean", "计数": "count", "最大": "max", "最小": "min"}[agg]
            
            # 自动汇总
            if color_col != "无":
                plot_df = df.groupby([x_col, color_col])[y_col].agg(agg_func).reset_index()
                fig = px.bar(plot_df, x=x_col, y=y_col, color=color_col, barmode='group',
                            title=f"{x_col} 的 {y_col}（{agg}）")
            else:
                plot_df = df.groupby(x_col)[y_col].agg(agg_func).reset_index().sort_values(y_col, ascending=False)
                fig = px.bar(plot_df, x=x_col, y=y_col, title=f"{x_col} 的 {y_col}（{agg}）",
                            color=y_col, color_continuous_scale='Purples')
            
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "📈 折线图":
            c1, c2 = st.columns(2)
            with c1:
                x_col = st.selectbox("X 轴", all_cols, key="line_x")
            with c2:
                y_cols = st.multiselect("Y 轴（可多选）", numeric_cols, default=numeric_cols[:1], key="line_y")
            
            if y_cols:
                plot_df = df.sort_values(x_col)
                fig = px.line(plot_df, x=x_col, y=y_cols, title=f"{x_col} 趋势",
                             markers=True)
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "🥧 饼图":
            c1, c2 = st.columns(2)
            with c1:
                name_col = st.selectbox("分类列", all_cols, key="pie_n")
            with c2:
                val_col = st.selectbox("数值列", numeric_cols, key="pie_v")
            
            plot_df = df.groupby(name_col)[val_col].sum().reset_index()
            fig = px.pie(plot_df, names=name_col, values=val_col, title=f"{name_col} 占比",
                        color_discrete_sequence=px.colors.sequential.Purples_r)
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "🔵 散点图":
            c1, c2, c3 = st.columns(3)
            with c1:
                x_col = st.selectbox("X", numeric_cols, key="sc_x")
            with c2:
                y_col = st.selectbox("Y", numeric_cols, key="sc_y", 
                                    index=1 if len(numeric_cols) > 1 else 0)
            with c3:
                color_col = st.selectbox("颜色", ["无"] + all_cols, key="sc_c")
            
            color = None if color_col == "无" else color_col
            fig = px.scatter(df, x=x_col, y=y_col, color=color,
                           title=f"{x_col} vs {y_col}", trendline="ols" if not color else None)
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "📦 箱线图":
            c1, c2 = st.columns(2)
            with c1:
                y_col = st.selectbox("数值列", numeric_cols, key="box_y")
            with c2:
                x_col = st.selectbox("分类列（可选）", ["无"] + text_cols, key="box_x")
            
            x = None if x_col == "无" else x_col
            fig = px.box(df, x=x, y=y_col, title=f"{y_col} 分布",
                        color_discrete_sequence=['#6C63FF'])
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)
        
        elif chart_type == "🌡️ 热力图":
            if len(numeric_cols) >= 2:
                corr = df[numeric_cols].corr().round(2)
                fig = px.imshow(corr, text_auto=True, title="数值列相关性",
                               color_continuous_scale='Purples', aspect='auto')
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("需要至少2个数值列")
    
    except Exception as e:
        st.error(friendly_error(e))





# ================================================================
#                        🏠 数据
# ================================================================
if menu == "🏠 数据":
    st.subheader("📋 数据查看与编辑")
    
    dt1, dt2 = st.tabs(["📊 数据视图", "🎨 条件格式"])
    
    with dt1:
        # 数据质量检测报告
        with st.expander("🔍 数据质量检测报告", expanded=True):
            quality = detect_data_quality(df)
            stats = quality["stats"]
            
            # 基本统计指标
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                st.metric("总行数", f"{stats['rows']:,}")
            with c2:
                st.metric("总列数", stats['columns'])
            with c3:
                st.metric("总单元格", f"{stats['total_cells']:,}")
            with c4:
                st.metric("缺失单元格", f"{stats['missing_cells']:,}")
            with c5:
                st.metric("缺失率", f"{stats['missing_rate']}%")
            
            # 问题提示
            if quality["errors"]:
                st.error(f"❌ 发现 {len(quality['errors'])} 个严重问题：")
                for e in quality["errors"]:
                    st.markdown(f"- {e}")
            
            if quality["warnings"]:
                st.warning(f"⚠️ 发现 {len(quality['warnings'])} 个警告：")
                for w in quality["warnings"]:
                    st.markdown(f"- {w}")
            
            if quality["suggestions"]:
                st.info(f"💡 建议操作：")
                for s in quality["suggestions"]:
                    st.markdown(f"- {s}")
            
            if not quality["errors"] and not quality["warnings"] and not quality["suggestions"]:
                st.success("🎉 数据质量良好，未发现问题")
        
        st.markdown("---")
        
        # 快捷操作
        qc1, qc2, qc3, qc4 = st.columns(4)
        with qc1:
            if st.button("🗑️ 删除空行", use_container_width=True):
                save_snapshot("删除空行")
                before = len(df)
                df = df.dropna(how='all').reset_index(drop=True)
                st.session_state.df = df
                st.toast(f"✅ 删除 {before - len(df)} 行")
                st.rerun()
        with qc2:
            if st.button("🔁 去除重复", use_container_width=True):
                save_snapshot("去除重复")
                before = len(df)
                df = df.drop_duplicates().reset_index(drop=True)
                st.session_state.df = df
                st.toast(f"✅ 删除 {before - len(df)} 条重复")
                st.rerun()
        with qc3:
            if st.button("✂️ 去除空格", use_container_width=True):
                save_snapshot("去空格")
                for c in text_cols:
                    df[c] = df[c].astype(str).str.strip()
                st.session_state.df = df
                st.toast("✅ 已清理")
                st.rerun()
        with qc4:
            if st.button("🧹 清除高亮", use_container_width=True):
                st.session_state.new_cols = []
                st.rerun()
        
        st.markdown("---")
        
        # 可编辑表格
        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", 
                                height=500, key="main_editor")
        
        if not edited.equals(df):
            if st.button("💾 保存修改", type="primary"):
                save_snapshot("编辑数据")
                st.session_state.df = edited
                st.toast("✅ 已保存")
                st.rerun()
    
    # ===== 条件格式 =====
    with dt2:
        st.markdown("##### 🎨 条件格式设置")
        
        format_type = st.radio("格式类型", 
            ["📊 色阶（数值列）", "🏷️ 图标集", "🎯 自定义条件", "🔔 高亮重复值"],
            horizontal=True, label_visibility="collapsed")
        
        st.markdown("---")
        
        if format_type == "📊 色阶（数值列）":
            col = st.selectbox("选择数值列", numeric_cols, key="cf_col")
            c1, c2 = st.columns(2)
            with c1:
                min_color = st.color_picker("最小值颜色", "#eff6ff", key="cf_min")
            with c2:
                max_color = st.color_picker("最大值颜色", "#4f46e5", key="cf_max")
            
            if st.button("✅ 应用色阶", type="primary", key="cf_b"):
                try:
                    def color_scale(val):
                        col_data = df[col]
                        if pd.isna(val):
                            return ''
                        min_val = col_data.min()
                        max_val = col_data.max()
                        if max_val == min_val:
                            return f'background-color: {min_color}'
                        ratio = (val - min_val) / (max_val - min_val)
                        r = int(int(min_color[1:3], 16) * (1-ratio) + int(max_color[1:3], 16) * ratio)
                        g = int(int(min_color[3:5], 16) * (1-ratio) + int(max_color[3:5], 16) * ratio)
                        b = int(int(min_color[5:7], 16) * (1-ratio) + int(max_color[5:7], 16) * ratio)
                        return f'background-color: rgb({r},{g},{b})'
                    
                    styled = df.style.applymap(color_scale, subset=[col])
                    st.dataframe(styled, use_container_width=True, height=500)
                except Exception as e:
                    st.error(friendly_error(e))
        
        elif format_type == "🏷️ 图标集":
            col = st.selectbox("选择数值列", numeric_cols, key="icon_col")
            icon_type = st.selectbox("图标类型", ["🏆 奖牌", "📈 箭头", "🔴🟡🟢 红绿灯"], key="icon_type")
            
            if st.button("✅ 应用图标", type="primary", key="icon_b"):
                try:
                    col_data = df[col]
                    q1, q2, q3 = col_data.quantile([0.25, 0.5, 0.75])
                    
                    def add_icon(val):
                        if pd.isna(val):
                            return ''
                        if icon_type == "🏆 奖牌":
                            if val >= q3:
                                return '🥇'
                            elif val >= q2:
                                return '🥈'
                            elif val >= q1:
                                return '🥉'
                            else:
                                return ''
                        elif icon_type == "📈 箭头":
                            if val >= q3:
                                return '⬆️'
                            elif val >= q2:
                                return '➡️'
                            else:
                                return '⬇️'
                        else:
                            if val >= q3:
                                return '🟢'
                            elif val >= q1:
                                return '🟡'
                            else:
                                return '🔴'
                    
                    new_col_name = f"{col}_图标"
                    df[new_col_name] = col_data.apply(add_icon)
                    mark_new_col(new_col_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_col_name}」")
                    st.rerun()
                except Exception as e:
                    st.error(friendly_error(e))
        
        elif format_type == "🎯 自定义条件":
            c1, c2 = st.columns(2)
            with c1:
                cond_col = st.selectbox("选择列", all_cols, key="cond_col")
            with c2:
                cond_op = st.selectbox("条件", [">", ">=", "<", "<=", "==", "!=", "包含", "为空"], key="cond_op")
            
            cond_val = ""
            if cond_op != "为空":
                cond_val = st.text_input("条件值", key="cond_val")
            
            highlight_color = st.color_picker("高亮颜色", "#fef3c7", key="hl_color")
            
            if st.button("✅ 预览效果", type="primary", key="cond_b"):
                try:
                    def highlight_cell(val, idx):
                        cell_val = df.iloc[idx, df.columns.get_loc(cond_col)]
                        try:
                            threshold = float(cond_val) if cond_val else 0
                            col_num = pd.to_numeric(cell_val, errors='coerce')
                            use_num = True
                        except:
                            threshold = cond_val
                            col_num = cell_val
                            use_num = False
                        
                        cd = col_num if use_num else cell_val
                        cond_map = {
                            ">": cd > threshold, ">=": cd >= threshold,
                            "<": cd < threshold, "<=": cd <= threshold,
                            "==": cd == threshold, "!=": cd != threshold,
                            "包含": str(cell_val).contains(str(cond_val), na=False),
                            "为空": pd.isna(cell_val) or (str(cell_val).strip() == ''),
                        }
                        if cond_map[cond_op]:
                            return f'background-color: {highlight_color}'
                        return ''
                    
                    styled = df.style.apply(lambda x: [highlight_cell(x.iloc[i], i) for i in range(len(x))], axis=1)
                    st.dataframe(styled, use_container_width=True, height=500)
                except Exception as e:
                    st.error(friendly_error(e))
        
        elif format_type == "🔔 高亮重复值":
            cols = st.multiselect("选择判重列（留空=所有列）", all_cols, key="dup_cols")
            dup_color = st.color_picker("高亮颜色", "#fef3c7", key="dup_color")
            
            if st.button("✅ 高亮重复", type="primary", key="dup_b"):
                try:
                    subset = cols if cols else None
                    duplicates = df.duplicated(subset=subset, keep=False)
                    
                    def highlight_dup(val, idx):
                        if duplicates.iloc[idx]:
                            return f'background-color: {dup_color}'
                        return ''
                    
                    styled = df.style.apply(lambda x: [highlight_dup(x.iloc[i], i) for i in range(len(x))], axis=1)
                    st.dataframe(styled, use_container_width=True, height=500)
                except Exception as e:
                    st.error(friendly_error(e))


# ================================================================
#                        🧹 清洗
# ================================================================
elif menu == "🧹 清洗":
    st.subheader("🧹 数据清洗")
    
    t1, t2, t3, t4, t5 = st.tabs(["缺失值", "列管理", "排序筛选", "类型转换", "⚡ 批量操作"])
    
    with t1:
        missing = df.isna().sum()
        miss_df = pd.DataFrame({
            '列名': missing.index, '缺失数': missing.values,
            '缺失率%': (missing / len(df) * 100).round(1).values
        })
        miss_df = miss_df[miss_df['缺失数'] > 0]
        
        if len(miss_df) > 0:
            st.dataframe(miss_df, use_container_width=True, hide_index=True, height=200)
            
            c1, c2 = st.columns(2)
            with c1:
                fill_cols = st.multiselect("选择列", miss_df['列名'].tolist(), key="fc")
            with c2:
                fill_method = st.selectbox("处理方式", [
                    "删除缺失行", "填充 0", "填充均值", "填充中位数", "填充众数",
                    "向下填充", "向上填充", "填充固定值"
                ], key="fm")
            
            fill_val = ""
            if fill_method == "填充固定值":
                fill_val = st.text_input("填充值", key="fv")
            
            if st.button("✅ 执行", type="primary", key="bf") and fill_cols:
                save_snapshot(f"{fill_method}")
                try:
                    if fill_method == "删除缺失行":
                        before = len(df)
                        df = df.dropna(subset=fill_cols).reset_index(drop=True)
                        st.toast(f"✅ 删除 {before-len(df)} 行")
                    else:
                        for col in fill_cols:
                            if fill_method == "填充 0":
                                df[col] = df[col].fillna(0)
                            elif fill_method == "填充均值" and pd.api.types.is_numeric_dtype(df[col]):
                                df[col] = df[col].fillna(df[col].mean())
                            elif fill_method == "填充中位数" and pd.api.types.is_numeric_dtype(df[col]):
                                df[col] = df[col].fillna(df[col].median())
                            elif fill_method == "填充众数":
                                m = df[col].mode()
                                if len(m):
                                    df[col] = df[col].fillna(m[0])
                            elif fill_method == "向下填充":
                                df[col] = df[col].ffill()
                            elif fill_method == "向上填充":
                                df[col] = df[col].bfill()
                            elif fill_method == "填充固定值":
                                df[col] = df[col].fillna(fill_val)
                        st.toast("✅ 处理完成")
                    st.session_state.df = df
                    st.rerun()
                except Exception as e:
                    st.error(friendly_error(e))
        else:
            st.success("🎉 没有缺失值")
    
    with t2:
        op = st.radio("操作", ["重命名", "删除", "调整顺序", "复制列"], horizontal=True, key="cm")
        
        if op == "重命名":
            c1, c2 = st.columns(2)
            with c1:
                rc = st.selectbox("选择列", all_cols, key="rc")
            with c2:
                nn = st.text_input("新名称", key="nn")
            if st.button("✅ 重命名", type="primary", key="brn") and nn:
                save_snapshot(f"重命名 {rc}")
                df = df.rename(columns={rc: nn})
                st.session_state.df = df
                st.toast(f"✅ {rc} → {nn}")
                st.rerun()
        
        elif op == "删除":
            dc = st.multiselect("选择要删除的列", all_cols, key="dc")
            if st.button("✅ 删除", type="primary", key="bdc") and dc:
                save_snapshot(f"删除列")
                df = df.drop(columns=dc)
                st.session_state.df = df
                st.toast(f"✅ 删除 {len(dc)} 列")
                st.rerun()
        
        elif op == "调整顺序":
            new_order = st.multiselect("按顺序选择所有列", all_cols, default=all_cols, key="co")
            if st.button("✅ 应用", type="primary", key="bco") and len(new_order) == len(all_cols):
                save_snapshot("调整顺序")
                df = df[new_order]
                st.session_state.df = df
                st.rerun()
        
        else:
            c1, c2 = st.columns(2)
            with c1:
                src = st.selectbox("源列", all_cols, key="cp_s")
            with c2:
                new_n = st.text_input("新列名", value=f"{src}_副本", key="cp_n")
            if st.button("✅ 复制", type="primary", key="b_cp") and new_n:
                save_snapshot(f"复制 {src}")
                df[new_n] = df[src].copy()
                mark_new_col(new_n)
                st.session_state.df = df
                st.toast("✅ 已复制")
                st.rerun()
    
    with t3:
        st.markdown("**排序**")
        sc = st.multiselect("排序列", all_cols, key="sc")
        if sc:
            orders = []
            cols_row = st.columns(len(sc))
            for i, col in enumerate(sc):
                with cols_row[i]:
                    o = st.radio(col, ["升序", "降序"], horizontal=True, key=f"so_{col}")
                    orders.append(o == "升序")
            if st.button("✅ 排序", type="primary", key="bs"):
                save_snapshot("排序")
                df = df.sort_values(by=sc, ascending=orders).reset_index(drop=True)
                st.session_state.df = df
                st.toast("✅ 完成")
                st.rerun()
        
        st.markdown("**筛选**")
        fc2 = st.selectbox("筛选列", all_cols, key="fc2")
        
        if df[fc2].dtype in ['object', 'category']:
            uv = df[fc2].dropna().unique().tolist()
            sv = st.multiselect("保留", uv, default=uv, key="sv")
            if st.button("✅ 筛选", type="primary", key="bf2") and sv:
                save_snapshot("筛选")
                df = df[df[fc2].isin(sv)].reset_index(drop=True)
                st.session_state.df = df
                st.toast(f"✅ 保留 {len(df)} 行")
                st.rerun()
        elif pd.api.types.is_numeric_dtype(df[fc2]):
            c1, c2 = st.columns(2)
            with c1:
                mn = st.number_input("最小值", value=float(df[fc2].min()), key="mn")
            with c2:
                mx = st.number_input("最大值", value=float(df[fc2].max()), key="mx")
            if st.button("✅ 筛选", type="primary", key="bf3"):
                save_snapshot("筛选")
                df = df[(df[fc2] >= mn) & (df[fc2] <= mx)].reset_index(drop=True)
                st.session_state.df = df
                st.toast(f"✅ 保留 {len(df)} 行")
                st.rerun()
    
    with t4:
        c1, c2 = st.columns(2)
        with c1:
            tc = st.selectbox("选择列", all_cols, key="tc")
            st.caption(f"当前: `{df[tc].dtype}`")
        with c2:
            tt = st.selectbox("转换为", ["文本", "整数", "小数", "日期"], key="tt")
        
        if st.button("✅ 转换", type="primary", key="bt"):
            save_snapshot(f"转换 {tc}")
            try:
                m = {"文本": lambda: df[tc].astype(str),
                     "整数": lambda: pd.to_numeric(df[tc], errors='coerce').astype('Int64'),
                     "小数": lambda: pd.to_numeric(df[tc], errors='coerce'),
                     "日期": lambda: pd.to_datetime(df[tc], errors='coerce')}
                df[tc] = m[tt]()
                st.session_state.df = df
                st.toast(f"✅ 已转换")
                st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
    
    # ===== 批量操作 =====
    with t5:
        st.markdown("##### ⚡ 一键批量处理")
        
        st.markdown("**📦 批量填充缺失值**")
        c1, c2 = st.columns(2)
        with c1:
            batch_fill_cols = st.multiselect("选择多列", all_cols, key="bfc")
        with c2:
            batch_fill_method = st.selectbox("填充方式", 
                ["填充 0", "填充均值", "填充中位数", "填充众数", "填充空字符串"], 
                key="bfm")
        if st.button("✅ 批量填充", type="primary", key="bbf") and batch_fill_cols:
            save_snapshot("批量填充")
            try:
                for col in batch_fill_cols:
                    if batch_fill_method == "填充 0":
                        df[col] = df[col].fillna(0)
                    elif batch_fill_method == "填充均值" and pd.api.types.is_numeric_dtype(df[col]):
                        df[col] = df[col].fillna(df[col].mean())
                    elif batch_fill_method == "填充中位数" and pd.api.types.is_numeric_dtype(df[col]):
                        df[col] = df[col].fillna(df[col].median())
                    elif batch_fill_method == "填充众数":
                        m = df[col].mode()
                        if len(m):
                            df[col] = df[col].fillna(m[0])
                    elif batch_fill_method == "填充空字符串":
                        df[col] = df[col].fillna("")
                st.session_state.df = df
                st.toast(f"✅ 已批量处理 {len(batch_fill_cols)} 列")
                st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
        
        st.markdown("---")
        
        st.markdown("**🔄 批量类型转换**")
        c1, c2 = st.columns(2)
        with c1:
            batch_type_cols = st.multiselect("选择多列", all_cols, key="btc")
        with c2:
            batch_type = st.selectbox("目标类型", ["文本", "整数", "小数", "日期"], key="btt")
        if st.button("✅ 批量转换", type="primary", key="bbt") and batch_type_cols:
            save_snapshot("批量转换")
            try:
                for col in batch_type_cols:
                    if batch_type == "文本":
                        df[col] = df[col].astype(str)
                    elif batch_type == "整数":
                        df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
                    elif batch_type == "小数":
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                    elif batch_type == "日期":
                        df[col] = pd.to_datetime(df[col], errors='coerce')
                st.session_state.df = df
                st.toast(f"✅ 已批量转换 {len(batch_type_cols)} 列")
                st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
        
        st.markdown("---")
        
        st.markdown("**🧹 批量文本处理**")
        c1, c2 = st.columns(2)
        with c1:
            batch_text_cols = st.multiselect("选择多列", text_cols, key="btxc")
        with c2:
            batch_text_op = st.selectbox("操作", 
                ["去首尾空格", "转小写", "转大写", "首字母大写", "去除所有空格"], 
                key="btxo")
        if st.button("✅ 批量处理", type="primary", key="bbxt") and batch_text_cols:
            save_snapshot("批量文本")
            try:
                for col in batch_text_cols:
                    s = df[col].astype(str)
                    if batch_text_op == "去首尾空格":
                        df[col] = s.str.strip()
                    elif batch_text_op == "转小写":
                        df[col] = s.str.lower()
                    elif batch_text_op == "转大写":
                        df[col] = s.str.upper()
                    elif batch_text_op == "首字母大写":
                        df[col] = s.str.title()
                    elif batch_text_op == "去除所有空格":
                        df[col] = s.str.replace(r'\s+', '', regex=True)
                st.session_state.df = df
                st.toast(f"✅ 已批量处理 {len(batch_text_cols)} 列")
                st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
        
        st.markdown("---")
        
        st.markdown("**🗑️ 批量删除**")
        batch_del_cols = st.multiselect("选择要删除的列", all_cols, key="bdc2")
        if st.button("🗑️ 批量删除列", type="primary", key="bbdc") and batch_del_cols:
            save_snapshot("批量删除")
            df = df.drop(columns=batch_del_cols)
            st.session_state.df = df
            st.toast(f"✅ 已删除 {len(batch_del_cols)} 列")
            st.rerun()
    
    show_data_preview(df)


# ================================================================
#                        🧮 计算（核心：极简公式引擎）
# ================================================================
elif menu == "🧮 计算":
    st.subheader("🧮 计算与公式")
    
    if not numeric_cols:
        st.warning("⚠️ 没有数值列，请先在「清洗」中转换列类型")
        st.stop()
    
    # === 核心：一个面板搞定所有计算 ===
    
    st.markdown("##### 🎯 选择运算类型")
    
    calc_type = st.radio(
        "calc",
        ["✏️ 两列运算", "📊 多列汇总", "🔢 列与数字", "📐 单列处理", "💼 业务公式"],
        horizontal=True,
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    # ============ 两列运算 ============
    if calc_type == "✏️ 两列运算":
        c1, c2, c3 = st.columns([2, 1, 2])
        with c1:
            col_a = st.selectbox("列 A", numeric_cols, key="tc_a")
        with c2:
            op = st.selectbox("运算", ["+", "-", "×", "÷"], key="tc_op")
        with c3:
            col_b = st.selectbox("列 B", numeric_cols, key="tc_b")
        
        # 实时预览
        try:
            ops = {"+": df[col_a]+df[col_b], "-": df[col_a]-df[col_b],
                   "×": df[col_a]*df[col_b], "÷": df[col_a]/df[col_b].replace(0, np.nan)}
            preview_val = ops[op].head(3).round(4).tolist()
            st.markdown(f"""
            <div class="formula-preview">
                📐 <b>{col_a} {op} {col_b}</b> → 前3行结果：
                <code>{preview_val[0]}</code>、<code>{preview_val[1]}</code>、<code>{preview_val[2]}</code>
            </div>
            """, unsafe_allow_html=True)
        except:
            pass
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col_a}{op}{col_b}", key="tc_n",
                                     label_visibility="collapsed", placeholder="新列名")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="tc_btn"):
                save_snapshot(f"{col_a}{op}{col_b}")
                df[new_name] = ops[op].round(4)
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    # ============ 多列汇总 ============
    elif calc_type == "📊 多列汇总":
        c1, c2 = st.columns([2, 1])
        with c1:
            cols = st.multiselect("选择列（多选）", numeric_cols, key="mc_c")
        with c2:
            agg = st.selectbox("汇总方式", ["求和", "平均值", "最大值", "最小值", "中位数"], key="mc_a")
        
        if cols:
            agg_func = {"求和": "sum", "平均值": "mean", "最大值": "max", "最小值": "min", "中位数": "median"}[agg]
            try:
                preview_val = getattr(df[cols], agg_func)(axis=1).head(3).round(4).tolist()
                st.markdown(f"""
                <div class="formula-preview">
                    📐 <b>{agg}({" , ".join(cols)})</b> → 前3行：
                    <code>{preview_val[0]}</code>、<code>{preview_val[1]}</code>、<code>{preview_val[2]}</code>
                </div>
                """, unsafe_allow_html=True)
            except:
                pass
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=agg, key="mc_n",
                                     label_visibility="collapsed", placeholder="新列名")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="mc_btn") and cols:
                save_snapshot(f"{agg} {','.join(cols)}")
                agg_func = {"求和": "sum", "平均值": "mean", "最大值": "max", "最小值": "min", "中位数": "median"}[agg]
                df[new_name] = getattr(df[cols], agg_func)(axis=1).round(4)
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    # ============ 列与数字 ============
    elif calc_type == "🔢 列与数字":
        c1, c2, c3 = st.columns([2, 1, 2])
        with c1:
            col = st.selectbox("列", numeric_cols, key="cn_c")
        with c2:
            op = st.selectbox("运算", ["+", "-", "×", "÷"], key="cn_op")
        with c3:
            num = st.number_input("数字", value=1.0, key="cn_n")
        
        try:
            ops = {"+": df[col]+num, "-": df[col]-num, "×": df[col]*num,
                   "÷": df[col]/num if num != 0 else df[col]*np.nan}
            preview_val = ops[op].head(3).round(4).tolist()
            st.markdown(f"""
            <div class="formula-preview">
                📐 <b>{col} {op} {num}</b> → 前3行：
                <code>{preview_val[0]}</code>、<code>{preview_val[1]}</code>、<code>{preview_val[2]}</code>
            </div>
            """, unsafe_allow_html=True)
        except:
            pass
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col}{op}{num}", key="cn_nm",
                                     label_visibility="collapsed", placeholder="新列名")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="cn_btn"):
                save_snapshot(f"{col}{op}{num}")
                df[new_name] = ops[op].round(4)
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    # ============ 单列处理 ============
    elif calc_type == "📐 单列处理":
        c1, c2 = st.columns([2, 2])
        with c1:
            col = st.selectbox("选择列", numeric_cols, key="sp_c")
        with c2:
            action = st.selectbox("操作", [
                "四舍五入", "向下取整", "向上取整", "绝对值",
                "排名（降序）", "排名（升序）", "累计求和",
                "占比%", "百分比格式", "金额大写"
            ], key="sp_a")
        
        digits = 2
        if action == "四舍五入":
            digits = st.number_input("保留小数位", 0, 10, 2, key="sp_d")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col}_{action}", key="sp_n",
                                     label_visibility="collapsed", placeholder="新列名")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="sp_btn"):
                save_snapshot(f"{action} {col}")
                try:
                    if action == "四舍五入":
                        df[new_name] = df[col].round(digits)
                    elif action == "向下取整":
                        df[new_name] = np.floor(df[col])
                    elif action == "向上取整":
                        df[new_name] = np.ceil(df[col])
                    elif action == "绝对值":
                        df[new_name] = df[col].abs()
                    elif action == "排名（降序）":
                        df[new_name] = df[col].rank(ascending=False, method='min').astype('Int64')
                    elif action == "排名（升序）":
                        df[new_name] = df[col].rank(ascending=True, method='min').astype('Int64')
                    elif action == "累计求和":
                        df[new_name] = df[col].cumsum()
                    elif action == "占比%":
                        total = df[col].sum()
                        df[new_name] = (df[col] / total * 100).round(2) if total != 0 else 0
                    elif action == "百分比格式":
                        df[new_name] = (df[col] * 100).round(2).astype(str) + '%'
                    elif action == "金额大写":
                        def to_rmb(n):
                            try:
                                n = round(float(n), 2)
                                units = ['', '拾', '佰', '仟', '万', '拾', '佰', '仟', '亿']
                                digits_cn = '零壹贰叁肆伍陆柒捌玖'
                                s = str(int(abs(n)*100))
                                result = ''
                                for i, d in enumerate(reversed(s)):
                                    if i == 0:
                                        result = digits_cn[int(d)] + '分' + result if int(d) else result
                                    elif i == 1:
                                        result = digits_cn[int(d)] + '角' + result if int(d) else result
                                    else:
                                        idx = i - 2
                                        if idx < len(units):
                                            result = digits_cn[int(d)] + units[idx] + result
                                return ('负' if n < 0 else '') + (result or '零') + '整'
                            except:
                                return str(n)
                        df[new_name] = df[col].apply(to_rmb)
                    
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
                except Exception as e:
                    st.error(friendly_error(e))
    
    # ============ 业务公式 ============
    else:
        scene = st.selectbox("常用业务场景", [
            "💰 含税金额（金额 × 1.13）",
            "💰 税额（金额 × 税率）",
            "💰 利润（收入 - 成本）",
            "💰 利润率%",
            "💰 折扣价（原价 × 折扣率）",
            "📊 增长率%（环比）",
            "📊 Z-Score 标准化",
            "📊 Min-Max 归一化",
        ], key="biz_s")
        
        if "含税金额" in scene:
            col = st.selectbox("金额列", numeric_cols, key="biz_1c")
            st.markdown(f'<div class="formula-preview">📐 <b>{col} × 1.13</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value="含税金额", key="biz_1n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_1b"):
                    save_snapshot("含税金额")
                    df[new_name] = (df[col] * 1.13).round(2)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif "税额" in scene:
            c1, c2 = st.columns(2)
            with c1:
                col = st.selectbox("金额列", numeric_cols, key="biz_2c")
            with c2:
                rate = st.number_input("税率%", value=13.0, key="biz_2r")
            st.markdown(f'<div class="formula-preview">📐 <b>{col} × {rate}%</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value="税额", key="biz_2n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_2b"):
                    save_snapshot("税额")
                    df[new_name] = (df[col] * rate / 100).round(2)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif scene.startswith("💰 利润（"):
            c1, c2 = st.columns(2)
            with c1:
                col_a = st.selectbox("收入列", numeric_cols, key="biz_3a")
            with c2:
                col_b = st.selectbox("成本列", numeric_cols, key="biz_3b")
            st.markdown(f'<div class="formula-preview">📐 <b>{col_a} - {col_b}</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value="利润", key="biz_3n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_3b2"):
                    save_snapshot("利润")
                    df[new_name] = (df[col_a] - df[col_b]).round(2)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif "利润率" in scene:
            c1, c2 = st.columns(2)
            with c1:
                col_a = st.selectbox("收入列", numeric_cols, key="biz_4a")
            with c2:
                col_b = st.selectbox("成本列", numeric_cols, key="biz_4b")
            st.markdown(f'<div class="formula-preview">📐 <b>({col_a} - {col_b}) / {col_a} × 100</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value="利润率%", key="biz_4n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_4b2"):
                    save_snapshot("利润率")
                    df[new_name] = ((df[col_a] - df[col_b]) / df[col_a].replace(0, np.nan) * 100).round(2)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif "折扣价" in scene:
            c1, c2 = st.columns(2)
            with c1:
                col_a = st.selectbox("原价列", numeric_cols, key="biz_5a")
            with c2:
                col_b = st.selectbox("折扣率列", numeric_cols, key="biz_5b")
            st.markdown(f'<div class="formula-preview">📐 <b>{col_a} × {col_b}</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value="折扣价", key="biz_5n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_5b2"):
                    save_snapshot("折扣价")
                    df[new_name] = (df[col_a] * df[col_b]).round(2)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif "增长率" in scene:
            col = st.selectbox("数值列", numeric_cols, key="biz_6c")
            st.markdown(f'<div class="formula-preview">📐 <b>本期/上期 - 1) × 100</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value=f"{col}_增长率%", key="biz_6n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_6b"):
                    save_snapshot("增长率")
                    df[new_name] = (df[col].pct_change() * 100).round(2)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif "Z-Score" in scene:
            col = st.selectbox("数值列", numeric_cols, key="biz_7c")
            st.markdown(f'<div class="formula-preview">📐 <b>(值 - 均值) / 标准差</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value=f"{col}_标准化", key="biz_7n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_7b"):
                    save_snapshot("Z-Score")
                    df[new_name] = ((df[col] - df[col].mean()) / df[col].std()).round(4)
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif "Min-Max" in scene:
            col = st.selectbox("数值列", numeric_cols, key="biz_8c")
            st.markdown(f'<div class="formula-preview">📐 <b>(值 - 最小值) / (最大值 - 最小值)</b></div>', unsafe_allow_html=True)
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value=f"{col}_归一化", key="biz_8n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="biz_8b"):
                    save_snapshot("Min-Max")
                    mn_v, mx_v = df[col].min(), df[col].max()
                    if mx_v != mn_v:
                        df[new_name] = ((df[col] - mn_v) / (mx_v - mn_v)).round(4)
                        mark_new_col(new_name)
                        st.session_state.df = df
                        st.toast(f"✅ 已生成「{new_name}」")
                        st.rerun()
                    else:
                        st.error("数据没有变化，无法归一化")
    
    show_data_preview(df)


# ================================================================
#                        📝 文本
# ================================================================
elif menu == "📝 文本":
    st.subheader("📝 文本处理")
    
    action = st.radio(
        "act",
        ["🔗 拼接", "✂️ 截取", "🔍 替换", "🧹 去空格", "🔠 大小写",
         "🎯 提取", "📋 分列", "📏 长度", "🔢 编号"],
        horizontal=True, label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    if action == "🔗 拼接":
        c1, c2 = st.columns([3, 1])
        with c1:
            cols = st.multiselect("选择要拼接的列", all_cols, key="cat_c")
        with c2:
            sep = st.text_input("分隔符", key="cat_s", placeholder="留空=直接拼")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value="拼接结果", key="cat_n",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="cat_b") and cols:
                save_snapshot("拼接")
                df[new_name] = df[cols].astype(str).agg(sep.join, axis=1)
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    elif action == "✂️ 截取":
        c1, c2, c3 = st.columns(3)
        with c1:
            col = st.selectbox("选择列", all_cols, key="cut_c")
        with c2:
            mode = st.selectbox("方式", ["从左", "从右", "中间"], key="cut_m")
        with c3:
            cnt = st.number_input("字符数", 1, 1000, 3, key="cut_cnt")
        
        sp = 0
        if mode == "中间":
            sp = st.number_input("起始位置（从0开始）", 0, 1000, 0, key="cut_s")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col}_截取", key="cut_n",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="cut_b"):
                save_snapshot("截取")
                s = df[col].astype(str)
                if mode == "从左":
                    df[new_name] = s.str[:cnt]
                elif mode == "从右":
                    df[new_name] = s.str[-cnt:]
                else:
                    df[new_name] = s.str[sp:sp+cnt]
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    elif action == "🔍 替换":
        c1, c2, c3 = st.columns(3)
        with c1:
            cols = st.multiselect("选择列", all_cols, key="rep_c")
        with c2:
            ft = st.text_input("查找", key="rep_f")
        with c3:
            rt = st.text_input("替换为", key="rep_r")
        regex = st.checkbox("使用正则", key="rep_re")
        
        if st.button("✅ 执行替换（覆盖原列）", type="primary", key="rep_b") and ft and cols:
            save_snapshot("替换")
            for c in cols:
                df[c] = df[c].astype(str).str.replace(ft, rt, regex=regex)
            st.session_state.df = df
            st.toast(f"✅ {len(cols)} 列已替换")
            st.rerun()
    
    elif action == "🧹 去空格":
        c1, c2 = st.columns([3, 2])
        with c1:
            cols = st.multiselect("选择列", all_cols, key="trm_c")
        with c2:
            mode = st.selectbox("模式", ["首尾空格", "所有空格", "多空格合并"], key="trm_m")
        
        if st.button("✅ 执行（覆盖原列）", type="primary", key="trm_b") and cols:
            save_snapshot("去空格")
            for c in cols:
                s = df[c].astype(str)
                if mode == "首尾空格":
                    df[c] = s.str.strip()
                elif mode == "所有空格":
                    df[c] = s.str.replace(r'\s+', '', regex=True)
                else:
                    df[c] = s.str.replace(r'\s+', ' ', regex=True).str.strip()
            st.session_state.df = df
            st.toast("✅ 完成")
            st.rerun()
    
    elif action == "🔠 大小写":
        c1, c2 = st.columns([3, 2])
        with c1:
            cols = st.multiselect("选择列", all_cols, key="case_c")
        with c2:
            mode = st.selectbox("方式", ["全部大写", "全部小写", "首字母大写"], key="case_m")
        
        if st.button("✅ 执行（覆盖原列）", type="primary", key="case_b") and cols:
            save_snapshot("大小写")
            for c in cols:
                s = df[c].astype(str)
                if mode == "全部大写":
                    df[c] = s.str.upper()
                elif mode == "全部小写":
                    df[c] = s.str.lower()
                else:
                    df[c] = s.str.title()
            st.session_state.df = df
            st.toast("✅ 完成")
            st.rerun()
    
    elif action == "🎯 提取":
        c1, c2 = st.columns(2)
        with c1:
            col = st.selectbox("选择列", all_cols, key="ext_c")
        with c2:
            mode = st.selectbox("提取类型", ["数字", "中文", "英文字母", "邮箱", "手机号", "身份证号", "自定义正则"], key="ext_m")
        
        pm = {"数字": r'(\d+\.?\d*)', "中文": r'([\u4e00-\u9fa5]+)', "英文字母": r'([a-zA-Z]+)',
              "邮箱": r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
              "手机号": r'(1[3-9]\d{9})', "身份证号": r'(\d{17}[\dXx])'}
        
        if mode == "自定义正则":
            pattern = st.text_input("正则表达式", key="ext_p")
        else:
            pattern = pm[mode]
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col}_提取", key="ext_n",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="ext_b") and pattern:
                save_snapshot("提取")
                df[new_name] = df[col].astype(str).str.extract(pattern, expand=False)
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    elif action == "📋 分列":
        c1, c2, c3 = st.columns(3)
        with c1:
            col = st.selectbox("选择列", all_cols, key="spl_c")
        with c2:
            sep = st.text_input("分隔符", value=",", key="spl_s")
        with c3:
            n = st.number_input("拆分列数", 2, 20, 3, key="spl_n")
        
        if st.button("✅ 分列", type="primary", key="spl_b") and sep:
            save_snapshot("分列")
            result = df[col].astype(str).str.split(sep, n=n-1, expand=True)
            for i in range(result.shape[1]):
                new_col = f"{col}_part{i+1}"
                df[new_col] = result[i]
                mark_new_col(new_col)
            st.session_state.df = df
            st.toast(f"✅ 拆分为 {result.shape[1]} 列")
            st.rerun()
    
    elif action == "📏 长度":
        col = st.selectbox("选择列", all_cols, key="len_c")
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col}_长度", key="len_n",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="len_b"):
                save_snapshot("长度")
                df[new_name] = df[col].astype(str).str.len()
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    elif action == "🔢 编号":
        c1, c2, c3 = st.columns(3)
        with c1:
            start = st.number_input("起始", 1, 999999, 1, key="num_s")
        with c2:
            prefix = st.text_input("前缀", value="NO.", key="num_p")
        with c3:
            pad = st.number_input("位数", 1, 10, 4, key="num_pad")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value="编号", key="num_n",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="num_b"):
                save_snapshot("编号")
                nums = range(start, start + len(df))
                df[new_name] = [f"{prefix}{str(n).zfill(pad)}" for n in nums]
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    show_data_preview(df)


# ================================================================
#                        📅 日期
# ================================================================
elif menu == "📅 日期":
    st.subheader("📅 日期处理")
    
    action = st.radio("act", ["📅 转换日期", "📆 提取年月日", "⏱️ 计算日期差", "➡️ 日期偏移"],
                     horizontal=True, label_visibility="collapsed")
    
    st.markdown("---")
    
    if action == "📅 转换日期":
        c1, c2 = st.columns(2)
        with c1:
            col = st.selectbox("选择列", all_cols, key="td_c")
        with c2:
            fmt = st.selectbox("格式", ["自动识别", "%Y-%m-%d", "%Y/%m/%d", "%Y%m%d",
                                       "%d/%m/%Y", "%Y年%m月%d日"], key="td_f")
        if st.button("✅ 转换（覆盖原列）", type="primary", key="td_b"):
            save_snapshot("转日期")
            try:
                if fmt == "自动识别":
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                else:
                    df[col] = pd.to_datetime(df[col], format=fmt, errors='coerce')
                st.session_state.df = df
                n_ok = df[col].notna().sum()
                st.toast(f"✅ 转换成功 {n_ok}/{len(df)} 行")
                st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
    
    else:
        date_cols = df.select_dtypes(include='datetime').columns.tolist()
        if not date_cols:
            st.info("💡 请先在「转换日期」中将文本列转为日期格式")
        elif action == "📆 提取年月日":
            c1, c2 = st.columns(2)
            with c1:
                col = st.selectbox("日期列", date_cols, key="de_c")
            with c2:
                parts = st.multiselect("提取", ["年", "月", "日", "星期", "季度", "周数", "年月"],
                                       default=["年", "月"], key="de_p")
            
            if st.button("✅ 生成", type="primary", key="de_b") and parts:
                save_snapshot("提取日期")
                dt = df[col].dt
                pm = {"年": dt.year, "月": dt.month, "日": dt.day, "星期": dt.day_name(),
                      "季度": dt.quarter, "周数": dt.isocalendar().week.astype('Int64'),
                      "年月": dt.strftime('%Y-%m')}
                for p in parts:
                    new_col = f"{col}_{p}"
                    df[new_col] = pm[p]
                    mark_new_col(new_col)
                st.session_state.df = df
                st.toast(f"✅ 已生成 {len(parts)} 列")
                st.rerun()
        
        elif action == "⏱️ 计算日期差":
            c1, c2, c3 = st.columns(3)
            with c1:
                d1 = st.selectbox("开始日期", date_cols, key="dd_1")
            with c2:
                d2 = st.selectbox("结束日期", date_cols, key="dd_2")
            with c3:
                unit = st.selectbox("单位", ["天", "小时", "月", "年"], key="dd_u")
            
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value=f"日期差_{unit}", key="dd_n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="dd_b"):
                    save_snapshot("日期差")
                    delta = df[d2] - df[d1]
                    um = {"天": delta.dt.days, "小时": (delta.dt.total_seconds()/3600).round(1),
                          "月": (delta.dt.days/30.44).round(1), "年": (delta.dt.days/365.25).round(2)}
                    df[new_name] = um[unit]
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
        
        elif action == "➡️ 日期偏移":
            c1, c2, c3 = st.columns(3)
            with c1:
                col = st.selectbox("日期列", date_cols, key="do_c")
            with c2:
                v = st.number_input("偏移量", value=7, key="do_v")
            with c2:
                unit = st.selectbox("单位", ["天", "周", "月", "年"], key="do_u")
            
            c1, c2 = st.columns([3, 1])
            with c1:
                new_name = st.text_input("新列名", value=f"{col}_偏移", key="do_n",
                                         label_visibility="collapsed")
            with c2:
                if st.button("✅ 生成", type="primary", use_container_width=True, key="do_b"):
                    save_snapshot("日期偏移")
                    v_int = int(v)
                    om = {"天": pd.Timedelta(days=v_int), "周": pd.Timedelta(weeks=v_int),
                          "月": pd.DateOffset(months=v_int), "年": pd.DateOffset(years=v_int)}
                    df[new_name] = df[col] + om[unit]
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
    
    show_data_preview(df)


# ================================================================
#                        🔎 查找
# ================================================================
elif menu == "🔎 查找":
    st.subheader("🔎 条件判断与查找")
    
    action = st.radio("act", ["❓ IF 条件", "🔀 多条件分类", "🔎 VLOOKUP 跨表", "🏷️ 去重"],
                     horizontal=True, label_visibility="collapsed")
    
    st.markdown("---")
    
    if action == "❓ IF 条件":
        c1, c2, c3 = st.columns(3)
        with c1:
            col = st.selectbox("判断列", all_cols, key="if_c")
        with c2:
            op = st.selectbox("条件", [">", ">=", "<", "<=", "==", "!=", "包含", "为空"], key="if_op")
        with c3:
            val = ""
            if op != "为空":
                val = st.text_input("条件值", key="if_v")
        
        c1, c2 = st.columns(2)
        with c1:
            tv = st.text_input("满足时显示", value="是", key="if_t")
        with c2:
            fv = st.text_input("不满足时显示", value="否", key="if_f")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value="判断结果", key="if_n",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="if_b"):
                save_snapshot("IF")
                col_data = df[col]
                try:
                    threshold = float(val) if val else 0
                    col_num = pd.to_numeric(col_data, errors='coerce')
                    use_num = True
                except:
                    threshold = val
                    col_num = col_data
                    use_num = False
                
                cd = col_num if use_num else col_data
                cond_map = {
                    ">": cd > threshold, ">=": cd >= threshold,
                    "<": cd < threshold, "<=": cd <= threshold,
                    "==": cd == threshold, "!=": cd != threshold,
                    "包含": col_data.astype(str).str.contains(str(val), na=False),
                    "为空": col_data.isna() | (col_data.astype(str).str.strip() == ''),
                }
                df[new_name] = np.where(cond_map[op], tv, fv)
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    elif action == "🔀 多条件分类":
        col = st.selectbox("判断列", all_cols, key="ifs_c")
        n_cond = st.number_input("条件数量", 2, 10, 3, key="ifs_n")
        
        conditions = []
        for i in range(int(n_cond)):
            c1, c2, c3 = st.columns(3)
            with c1:
                op = st.selectbox(f"#{i+1} 运算符", [">=", ">", "<=", "<", "==", "包含"], key=f"ifs_o_{i}")
            with c2:
                v = st.text_input(f"#{i+1} 值", key=f"ifs_v_{i}")
            with c3:
                r = st.text_input(f"#{i+1} 结果", key=f"ifs_r_{i}")
            conditions.append((op, v, r))
        
        default = st.text_input("默认值（都不满足时）", value="其他", key="ifs_d")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value="分类", key="ifs_nm",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="ifs_b"):
                save_snapshot("多条件")
                result = pd.Series(default, index=df.index)
                for op, v, r in reversed(conditions):
                    if not v or not r:
                        continue
                    try:
                        vn = float(v)
                        cd = pd.to_numeric(df[col], errors='coerce')
                    except:
                        vn = v
                        cd = df[col].astype(str)
                    
                    om = {">=": cd >= vn, ">": cd > vn, "<=": cd <= vn, "<": cd < vn,
                          "==": cd == vn, "包含": df[col].astype(str).str.contains(str(v), na=False)}
                    result[om[op]] = r
                
                df[new_name] = result
                mark_new_col(new_name)
                st.session_state.df = df
                st.toast(f"✅ 已生成「{new_name}」")
                st.rerun()
    
    elif action == "🔎 VLOOKUP 跨表":
        lf = st.file_uploader("上传查找表", type=["xlsx", "xls", "csv"], key="vl_f")
        
        if lf:
            try:
                df_lk = pd.read_csv(lf) if lf.name.endswith('.csv') else pd.read_excel(lf)
                
                st.caption("查找表预览：")
                st.dataframe(df_lk.head(5), use_container_width=True, height=180)
                
                c1, c2 = st.columns(2)
                with c1:
                    mk = st.selectbox("主表匹配列", all_cols, key="vl_mk")
                with c2:
                    lk = st.selectbox("查找表匹配列", df_lk.columns.tolist(), key="vl_lk")
                
                rc = st.multiselect("要带回来的列", [c for c in df_lk.columns if c != lk], key="vl_rc")
                
                if st.button("✅ 执行 VLOOKUP", type="primary", key="vl_b") and rc:
                    save_snapshot("VLOOKUP")
                    lk_sub = df_lk[[lk] + rc].drop_duplicates(subset=lk)
                    rename_map = {col: f"{col}_查找" for col in rc if col in df.columns}
                    if rename_map:
                        lk_sub = lk_sub.rename(columns=rename_map)
                    df = df.merge(lk_sub, left_on=mk, right_on=lk, how='left')
                    if lk != mk and lk in df.columns:
                        df = df.drop(columns=[lk])
                    
                    for col in rc:
                        final_col = rename_map.get(col, col)
                        mark_new_col(final_col)
                    
                    st.session_state.df = df
                    st.toast(f"✅ 匹配 {len(rc)} 列")
                    st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
    
    elif action == "🏷️ 去重":
        c1, c2 = st.columns([3, 2])
        with c1:
            cols = st.multiselect("判重列（留空=所有列）", all_cols, key="dp_c")
        with c2:
            mode = st.selectbox("操作", ["标记重复", "删除（保留首条）", "删除（保留末条）"], key="dp_m")
        
        if st.button("✅ 执行", type="primary", key="dp_b"):
            save_snapshot(f"去重 {mode}")
            subset = cols if cols else None
            if mode == "标记重复":
                df["是否重复"] = df.duplicated(subset=subset, keep=False).map({True: "重复", False: "唯一"})
                mark_new_col("是否重复")
            elif "首条" in mode:
                b = len(df)
                df = df.drop_duplicates(subset=subset, keep='first').reset_index(drop=True)
                st.toast(f"✅ {b}→{len(df)} 行")
            else:
                b = len(df)
                df = df.drop_duplicates(subset=subset, keep='last').reset_index(drop=True)
                st.toast(f"✅ {b}→{len(df)} 行")
            st.session_state.df = df
            st.rerun()
    
    show_data_preview(df)


# ================================================================
#                        📊 汇总
# ================================================================
elif menu == "📊 汇总":
    st.subheader("📊 数据汇总")
    
    action = st.radio("act", ["📊 分组汇总（透视表）", "🎯 条件汇总", "📈 描述统计", "📦 数值分箱"],
                     horizontal=True, label_visibility="collapsed")
    
    st.markdown("---")
    
    if action == "📊 分组汇总（透视表）":
        st.markdown("**📐 高级透视表设置**")
        
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            row_group = st.multiselect("行分组", all_cols, key="pv_row")
        with c2:
            col_group = st.multiselect("列分组（可选）", all_cols, key="pv_col")
        with c3:
            value_cols = st.multiselect("汇总列（数值）", numeric_cols, key="pv_val")
        
        agg_funcs = st.multiselect("汇总方式", 
            ["求和", "平均值", "计数", "最大值", "最小值", "标准差", "方差", "中位数"],
            default=["求和"], key="pv_agg")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            sort_by = st.selectbox("排序依据", ["无"] + value_cols, key="pv_sort")
        with c2:
            sort_order = st.selectbox("排序", ["降序", "升序"], key="pv_order")
        
        show_total = st.checkbox("显示汇总行", value=True, key="pv_total")
        
        if st.button("✅ 生成透视表", type="primary", key="pv_b") and row_group and value_cols:
            try:
                fm = {"求和": "sum", "平均值": "mean", "计数": "count", "最大值": "max", 
                      "最小值": "min", "标准差": "std", "方差": "var", "中位数": "median"}
                funcs = [fm[f] for f in agg_funcs]
                
                # 构建透视表
                if col_group:
                    pivot = df.pivot_table(
                        index=row_group,
                        columns=col_group,
                        values=value_cols,
                        aggfunc=funcs,
                        margins=show_total,
                        margins_name="合计"
                    ).round(2)
                else:
                    pivot = df.groupby(row_group)[value_cols].agg(funcs).round(2)
                
                # 处理多级列名
                if isinstance(pivot.columns, pd.MultiIndex):
                    pivot.columns = ['_'.join(str(c) for c in col).strip() for col in pivot.columns]
                
                pivot = pivot.reset_index()
                
                # 排序
                if sort_by != "无":
                    sorted_cols = [c for c in pivot.columns if sort_by in c]
                    if sorted_cols:
                        pivot = pivot.sort_values(sorted_cols[0], ascending=(sort_order == "升序"))
                
                st.markdown("##### 📋 汇总结果")
                st.dataframe(pivot, use_container_width=True, height=400)
                st.session_state['pivot_table'] = pivot
                
                # 导出选项
                csv = pivot.to_csv(index=False).encode('utf-8-sig')
                st.download_button("⬇️ 下载透视表", csv,
                                  file_name=f"透视表_{datetime.now().strftime('%H%M%S')}.csv",
                                  mime="text/csv")
                
                if st.checkbox("合并回主表", key="pv_m"):
                    save_snapshot("合并汇总")
                    df = df.merge(pivot, on=row_group, how='left', suffixes=('', '_汇总'))
                    st.session_state.df = df
                    st.toast("✅ 已合并")
                    st.rerun()
            except Exception as e:
                st.error(friendly_error(e))
    
    elif action == "🎯 条件汇总":
        c1, c2, c3 = st.columns(3)
        with c1:
            gc = st.selectbox("分组列", all_cols, key="si_g")
        with c2:
            vc = st.selectbox("汇总列", numeric_cols, key="si_v")
        with c3:
            fn = st.selectbox("方式", ["求和", "平均值", "计数", "最大值", "最小值"], key="si_f")
        
        if st.button("✅ 生成", type="primary", key="si_b"):
            save_snapshot(f"{fn} by {gc}")
            fm = {"求和": "sum", "平均值": "mean", "计数": "count", "最大值": "max", "最小值": "min"}
            new_col = f"{vc}_{fn}_按{gc}"
            df[new_col] = df.groupby(gc)[vc].transform(fm[fn]).round(2)
            mark_new_col(new_col)
            st.session_state.df = df
            st.toast(f"✅ 已生成「{new_col}」")
            st.rerun()
    
    elif action == "📈 描述统计":
        if numeric_cols:
            stats = df[numeric_cols].describe().round(2).T
            stats.columns = ['计数', '均值', '标准差', '最小', '25%', '中位数', '75%', '最大']
            stats['总和'] = df[numeric_cols].sum().round(2)
            stats['缺失'] = df[numeric_cols].isna().sum()
            st.dataframe(stats, use_container_width=True, height=400)
            
            if st.button("💾 添加到导出文件"):
                st.session_state['stats_table'] = stats.reset_index().rename(columns={'index': '列名'})
                st.toast("✅ 已添加")
        else:
            st.warning("无数值列")
    
    elif action == "📦 数值分箱":
        c1, c2 = st.columns(2)
        with c1:
            col = st.selectbox("选择列", numeric_cols, key="bn_c")
        with c2:
            mode = st.selectbox("方式", ["等距分箱", "自定义分界", "等频分箱"], key="bn_m")
        
        nb = 5
        be = "0,60,80,100"
        if mode == "等距分箱":
            nb = st.number_input("分段数", 2, 20, 5, key="bn_n")
        elif mode == "自定义分界":
            be = st.text_input("分界点（逗号分隔）", value="0,60,80,100", key="bn_e")
        else:
            nb = st.number_input("分组数", 2, 20, 5, key="bn_q")
        
        bl = st.text_input("自定义标签（可选，逗号分隔）", key="bn_l")
        
        c1, c2 = st.columns([3, 1])
        with c1:
            new_name = st.text_input("新列名", value=f"{col}_分段", key="bn_nm",
                                     label_visibility="collapsed")
        with c2:
            if st.button("✅ 生成", type="primary", use_container_width=True, key="bn_b"):
                save_snapshot("分箱")
                try:
                    labels = [x.strip() for x in bl.split(",")] if bl else None
                    if mode == "等距分箱":
                        df[new_name] = pd.cut(df[col], bins=nb, labels=labels)
                    elif mode == "自定义分界":
                        edges = [float(x.strip()) for x in be.split(",")]
                        df[new_name] = pd.cut(df[col], bins=edges, labels=labels, include_lowest=True)
                    else:
                        df[new_name] = pd.qcut(df[col], q=nb, labels=labels, duplicates='drop')
                    mark_new_col(new_name)
                    st.session_state.df = df
                    st.toast(f"✅ 已生成「{new_name}」")
                    st.rerun()
                except Exception as e:
                    st.error(friendly_error(e))
    
    show_data_preview(df)


# ================================================================
#                        📜 历史
# ================================================================
elif menu == "📜 历史":
    st.subheader("📜 操作历史")
    
    if st.session_state.op_log:
        st.caption(f"共 {len(st.session_state.op_log)} 个操作")
        
        for i, op in enumerate(reversed(st.session_state.op_log)):
            st.markdown(f"""
            <div style="background:rgba(108,99,255,0.05); border-left:3px solid #6C63FF;
                        padding:10px 14px; margin:6px 0; border-radius:8px;">
                <span style="color:#a78bfa; font-size:0.85rem;">#{len(st.session_state.op_log)-i} · {op['time']}</span><br>
                {op['desc']}
            </div>
            """, unsafe_allow_html=True)
        
        if st.button("🗑️ 清空历史"):
            st.session_state.op_log = []
            st.rerun()
    else:
        st.info("暂无操作记录")


# ================================================================
#                        💾 导出
# ================================================================
st.markdown("---")
st.markdown("### 💾 导出数据")

df = st.session_state.df

with st.expander("⚙️ 导出选项", expanded=False):
    
    # === 列选择 ===
    col_mode = st.radio("列范围", ["全部列", "选择列", "排除列"], horizontal=True, key="exp_cmode")
    
    if col_mode == "选择列":
        export_cols = st.multiselect("选择要导出的列", df.columns.tolist(),
                                      default=df.columns.tolist(), key="exp_cols")
    elif col_mode == "排除列":
        exc = st.multiselect("选择要排除的列", df.columns.tolist(), key="exp_exc")
        export_cols = [c for c in df.columns if c not in exc]
    else:
        export_cols = df.columns.tolist()
    
    # === 行选择 ===
    row_mode = st.radio("行范围", ["全部行", "前N行", "后N行", "指定范围"], horizontal=True, key="exp_rmode")
    
    export_df = df[export_cols].copy() if export_cols else df.copy()
    
    if row_mode == "前N行":
        n = st.number_input("前", 1, len(df), min(100, len(df)), key="exp_h")
        export_df = export_df.head(int(n))
    elif row_mode == "后N行":
        n = st.number_input("后", 1, len(df), min(100, len(df)), key="exp_t")
        export_df = export_df.tail(int(n))
    elif row_mode == "指定范围":
        c1, c2 = st.columns(2)
        with c1:
            s = st.number_input("从", 1, len(df), 1, key="exp_s")
        with c2:
            e = st.number_input("到", 1, len(df), min(100, len(df)), key="exp_e")
        if s <= e:
            export_df = export_df.iloc[int(s)-1:int(e)]

c1, c2, c3 = st.columns([3, 2, 1])
with c1:
    out_name = st.text_input("文件名", value=generate_filename(), 
                              key="out_n", label_visibility="collapsed")
with c2:
    out_fmt = st.selectbox("格式", ["Excel (.xlsx)", "CSV (.csv)"], 
                            key="out_f", label_visibility="collapsed")
with c3:
    inc_idx = st.checkbox("行号", False, key="out_i")

# 默认是全部数据
if 'export_df' not in dir():
    export_df = df

st.caption(f"📦 将导出 **{len(export_df):,} 行 × {len(export_df.columns)} 列**")

if out_fmt == "Excel (.xlsx)":
    export_sheets = {"数据": export_df}
    if 'pivot_table' in st.session_state:
        export_sheets["汇总表"] = st.session_state.pivot_table
    if 'stats_table' in st.session_state:
        export_sheets["统计"] = st.session_state.stats_table
    
    try:
        excel_data = df_to_excel_optimized(export_sheets, index=inc_idx)
        st.download_button("⬇️ 下载 Excel", data=excel_data,
                           file_name=f"{out_name}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True)
    except Exception as e:
        st.error(friendly_error(e))
else:
    csv = export_df.to_csv(index=inc_idx).encode('utf-8-sig')
    st.download_button("⬇️ 下载 CSV", data=csv,
                       file_name=f"{out_name}.csv", mime="text/csv",
                       type="primary", use_container_width=True)
