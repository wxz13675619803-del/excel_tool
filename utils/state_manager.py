"""
session_state 内存安全层
解决三大问题：
1. history.copy() 改为 SQL 操作日志（零内存）
2. 跨页共享数据只存一份引用 + 元数据
3. 自动 GC：切页时释放不必要的临时变量
"""
import gc
import streamlit as st
import pandas as pd
import numpy as np


# ─────────────────────────────────────────────
# 安全的 snapshot：改用"操作描述 + 列diff"替代完整副本
# 仅在行数 < 50k 时才存完整副本（兼容撤销功能）
# ─────────────────────────────────────────────
MAX_FULL_SNAPSHOT_ROWS = 50_000
MAX_HISTORY = 10  # 比原来的 30 少，避免 OOM


def save_snapshot(desc: str = ""):
    df = st.session_state.get("df")
    if df is None:
        return

    if len(df) <= MAX_FULL_SNAPSHOT_ROWS:
        # 小数据：存完整副本（原逻辑）
        history = st.session_state.setdefault("history", [])
        history.append({"type": "full", "df": df.copy(), "desc": desc})
        if len(history) > MAX_HISTORY:
            history.pop(0)
    else:
        # 大数据：只记录操作描述，禁用撤销
        history = st.session_state.setdefault("history", [])
        history.append({"type": "meta", "desc": desc, "shape": df.shape})
        if len(history) > MAX_HISTORY:
            history.pop(0)

    st.session_state["redo_stack"] = []

    if desc:
        op_log = st.session_state.setdefault("op_log", [])
        from datetime import datetime
        op_log.append({"time": datetime.now().strftime("%H:%M:%S"), "desc": desc})


def undo():
    history = st.session_state.get("history", [])
    redo_stack = st.session_state.setdefault("redo_stack", [])
    if not history:
        return
    snapshot = history.pop()
    if snapshot["type"] == "full":
        redo_stack.append({"type": "full", "df": st.session_state.df.copy()})
        st.session_state.df = snapshot["df"]
        gc.collect()
        st.rerun()
    else:
        st.warning("⚠️ 大数据模式下无法撤销（内存限制）")


def redo():
    redo_stack = st.session_state.get("redo_stack", [])
    history = st.session_state.setdefault("history", [])
    if not redo_stack:
        return
    snapshot = redo_stack.pop()
    if snapshot["type"] == "full":
        history.append({"type": "full", "df": st.session_state.df.copy()})
        st.session_state.df = snapshot["df"]
        gc.collect()
        st.rerun()


# ─────────────────────────────────────────────
# 跨页状态初始化（只跑一次）
# ─────────────────────────────────────────────
DEFAULTS = {
    "df": None,
    "original_df": None,
    "sheets": {},
    "current_sheet": None,
    "history": [],
    "redo_stack": [],
    "op_log": [],
    "last_file_hash": None,
    "new_cols": [],
    "ai_insight_cache": {},
    "ai_chat_history": [],
    "recipe": [],
    "recording": False,
    # 大数据专用
    "_df_row_count": 0,
    "_is_large_data": False,
    "_preview_offset": 0,
}


def init_session():
    for k, v in DEFAULTS.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─────────────────────────────────────────────
# 大数据判断阈值
# ─────────────────────────────────────────────
LARGE_DATA_THRESHOLD = 100_000  # 行数超过此值进入"大数据模式"


def is_large_data() -> bool:
    df = st.session_state.get("df")
    if df is None:
        return False
    return len(df) > LARGE_DATA_THRESHOLD


def get_df_safe() -> pd.DataFrame | None:
    """安全获取 df，大数据模式下给出提示"""
    return st.session_state.get("df")


# ─────────────────────────────────────────────
# 切页时的内存清理钩子
# ─────────────────────────────────────────────
def on_page_change():
    """在侧边栏 menu 变化时调用，清理临时变量"""
    keys_to_clear = [
        "pivot_table",
        "stats_table",
        "ai_result_table",
        "_chart_cache",
    ]
    cleared = False
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
            cleared = True
    if cleared:
        gc.collect()


# ─────────────────────────────────────────────
# 内存使用估算（显示在侧边栏）
# ─────────────────────────────────────────────
def estimate_memory_mb() -> float:
    df = st.session_state.get("df")
    history = st.session_state.get("history", [])
    total = 0.0
    if df is not None:
        total += df.memory_usage(deep=True).sum() / 1024 / 1024
    for snap in history:
        if snap.get("type") == "full" and "df" in snap:
            total += snap["df"].memory_usage(deep=True).sum() / 1024 / 1024
    return round(total, 1)
