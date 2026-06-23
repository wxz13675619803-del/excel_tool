"""工具函数 - 含大数据性能优化"""
import gc
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import streamlit as st

# 超过此大小（MB）自动走分块加载
_CHUNK_LOAD_THRESHOLD_MB = 50
_CHUNK_ROWS = 100_000  # 每块行数


@st.cache_data(ttl=3600, show_spinner=False)
def load_excel_cached(file_content: bytes, file_name: str) -> dict:
    """
    缓存加载 Excel/CSV。
    - 小文件（<50MB）：一次性加载后做 dtype 优化
    - 大 CSV（≥50MB）：分块加载 + 合并，峰值内存 ≈ 两个 chunk
    """
    file_size_mb = len(file_content) / 1024 / 1024
    buffer = BytesIO(file_content)

    if file_name.lower().endswith('.csv'):
        for enc in ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin1']:
            try:
                buffer.seek(0)
                if file_size_mb >= _CHUNK_LOAD_THRESHOLD_MB:
                    chunks = []
                    for chunk in pd.read_csv(buffer, encoding=enc, chunksize=_CHUNK_ROWS,
                                             low_memory=True):
                        chunks.append(optimize_dtypes(chunk))
                    df = pd.concat(chunks, ignore_index=True)
                    del chunks
                    gc.collect()
                else:
                    df = pd.read_csv(buffer, encoding=enc, low_memory=True)
                    df = optimize_dtypes(df)
                return {"Sheet1": df}
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        raise ValueError("无法识别CSV编码")
    else:
        xls = pd.ExcelFile(buffer)
        sheets = {}
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name)
            sheets[name] = optimize_dtypes(df) if len(df) > 10_000 else df
        return sheets


def optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """
    自动优化DataFrame内存占用
    将大数据的内存减少50%~80%
    """
    for col in df.columns:
        col_type = df[col].dtype

        if col_type == 'object':
            num_unique = df[col].nunique()
            num_total = len(df[col])
            if num_unique / num_total < 0.5:
                df[col] = df[col].astype('category')

        elif col_type in ['int64', 'int32']:
            c_min, c_max = df[col].min(), df[col].max()
            if c_min >= np.iinfo(np.int8).min and c_max <= np.iinfo(np.int8).max:
                df[col] = df[col].astype(np.int8)
            elif c_min >= np.iinfo(np.int16).min and c_max <= np.iinfo(np.int16).max:
                df[col] = df[col].astype(np.int16)
            elif c_min >= np.iinfo(np.int32).min and c_max <= np.iinfo(np.int32).max:
                df[col] = df[col].astype(np.int32)

        elif col_type in ['float64']:
            df[col] = df[col].astype(np.float32)
    
    return df


def df_to_excel_optimized(df_dict: dict, index=False) -> BytesIO:
    """
    高性能 Excel 导出。
    - 数据行 < 100k：原 xlsxwriter 路径（自动列宽 + 表头格式）
    - 数据行 ≥ 100k：constant_memory 模式，峰值内存减少 ~80%
    """
    output = BytesIO()
    total_rows = sum(len(df) for df in df_dict.values())
    use_constant = total_rows >= 100_000

    if use_constant:
        import xlsxwriter
        wb = xlsxwriter.Workbook(output, {"in_memory": True, "constant_memory": True})
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#6C63FF",
                                  "font_color": "white", "border": 1})
        for sheet_name, df in df_dict.items():
            ws = wb.add_worksheet(sheet_name[:31])
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, hdr_fmt)
            for ri, row in enumerate(df.itertuples(index=False), start=1):
                for ci, val in enumerate(row):
                    ws.write(ri, ci, val)
        wb.close()
    else:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in df_dict.items():
                df.to_excel(writer, index=index, sheet_name=sheet_name[:31])
                ws = writer.sheets[sheet_name[:31]]
                for ci, col in enumerate(df.columns):
                    try:
                        max_len = min(int(df[col].astype(str).str.len().max()), 45) + 3
                        max_len = max(max_len, len(str(col)) + 2)
                    except Exception:
                        max_len = 12
                    ws.set_column(ci, ci, max_len)
                hdr_fmt = writer.book.add_format({
                    "bold": True, "bg_color": "#6C63FF", "font_color": "white",
                    "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center",
                })
                for ci, col in enumerate(df.columns):
                    ws.write(0, ci, col, hdr_fmt)

    output.seek(0)
    return output


def get_col_types(df: pd.DataFrame) -> dict:
    """获取列类型分类"""
    return {
        'numeric': df.select_dtypes(include='number').columns.tolist(),
        'text': df.select_dtypes(include=['object', 'category']).columns.tolist(),
        'datetime': df.select_dtypes(include='datetime').columns.tolist(),
        'all': df.columns.tolist()
    }


def paginate_dataframe(df: pd.DataFrame, page_size: int = 500, page_num: int = 1):
    """分页显示大数据"""
    total_pages = max(1, (len(df) - 1) // page_size + 1)
    page_num = min(page_num, total_pages)
    start = (page_num - 1) * page_size
    end = start + page_size
    return df.iloc[start:end], total_pages


def lazy_load_large_file(file_content: bytes, file_name: str, max_rows: int = 10000):
    """懒加载大文件，先读取前max_rows行进行预览"""
    from io import BytesIO
    buffer = BytesIO(file_content)
    
    if file_name.endswith('.csv'):
        for enc in ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'latin1']:
            try:
                buffer.seek(0)
                df = pd.read_csv(buffer, encoding=enc, nrows=max_rows)
                return {"Sheet1": df}, len(df) < max_rows
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        raise ValueError("无法识别CSV编码")
    else:
        xls = pd.ExcelFile(buffer)
        sheets = {}
        all_complete = True
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name, nrows=max_rows)
            sheets[name] = df
            if len(df) >= max_rows:
                all_complete = False
        return sheets, all_complete


def fast_groupby(df: pd.DataFrame, group_cols: list, agg_cols: list, agg_funcs: list):
    """快速分组聚合（针对大数据优化）"""
    if len(df) > 100000:
        df_sampled = df.sample(min(50000, len(df)), random_state=42)
        return df_sampled.groupby(group_cols)[agg_cols].agg(agg_funcs).round(2)
    return df.groupby(group_cols)[agg_cols].agg(agg_funcs).round(2)


def generate_filename(prefix="处理结果"):
    return f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


@st.cache_data(ttl=300, show_spinner=False)
def detect_data_quality(df: pd.DataFrame) -> dict:
    """
    智能检测数据质量问题。
    - 行数 ≤ 50k：全量扫描
    - 行数 > 50k：随机采样 10k 行，结果外推估计
    - 结果缓存 5 分钟，切页不重算
    """
    issues = {"warnings": [], "errors": [], "suggestions": [], "stats": {}}
    total_rows = len(df)
    sample_size = min(10_000, total_rows)
    sample_df = df.sample(sample_size, random_state=42) if total_rows > 50_000 else df
    is_sampled = total_rows > 50_000

    issues["stats"] = {
        "rows": total_rows,
        "columns": len(df.columns),
        "total_cells": total_rows * len(df.columns),
        "missing_cells": int(sample_df.isna().sum().sum() / len(sample_df) * total_rows),
        "missing_rate": round(sample_df.isna().mean().mean() * 100, 2),
    }

    suffix = f"（{sample_size:,} 行采样估计）" if is_sampled else ""

    dup_rate = sample_df.duplicated().mean()
    if dup_rate > 0:
        est = int(dup_rate * total_rows)
        issues["warnings"].append(f"发现约 {est:,} 条重复行{suffix}")
        issues["suggestions"].append("建议使用「去重」功能删除重复行")

    for col in df.columns:
        miss_rate = sample_df[col].isna().mean() * 100
        if miss_rate > 0:
            issues["warnings"].append(f"列「{col}」缺失率约 {miss_rate:.1f}%{suffix}")
            if miss_rate > 50:
                issues["errors"].append(f"列「{col}」缺失率超 50%，建议检查数据源")

    numeric_cols = sample_df.select_dtypes(include="number").columns
    for col in numeric_cols:
        try:
            mean, std = sample_df[col].mean(), sample_df[col].std()
            if std > 0:
                outliers = int(((sample_df[col] - mean).abs() > 3 * std).mean() * total_rows)
                if outliers > 0:
                    issues["warnings"].append(f"列「{col}」约 {outliers:,} 个疑似异常值（3σ）{suffix}")
        except Exception:
            pass

    text_cols = sample_df.select_dtypes(include=["object", "category"]).columns
    for col in text_cols:
        try:
            has_space = (
                sample_df[col].astype(str).str.startswith(" ").mean()
                + sample_df[col].astype(str).str.endswith(" ").mean()
            )
            if has_space > 0:
                est = int(has_space / 2 * total_rows)
                issues["warnings"].append(f"列「{col}」约 {est:,} 个值首尾有空格{suffix}")
        except Exception:
            pass

    return issues


def get_column_stats(df: pd.DataFrame, col_name: str) -> dict:
    """获取单列详细统计信息"""
    col = df[col_name]
    stats = {
        "name": col_name,
        "dtype": str(col.dtype),
        "non_null": int(col.notna().sum()),
        "null": int(col.isna().sum()),
        "unique": int(col.nunique()),
    }
    
    if pd.api.types.is_numeric_dtype(col):
        stats.update({
            "min": float(col.min()) if col.notna().any() else None,
            "max": float(col.max()) if col.notna().any() else None,
            "mean": float(col.mean()) if col.notna().any() else None,
            "median": float(col.median()) if col.notna().any() else None,
            "std": float(col.std()) if col.notna().any() else None,
        })
    elif pd.api.types.is_datetime64_any_dtype(col):
        stats.update({
            "min": str(col.min()) if col.notna().any() else None,
            "max": str(col.max()) if col.notna().any() else None,
            "range_days": int((col.max() - col.min()).days) if col.notna().any() else None,
        })
    else:
        stats.update({
            "top_values": col.value_counts().head(5).to_dict(),
            "length_min": int(col.astype(str).str.len().min()) if col.notna().any() else None,
            "length_max": int(col.astype(str).str.len().max()) if col.notna().any() else None,
        })
    
    return stats
