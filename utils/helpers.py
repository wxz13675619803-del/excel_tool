"""工具函数 - 含大数据性能优化"""
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import streamlit as st


@st.cache_data(ttl=3600, show_spinner=False)
def load_excel_cached(file_content: bytes, file_name: str) -> dict:
    """
    缓存加载Excel文件
    使用文件内容的hash作为缓存key，避免重复读取
    """
    from io import BytesIO
    buffer = BytesIO(file_content)
    
    if file_name.endswith('.csv'):
        # CSV自动检测编码
        for enc in ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'latin1']:
            try:
                buffer.seek(0)
                df = pd.read_csv(buffer, encoding=enc)
                return {"Sheet1": df}
            except (UnicodeDecodeError, pd.errors.ParserError):
                continue
        raise ValueError("无法识别CSV编码")
    else:
        xls = pd.ExcelFile(buffer)
        sheets = {}
        for name in xls.sheet_names:
            sheets[name] = pd.read_excel(xls, sheet_name=name)
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
    """高性能Excel导出（支持多sheet + 自动列宽）"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=index, sheet_name=sheet_name[:31])
            worksheet = writer.sheets[sheet_name[:31]]
            
            # 自动列宽
            for col_idx, col in enumerate(df.columns):
                try:
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(str(col))
                    ) + 3
                    worksheet.set_column(col_idx, col_idx, min(max_len, 45))
                except:
                    worksheet.set_column(col_idx, col_idx, 12)
            
            # 表头格式
            header_format = writer.book.add_format({
                'bold': True,
                'bg_color': '#6C63FF',
                'font_color': 'white',
                'border': 1,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
            })
            for col_idx, col in enumerate(df.columns):
                worksheet.write(0, col_idx, col, header_format)
    
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


def detect_data_quality(df: pd.DataFrame) -> dict:
    """智能检测数据质量问题"""
    issues = {
        "warnings": [],
        "errors": [],
        "suggestions": [],
        "stats": {}
    }
    
    # 基本统计
    issues["stats"] = {
        "rows": len(df),
        "columns": len(df.columns),
        "total_cells": len(df) * len(df.columns),
        "missing_cells": int(df.isna().sum().sum()),
        "missing_rate": round(df.isna().sum().sum() / (len(df) * len(df.columns)) * 100, 2),
    }
    
    # 检测重复行
    duplicate_count = df.duplicated().sum()
    if duplicate_count > 0:
        issues["warnings"].append(f"发现 {duplicate_count} 条重复行")
        issues["suggestions"].append("建议使用「去重」功能删除重复行")
    
    # 检测缺失值
    for col in df.columns:
        missing = df[col].isna().sum()
        if missing > 0:
            rate = round(missing / len(df) * 100, 1)
            issues["warnings"].append(f"列「{col}」有 {missing} 个缺失值（占 {rate}%）")
            if rate > 50:
                issues["errors"].append(f"列「{col}」缺失率超过50%，建议检查数据源")
    
    # 检测数据类型问题
    for col in df.columns:
        # 文本列中可能存在数字
        if df[col].dtype == 'object':
            # 检查是否可以转换为数值
            try:
                num_test = pd.to_numeric(df[col], errors='coerce')
                numeric_rate = num_test.notna().sum() / len(df)
                if numeric_rate > 0.8 and numeric_rate < 1.0:
                    issues["warnings"].append(f"列「{col}」大部分是数字但被识别为文本，{int((1-numeric_rate)*100)}% 的值无法转换")
                    issues["suggestions"].append(f"建议将列「{col}」转换为数值类型")
                elif numeric_rate == 1.0:
                    issues["suggestions"].append(f"列「{col}」全部是数字，建议转换为数值类型")
            except:
                pass
        
        # 检测日期格式问题
        try:
            date_test = pd.to_datetime(df[col], errors='coerce')
            date_rate = date_test.notna().sum() / len(df)
            if date_rate > 0.8 and date_rate < 1.0 and df[col].dtype != 'datetime64[ns]':
                issues["warnings"].append(f"列「{col}」大部分是日期但被识别为文本")
                issues["suggestions"].append(f"建议将列「{col}」转换为日期类型")
        except:
            pass
    
    # 检测异常值（数值列）
    numeric_cols = df.select_dtypes(include='number').columns
    for col in numeric_cols:
        try:
            mean = df[col].mean()
            std = df[col].std()
            if std > 0:
                outliers = ((df[col] - mean).abs() > 3 * std).sum()
                if outliers > 0:
                    issues["warnings"].append(f"列「{col}」发现 {outliers} 个疑似异常值（超出3σ范围）")
        except:
            pass
    
    # 检测文本列空格问题
    text_cols = df.select_dtypes(include=['object', 'category']).columns
    for col in text_cols:
        try:
            has_space = df[col].astype(str).str.startswith(' ').sum() + df[col].astype(str).str.endswith(' ').sum()
            if has_space > 0:
                issues["warnings"].append(f"列「{col}」发现 {has_space} 个值首尾有空格")
                issues["suggestions"].append(f"建议对列「{col}」执行去空格操作")
        except:
            pass
    
    # 检测列名问题
    for col in df.columns:
        if ' ' in col:
            issues["suggestions"].append(f"列名「{col}」包含空格，建议重命名")
        if col.startswith(('(', ')', '[', ']')):
            issues["warnings"].append(f"列名「{col}」包含特殊字符，可能导致公式问题")
    
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
