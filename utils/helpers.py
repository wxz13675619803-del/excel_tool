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


def generate_filename(prefix="处理结果"):
    return f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
