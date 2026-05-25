"""通用工具函数"""
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime


def load_excel(uploaded_file) -> dict:
    """加载Excel文件，返回所有sheet"""
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for sheet_name in xls.sheet_names:
        sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
    return sheets


def df_to_excel(df_dict: dict, index=False) -> BytesIO:
    """将多个DataFrame导出为Excel（多sheet）"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=index, sheet_name=sheet_name)
            # 自动调整列宽
            worksheet = writer.sheets[sheet_name]
            for col_idx, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 3
                worksheet.set_column(col_idx, col_idx, min(max_len, 50))
    output.seek(0)
    return output


def get_col_types(df: pd.DataFrame) -> dict:
    """获取列的类型分类"""
    return {
        'numeric': df.select_dtypes(include='number').columns.tolist(),
        'text': df.select_dtypes(include='object').columns.tolist(),
        'datetime': df.select_dtypes(include='datetime').columns.tolist(),
        'all': df.columns.tolist()
    }


def generate_filename(prefix="处理结果"):
    """生成带时间戳的文件名"""
    return f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"


def safe_eval_formula(df: pd.DataFrame, formula: str) -> pd.Series:
    """安全的公式解析器"""
    from simpleeval import simple_eval, EvalWithCompoundTypes
    
    # 替换列名为实际数据
    local_vars = {col: df[col] for col in df.columns}
    local_vars['df'] = df
    local_vars['np'] = np
    local_vars['pd'] = pd
    
    try:
        result = eval(formula, {"__builtins__": {}}, local_vars)
        if isinstance(result, pd.Series):
            return result
        else:
            return pd.Series([result] * len(df), index=df.index)
    except Exception as e:
        raise ValueError(f"公式解析错误: {str(e)}")