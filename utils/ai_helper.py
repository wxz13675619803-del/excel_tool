"""AI助手模块 - 接入DeepSeek/通义千问"""
import streamlit as st
import pandas as pd
import json
from openai import OpenAI


def get_ai_client():
    """获取AI客户端（优先用Streamlit secrets，否则用环境变量）"""
    try:
        api_key = st.secrets.get("DEEPSEEK_API_KEY", "")
        base_url = st.secrets.get("AI_BASE_URL", "https://api.deepseek.com/v1")
        model = st.secrets.get("AI_MODEL", "deepseek-chat")
    except:
        api_key = "sk-193491fb2a94465e9ff85072fe8db692"
        base_url = "https://api.deepseek.com/v1"
        model = "deepseek-chat"
    
    if not api_key:
        return None, None
    
    client = OpenAI(api_key=api_key, base_url=base_url)
    return client, model


def df_summary_for_ai(df: pd.DataFrame, max_rows=5) -> str:
    """生成数据摘要给AI（避免发送整个DataFrame）"""
    summary = {
        "总行数": len(df),
        "总列数": len(df.columns),
        "列信息": []
    }
    for col in df.columns:
        col_info = {
            "列名": col,
            "类型": str(df[col].dtype),
            "非空数": int(df[col].notna().sum()),
            "唯一值数": int(df[col].nunique()),
            "示例值": df[col].dropna().head(3).astype(str).tolist()
        }
        if pd.api.types.is_numeric_dtype(df[col]):
            col_info["最小"] = float(df[col].min()) if df[col].notna().any() else None
            col_info["最大"] = float(df[col].max()) if df[col].notna().any() else None
            col_info["均值"] = float(df[col].mean()) if df[col].notna().any() else None
        summary["列信息"].append(col_info)
    return json.dumps(summary, ensure_ascii=False, indent=2)


def ai_insight(df: pd.DataFrame) -> str:
    """AI数据洞察 - 上传后自动分析"""
    client, model = get_ai_client()
    if not client:
        return "⚠️ 未配置AI密钥"
    
    summary = df_summary_for_ai(df)
    
    prompt = f"""你是Excel数据分析专家。请分析下面的数据，用通俗易懂的中文给出洞察。

数据摘要：
{summary}

请按以下格式输出（要简洁，每点不超过30字）：

📋 **数据是什么**
- （1句话描述数据内容）

🔍 **关键发现**
- （3-5条数据特征和异常）

💡 **建议操作**
- （3-5条具体可执行的操作建议，要明确说出用哪些列做什么）

⚠️ **注意事项**
- （数据质量问题）
"""
    
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=800,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"❌ AI调用失败: {e}"


def ai_chat_to_code(user_query: str, df: pd.DataFrame) -> dict:
    """AI对话 → 转换为可执行的Pandas代码"""
    client, model = get_ai_client()
    if not client:
        return {"error": "未配置AI密钥"}
    
    summary = df_summary_for_ai(df)
    
    prompt = f"""你是Pandas专家。用户上传了一个DataFrame（变量名df），现在用自然语言描述操作需求。
请生成可执行的Python代码。

数据摘要：
{summary}

用户需求：{user_query}

要求：
1. 只返回JSON，格式如下：
{{
  "explanation": "用一句话告诉用户你要做什么（中文）",
  "code": "实际的Pandas代码（操作df，结果赋值给df本身，或保存为变量result）",
  "result_type": "modify_df 或 new_table 或 chart"
}}

2. code中使用的列名必须和数据中的列名完全一致
3. 不要导入任何库（df、pd、np都已经可用）
4. 不要使用print
5. 如果是修改主表，操作df本身；如果是生成新汇总表，保存为result变量
6. 严禁使用exec、eval、import、os、sys等危险函数

示例：
用户："算每个销售员的总销售额"
返回：
{{
  "explanation": "按销售员分组，汇总销售额",
  "code": "result = df.groupby('销售员')['销售额'].sum().reset_index().sort_values('销售额', ascending=False)",
  "result_type": "new_table"
}}
"""
    
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=600,
            response_format={"type": "json_object"}
        )
        result = json.loads(response.choices[0].message.content)
        return result
    except Exception as e:
        return {"error": str(e)}


def ai_explain_result(df_before: pd.DataFrame, df_after: pd.DataFrame, operation: str) -> str:
    """AI解释操作结果"""
    client, model = get_ai_client()
    if not client:
        return ""
    
    changes = {
        "操作": operation,
        "原数据": f"{len(df_before)}行 × {len(df_before.columns)}列",
        "新数据": f"{len(df_after)}行 × {len(df_after.columns)}列",
        "新增的列": list(set(df_after.columns) - set(df_before.columns)),
        "删除的列": list(set(df_before.columns) - set(df_after.columns)),
    }
    
    prompt = f"""用1-2句话简洁解释这个数据操作的结果：
{json.dumps(changes, ensure_ascii=False)}

要求：用通俗语言，不要太技术化。"""
    
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
            max_tokens=150,
        )
        return response.choices[0].message.content
    except:
        return ""


def safe_exec_pandas_code(code: str, df: pd.DataFrame):
    """安全执行Pandas代码"""
    import numpy as np
    
    # 安全检查
    dangerous = ['exec', 'eval', 'import', '__', 'open(', 'os.', 'sys.', 'subprocess']
    for d in dangerous:
        if d in code:
            return None, None, f"❌ 检测到不安全代码: {d}"
    
    try:
        local_vars = {"df": df.copy(), "pd": pd, "np": np, "result": None}
        exec(code, {"__builtins__": {}}, local_vars)
        
        new_df = local_vars.get("df")
        result = local_vars.get("result")
        
        return new_df, result, None
    except Exception as e:
        return None, None, f"❌ 执行失败: {str(e)[:200]}"
