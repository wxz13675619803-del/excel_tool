"""AI助手模块 - 接入DeepSeek/通义千问"""
import streamlit as st
import pandas as pd
import json
import traceback
from openai import OpenAI
from openai import APIError, AuthenticationError, RateLimitError



# streamlit run app.py --theme.primaryColor "#667eea" --theme.backgroundColor "#ffffff"
# # 你的智谱 API Key（就是你贴的那个）
# DEEPSEEK_API_KEY = "bf398e8906eb490ea333590588582a1c.QAnxGk85yKKo5zhv"

# # ✅ 智谱 OpenAI 兼容 Base URL（不是 trialcenter！）
# AI_BASE_URL = "https://open.bigmodel.cn/api/paas/v4"

# # ✅ 智谱支持的模型（推荐用这个）
# AI_MODEL = "glm-4-flash"        # 快 + 有免费额度
# # 或 "glm-4-plus" / "glm-4"  （消耗额度更多）

def get_ai_client():
    """获取AI客户端（优先用 Streamlit secrets，否则用环境变量）"""
    try:
        api_key = st.secrets.get("DEEPSEEK_API_KEY", "bf398e8906eb490ea333590588582a1c.QAnxGk85yKKo5zhv")
        base_url = st.secrets.get("AI_BASE_URL", "https://open.bigmodel.cn/api/paas/v4")
        model = st.secrets.get("AI_MODEL", "glm-4-flash")
    except:
        api_key = ""
        base_url = "https://open.bigmodel.cn/api/paas/v4"
        model = "glm-4-flash"
    
    # 检查环境变量
    if not api_key:
        import os
        api_key = os.environ.get("DEEPSEEK_API_KEY", "")
    
    if not api_key:
        return None, None
    
    try:
        client = OpenAI(api_key=api_key, base_url=base_url)
        return client, model
    except Exception as e:
        st.session_state['ai_error'] = f"初始化失败: {str(e)}"
        return None, None


def handle_ai_error(e: Exception) -> str:
    """统一处理AI错误，返回友好提示"""
    error_map = {
        AuthenticationError: "❌ API密钥无效，请检查密钥配置",
        RateLimitError: "❌ 请求过于频繁，请稍后再试",
        APIError: "❌ AI服务暂时不可用，请稍后再试",
    }
    
    for error_type, message in error_map.items():
        if isinstance(e, error_type):
            return message
    
    # 通用错误处理
    if "network" in str(e).lower() or "timeout" in str(e).lower():
        return "❌ 网络连接失败，请检查网络"
    
    return f"❌ AI调用失败: {str(e)[:100]}"


def local_insight(df: pd.DataFrame) -> str:
    """本地数据洞察（当AI不可用时的回退方案）"""
    insights = []
    
    insights.append("📋 **数据概览**")
    insights.append(f"- 共 {len(df)} 行，{len(df.columns)} 列")
    
    # 数值列统计
    numeric_cols = df.select_dtypes(include='number').columns
    if len(numeric_cols) > 0:
        insights.append("\n📊 **数值列统计**")
        for col in numeric_cols[:3]:
            col_data = df[col]
            insights.append(f"- 「{col}」: 均值={col_data.mean():.2f}, 最大={col_data.max():.2f}, 最小={col_data.min():.2f}")
    
    # 文本列统计
    text_cols = df.select_dtypes(include=['object', 'category']).columns
    if len(text_cols) > 0:
        insights.append("\n📝 **文本列统计**")
        for col in text_cols[:3]:
            col_data = df[col]
            insights.append(f"- 「{col}」: {col_data.nunique()} 个唯一值")
    
    # 缺失值检测
    missing = df.isna().sum().sum()
    if missing > 0:
        insights.append(f"\n⚠️ **数据质量提醒**")
        insights.append(f"- 发现 {missing} 个缺失值")
    
    # 重复值检测
    duplicates = df.duplicated().sum()
    if duplicates > 0:
        insights.append(f"- 发现 {duplicates} 条重复行")
    
    return "\n".join(insights)


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
    """AI数据洞察 - 上传后自动分析（含本地回退）"""
    client, model = get_ai_client()
    
    # 如果AI不可用，使用本地回退
    if not client:
        return "⚠️ AI服务未配置，以下是本地分析结果：\n\n" + local_insight(df)
    
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
        # AI调用失败，使用本地回退
        error_msg = handle_ai_error(e)
        return f"{error_msg}\n\n💡 使用本地分析：\n\n{local_insight(df)}"


def ai_chat_to_code(user_query: str, df: pd.DataFrame) -> dict:
    """AI对话 → 转换为可执行的Pandas代码（含错误处理和回退）"""
    client, model = get_ai_client()
    if not client:
        return {"error": "AI服务未配置，请在Secrets中添加DEEPSEEK_API_KEY"}
    
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

可用列名：{', '.join(df.columns.tolist())}

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
        
        # 解析JSON
        content = response.choices[0].message.content
        
        # 尝试清理可能的markdown代码块标记
        if content.startswith('```'):
            content = content[3:].replace('```', '').strip()
        
        result = json.loads(content)
        return result
        
    except json.JSONDecodeError:
        return {"error": "AI返回格式错误，请重试"}
    except Exception as e:
        error_msg = handle_ai_error(e)
        return {"error": error_msg}


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
