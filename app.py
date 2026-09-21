import streamlit as st
import requests
import docx
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import re
import time
from datetime import datetime

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(page_title="关键词深度研究档案生成器", layout="wide")
st.title("📚 关键词深度研究档案生成器")
st.markdown("输入任意关键词，自动生成结构化调研报告（Word 文档）")

# ============================================================
# API 密钥读取逻辑（优先从 Streamlit Secrets 读取）
# ============================================================
def get_secret(key_name):
    try:
        return st.secrets[key_name]
    except Exception:
        return None

deepseek_key = get_secret("DEEPSEEK_API_KEY")
bocha_key = get_secret("BOCHA_API_KEY")

# 如果 secrets 里没有，显示侧边栏输入框（本地测试用）
if not deepseek_key or not bocha_key:
    st.sidebar.header("🔑 API 密钥配置")
    if not deepseek_key:
        deepseek_key = st.sidebar.text_input("DeepSeek API Key", type="password")
    if not bocha_key:
        bocha_key = st.sidebar.text_input("Bocha Search API Key", type="password")
    st.sidebar.info("密钥仅保存在当前会话中。")

# ============================================================
# 博查搜索 API
# ============================================================
def bocha_search(query: str, api_key: str, count: int = 10) -> list:
    """调用博查搜索 API，返回搜索结果列表"""
    if not api_key:
        return []
    
    url = "https://api.bochaai.com/v1/web-search"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "query": query,
        "count": count,
        "summary": True,       # 开启摘要，便于 AI 提取关键信息
        "freshness": "noLimit" # 不限时间，获取最新结果
    }
    
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        
        results = []
        web_pages = data.get("data", {}).get("webPages", {}).get("value", [])
        for item in web_pages:
            results.append({
                "title": item.get("name", ""),
                "url": item.get("url", ""),
                "snippet": item.get("summary", "") or item.get("snippet", ""),
            })
        return results
    except Exception as e:
        st.error(f"博查搜索失败：{e}")
        return []

# ============================================================
# DeepSeek 生成 API
# ============================================================
def deepseek_generate(prompt: str, api_key: str) -> str:
    if not api_key:
        return ""
    url = "https://api.deepseek.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 4000
    }
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=90)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        st.error(f"DeepSeek API 调用失败：{e}")
        return ""

# ============================================================
# Word 文档生成
# ============================================================
def create_word_doc(title: str, content: str) -> BytesIO:
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    heading = doc.add_heading(title, level=0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph(f"生成日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    paragraphs = content.split('\n\n')
    main_title_pattern = re.compile(r'^[一二三四五六七八九十]、')

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue
        if main_title_pattern.match(para):
            doc.add_heading(para, level=2)
        else:
            p = doc.add_paragraph(para)
            if re.match(r'^(\d+\.|\•|\-|\—)', para) and p.runs:
                p.runs[0].bold = True

    doc.add_page_break()
    doc.add_heading("附录：数据来源", level=2)
    doc.add_paragraph("本报告基于博查搜索返回的公开结果，并结合 DeepSeek 模型知识生成。")
    doc.add_paragraph("具体引用链接请查阅报告中提及的 URL。")

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ============================================================
# 构建提示词
# ============================================================
def build_prompt(keyword: str, results: list) -> str:
    if results:
        search_text = ""
        for idx, r in enumerate(results, 1):
            search_text += f"{idx}. {r['title']}\n   URL: {r['url']}\n   摘要: {r['snippet']}\n\n"
    else:
        search_text = "（无搜索结果，请依靠你的知识进行撰写）"

    return f"""你是一位专业的深度调研分析师。请针对关键词「{keyword}」，基于以下搜索结果并结合你对该关键词的全面知识，撰写一份结构严谨的深度研究报告。

报告必须包含以下六个维度，每个维度作为一级标题，标题格式为“一、来源出处”、“二、含义变迁”、“三、相关新闻与社会事件”、“四、统计数据”、“五、学术文献与行业报告”、“六、哲学思想关联”。请确保每个部分内容充实，逻辑清晰，并尽量引用搜索结果中的具体信息（如有），在行文中用括号标注来源。

搜索结果如下：
{search_text}

请开始撰写报告，直接输出正文。正文使用中文，每部分之间用空行分隔。一级标题必须严格以“一、”“二、”……开头。"""

# ============================================================
# 主界面逻辑
# ============================================================
keyword = st.text_input("请输入要研究的关键词", placeholder="例如：算法推荐、信息茧房、网络暴力...")
generate_btn = st.button("🚀 生成报告", type="primary")

if generate_btn and keyword.strip():
    if not deepseek_key or not bocha_key:
        st.warning("请先配置 API 密钥！")
        st.stop()

    with st.spinner("正在搜索相关信息..."):
        results = bocha_search(keyword.strip(), bocha_key, count=10)

    with st.spinner("正在综合分析并撰写报告（约 30-60 秒）..."):
        prompt = build_prompt(keyword.strip(), results)
        report = deepseek_generate(prompt, deepseek_key)

    if not report:
        st.error("报告生成失败，请检查 API 密钥或网络。")
        st.stop()

    with st.expander("📄 查看报告预览"):
        st.text_area("报告内容", report, height=400)

    with st.spinner("正在生成 Word 文档..."):
        doc_bytes = create_word_doc(f"「{keyword.strip()}」深度研究档案", report)

    st.success("✅ 报告生成完成！")
    st.download_button(
        label="📥 下载 Word 文档 (.docx)",
        data=doc_bytes,
        file_name=f"关键词深度研究档案_{keyword.strip()}_{time.strftime('%Y%m%d_%H%M%S')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
elif generate_btn:
    st.warning("请输入关键词！")

st.markdown("---")
st.caption("提示：报告基于博查搜索结果与 DeepSeek 模型知识，仅供参考。")
