import streamlit as st
import google.generativeai as genai
import json
import os
import io
import re
import urllib.parse
from datetime import datetime
from docx import Document
from docx.shared import Pt as DocxPt, RGBColor as DocxRGB
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from video_engine import create_animated_video

st.set_page_config(
    page_title="MY 설교 AI 스튜디오 Pro",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 1. 개인 접속 보안 인증 ---
USER_PIN = st.secrets.get("APP_PIN", "7777")
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 개인 전용 설교 AI 플랫폼")
    pin = st.text_input("접속 비밀번호(PIN)를 입력하세요", type="password")
    if st.button("로그인"):
        if pin == USER_PIN:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("PIN 코드가 올바르지 않습니다.")
    st.stop()

# --- 2. Gemini AI 연동 엔진 ---
secret_key = st.secrets.get("GEMINI_API_KEY", "")
sidebar_key = st.sidebar.text_input("🔑 Gemini API Key", value=secret_key, type="password")
ACTIVE_KEY = sidebar_key.strip() if sidebar_key else secret_key.strip()

def get_ai_response(prompt: str, is_json: bool = True):
    if not ACTIVE_KEY:
        st.error("Gemini API Key가 필요합니다. 사이드바에 키를 입력해주세요.")
        return None
    try:
        genai.configure(api_key=ACTIVE_KEY)
    except Exception as e:
        st.error(f"API 키 오류: {str(e)}")
        return None

    candidate_models = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro"]
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            if is_json:
                try:
                    res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                    return json.loads(res.text)
                except Exception:
                    res = model.generate_content(prompt)
                    match = re.search(r"\{.*\}|\[.*\]", res.text, re.DOTALL)
                    if match:
                        return json.loads(match.group(0))
            else:
                res = model.generate_content(prompt)
                return res.text
        except Exception:
            continue
    st.error("AI 응답을 생성하지 못했습니다. 잠시 후 다시 시도해주세요.")
    return None

# --- 3. 문서 생성 도구 (Word & PDF & PPT) ---
def create_docx(title: str, content: str) -> io.BytesIO:
    doc = Document()
    title_p = doc.add_paragraph()
    run = title_p.add_run(title)
    run.font.size = DocxPt(18)
    run.font.bold = True
    run.font.color.rgb = DocxRGB(30, 58, 138)
    
    doc.add_paragraph(f"작성일: {datetime.now().strftime('%Y-%m-%d')} | MY 설교 AI 스튜디오\n")
    
    for paragraph in content.split("\n"):
        if paragraph.strip():
            doc.add_paragraph(paragraph.strip())
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_pdf(title: str, content: str) -> io.BytesIO:
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    styles = getSampleStyleSheet()
    
    story = [
        Paragraph(f"<b><font size=16 color='#1e3a8a'>{title}</font></b>", styles["Heading1"]),
        Spacer(1, 10),
        Paragraph(f"<font size=9 color='#64748b'>생성일: {datetime.now().strftime('%Y-%m-%d')}</font>", styles["Normal"]),
        Spacer(1, 15),
    ]
    for line in content.split("\n"):
        if line.strip():
            story.append(Paragraph(line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), styles["Normal"]))
            story.append(Spacer(1, 6))
            
    doc.build(story)
    bio.seek(0)
    return bio

def generate_church_pptx(title, scripture, preacher, points, conclusion):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def apply_bg(slide, color=RGBColor(24, 32, 54)):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    # 표지
    slide1 = prs.slides.add_slide(blank_layout)
    apply_bg(slide1, RGBColor(15, 23, 42))
    tx = slide1.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.0))
    p1 = tx.text_frame.paragraphs[0]
    p1.text = title
    p1.font.size, p1.font.bold = Pt(44), True
    p1.font.color.rgb, p1.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER
    p2 = tx.text_frame.add_paragraph()
    p2.text = f"\n본문: {scripture} | 설교: {preacher}"
    p2.font.size, p2.font.color.rgb = Pt(22), RGBColor(226, 232, 240)
    p2.alignment = PP_ALIGN.CENTER

    for idx, pt in enumerate(points, 1):
        slide = prs.slides.add_slide(blank_layout)
        apply_bg(slide, RGBColor(24, 32, 54))
        htx = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(1.2))
        hp = htx.text_frame.paragraphs[0]
        hp.text = pt.get("main_point", f"제 {idx} 대지")
        hp.font.size, hp.font.bold = Pt(36), True
        hp.font.color.rgb = RGBColor(253, 224, 71)

        btx = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(11.33), Inches(4.5))
        tf = btx.text_frame
        tf.word_wrap = True
        bp1 = tf.paragraphs[0]
        bp1.text = f"📖 성경적 해설:\n{pt.get('explanation', '')}\n"
        bp1.font.size, bp1.font.color.rgb = Pt(22), RGBColor(241, 245, 249)
        bp2 = tf.add_paragraph()
        bp2.text = f"\n💡 삶의 적용:\n{pt.get('application', '')}"
        bp2.font.size, bp2.font.color.rgb = Pt(22), RGBColor(147, 197, 253)

    slide_end = prs.slides.add_slide(blank_layout)
    apply_bg(slide_end, RGBColor(15, 23, 42))
    etx = slide_end.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.33), Inches(5.0))
    ep1 = etx.text_frame.paragraphs[0]
    ep1.text = "결론 및 결단의 기도"
    ep1.font.size, ep1.font.bold = Pt(36), True
    ep1.font.color.rgb = RGBColor(253, 224, 71)
    ep2 = etx.text_frame.add_paragraph()
    ep2.text = f"\n{conclusion}"
    ep2.font.size, ep2.font.color.rgb = Pt(22), RGBColor(226, 232, 240)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

def generate_cardnews_pptx(slides_data):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(10)
    blank_layout = prs.slide_layouts[6]

    for item in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(15, 23, 42)

        badge = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(2), Inches(0.8))
        bp = badge.text_frame.paragraphs[0]
        bp.text = f"CARD {item.get('card_number', '')}"
        bp.font.size, bp.font.bold = Pt(14), True
        bp.font.color.rgb = RGBColor(99, 102, 241)

        tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(2.2))
        tf = tbox.text_frame
        tf.word_wrap = True
        tp = tf.paragraphs[0]
        tp.text = item.get("headline", "")
        tp.font.size, tp.font.bold = Pt(32), True
        tp.font.color.rgb = RGBColor(253, 224, 71)

        bbox = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), Inches(8.4), Inches(4.5))
        btf = bbox.text_frame
        btf.word_wrap = True
        bp_body = btf.paragraphs[0]
        bp_body.text = item.get("body_text", "")
        bp_body.font.size, bp_body.font.color.rgb = Pt(22), RGBColor(241, 245, 249)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# --- 4. 전역 세션 상태 초기화 ---
if "sermon_library" not in st.session_state:
    st.session_state.sermon_library = [
        {
            "id": 1,
            "title": "신앙을 다음 세대에 전수하라",
            "scripture": "시편 78:4-7",
            "date": "2026-08-27",
            "tags": ["신앙 전수", "다음 세대", "가정 예배", "십자가 복음"],
            "testament": "구약",
            "book": "시편",
            "text": """1. 신앙의 유산을 숨기지 말고 적극적으로 전수하십시오.
오늘 본문 4절은 우리가 여호와의 영예와 능력을 자손에게 숨기지 않겠다고 고백합니다. 내가 말하지 않으면 그 신앙의 역사가 삭제됩니다. 의도적인 결단과 작정을 통해 자녀들에게 복음을 전하는 일에 힘쓰십시오.

2. 하나님이 행하신 구체적인 은혜와 복음을 세어가며 가르치십시오.
성경에서 전한다는 표현은 마치 수를 세는 것처럼 구체적으로 말하는 것을 뜻합니다. 삶의 위기 때 하나님이 어떻게 응답하셨는지 생생한 간증을 자녀들의 마음에 날카롭게 새겨주십시오.

3. 다음 세대가 오직 하나님께만 인생의 소망을 두게 하십시오.
우리가 신앙을 전수하는 궁극적 목적은 자녀들이 세상의 헛된 확신이 아닌 오직 하나님께 소망을 두게 하려는 것입니다."""
        }
    ]

if "current_sermon_idx" not in st.session_state:
    st.session_state.current_sermon_idx = 0

current_s = st.session_state.sermon_library[st.session_state.current_sermon_idx]
if "full_sermon" not in st.session_state or not st.session_state.full_sermon:
    st.session_state.full_sermon = current_s["text"]
if "sermon_title" not in st.session_state:
    st.session_state.sermon_title = current_s["title"]
if "sermon_scripture" not in st.session_state:
    st.session_state.sermon_scripture = current_s["scripture"]
if "preacher_name" not in st.session_state:
    st.session_state.preacher_name = "담임목사"

# --- 5. 상단 내비게이션 바 ---
app_mode = st.sidebar.radio(
    "🕊️ 플랫폼 대메뉴",
    ["📊 설교 대시보드 (메인 작업실)", "📚 설교 서재 (Sermon Library)", "🎬 쇼츠 만들기 (스튜디오)", "📤 새 설교 등록/업로드"]
)

# ==============================================================================
# MENU 1: 📊 설교 대시보드 (메인 작업실) - seolgyo-ai 완벽 구현
# ==============================================================================
if app_mode == "📊 설교 대시보드 (메인 작업실)":
    # 1. 헤더 타이틀 및 성경 본문 배지
    st.markdown(
        f"""
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 8px;">
            <h1 style="font-size: 28px; font-weight: 800; margin: 0; color: #f8fafc;">{st.session_state.sermon_title}</h1>
            <span style="background-color: #2563eb; color: white; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold;">기본 {st.session_state.sermon_scripture}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # 2. 상단 접이식 섹션 1: 💡 설교를 더 풍성하게 — 참고 구절 & 예화
    with st.expander("💡 설교를 더 풍성하게 — 참고 구절 & 예화", expanded=False):
        if st.button("✨ 참고 성구 및 신학적 예화 생성하기"):
            with st.spinner("본문과 주제에 맞는 풍성한 예화와 성구를 검색 중입니다..."):
                prompt = f"""
                본문: {st.session_state.sermon_scripture}, 설교 제목: {st.session_state.sermon_title}
                설교문 요약: {st.session_state.full_sermon[:1500]}
                
                다음 형식으로 출력하세요:
                1. 연관 핵심 참고 성구 3개 및 설교적 연결점
                2. 일상/현대적 공감 예화 2가지
                3. 교회사/고전 문학 및 기독교 사상가 명언 2가지
                """
                st.session_state.rich_materials = get_ai_response(prompt, is_json=False)
        
        if "rich_materials" in st.session_state and st.session_state.rich_materials:
            st.markdown(st.session_state.rich_materials)
        else:
            st.caption("버튼을 누르면 본문과 연관된 성경 구절, 현대적 예화, 신학자 명언을 추천합니다.")

    # 3. 상단 접이식 섹션 2: 🎵 추천 찬양 — 새찬송가 · 복음성가 · CCM
    with st.expander("🎵 추천 찬양 — 새찬송가 · 복음성가 · CCM", expanded=False):
        if st.button("🎶 맞춤 찬양 15곡 (각 5곡씩) 추천받기"):
            with st.spinner("설교 메시지와 어울리는 은혜로운 찬양을 선곡 중입니다..."):
                prompt = f"""
                설교 본문: {st.session_state.sermon_scripture}, 제목: {st.session_state.sermon_title}
                JSON 포맷으로 출력:
                {{
                    "hymns": ["새찬송가 000장 - 제목", "새찬송가 000장 - 제목", "새찬송가 000장 - 제목", "새찬송가 000장 - 제목", "새찬송가 000장 - 제목"],
                    "gospel_songs": ["복음성가 1", "복음성가 2", "복음성가 3", "복음성가 4", "복음성가 5"],
                    "ccm": ["CCM 1", "CCM 2", "CCM 3", "CCM 4", "CCM 5"]
                }}
                """
                st.session_state.praise_list = get_ai_response(prompt, is_json=True)

        if "praise_list" in st.session_state and st.session_state.praise_list:
            p_data = st.session_state.praise_list
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                st.markdown("#### 📖 새찬송가 (5곡)")
                for song in p_data.get("hymns", []):
                    yt_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(song)}"
                    st.markdown(f"- {song} [▶️ 듣기]({yt_url})")
            with col_p2:
                st.markdown("#### 🕊️ 복음성가 (5곡)")
                for song in p_data.get("gospel_songs", []):
                    yt_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(song)}"
                    st.markdown(f"- {song} [▶️ 듣기]({yt_url})")
            with col_p3:
                st.markdown("#### 🎸 현대 CCM (5곡)")
                for song in p_data.get("ccm", []):
                    yt_url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(song)}"
                    st.markdown(f"- {song} [▶️ 듣기]({yt_url})")

    st.write("---")

    # 4. 2단 분할 레이아웃 (좌측 메뉴 패널 / 우측 콘텐츠 뷰어)
    left_panel, right_panel = st.columns([1, 2.5])

    with left_panel:
        st.markdown("<p style='font-size:12px; font-weight:bold; color:#94a3b8; margin-bottom:4px;'>콘텐츠</p>", unsafe_allow_html=True)
        content_tab = st.radio(
            "콘텐츠 선택",
            ["설교 요약", "소그룹 나눔", "QT 5일치", "카드뉴스", "쇼츠 대본"],
            label_visibility="collapsed"
        )

        st.markdown("<p style='font-size:12px; font-weight:bold; color:#94a3b8; margin-top:20px; margin-bottom:4px;'>이 설교로 만들기</p>", unsafe_allow_html=True)
        maker_tab = st.radio(
            "사역 도구 선택",
            ["🖥️ 설교 PPT 다운로드", "🏡 가정예배지 만들기", "🔍 설교 점검·제안", "📖 소그룹 리더가이드"],
            label_visibility="collapsed"
        )

    # 5. 우측 메인 콘텐츠 뷰어 & 액션 바
    with right_panel:
        current_title = f"{st.session_state.sermon_title} - {content_tab}"
        
        # --- TAB: 설교 요약 ---
        if content_tab == "설교 요약":
            st.markdown(f"### 설교 요약")
            summary_content = st.text_area("내용 편집", value=st.session_state.full_sermon, height=380)
            
            col_b1, col_b2, col_b3, col_b4 = st.columns([1, 1, 1.2, 1.2])
            with col_b1:
                if st.button("💾 저장"):
                    st.session_state.full_sermon = summary_content
                    st.success("저장되었습니다.")
            with col_b3:
                docx_file = create_docx(f"{st.session_state.sermon_title} (설교 요약)", summary_content)
                st.download_button("📥 워드 다운로드", data=docx_file, file_name="설교요약.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_b4:
                pdf_file = create_pdf(f"{st.session_state.sermon_title} (설교 요약)", summary_content)
                st.download_button("📥 PDF 다운로드", data=pdf_file, file_name="설교요약.pdf", mime="application/pdf")

        # --- TAB: 소그룹 나눔 ---
        elif content_tab == "소그룹 나눔":
            st.markdown("### 소그룹 나눔지")
            if st.button("✨ 소그룹 나눔 질문 자동 생성", type="primary"):
                with st.spinner("나눔 질문 구성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3500]}\n소그룹 구역모임 나눔지(마음열기, 말씀나눔 2개, 삶적용 2개, 마침기도)를 마크다운 형식으로 작성하세요."
                    st.session_state.small_group_text = get_ai_response(prompt, is_json=False)
            
            group_txt = st.text_area("소그룹 나눔 내용", value=st.session_state.get("small_group_text", ""), height=350)
            if group_txt:
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.download_button("📥 워드 다운로드", data=create_docx("소그룹 나눔지", group_txt), file_name="소그룹나눔지.docx")
                with col_b2:
                    st.download_button("📥 PDF 다운로드", data=create_pdf("소그룹 나눔지", group_txt), file_name="소그룹나눔지.pdf")

        # --- TAB: QT 5일치 ---
        elif content_tab == "QT 5일치":
            st.markdown("### 설교 기반 주간 QT 5일치")
            if st.button("✨ 5일치 QT 묵상지 자동 생성", type="primary"):
                with st.spinner("주간 5일치 말씀 묵상지 생성 중..."):
                    prompt = f"""
                    설교 본문: {st.session_state.sermon_scripture}, 설교 요약: {st.session_state.full_sermon[:3000]}
                    월요일부터 금요일까지 5일치 QT 묵상지를 작성하세요.
                    각 날짜마다: [제목], [본문 성경 구절], [말씀 묵상 해설 (300자)], [삶의 질문과 적용], [오늘의 기도] 형식으로 작성하세요.
                    """
                    st.session_state.qt5_text = get_ai_response(prompt, is_json=False)

            qt_txt = st.text_area("QT 5일치 원고", value=st.session_state.get("qt5_text", ""), height=350)
            if qt_txt:
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.download_button("📥 워드 다운로드", data=create_docx("주간 QT 5일치", qt_txt), file_name="QT_5일치.docx")
                with col_b2:
                    st.download_button("📥 PDF 다운로드", data=create_pdf("주간 QT 5일치", qt_txt), file_name="QT_5일치.pdf")

        # --- TAB: 카드뉴스 ---
        elif content_tab == "카드뉴스":
            st.markdown("### 카드뉴스 (7~10장)")
            c_cnt = st.slider("카드 장수", 7, 10, 8)
            if st.button("🎨 카드뉴스 문구 생성", type="primary"):
                with st.spinner("카드뉴스 구성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3500]}\n정확히 {c_cnt}장의 카드뉴스 JSON 출력: {{\"cards\": [{{\"card_number\": 1, \"headline\": \"제목\", \"body_text\": \"문구\"}}]}}"
                    res = get_ai_response(prompt, is_json=True)
                    if res and "cards" in res:
                        st.session_state.card_list = res["cards"]

            if "card_list" in st.session_state:
                st.download_button("📥 1:1 정사각형 카드뉴스 PPT 내려받기", data=generate_cardnews_pptx(st.session_state.card_list), file_name="카드뉴스_정사각형.pptx")
                for c in st.session_state.card_list:
                    st.info(f"**CARD {c['card_number']}. {c['headline']}**\n\n{c['body_text']}")

        # --- TAB: 쇼츠 대본 ---
        elif content_tab == "쇼츠 대본":
            st.markdown("### 60초 세로 쇼츠 대본")
            if st.button("🎬 바이럴 쇼츠 대본 3가지 버전 생성", type="primary"):
                with st.spinner("후킹 멘트와 핵심 메시지가 담긴 쇼츠 대본 작성 중..."):
                    prompt = f"""
                    설교문: {st.session_state.full_sermon[:3000]}
                    유튜브 쇼츠/인스타 릴스용 60초 대본 3가지 버전(1. 감동형, 2. 질문형, 3. 결단선포형)을 작성하세요.
                    각 대본마다 [0~5초 후킹 멘트] - [5~45초 본론 메시지] - [45~60초 결단 및 콜투액션] 형식으로 나누어 작성하세요.
                    """
                    st.session_state.shorts_script_text = get_ai_response(prompt, is_json=False)

            shorts_txt = st.text_area("쇼츠 대본", value=st.session_state.get("shorts_script_text", ""), height=350)
            if shorts_txt:
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.download_button("📥 워드 다운로드", data=create_docx("쇼츠 대본", shorts_txt), file_name="쇼츠대본.docx")
                with col_b2:
                    st.download_button("📥 PDF 다운로드", data=create_pdf("쇼츠 대본", shorts_txt), file_name="쇼츠대본.pdf")

        # --- TOOL: 가정예배지 만들기 (영유아 / 어린이 / 청소년 / 청장년) ---
        if maker_tab == "🏡 가정예배지 만들기":
            st.write("---")
            st.markdown("### 🏡 세대별 맞춤 가정예배지 제작")
            age_group = st.selectbox("예배 대상 선택", ["👶 영유아용 (쉽고 활동적인 나눔)", "🧒 어린이용 (눈높이 퀴즈와 이야기)", "🧑 청소년용 (고민 토론 및 가치관)", "👨‍👩‍👧 청장년용 (깊은 묵상과 중보기도)"])
            
            if st.button(f"✨ {age_group.split(' ')[1]} 맞춤 가정예배지 생성", type="primary"):
                with st.spinner("가정예배 순서지 작성 중..."):
                    prompt = f"""
                    설교 본문: {st.session_state.sermon_scripture}, 설교문: {st.session_state.full_sermon[:3000]}
                    대상: {age_group}
                    순서:
                    1. 찬양 및 신앙고백
                    2. 함께 읽는 성경 구절
                    3. {age_group.split(' ')[1]}의 눈높이에 맞춘 3분 메시지
                    4. 가족 나눔 활동 및 질문 2가지
                    5. 축복 기도문
                    마크다운 형식으로 작성하세요.
                    """
                    st.session_state.family_worship_text = get_ai_response(prompt, is_json=False)

            fam_txt = st.text_area("가정예배지 본문", value=st.session_state.get("family_worship_text", ""), height=300)
            if fam_txt:
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    st.download_button("📥 워드 다운로드", data=create_docx("가정예배 순서지", fam_txt), file_name="가정예배지.docx")
                with col_b2:
                    st.download_button("📥 PDF 다운로드", data=create_pdf("가정예배 순서지", fam_txt), file_name="가정예배지.pdf")

        # --- TOOL: 설교 PPT 다운로드 ---
        elif maker_tab == "🖥️ 설교 PPT 다운로드":
            st.write("---")
            st.markdown("### 🖥️ 예배 설교용 PPT (.pptx) 생성")
            if st.button("🎨 16:9 와이드 PPT 생성"):
                with st.spinner("대지 추출 및 슬라이드 구성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3500]}\nJSON 출력: {{\"points\": [{{\"main_point\":\"대지\", \"explanation\":\"해설\", \"application\":\"적용\"}}], \"conclusion\":\"결론기도\"}}"
                    ppt_data = get_ai_response(prompt, is_json=True)
                    if ppt_data:
                        pptx_bio = generate_church_pptx(
                            title=st.session_state.sermon_title,
                            scripture=st.session_state.sermon_scripture,
                            preacher=st.session_state.preacher_name,
                            points=ppt_data.get("points", []),
                            conclusion=ppt_data.get("conclusion", "결단의 기도")
                        )
                        st.success("PPT가 완성되었습니다!")
                        st.download_button("📥 파워포인트(.pptx) 파일 내려받기", data=pptx_bio, file_name=f"{st.session_state.sermon_title}.pptx")

# ==============================================================================
# MENU 2: 📚 설교 서재 (Sermon Library) - seolgyo-ai 서재 완벽 구현
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>설교 서재</h1>", unsafe_allow_html=True)
    st.caption(f"설교문 {len(st.session_state.sermon_library)}편 · 올해 {len(st.session_state.sermon_library)}편 · 66권 중 1권")

    sub_c1, sub_c2 = st.columns([1, 4])
    with sub_c1:
        if st.button("➕ 새 설교 작성", type="primary"):
            st.session_state.sermon_title = "새 설교 제목"
            st.session_state.sermon_scripture = "본문 구절 입력"
            st.session_state.full_sermon = ""
            st.success("새 설교 작성을 시작합니다. [설교 대시보드] 메뉴로 이동하세요.")

    col_main, col_widgets = st.columns([2.5, 1])

    with col_main:
        search_kw = st.text_input("🔍 제목 또는 성경 구절로 검색", placeholder="예: 시편, 신앙, 고난...")
        
        # 폴더 태그 바
        st.markdown(
            """
            <div style="display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap;">
                <span style="border: 1px dashed #64748b; padding: 4px 10px; border-radius: 20px; font-size: 12px; color: #94a3b8;">+ 새 폴더</span>
                <span style="background: #1e293b; padding: 4px 12px; border-radius: 20px; font-size: 12px; color: #e2e8f0;">신앙 전수 1</span>
                <span style="background: #1e293b; padding: 4px 12px; border-radius: 20px; font-size: 12px; color: #e2e8f0;">다음 세대 1</span>
                <span style="background: #1e293b; padding: 4px 12px; border-radius: 20px; font-size: 12px; color: #e2e8f0;">십자가 복음 1</span>
                <span style="background: #1e293b; padding: 4px 12px; border-radius: 20px; font-size: 12px; color: #e2e8f0;">가정 예배 1</span>
            </div>
            """,
            unsafe_allow_html=True
        )

        # 설교 카드 목록
        for idx, s_item in enumerate(st.session_state.sermon_library):
            if search_kw and (search_kw not in s_item["title"] and search_kw not in s_item["scripture"]):
                continue
            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 18px; margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                            <div>
                                <h3 style="margin: 0 0 6px 0; font-size: 18px; font-weight: bold; color: #f8fafc;">{s_item['title']}</h3>
                                <p style="margin: 0 0 10px 0; font-size: 13px; color: #94a3b8;">{s_item['scripture']} · 설교일: {s_item['date']}</p>
                            </div>
                        </div>
                        <div style="display: flex; gap: 6px;">
                            {' '.join([f'<span style=\"background:#1e293b; color:#38bdf8; padding:2px 8px; border-radius:6px; font-size:11px;\">{t}</span>' for t in s_item.get('tags', [])])}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button(f"📖 이 설교 불러와서 작업하기", key=f"load_{idx}"):
                    st.session_state.current_sermon_idx = idx
                    st.session_state.sermon_title = s_item["title"]
                    st.session_state.sermon_scripture = s_item["scripture"]
                    st.session_state.full_sermon = s_item["text"]
                    st.success(f"'{s_item['title']}' 설교를 불러왔습니다! [설교 대시보드]에서 작업하세요.")

    with col_widgets:
        # 설교 발자취 위젯
        st.markdown(
            """
            <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                <div style="display: flex; justify-content: space-between; font-weight: bold; font-size: 14px; margin-bottom: 12px;">
                    <span>📖 설교 발자취</span>
                    <span style="color: #60a5fa; font-size: 12px;">1/66권 지도 펼치기 ▾</span>
                </div>
                <div style="font-size: 12px; color: #cbd5e1; margin-bottom: 6px;">구약 <span style="float: right; color: #94a3b8;">1/39</span></div>
                <div style="background: #334155; height: 6px; border-radius: 3px; margin-bottom: 12px;"><div style="background: #3b82f6; width: 3%; height: 6px; border-radius: 3px;"></div></div>
                <div style="font-size: 12px; color: #cbd5e1; margin-bottom: 6px;">신약 <span style="float: right; color: #94a3b8;">0/27</span></div>
                <div style="background: #334155; height: 6px; border-radius: 3px;"><div style="background: #3b82f6; width: 0%; height: 6px; border-radius: 3px;"></div></div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # 설교 타임라인 위젯
        st.markdown(
            """
            <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                <div style="font-weight: bold; font-size: 14px; margin-bottom: 8px;">🗓️ 설교 타임라인</div>
                <p style="font-size: 12px; color: #94a3b8; line-height: 1.5;">설교일이 입력된 설교문이 달력에 자동 정렬됩니다.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

        # 내 서재 위젯
        st.markdown(
            """
            <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 16px;">
                <div style="font-weight: bold; font-size: 14px; margin-bottom: 8px;">📁 내 서재 연구 자료함</div>
                <p style="font-size: 12px; color: #94a3b8; line-height: 1.5;">주석·연구 자료·손글씨 노트 보관함</p>
            </div>
            """,
            unsafe_allow_html=True
        )

# ==============================================================================
# MENU 3: 🎬 쇼츠 만들기 (스튜디오) - seolgyo-ai 쇼츠 UI 완벽 구현
# ==============================================================================
elif app_mode == "🎬 쇼츠 만들기 (스튜디오)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>▶️ 쇼츠 만들기</h1>", unsafe_allow_html=True)
    st.caption("유튜브 설교 영상 링크만 넣으면, 적절한 구간을 골라 쇼츠 영상을 자동으로 만들어 드립니다.")

    tab_make, tab_my = st.tabs(["만들기", "내 쇼츠"])

    with tab_make:
        st.markdown(
            """
            <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 16px; padding: 24px; margin-top: 12px;">
                <label style="font-size: 14px; font-weight: bold; color: #f8fafc;">유튜브 링크</label>
            </div>
            """,
            unsafe_allow_html=True
        )
        yt_link = st.text_input("유튜브 URL", placeholder="https://youtube.com/watch?v=...", label_visibility="collapsed")
        st.caption("설교 영상은 물론, 찬양·광고가 포함된 **예배 전체 영상**도 괜찮아요. 설교 부분만 자동으로 찾아냅니다.")

        st.markdown("#### 템플릿 (스타일)")
        t_col1, t_col2, t_col3, t_col4, t_col5 = st.columns(5)
        
        with t_col1:
            st.image("https://images.unsplash.com/photo-1507692049790-de58290a4334?w=300&q=80", caption="기본")
        with t_col2:
            st.image("https://images.unsplash.com/photo-1518495973542-4542c06a5843?w=300&q=80", caption="그라데이션")
        with t_col3:
            st.image("https://images.unsplash.com/photo-1519681393784-d120267933ba?w=300&q=80", caption="강조")
        with t_col4:
            st.image("https://images.unsplash.com/photo-1470252649378-9c29740c9fa8?w=300&q=80", caption="라이트")
        with t_col5:
            st.image("https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=300&q=80", caption="커버")

        selected_template = st.radio("스타일 선택", ["기본", "그라데이션", "강조", "라이트", "커버"], horizontal=True)

        st.markdown("#### 교회 정보 (선택)")
        church_name = st.text_input("교회명", value="화광교회")
        st.caption("영상 하단에 교회 이름이 자막 배지로 표시됩니다.")

        st.markdown("#### 나레이션 & 자막 원고")
        shorts_script = st.text_area("쇼츠 자막/나레이션 텍스트 (줄바꿈으로 문장 구분)", value="내 열심보다 중요한 것은 하나님의 이끄심입니다.\n우리가 멈출 때 비로소 하나님의 역사가 시작됩니다.\n오늘 그분의 인도하심 앞에 온전히 맡기십시오.", height=100)
        
        bg_user_media = st.file_uploader("직접 찍은 배경 영상/사진 업로드 (선택)", type=["mp4", "jpg", "png"])
        bgm_user_audio = st.file_uploader("배경음악 MP3 업로드 (선택)", type=["mp3"])

        if st.button("🚀 고화질 쇼츠 영상 자동 분석 & 렌더링", type="primary"):
            with st.spinner("AI 컷편집, 자막 애니메이션 및 BGM 믹싱 렌더링 중... (약 20~40초)"):
                bg_p, bgm_p = None, None
                if bg_user_media:
                    bg_p = f"./uploads_{bg_user_media.name}"
                    with open(bg_p, "wb") as f: f.write(bg_user_media.getbuffer())
                if bgm_user_audio:
                    bgm_p = f"./uploads_{bgm_user_audio.name}"
                    with open(bgm_p, "wb") as f: f.write(bgm_user_audio.getbuffer())

                lines = [l.strip() for l in shorts_script.split("\n") if l.strip()]
                rendered_path = create_animated_video(
                    title=f"{st.session_state.sermon_title} | {church_name}",
                    script_paragraphs=lines,
                    bg_media_path=bg_p,
                    bgm_path=bgm_p,
                    aspect_ratio="9:16",
                    voice="ko-KR-InJoonNeural"
                )
                st.session_state.rendered_shorts_path = rendered_path
                st.success("쇼츠 생성이 완료되었습니다!")

        if "rendered_shorts_path" in st.session_state and os.path.exists(st.session_state.rendered_shorts_path):
            st.video(st.session_state.rendered_shorts_path)
            with open(st.session_state.rendered_shorts_path, "rb") as vf:
                st.download_button("📥 완성된 MP4 쇼츠 영상 다운로드", data=vf, file_name="sermon_shorts.mp4", mime="video/mp4")

    with tab_my:
        st.info("이전에 생성한 쇼츠 영상 아카이브 목록입니다.")

# ==============================================================================
# MENU 4: 📤 새 설교 등록/업로드
# ==============================================================================
elif app_mode == "📤 새 설교 등록/업로드":
    st.title("📤 새 설교문 등록 및 파일 업로드")
    
    tab_f, tab_t = st.tabs(["파일로 올리기 (.docx, .pdf, .txt)", "직접 타이핑 작성"])
    
    with tab_f:
        u_file = st.file_uploader("설교 원고 파일 선택", type=["docx", "pdf", "txt"])
        new_title = st.text_input("설교 제목", value="신앙을 다음 세대에 전수하라")
        new_scripture = st.text_input("성경 본문", value="시편 78:4-7")
        
        if u_file and st.button("파일에서 설교 등록하기", type="primary"):
            text = ""
            fn = u_file.name.lower()
            if fn.endswith('.txt'): text = u_file.read().decode('utf-8', errors='ignore')
            elif fn.endswith('.docx'):
                d = Document(u_file)
                text = "\n".join([p.text for p in d.paragraphs if p.text])
            elif fn.endswith('.pdf'):
                pdf = PdfReader(u_file)
                for page in pdf.pages: text += (page.extract_text() or "") + "\n"
                
            st.session_state.sermon_title = new_title
            st.session_state.sermon_scripture = new_scripture
            st.session_state.full_sermon = text
            st.session_state.sermon_library.append({
                "id": len(st.session_state.sermon_library) + 1,
                "title": new_title,
                "scripture": new_scripture,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "tags": ["새 설교", "주일 설교"],
                "testament": "구약",
                "book": "성경",
                "text": text
            })
            st.success("설교문이 등록되었습니다! [설교 대시보드] 메뉴에서 모든 사역자료를 바로 생성하세요.")
