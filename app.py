import streamlit as st
import google.generativeai as genai
import json
import os
import io
import re
import asyncio
import edge_tts
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
from reportlab.lib.styles import getSampleStyleSheet
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

# --- 3. 공통 파일 변환 엔진 (Word / PPT / PDF / TXT / Voice MP3) ---
def create_docx(title: str, content: str) -> io.BytesIO:
    doc = Document()
    tp = doc.add_paragraph()
    run = tp.add_run(title)
    run.font.size, run.font.bold = DocxPt(18), True
    run.font.color.rgb = DocxRGB(30, 58, 138)
    doc.add_paragraph(f"작성일: {datetime.now().strftime('%Y-%m-%d')} | MY 설교 AI 스튜디오\n")
    for line in content.split("\n"):
        if line.strip(): doc.add_paragraph(line.strip())
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

def create_txt(title: str, content: str) -> io.BytesIO:
    text_data = f"[{title}]\n작성일: {datetime.now().strftime('%Y-%m-%d')}\n\n{content}"
    return io.BytesIO(text_data.encode("utf-8"))

def create_document_pptx(title: str, content: str) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # 1. 표지 슬라이드
    slide1 = prs.slides.add_slide(blank_layout)
    fill1 = slide1.background.fill
    fill1.solid()
    fill1.fore_color.rgb = RGBColor(15, 23, 42)
    
    tbox = slide1.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.33), Inches(2.5))
    p = tbox.text_frame.paragraphs[0]
    p.text = title
    p.font.size, p.font.bold = Pt(40), True
    p.font.color.rgb, p.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER
    
    # 2. 본문 슬라이드 분할 (250자 단위 슬라이드 생성)
    paragraphs = [p.strip() for p in content.split("\n") if p.strip()]
    chunk = ""
    for para in paragraphs:
        chunk += para + "\n\n"
        if len(chunk) > 280:
            slide = prs.slides.add_slide(blank_layout)
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(24, 32, 54)
            
            # 상단 제목 바
            htx = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.33), Inches(0.8))
            hp = htx.text_frame.paragraphs[0]
            hp.text = title
            hp.font.size, hp.font.bold = Pt(22), True
            hp.font.color.rgb = RGBColor(147, 197, 253)
            
            # 본문 내용
            btx = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.33), Inches(5.2))
            tf = btx.text_frame
            tf.word_wrap = True
            bp = tf.paragraphs[0]
            bp.text = chunk.strip()
            bp.font.size = Pt(20)
            bp.font.color.rgb = RGBColor(241, 245, 249)
            chunk = ""
            
    if chunk:
        slide = prs.slides.add_slide(blank_layout)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(24, 32, 54)
        htx = slide.shapes.add_textbox(Inches(1.0), Inches(0.6), Inches(11.33), Inches(0.8))
        hp = htx.text_frame.paragraphs[0]
        hp.text = title
        hp.font.size, hp.font.bold = Pt(22), True
        hp.font.color.rgb = RGBColor(147, 197, 253)
        
        btx = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.33), Inches(5.2))
        tf = btx.text_frame
        tf.word_wrap = True
        bp = tf.paragraphs[0]
        bp.text = chunk.strip()
        bp.font.size = Pt(20)
        bp.font.color.rgb = RGBColor(241, 245, 249)

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
        fill = slide.background.fill
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
        bp_body.font.size = Pt(22)
        bp_body.font.color.rgb = RGBColor(241, 245, 249)
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

async def generate_voiceover_audio(text: str, voice: str = "ko-KR-InJoonNeural") -> str:
    out_path = "./outputs/voiceover_temp.mp3"
    os.makedirs("./outputs", exist_ok=True)
    communicate = edge_tts.Communicate(text, voice)
    await communicate.save(out_path)
    return out_path

# --- 4. 전 섹션 공통 다운로드 액션 바 함수 (Word / PPT / PDF / TXT) ---
def render_all_download_buttons(title: str, content: str, key_prefix: str):
    if not content or not content.strip():
        return
    st.markdown("<div style='margin-top: 10px; margin-bottom: 6px; font-size: 12px; font-weight: bold; color: #94a3b8;'>📥 전체 포맷 다운로드</div>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.download_button(
            "📥 워드 (.docx)",
            data=create_docx(title, content),
            file_name=f"{title}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"{key_prefix}_docx"
        )
    with c2:
        st.download_button(
            "📥 PPT (.pptx)",
            data=create_document_pptx(title, content),
            file_name=f"{title}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key=f"{key_prefix}_pptx"
        )
    with c3:
        st.download_button(
            "📥 PDF (.pdf)",
            data=create_pdf(title, content),
            file_name=f"{title}.pdf",
            mime="application/pdf",
            key=f"{key_prefix}_pdf"
        )
    with c4:
        st.download_button(
            "📥 텍스트 (.txt)",
            data=create_txt(title, content),
            file_name=f"{title}.txt",
            mime="text/plain",
            key=f"{key_prefix}_txt"
        )

# --- 5. 전역 세션 초기화 ---
if "sermon_library" not in st.session_state:
    st.session_state.sermon_library = [
        {
            "id": 1,
            "title": "신앙을 다음 세대에 전수하라",
            "scripture": "시편 78:4-7",
            "date": "2026-08-27",
            "tags": ["신앙 전수", "다음 세대", "가정 예배"],
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

# --- 6. 사이드바 메인 내비게이션 ---
app_mode = st.sidebar.radio(
    "🕊️ 플랫폼 대메뉴",
    [
        "📊 설교 대시보드 (메인 작업실)",
        "🎙️ AI 보이스오버 스튜디오",
        "🎬 쇼츠 만들기 (스튜디오)",
        "📚 설교 서재 (Sermon Library)",
        "📤 새 설교 등록/업로드"
    ]
)

# ==============================================================================
# 1. 📊 설교 대시보드 (메인 작업실)
# ==============================================================================
if app_mode == "📊 설교 대시보드 (메인 작업실)":
    st.markdown(
        f"""
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 8px;">
            <h1 style="font-size: 28px; font-weight: 800; margin: 0; color: #f8fafc;">{st.session_state.sermon_title}</h1>
            <span style="background-color: #2563eb; color: white; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: bold;">기본 {st.session_state.sermon_scripture}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    # 접이식 섹션 1: 참고 구절 & 예화
    with st.expander("💡 설교를 더 풍성하게 — 참고 구절 & 예화", expanded=False):
        if st.button("✨ 참고 성구 및 신학적 예화 생성하기"):
            with st.spinner("본문과 연관된 성구와 예화를 분석 중입니다..."):
                prompt = f"""
                본문: {st.session_state.sermon_scripture}, 설교 제목: {st.session_state.sermon_title}
                설교문: {st.session_state.full_sermon[:1500]}
                1. 연관 핵심 참고 성구 3개 및 설교적 연결점
                2. 일상/현대적 공감 예화 2가지
                3. 교회사/고전 문학 및 기독교 사상가 명언 2가지
                """
                st.session_state.rich_materials = get_ai_response(prompt, is_json=False)
        
        if "rich_materials" in st.session_state and st.session_state.rich_materials:
            st.markdown(st.session_state.rich_materials)
            render_all_download_buttons("참고구절_및_예화", st.session_state.rich_materials, "rich_mat")

    # 접이식 섹션 2: 추천 찬양
    with st.expander("🎵 추천 찬양 — 새찬송가 · 복음성가 · CCM (각 5곡 검색 연결)", expanded=False):
        if st.button("🎶 맞춤 찬양 15곡 추천받기"):
            with st.spinner("설교 메시지와 어울리는 찬양을 선곡 중입니다..."):
                prompt = f"""
                설교 본문: {st.session_state.sermon_scripture}, 제목: {st.session_state.sermon_title}
                JSON 포맷:
                {{
                    "hymns": ["새찬송가 000장 - 제목", "새찬송가 000장 - 제목", "새찬송가 000장 - 제목", "새찬송가 000장 - 제목", "새찬송가 000장 - 제목"],
                    "gospel_songs": ["복음성가 1", "복음성가 2", "복음성가 3", "복음성가 4", "복음성가 5"],
                    "ccm": ["CCM 1", "CCM 2", "CCM 3", "CCM 4", "CCM 5"]
                }}
                """
                st.session_state.praise_list = get_ai_response(prompt, is_json=True)

        if "praise_list" in st.session_state and st.session_state.praise_list:
            p_data = st.session_state.praise_list
            cp1, cp2, cp3 = st.columns(3)
            with cp1:
                st.markdown("#### 📖 새찬송가 (5곡)")
                for song in p_data.get("hymns", []):
                    st.markdown(f"- {song} [🔍 검색](https://www.google.com/search?q={urllib.parse.quote(song)}) | [▶️ 듣기](https://www.youtube.com/results?search_query={urllib.parse.quote(song)})")
            with cp2:
                st.markdown("#### 🕊️ 복음성가 (5곡)")
                for song in p_data.get("gospel_songs", []):
                    st.markdown(f"- {song} [🔍 검색](https://www.google.com/search?q={urllib.parse.quote(song)}) | [▶️ 듣기](https://www.youtube.com/results?search_query={urllib.parse.quote(song)})")
            with cp3:
                st.markdown("#### 🎸 현대 CCM (5곡)")
                for song in p_data.get("ccm", []):
                    st.markdown(f"- {song} [🔍 검색](https://www.google.com/search?q={urllib.parse.quote(song)}) | [▶️ 듣기](https://www.youtube.com/results?search_query={urllib.parse.quote(song)})")

    st.write("---")

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
            ["🏡 세대별 가정예배지", "🔍 설교 점검 및 제안", "📖 소그룹 리더가이드"],
            label_visibility="collapsed"
        )

    with right_panel:
        # TAB 1: 설교 요약
        if content_tab == "설교 요약":
            st.markdown("### 설교 요약")
            s_edit = st.text_area("설교문/요약 편집", value=st.session_state.full_sermon, height=360)
            if st.button("💾 본문 저장"):
                st.session_state.full_sermon = s_edit
                st.success("저장되었습니다.")
            render_all_download_buttons(f"{st.session_state.sermon_title}_설교요약", s_edit, "sermon_sum")

        # TAB 2: 소그룹 나눔
        elif content_tab == "소그룹 나눔":
            st.markdown("### 소그룹 나눔지")
            if st.button("✨ 소그룹 나눔 질문 자동 생성", type="primary"):
                with st.spinner("소그룹 나눔지 작성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3500]}\n소그룹 구역모임 나눔지(마음열기, 말씀나눔 2개, 삶적용 2개, 합심기도문)를 깔끔한 텍스트로 작성하세요."
                    st.session_state.small_group_text = get_ai_response(prompt, is_json=False)
            group_txt = st.text_area("소그룹 나눔 내용", value=st.session_state.get("small_group_text", ""), height=350)
            render_all_download_buttons(f"{st.session_state.sermon_title}_소그룹나눔지", group_txt, "sm_grp")

        # TAB 3: QT 5일치
        elif content_tab == "QT 5일치":
            st.markdown("### 설교 기반 주간 QT 5일치")
            if st.button("✨ 5일치 QT 묵상지 자동 생성", type="primary"):
                with st.spinner("월~금 5일치 말씀 묵상지 생성 중..."):
                    prompt = f"""
                    본문: {st.session_state.sermon_scripture}, 설교문: {st.session_state.full_sermon[:3000]}
                    월요일부터 금요일까지 5일치 QT 묵상지를 작성하세요.
                    각 일자별: [제목], [본문 성구], [묵상 글], [적용 질문], [오늘의 기도]
                    """
                    st.session_state.qt5_text = get_ai_response(prompt, is_json=False)
            qt_txt = st.text_area("QT 원고", value=st.session_state.get("qt5_text", ""), height=350)
            render_all_download_buttons(f"{st.session_state.sermon_title}_주간QT5일치", qt_txt, "qt5")

        # TAB 4: 카드뉴스
        elif content_tab == "카드뉴스":
            st.markdown("### 카드뉴스 (7~10장)")
            c_cnt = st.slider("카드 장수", 7, 10, 8)
            if st.button("🎨 카드뉴스 문구 생성", type="primary"):
                with st.spinner("카드뉴스 구조화 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3500]}\n정확히 {c_cnt}장의 카드뉴스 JSON 출력: {{\"cards\": [{{\"card_number\": 1, \"headline\": \"제목\", \"body_text\": \"문구\"}}]}}"
                    res = get_ai_response(prompt, is_json=True)
                    if res and "cards" in res: st.session_state.card_list = res["cards"]

            if "card_list" in st.session_state:
                st.download_button("📥 1:1 정사각형 카드뉴스 PPT 내려받기", data=generate_cardnews_pptx(st.session_state.card_list), file_name="카드뉴스_1대1.pptx")
                card_all_text = "\n\n".join([f"CARD {c['card_number']}. {c['headline']}\n{c['body_text']}" for c in st.session_state.card_list])
                render_all_download_buttons(f"{st.session_state.sermon_title}_카드뉴스", card_all_text, "cd_news")
                for c in st.session_state.card_list:
                    st.info(f"**CARD {c['card_number']}. {c['headline']}**\n\n{c['body_text']}")

        # TAB 5: 쇼츠 대본
        elif content_tab == "쇼츠 대본":
            st.markdown("### 60초 세로 쇼츠 대본")
            if st.button("🎬 바이럴 쇼츠 대본 3종 생성", type="primary"):
                with st.spinner("후킹 멘트와 핵심 메시지 작성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3000]}\n쇼츠/릴스용 60초 대본 3가지 버전(감동형, 질문형, 결단선포형)을 [후킹]-[본론]-[결단] 구조로 작성하세요."
                    st.session_state.shorts_script_text = get_ai_response(prompt, is_json=False)
            sh_txt = st.text_area("쇼츠 대본", value=st.session_state.get("shorts_script_text", ""), height=350)
            render_all_download_buttons(f"{st.session_state.sermon_title}_쇼츠대본", sh_txt, "sh_script")

        # TOOL: 세대별 맞춤 가정예배지
        if maker_tab == "🏡 세대별 가정예배지":
            st.write("---")
            st.markdown("### 🏡 세대별 맞춤 가정예배지")
            age_group = st.selectbox("예배 대상 선택", ["👶 영유아용 (쉽고 활동적인 나눔)", "🧒 어린이용 (눈높이 퀴즈와 이야기)", "🧑 청소년용 (고민 토론 및 가치관)", "👨‍👩‍👧 청장년용 (깊은 묵상과 중보기도)"])
            if st.button(f"✨ {age_group.split(' ')[1]} 가정예배지 생성", type="primary"):
                with st.spinner("가정예배 순서지 작성 중..."):
                    prompt = f"본문: {st.session_state.sermon_scripture}, 설교문: {st.session_state.full_sermon[:3000]}\n대상: {age_group}\n1.찬양 2.말씀 3.눈높이 3분 메시지 4.가족 나눔 2개 5.축복기도"
                    st.session_state.family_worship_text = get_ai_response(prompt, is_json=False)
            fam_txt = st.text_area("가정예배지 본문", value=st.session_state.get("family_worship_text", ""), height=300)
            render_all_download_buttons(f"{st.session_state.sermon_title}_가정예배지_{age_group.split(' ')[1]}", fam_txt, "fam_worship")

# ==============================================================================
# 2. 🎙️ AI 보이스오버 스튜디오 (Voice-over)
# ==============================================================================
elif app_mode == "🎙️ AI 보이스오버 스튜디오":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>🎙️ AI 보이스오버 (Voice-over) 스튜디오</h1>", unsafe_allow_html=True)
    st.caption("설교 원고나 요약문을 자연스러운 신경망 음성으로 합성하여 실시간 청취 및 고음질 MP3 파일로 다운로드합니다.")

    vo_col1, vo_col2 = st.columns([1.5, 1])
    with vo_col1:
        vo_text = st.text_area("음성으로 변환할 설교 텍스트", value=st.session_state.full_sermon, height=350)
        vo_c1, vo_c2 = st.columns(2)
        with vo_c1:
            vo_voice = st.selectbox("성우 보이스 선택", ["인준 (남성 - 차분하고 신뢰감 있는 톤)", "선희 (여성 - 맑고 또렷한 톤)"])
        with vo_c2:
            voice_id = "ko-KR-InJoonNeural" if "인준" in vo_voice else "ko-KR-SunHiNeural"
        
        if st.button("🎙️ 고음질 음성(TTS) 즉시 생성하기", type="primary"):
            if not vo_text.strip():
                st.warning("변환할 텍스트를 입력해주세요.")
            else:
                with st.spinner("AI 신경망 음성 생성 중... (Edge-TTS)"):
                    out_audio = asyncio.run(generate_voiceover_audio(vo_text, voice_id))
                    st.session_state.vo_audio_path = out_audio
                    st.success("보이스오버 음성 생성이 완료되었습니다!")

    with vo_col2:
        st.markdown("### 🎧 음성 미리듣기 및 파일 다운로드")
        if "vo_audio_path" in st.session_state and os.path.exists(st.session_state.vo_audio_path):
            st.audio(st.session_state.vo_audio_path)
            with open(st.session_state.vo_audio_path, "rb") as af:
                st.download_button(
                    label="📥 MP3 오디오 파일 다운로드",
                    data=af,
                    file_name=f"{st.session_state.sermon_title}_voiceover.mp3",
                    mime="audio/mp3"
                )
        else:
            st.info("왼쪽에서 버튼을 누르면 이곳에 플레이어와 MP3 다운로드 버튼이 나타납니다.")

# ==============================================================================
# 3. 🎬 쇼츠 만들기 (스튜디오) - 소스 링크, 5개 제목 추천, #태그, 300MB 업로드
# ==============================================================================
elif app_mode == "🎬 쇼츠 만들기 (스튜디오)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>▶️ 쇼츠 만들기 스튜디오</h1>", unsafe_allow_html=True)
    st.caption("클릭을 부르는 제목 추천, 해시태그 자동 생성, 무료 미디어 소스 연동 및 300MB 대용량 영상 합성을 지원합니다.")

    # 1. 무료 미디어 소스 바로가기 바
    st.markdown("#### 🌐 무료 미디어/BGM 소스 바로가기")
    src_c1, src_c2, src_c3, src_c4 = st.columns(4)
    with src_c1:
        st.link_button("🎥 픽사베이 (Pixabay 무료 영상)", "https://pixabay.com/ko/videos/")
    with src_c2:
        st.link_button("📸 펙셀스 (Pexels 무료 비디오)", "https://www.pexels.com/ko-kr/videos/")
    with src_c3:
        st.link_button("🎵 픽사베이 무료 음악(BGM)", "https://pixabay.com/ko/music/")
    with src_c4:
        st.link_button("🎶 유튜브 오디오 라이브러리", "https://studio.youtube.com/channel/UC/music")

    st.write("---")

    # 2. AI 5가지 추천 제목 & 해시태그 생성
    st.markdown("#### 💡 확 끌리는 5가지 쇼츠 제목 & #해시태그 추천")
    if st.button("✨ 조회수 폭발 5가지 쇼츠 제목 & 해시태그 뽑기"):
        with st.spinner("바이럴 제목 및 알고리즘 태그 분석 중..."):
            prompt = f"""
            설교 본문: {st.session_state.sermon_scripture}, 제목: {st.session_state.sermon_title}
            설교 요약: {st.session_state.full_sermon[:2000]}
            유튜브 쇼츠 및 인스타그램 릴스에서 클릭률을 극대화할 수 있는 제목 5개와 해시태그 8개를 JSON으로 작성하세요:
            {{
                "titles": [
                    "1. 호기심 자극형 제목",
                    "2. 공감/위로형 제목",
                    "3. 질문형 제목",
                    "4. 강한 결단 선포형 제목",
                    "5. 반전/핵심 통찰형 제목"
                ],
                "hashtags": ["#쇼츠", "#말씀", "#은혜", "#기도", "#설교", "#믿음", "#기독교", "#크리스천"]
            }}
            """
            st.session_state.shorts_rec = get_ai_response(prompt, is_json=True)

    selected_title = st.session_state.sermon_title
    if "shorts_rec" in st.session_state and st.session_state.shorts_rec:
        rec = st.session_state.shorts_rec
        st.markdown("**추천 제목 선택 (클릭하면 바로 적용됩니다):**")
        t_choice = st.radio("추천 제목 목록", rec.get("titles", []), horizontal=False)
        if t_choice:
            selected_title = re.sub(r"^\d+\.\s*", "", t_choice)
        st.markdown("**자동 추천 #해시태그:** " + " ".join(rec.get("hashtags", [])))

    st.write("---")

    # 3. 쇼츠 영상 파라미터 세팅
    sh_col1, sh_col2 = st.columns([1.2, 1])
    with sh_col1:
        v_title = st.text_input("쇼츠 최종 제목 (직접 수정 가능)", value=selected_title)
        v_script = st.text_area("자막 및 나레이션 대본 (줄바꿈으로 문장 구분)", value="내 열심보다 중요한 것은 하나님의 이끄심입니다.\n우리가 멈출 때 비로소 하나님의 역사가 시작됩니다.\n오늘 그분의 인도하심 앞에 온전히 맡기십시오.", height=120)
        
        c_v1, c_v2 = st.columns(2)
        with c_v1:
            v_ratio = st.radio("화면 비율", ["9:16 (세로 쇼츠/릴스)", "16:9 (가로 1~2분 영상)"])
        with c_v2:
            v_voice = st.selectbox("성우 목소리", ["인준 (남성 - 차분함)", "선희 (여성 - 또렷함)"])

        st.markdown("#### 📁 미디어 업로드 (최대 300MB 지원)")
        bg_media = st.file_uploader("배경 동영상/사진 업로드 (최대 300MB)", type=["mp4", "mov", "jpg", "png"])
        bgm_media = st.file_uploader("배경음악 MP3 업로드 (최대 300MB)", type=["mp3", "wav"])

        if st.button("🚀 고화질 비디오 자동 합성 & 렌더링 시작", type="primary"):
            with st.spinner("TTS 음성 합성, 자막 애니메이션 및 BGM 믹싱 렌더링 중... (약 20~40초)"):
                bg_p, bgm_p = None, None
                if bg_media:
                    bg_p = f"./uploads_{bg_media.name}"
                    with open(bg_p, "wb") as f: f.write(bg_media.getbuffer())
                if bgm_media:
                    bgm_p = f"./uploads_{bgm_media.name}"
                    with open(bgm_p, "wb") as f: f.write(bgm_media.getbuffer())

                lines = [l.strip() for l in v_script.split("\n") if l.strip()]
                voice_id = "ko-KR-InJoonNeural" if "인준" in v_voice else "ko-KR-SunHiNeural"
                ratio_val = "9:16" if "9:16" in v_ratio else "16:9"

                rendered_out = create_animated_video(
                    title=v_title,
                    script_paragraphs=lines,
                    bg_media_path=bg_p,
                    bgm_path=bgm_p,
                    aspect_ratio=ratio_val,
                    voice=voice_id
                )
                st.session_state.rendered_shorts_out = rendered_out
                st.success("영상 렌더링이 완료되었습니다!")

    with sh_col2:
        st.markdown("### 🎬 완성된 쇼츠 미리보기 및 다운로드")
        if "rendered_shorts_out" in st.session_state and os.path.exists(st.session_state.rendered_shorts_out):
            st.video(st.session_state.rendered_shorts_out)
            with open(st.session_state.rendered_shorts_out, "rb") as vf:
                st.download_button(
                    label="📥 완성된 MP4 비디오 파일 다운로드",
                    data=vf,
                    file_name="sermon_shorts_final.mp4",
                    mime="video/mp4"
                )
        else:
            st.info("왼쪽에서 렌더링 버튼을 누르면 이곳에서 영상을 즉시 확인하고 다운로드할 수 있습니다.")

# ==============================================================================
# 4. 📚 설교 서재 (Sermon Library)
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>설교 서재</h1>", unsafe_allow_html=True)
    st.caption(f"설교문 {len(st.session_state.sermon_library)}편 보관 중")

    col_main, col_side = st.columns([2.5, 1])
    with col_main:
        search_kw = st.text_input("🔍 제목 또는 성경 구절로 검색", placeholder="예: 시편, 신앙, 소망...")
        for idx, s_item in enumerate(st.session_state.sermon_library):
            if search_kw and (search_kw not in s_item["title"] and search_kw not in s_item["scripture"]):
                continue
            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 18px; margin-bottom: 12px;">
                        <h3 style="margin: 0 0 6px 0; font-size: 18px; font-weight: bold; color: #f8fafc;">{s_item['title']}</h3>
                        <p style="margin: 0 0 10px 0; font-size: 13px; color: #94a3b8;">{s_item['scripture']} · 등록일: {s_item['date']}</p>
                        <div>{' '.join([f'<span style=\"background:#1e293b; color:#38bdf8; padding:2px 8px; border-radius:6px; font-size:11px;\">#{t}</span>' for t in s_item.get('tags', [])])}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button("📖 이 설교 불러와서 작업하기", key=f"lib_load_{idx}"):
                    st.session_state.current_sermon_idx = idx
                    st.session_state.sermon_title = s_item["title"]
                    st.session_state.sermon_scripture = s_item["scripture"]
                    st.session_state.full_sermon = s_item["text"]
                    st.success(f"'{s_item['title']}' 설교를 불러왔습니다! [설교 대시보드]로 이동하세요.")

    with col_side:
        st.markdown(
            """
            <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                <div style="font-weight: bold; font-size: 14px; margin-bottom: 8px;">📖 설교 발자취</div>
                <div style="font-size: 12px; color: #cbd5e1;">구약 (1/39) · 신약 (0/27)</div>
            </div>
            """,
            unsafe_allow_html=True
        )

# ==============================================================================
# 5. 📤 새 설교 등록/업로드 (직접 타이핑 작성 100% 구현)
# ==============================================================================
elif app_mode == "📤 새 설교 등록/업로드":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>📤 새 설교문 등록 및 원고 작성</h1>", unsafe_allow_html=True)
    
    tab_type, tab_file = st.tabs(["✍️ 직접 타이핑 작성", "📁 파일 업로드 (.docx, .pdf, .txt)"])
    
    # 1. 직접 타이핑 작성 탭
    with tab_type:
        st.markdown("#### 강단 선포용 설교 원고 직접 작성")
        tc1, tc2, tc3 = st.columns([2, 1.5, 1])
        with tc1:
            t_title = st.text_input("설교 제목", placeholder="예: 광야에서 만나는 하나님의 은혜")
        with tc2:
            t_scripture = st.text_input("성경 본문", placeholder="예: 출애굽기 16:1-12")
        with tc3:
            t_preacher = st.text_input("설교자 성함", value=st.session_state.preacher_name)
            
        t_tags = st.text_input("설교 태그 (쉼표로 구분)", value="주일설교, 은혜, 광야, 기도")
        t_content = st.text_area("설교문 본문 전문", height=400, placeholder="설교 원고 전문을 이곳에 직접 자유롭게 입력하세요...")
        
        st.caption(f"현재 글자 수: **{len(t_content):,}자** (공백 포함)")
        
        if st.button("💾 새 설교로 등록 및 대시보드 동기화", type="primary"):
            if not t_title.strip() or not t_content.strip():
                st.warning("설교 제목과 본문 내용을 모두 입력해주세요.")
            else:
                tag_list = [t.strip() for t in t_tags.split(",") if t.strip()]
                new_entry = {
                    "id": len(st.session_state.sermon_library) + 1,
                    "title": t_title.strip(),
                    "scripture": t_scripture.strip(),
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "tags": tag_list,
                    "text": t_content.strip()
                }
                st.session_state.sermon_library.append(new_entry)
                st.session_state.current_sermon_idx = len(st.session_state.sermon_library) - 1
                st.session_state.sermon_title = t_title.strip()
                st.session_state.sermon_scripture = t_scripture.strip()
                st.session_state.preacher_name = t_preacher.strip()
                st.session_state.full_sermon = t_content.strip()
                st.success(f"'{t_title}' 설교가 서재에 등록되었습니다! [설교 대시보드] 메뉴에서 모든 사역자료(QT, 요약, 소그룹, 영상)를 바로 확인하세요.")

    # 2. 파일 업로드 탭
    with tab_file:
        st.markdown("#### 설교문 파일 업로드")
        u_file = st.file_uploader("워드, PDF, 텍스트 파일 선택", type=["docx", "pdf", "txt"])
        f_title = st.text_input("설교 제목 (파일 업로드용)", value="새로 등록할 설교 제목")
        f_scripture = st.text_input("성경 본문 (파일 업로드용)", value="본문 구절")
        
        if u_file and st.button("📂 파일 읽어와서 서재에 등록하기", type="primary"):
            text = ""
            fn = u_file.name.lower()
            if fn.endswith('.txt'): text = u_file.read().decode('utf-8', errors='ignore')
            elif fn.endswith('.docx'):
                d = Document(u_file)
                text = "\n".join([p.text for p in d.paragraphs if p.text])
            elif fn.endswith('.pdf'):
                pdf = PdfReader(u_file)
                for page in pdf.pages: text += (page.extract_text() or "") + "\n"

            st.session_state.sermon_title = f_title
            st.session_state.sermon_scripture = f_scripture
            st.session_state.full_sermon = text
            st.session_state.sermon_library.append({
                "id": len(st.session_state.sermon_library) + 1,
                "title": f_title,
                "scripture": f_scripture,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "tags": ["파일등록", "설교문"],
                "text": text
            })
            st.success("파일 등록이 완료되었습니다! [설교 대시보드]로 이동하세요.")
