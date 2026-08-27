import streamlit as st
import google.generativeai as genai
import json
import os
import io
import re
import asyncio
import edge_tts
import urllib.parse
import urllib.request
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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from video_engine import create_animated_video

st.set_page_config(
    page_title="MY 설교 AI 스튜디오 Pro",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 툴바 버튼 CSS 스타일링 (3번째 이미지처럼 가로 밀착 정렬)
st.markdown("""
<style>
    div[data-testid="column"] button {
        width: 100% !important;
        padding: 4px 6px !important;
        font-size: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 1. 개인 보안 접속 인증 ---
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

# --- 2. Gemini AI 연동 엔진 (오류 해결 & 순수 한국어 강제) ---
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
        st.error(f"API 키 설정 오류: {str(e)}")
        return None

    # 한국어 전용 시스템 지침 추가
    korean_system_prompt = (
        "당신은 신뢰할 수 있는 한국 교회의 목회 동역자 AI입니다. "
        "모든 응답은 반드시 100% 순수 한국어로만 작성하세요. "
        "영문 생각 과정(Drafting, Concept, Focus, Idea 등)이나 번역 초안, 영어 메모는 절대 출력하지 마세요."
    )
    full_prompt = f"{korean_system_prompt}\n\n[요청 작업]\n{prompt}"

    candidate_models = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro"]
    try:
        live_models = [m.name.replace("models/", "") for m in genai.list_models() if "generateContent" in m.supported_generation_methods]
        for lm in live_models:
            if lm not in candidate_models:
                candidate_models.append(lm)
    except Exception:
        pass

    last_error_msg = ""
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            if is_json:
                try:
                    res = model.generate_content(full_prompt, generation_config={"response_mime_type": "application/json"})
                    return json.loads(res.text)
                except Exception:
                    res = model.generate_content(full_prompt)
                    match = re.search(r"\{.*\}|\[.*\]", res.text, re.DOTALL)
                    if match:
                        return json.loads(match.group(0))
            else:
                res = model.generate_content(full_prompt)
                if res.text:
                    # 영문 메타데이터 블록 제거 정제
                    cleaned_text = re.sub(r"\*?\*?(Drafting|Concept|Focus|Selection|Idea \d+):\*?\*?.*?\n", "", res.text)
                    return cleaned_text.strip()
        except Exception as e:
            last_error_msg = str(e)
            continue

    st.error(f"AI 호출 오류 상세: {last_error_msg}")
    return None

# --- 3. PDF 한글 폰트 자동 등록 엔진 (깨짐 현상 완벽 해결) ---
PDF_FONT_NAME = "Helvetica"

def setup_korean_font():
    global PDF_FONT_NAME
    possible_paths = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf",
        "C:/Windows/Fonts/gulim.ttc",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc"
    ]
    for path in possible_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("NanumKorean", path))
                PDF_FONT_NAME = "NanumKorean"
                return
            except Exception:
                pass

    local_font = "./NanumGothic.ttf"
    if not os.path.exists(local_font):
        try:
            url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
            urllib.request.urlretrieve(url, local_font)
        except Exception:
            pass
    if os.path.exists(local_font):
        try:
            pdfmetrics.registerFont(TTFont("NanumKorean", local_font))
            PDF_FONT_NAME = "NanumKorean"
        except Exception:
            pass

setup_korean_font()

# --- 4. 문서 변환 엔진 (Word / PPT / PDF / TXT) ---
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
    doc = SimpleDocTemplate(
        bio,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )
    
    title_style = ParagraphStyle(
        name="K_Title",
        fontName=PDF_FONT_NAME,
        fontSize=15,
        leading=20,
        textColor="#1e3a8a",
        spaceAfter=8
    )
    meta_style = ParagraphStyle(
        name="K_Meta",
        fontName=PDF_FONT_NAME,
        fontSize=8,
        leading=12,
        textColor="#64748b",
        spaceAfter=12
    )
    body_style = ParagraphStyle(
        name="K_Body",
        fontName=PDF_FONT_NAME,
        fontSize=9.5,
        leading=15,
        textColor="#1e293b",
        spaceAfter=5
    )

    story = [
        Paragraph(f"<b>{title}</b>", title_style),
        Paragraph(f"생성일: {datetime.now().strftime('%Y-%m-%d')} | MY 설교 AI 스튜디오", meta_style),
        Spacer(1, 8),
    ]

    for line in content.split("\n"):
        clean = line.strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if clean:
            clean = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean)
            story.append(Paragraph(clean, body_style))
        else:
            story.append(Spacer(1, 4))

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
    
    # 표지
    slide1 = prs.slides.add_slide(blank_layout)
    fill1 = slide1.background.fill
    fill1.solid()
    fill1.fore_color.rgb = RGBColor(15, 23, 42)
    tbox = slide1.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.33), Inches(2.5))
    p = tbox.text_frame.paragraphs[0]
    p.text = title
    p.font.size, p.font.bold = Pt(38), True
    p.font.color.rgb, p.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER
    
    # 본문 슬라이드
    paragraphs = [p.strip() for p in content.split("\n") if p.strip()]
    chunk = ""
    for para in paragraphs:
        chunk += para + "\n\n"
        if len(chunk) > 280:
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

# --- 5. 모든 섹션 상단 통일 툴바 컴포넌트 [수정 / 복사 / 워드 / PDF / PPT / txt] ---
def render_section_top_toolbar(title: str, content: str, state_key: str):
    col_t, col_btns = st.columns([1.3, 2.7])
    with col_t:
        st.markdown(f"<h3 style='margin: 0; padding: 0; font-size: 20px; font-weight: 800; line-height: 1.3;'>{title}</h3>", unsafe_allow_html=True)
    with col_btns:
        if content and content.strip():
            c_edit, c_copy, c_doc, c_pdf, c_ppt, c_txt = st.columns([1, 1, 1.1, 1.1, 1.1, 1])
            with c_edit:
                if st.button("✏️ 수정", key=f"edit_btn_{state_key}"):
                    st.session_state[f"edit_mode_{state_key}"] = not st.session_state.get(f"edit_mode_{state_key}", False)
            with c_copy:
                if st.button("📋 복사", key=f"copy_btn_{state_key}"):
                    st.session_state[f"show_copy_{state_key}"] = not st.session_state.get(f"show_copy_{state_key}", False)
            with c_doc:
                st.download_button("📥 워드", data=create_docx(title, content), file_name=f"{title}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_docx_{state_key}")
            with c_pdf:
                st.download_button("📥 PDF", data=create_pdf(title, content), file_name=f"{title}.pdf", mime="application/pdf", key=f"dl_pdf_{state_key}")
            with c_ppt:
                st.download_button("📥 PPT", data=create_document_pptx(title, content), file_name=f"{title}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"dl_pptx_{state_key}")
            with c_txt:
                st.download_button("📥 txt", data=create_txt(title, content), file_name=f"{title}.txt", mime="text/plain", key=f"dl_txt_{state_key}")

    if st.session_state.get(f"show_copy_{state_key}", False):
        st.info("💡 아래 상자의 텍스트를 드래그하거나 우측 상단 복사 아이콘을 클릭하세요:")
        st.code(content, language="text")

# --- 6. 성경 66권 목록 정의 ---
BIBLE_BOOKS = [
    "창세기", "출애굽기", "레위기", "민수기", "신명기", "여호수아", "사사기", "룻기", "사무엘상", "사무엘하",
    "열왕기상", "열왕기하", "역대상", "역대하", "에스라", "느헤미야", "에스더", "욥기", "시편", "잠언",
    "전도서", "아가", "이사야", "예레미야", "예레미야애가", "에스겔", "다니엘", "호세아", "요엘", "아모스",
    "오바댜", "요나", "미가", "나훔", "하박국", "스바냐", "학개", "스가랴", "말라기",
    "마태복음", "마가복음", "누가복음", "요한복음", "사도행전", "로마서", "고린도전서", "고린도후서", "갈라디아서", "에베소서",
    "빌립보서", "골로새서", "데살로니가전서", "데살로니가후서", "디모데전서", "디모데후서", "디도서", "빌레몬서", "히브리서", "야고보서",
    "베드로전서", "베드로후서", "요한일서", "요한이서", "요한삼서", "유다서", "요한계시록"
]

# --- 7. 전역 세션 초기화 ---
if "sermon_library" not in st.session_state:
    st.session_state.sermon_library = [
        {
            "id": 1,
            "title": "신앙을 다음 세대에 전수하라",
            "scripture": "시편 78:4-7",
            "theology": "개혁주의/장로교",
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

# --- 8. 메인 내비게이션 바 ---
app_mode = st.sidebar.radio(
    "🕊️ 플랫폼 대메뉴",
    [
        "📊 설교 대시보드 (메인 작업실)",
        "📤 새 설교 등록/원고작성",
        "🎙️ AI 보이스오버 스튜디오",
        "🎬 쇼츠 만들기 (스튜디오)",
        "📚 설교 서재 (Sermon Library)"
    ]
)

# ==============================================================================
# 1. 📊 설교 대시보드 (메인 작업실) - 모든 섹션 상단 6종 툴바 탑재
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

    # 상단 접이식 바 1: 참고 성구 & 예화
    with st.expander("💡 설교를 더 풍성하게 — 참고 구절 & 예화", expanded=False):
        if st.button("✨ 참고 성구 및 신학적 예화 생성하기"):
            with st.spinner("본문과 연관된 성구와 예화를 분석 중입니다..."):
                prompt = f"""
                본문: {st.session_state.sermon_scripture}, 제목: {st.session_state.sermon_title}
                설교문: {st.session_state.full_sermon[:1500]}
                
                오직 100% 한국어로만 작성하세요:
                1. 연관 핵심 참고 성구 3개 및 설교적 연결점
                2. 일상/현대적 공감 예화 2가지
                3. 교회사/고전 문학 및 기독교 사상가 명언 2가지
                """
                st.session_state.rich_materials = get_ai_response(prompt, is_json=False)
        
        if "rich_materials" in st.session_state and st.session_state.rich_materials:
            render_section_top_toolbar(f"{st.session_state.sermon_title}_참고성구및예화", st.session_state.rich_materials, "rich_mat")
            st.markdown(st.session_state.rich_materials)

    # 상단 접이식 바 2: 추천 찬양
    with st.expander("🎵 추천 찬양 — 새찬송가 · 복음성가 · CCM (각 5곡 검색 연결)", expanded=False):
        if st.button("🎶 맞춤 찬양 15곡 추천받기"):
            with st.spinner("설교 메시지와 어울리는 찬양을 선곡 중입니다..."):
                prompt = f"""
                본문: {st.session_state.sermon_scripture}, 제목: {st.session_state.sermon_title}
                JSON: {{"hymns": ["새찬송가 000장 - 제목 5곡"], "gospel_songs": ["복음성가 5곡"], "ccm": ["CCM 5곡"]}}
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

    # 2단 레이아웃 (좌측 메뉴 / 우측 뷰어)
    left_panel, right_panel = st.columns([1, 2.5])
    with left_panel:
        st.markdown("<p style='font-size:12px; font-weight:bold; color:#94a3b8; margin-bottom:4px;'>콘텐츠</p>", unsafe_allow_html=True)
        content_sel = st.radio(
            "콘텐츠 선택",
            ["설교 요약", "소그룹 나눔", "QT 5일치", "카드뉴스", "쇼츠 대본"],
            label_visibility="collapsed"
        )
        st.markdown("<p style='font-size:12px; font-weight:bold; color:#94a3b8; margin-top:20px; margin-bottom:4px;'>이 설교로 만들기</p>", unsafe_allow_html=True)
        maker_sel = st.radio(
            "이 설교로 만들기 선택",
            ["선택 안 함", "🏡 세대별 가정예배지", "🔍 설교 점검 및 제안", "📖 소그룹 리더가이드"],
            label_visibility="collapsed"
        )

    # 우측 뷰어 패널 (제목 + 상단 툴바 + 수정/저장 기능)
    with right_panel:
        active_view = maker_sel if maker_sel != "선택 안 함" else content_sel

        # 1. 설교 요약
        if active_view == "설교 요약":
            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교요약", st.session_state.full_sermon, "sermon_sum")
            st.caption("설교문 본문/요약")
            if st.session_state.get("edit_mode_sermon_sum", False):
                s_edit = st.text_area("설교문 편집", value=st.session_state.full_sermon, height=380, key="edit_sum_area")
                if st.button("💾 본문 저장", key="save_full_sermon"):
                    st.session_state.full_sermon = s_edit
                    st.session_state.edit_mode_sermon_sum = False
                    st.success("저장되었습니다.")
                    st.rerun()
            else:
                st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{st.session_state.full_sermon}</div>", unsafe_allow_html=True)

        # 2. 소그룹 나눔
        elif active_view == "소그룹 나눔":
            grp_txt = st.session_state.get("small_group_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_소그룹나눔지", grp_txt, "sm_grp")
            
            if st.button("✨ 소그룹 나눔 질문 자동 생성", type="primary"):
                with st.spinner("소그룹 나눔지 작성 중..."):
                    prompt = f"""
                    설교 본문: {st.session_state.sermon_scripture}
                    설교 요약: {st.session_state.full_sermon[:3500]}
                    
                    [필수 작성 지침]
                    - 영문 생각 과정(Drafting, Concept, Focus 등)은 절대 출력하지 말고, 오직 100% 한국어로 완성된 나눔지만 작성하세요.
                    - 구성 형식:
                      1. 마음 열기 (아이스브레이크 일상 나눔 질문)
                      2. 말씀 속으로 (본문 및 설교 핵심 이해 질문 2개)
                      3. 삶 속으로 (구체적 실천 및 삶의 적용 질문 2개)
                      4. 마침 합심 기도문
                    """
                    st.session_state.small_group_text = get_ai_response(prompt, is_json=False)
                    st.rerun()

            if grp_txt:
                if st.session_state.get("edit_mode_sm_grp", False):
                    edited_grp = st.text_area("소그룹 나눔지 편집", value=grp_txt, height=350, key="edit_grp_area")
                    if st.button("💾 저장", key="save_grp_btn"):
                        st.session_state.small_group_text = edited_grp
                        st.session_state.edit_mode_sm_grp = False
                        st.success("저장되었습니다.")
                        st.rerun()
                else:
                    st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{grp_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 소그룹 나눔지를 생성하세요.")

        # 3. QT 5일치
        elif active_view == "QT 5일치":
            qt_txt = st.session_state.get("qt5_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_주간QT5일치", qt_txt, "qt5")

            if st.button("✨ 5일치 QT 묵상지 자동 생성", type="primary"):
                with st.spinner("주간 5일치 QT 작성 중..."):
                    prompt = f"""
                    본문: {st.session_state.sermon_scripture}, 설교문: {st.session_state.full_sermon[:3000]}
                    오직 100% 한국어로 월~금 5일치 QT 묵상지를 작성하세요.
                    각 일자별 형식: [월요일~금요일] / [제목] / [성구] / [말씀 묵상 해설] / [적용 질문] / [오늘의 기도]
                    """
                    st.session_state.qt5_text = get_ai_response(prompt, is_json=False)
                    st.rerun()

            if qt_txt:
                if st.session_state.get("edit_mode_qt5", False):
                    edited_qt = st.text_area("QT 5일치 편집", value=qt_txt, height=350, key="edit_qt_area")
                    if st.button("💾 저장", key="save_qt_btn"):
                        st.session_state.qt5_text = edited_qt
                        st.session_state.edit_mode_qt5 = False
                        st.success("저장되었습니다.")
                        st.rerun()
                else:
                    st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{qt_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 5일치 QT를 생성하세요.")

        # 4. 카드뉴스
        elif active_view == "카드뉴스":
            card_all_text = "\n\n".join([f"CARD {c['card_number']}. {c['headline']}\n{c['body_text']}" for c in st.session_state.get("card_list", [])]) if "card_list" in st.session_state else ""
            render_section_top_toolbar(f"{st.session_state.sermon_title}_카드뉴스", card_all_text, "cd_news")
            
            c_cnt = st.slider("카드 장수", 7, 10, 8)
            if st.button("🎨 카드뉴스 문구 생성", type="primary"):
                with st.spinner("카드뉴스 구성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3500]}\n정확히 {c_cnt}장의 카드뉴스 JSON 출력 (100% 한국어): {{\"cards\": [{{\"card_number\": 1, \"headline\": \"제목\", \"body_text\": \"문구\"}}]}}"
                    res = get_ai_response(prompt, is_json=True)
                    if res and "cards" in res:
                        st.session_state.card_list = res["cards"]
                        st.rerun()

            if "card_list" in st.session_state:
                st.download_button("📥 1:1 정사각형 카드뉴스 PPT 내려받기", data=generate_cardnews_pptx(st.session_state.card_list), file_name="카드뉴스_1대1.pptx")
                for c in st.session_state.card_list:
                    st.info(f"**CARD {c['card_number']}. {c['headline']}**\n\n{c['body_text']}")

        # 5. 쇼츠 대본
        elif active_view == "쇼츠 대본":
            sh_txt = st.session_state.get("shorts_script_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_쇼츠대본", sh_txt, "sh_script")

            if st.button("🎬 바이럴 쇼츠 대본 3종 생성", type="primary"):
                with st.spinner("쇼츠 대본 작성 중..."):
                    prompt = f"설교문: {st.session_state.full_sermon[:3000]}\n오직 100% 한국어로 쇼츠/릴스용 60초 대본 3가지 버전(감동형, 질문형, 결단선포형)을 [후킹]-[본론]-[결단] 구조로 작성하세요."
                    st.session_state.shorts_script_text = get_ai_response(prompt, is_json=False)
                    st.rerun()

            if sh_txt:
                if st.session_state.get("edit_mode_sh_script", False):
                    edited_sh = st.text_area("쇼츠 대본 편집", value=sh_txt, height=350, key="edit_sh_area")
                    if st.button("💾 저장", key="save_sh_btn"):
                        st.session_state.shorts_script_text = edited_sh
                        st.session_state.edit_mode_sh_script = False
                        st.success("저장되었습니다.")
                        st.rerun()
                else:
                    st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{sh_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 쇼츠 대본을 생성하세요.")

        # 6. 세대별 가정예배지
        elif active_view == "🏡 세대별 가정예배지":
            age_group = st.selectbox("예배 대상 선택", ["👶 영유아용", "🧒 어린이용", "🧑 청소년용", "👨‍👩‍👧 청장년용"])
            fam_txt = st.session_state.get(f"family_worship_{age_group}", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_가정예배지_{age_group}", fam_txt, f"fam_{age_group}")

            if st.button(f"✨ {age_group} 맞춤 가정예배지 생성", type="primary"):
                with st.spinner("가정예배 순서지 작성 중..."):
                    prompt = f"본문: {st.session_state.sermon_scripture}, 설교문: {st.session_state.full_sermon[:3000]}\n대상: {age_group}\n100% 한국어로 작성: 1.찬양 2.말씀 3.눈높이 3분 메시지 4.가족 나눔 2개 5.축복기도문"
                    st.session_state[f"family_worship_{age_group}"] = get_ai_response(prompt, is_json=False)
                    st.rerun()

            if fam_txt:
                if st.session_state.get(f"edit_mode_fam_{age_group}", False):
                    edited_fam = st.text_area("가정예배지 편집", value=fam_txt, height=320, key=f"edit_fam_{age_group}")
                    if st.button("💾 저장", key=f"save_fam_{age_group}"):
                        st.session_state[f"family_worship_{age_group}"] = edited_fam
                        st.session_state[f"edit_mode_fam_{age_group}"] = False
                        st.success("저장되었습니다.")
                        st.rerun()
                else:
                    st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{fam_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption(f"위 버튼을 눌러 {age_group} 맞춤 가정예배지를 생성하세요.")

        # 7. 설교 점검 및 제안
        elif active_view == "🔍 설교 점검 및 제안":
            audit_txt = st.session_state.get("sermon_audit_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교점검및제안", audit_txt, "sermon_audit")

            if st.button("🔍 설교 전달력 및 신학적 완성도 정밀 점검", type="primary"):
                with st.spinner("설교 분석 및 피드백 작성 중..."):
                    prompt = f"""
                    설교 본문: {st.session_state.sermon_scripture}, 제목: {st.session_state.sermon_title}
                    설교 원고: {st.session_state.full_sermon[:4000]}
                    100% 한국어로 5가지 전문 피드백 작성:
                    1. 본문 주해의 정확성 및 성경 중심성
                    2. 대지 전개 및 구조 평가
                    3. 예화 및 삶의 적용 적절성
                    4. 스피치 전달력 개선 제안
                    5. 총평 및 3가지 권고사항
                    """
                    st.session_state.sermon_audit_text = get_ai_response(prompt, is_json=False)
                    st.rerun()

            if audit_txt:
                if st.session_state.get("edit_mode_sermon_audit", False):
                    edited_audit = st.text_area("설교 점검 내용 편집", value=audit_txt, height=350, key="edit_audit_area")
                    if st.button("💾 저장", key="save_audit_btn"):
                        st.session_state.sermon_audit_text = edited_audit
                        st.session_state.edit_mode_sermon_audit = False
                        st.success("저장되었습니다.")
                        st.rerun()
                else:
                    st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{audit_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 설교 점검 리포트를 생성하세요.")

        # 8. 소그룹 리더가이드
        elif active_view == "📖 소그룹 리더가이드":
            ldr_txt = st.session_state.get("leader_guide_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_소그룹리더가이드", ldr_txt, "ldr_guide")

            if st.button("📖 구역장/순장용 소그룹 리더가이드 생성", type="primary"):
                with st.spinner("소그룹 인도자용 심화 가이드 작성 중..."):
                    prompt = f"""
                    설교 본문: {st.session_state.sermon_scripture}, 설교 요약: {st.session_state.full_sermon[:3500]}
                    100% 한국어로 구역장/셀리더용 리더가이드를 작성하세요:
                    1. 모임의 핵심 목표
                    2. 본문 배경 및 신학적 핵심 요약 (리더용 심화 해설)
                    3. 나눔 질문별 예상 답변 및 인도 팁
                    4. 침묵/돌발 질문 대처 요령
                    5. 소그룹 중보기도 제목 3가지
                    """
                    st.session_state.leader_guide_text = get_ai_response(prompt, is_json=False)
                    st.rerun()

            if ldr_txt:
                if st.session_state.get("edit_mode_ldr_guide", False):
                    edited_ldr = st.text_area("리더가이드 편집", value=ldr_txt, height=350, key="edit_ldr_area")
                    if st.button("💾 저장", key="save_ldr_btn"):
                        st.session_state.leader_guide_text = edited_ldr
                        st.session_state.edit_mode_ldr_guide = False
                        st.success("저장되었습니다.")
                        st.rerun()
                else:
                    st.markdown(f"<div style='background-color:#0f172a; border:1px solid #334155; border-radius:12px; padding:18px; line-height:1.7; white-space:pre-wrap; color:#f1f5f9;'>{ldr_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 소그룹 리더가이드를 생성하세요.")

# ==============================================================================
# 2. 📤 새 설교 등록/원고작성 (직접 타이핑 / 파일 업로드 / AI 강해설교문 3종)
# ==============================================================================
elif app_mode == "📤 새 설교 등록/원고작성":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>📤 새 설교 등록 및 원고 작성</h1>", unsafe_allow_html=True)
    st.caption("직접 타이핑, 파일 업로드, 또는 개혁주의/복음주의/장로교 관점의 AI 강해설교문 자동 생성으로 새 설교를 등록하세요.")

    tab_type, tab_file, tab_ai = st.tabs([
        "✍️ 직접 타이핑 작성",
        "📁 파일 업로드 (.docx, .pdf, .txt)",
        "📖 AI 강해설교문 생성 (개혁주의/복음주의/장로교)"
    ])

    with tab_type:
        st.markdown("#### 강단 선포용 설교 원고 직접 작성")
        tc1, tc2, tc3 = st.columns([2, 1.5, 1])
        with tc1: t_title = st.text_input("설교 제목", placeholder="예: 광야에서 만나는 하나님의 은혜", key="type_title")
        with tc2: t_scripture = st.text_input("성경 본문", placeholder="예: 출애굽기 16:1-12", key="type_scrip")
        with tc3: t_preacher = st.text_input("설교자 성함", value=st.session_state.preacher_name, key="type_preach")
            
        t_tags = st.text_input("설교 태그 (쉼표 구분)", value="주일설교, 은혜, 광야", key="type_tags")
        t_content = st.text_area("설교문 본문 전문", height=380, placeholder="설교 원고를 직접 입력하세요...", key="type_text")
        st.caption(f"글자 수: **{len(t_content):,}자**")
        
        if st.button("💾 새 설교로 등록 및 대시보드 동기화", type="primary", key="save_type_sermon"):
            if not t_title.strip() or not t_content.strip():
                st.warning("설교 제목과 본문을 입력해주세요.")
            else:
                new_entry = {
                    "id": len(st.session_state.sermon_library) + 1,
                    "title": t_title.strip(),
                    "scripture": t_scripture.strip(),
                    "theology": "직접작성",
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "tags": [t.strip() for t in t_tags.split(",") if t.strip()],
                    "text": t_content.strip()
                }
                st.session_state.sermon_library.append(new_entry)
                st.session_state.current_sermon_idx = len(st.session_state.sermon_library) - 1
                st.session_state.sermon_title = t_title.strip()
                st.session_state.sermon_scripture = t_scripture.strip()
                st.session_state.preacher_name = t_preacher.strip()
                st.session_state.full_sermon = t_content.strip()
                st.success(f"'{t_title}' 설교가 등록되었습니다! [📊 설교 대시보드]로 이동하여 확인하세요.")

    with tab_file:
        st.markdown("#### 설교문 파일 업로드 (.docx, .pdf, .txt)")
        u_file = st.file_uploader("설교 파일 선택", type=["docx", "pdf", "txt"], key="up_sermon_file")
        f_title = st.text_input("설교 제목", value="업로드 설교", key="up_title")
        f_scripture = st.text_input("성경 본문", value="본문 구절", key="up_scrip")
        
        if u_file and st.button("📂 파일 읽어와서 서재에 등록하기", type="primary", key="save_up_sermon"):
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
                "theology": "파일업로드",
                "date": datetime.now().strftime("%Y-%m-%d"),
                "tags": ["파일등록"],
                "text": text
            })
            st.success("파일 등록이 완료되었습니다! [📊 설교 대시보드]로 이동하세요.")

    with tab_ai:
        st.markdown("#### 📖 성경 본문 선택 기반 정통 강해설교문 자동 생성")
        ai_c1, ai_c2, ai_c3 = st.columns([1.2, 1.5, 1.3])
        with ai_c1: sel_book = st.selectbox("성경 66권 선택", BIBLE_BOOKS, index=44)
        with ai_c2: sel_chap_verse = st.text_input("장 및 절 입력", value="8장 28절~39절")
        with ai_c3:
            theology_choice = st.selectbox(
                "신학적 관점 선택",
                [
                    "개혁주의 (Reformed - 칼빈주의/하나님 주권)",
                    "장로교 정통 (Presbyterian - 웨스트민스터/구속사)",
                    "복음주의 (Evangelical - 십자가/은혜/복음선포)"
                ]
            )

        ai_t1, ai_t2 = st.columns([2, 1])
        with ai_t1: ai_sermon_topic = st.text_input("설교 주제 / 강조 포인트 (선택)", value="고난 속에서도 흔들리지 않는 하나님의 영원한 사랑과 구원의 확신")
        with ai_t2: sermon_style = st.selectbox("설교 형태", ["3대지 본문중심 강해설교", "구속사적 복음설교", "원어 주해 중심 강해설교"])

        full_scripture_str = f"{sel_book} {sel_chap_verse}"

        if st.button("🚀 정통 강해설교문 전문 자동 작성 시작 (25~30분 분량)", type="primary"):
            with st.spinner(f"[{theology_choice.split(' ')[0]}] 관점으로 {full_scripture_str} 본문 강해설교문 작성 중..."):
                prompt = f"""
                당신은 한국 교회의 탁월한 복음주의/개혁주의/장로교 설교학 교수이자 목회자입니다.
                - 성경 본문: {full_scripture_str}
                - 주제: {ai_sermon_topic}
                - 신학적 관점: {theology_choice}
                - 설교 형태: {sermon_style}
                
                오직 100% 한국어로 25~30분 분량의 완성된 '설교문 전문(Full Manuscript)'을 서론-본론(3대지)-결론 및 결단의 기도 구조로 작성하세요.
                각 대지마다 본문 주해 해설과 청중의 삶에 직결되는 현실적 예화, 구체적 적용 방안을 풍성히 포함하세요.
                어조: 목양적이고 은혜로운 구어체 선포형 어투(~합니다, ~바랍니다).
                """
                generated_sermon = get_ai_response(prompt, is_json=False)
                if generated_sermon:
                    st.session_state.temp_generated_sermon = generated_sermon
                    st.session_state.temp_ai_title = f"{sel_book} 강해: {ai_sermon_topic}"
                    st.session_state.temp_ai_scrip = full_scripture_str
                    st.success("강해설교문 전문이 성공적으로 완성되었습니다!")

        if "temp_generated_sermon" in st.session_state:
            st.write("---")
            render_section_top_toolbar(st.session_state.temp_ai_title, st.session_state.temp_generated_sermon, "ai_gen_sermon")
            st.session_state.temp_generated_sermon = st.text_area("작성된 강해설교문 검토 및 수정", value=st.session_state.temp_generated_sermon, height=450)
            
            if st.button("✅ 이 설교문을 내 서재와 대시보드에 최종 등록하기", type="primary"):
                new_entry = {
                    "id": len(st.session_state.sermon_library) + 1,
                    "title": st.session_state.temp_ai_title,
                    "scripture": st.session_state.temp_ai_scrip,
                    "theology": theology_choice.split(' ')[0],
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "tags": [sel_book, theology_choice.split(' ')[0], "강해설교"],
                    "text": st.session_state.temp_generated_sermon
                }
                st.session_state.sermon_library.append(new_entry)
                st.session_state.current_sermon_idx = len(st.session_state.sermon_library) - 1
                st.session_state.sermon_title = st.session_state.temp_ai_title
                st.session_state.sermon_scripture = st.session_state.temp_ai_scrip
                st.session_state.full_sermon = st.session_state.temp_generated_sermon
                st.success("설교문이 등록되었습니다! [📊 설교 대시보드] 메뉴에서 확인하세요.")

# ==============================================================================
# 3. 🎙️ AI 보이스오버 스튜디오
# ==============================================================================
elif app_mode == "🎙️ AI 보이스오버 스튜디오":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>🎙️ AI 보이스오버 스튜디오</h1>", unsafe_allow_html=True)
    st.caption("설교문이나 요약문을 신경망 음성으로 합성하여 실시간 청취 및 고음질 MP3로 다운로드합니다.")

    vo_col1, vo_col2 = st.columns([1.5, 1])
    with vo_col1:
        vo_text = st.text_area("음성 변환 텍스트", value=st.session_state.full_sermon, height=350)
        vo_voice = st.selectbox("성우 보이스 선택", ["인준 (남성 - 차분하고 신뢰감 있는 톤)", "선희 (여성 - 맑고 또렷한 톤)"])
        voice_id = "ko-KR-InJoonNeural" if "인준" in vo_voice else "ko-KR-SunHiNeural"
        
        if st.button("🎙️ 고음질 음성(TTS) 즉시 생성하기", type="primary"):
            if not vo_text.strip():
                st.warning("텍스트를 입력해주세요.")
            else:
                with st.spinner("AI 음성 생성 중..."):
                    out_audio = asyncio.run(generate_voiceover_audio(vo_text, voice_id))
                    st.session_state.vo_audio_path = out_audio
                    st.success("보이스오버 음성 생성이 완료되었습니다!")

    with vo_col2:
        st.markdown("### 🎧 음성 플레이어 & 다운로드")
        if "vo_audio_path" in st.session_state and os.path.exists(st.session_state.vo_audio_path):
            st.audio(st.session_state.vo_audio_path)
            with open(st.session_state.vo_audio_path, "rb") as af:
                st.download_button("📥 MP3 오디오 파일 다운로드", data=af, file_name=f"{st.session_state.sermon_title}_voice.mp3", mime="audio/mp3")
        else:
            st.info("왼쪽에서 버튼을 누르면 이곳에 재생 플레이어가 나타납니다.")

# ==============================================================================
# 4. 🎬 쇼츠 만들기 (스튜디오)
# ==============================================================================
elif app_mode == "🎬 쇼츠 만들기 (스튜디오)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>▶️ 쇼츠 만들기 스튜디오</h1>", unsafe_allow_html=True)
    
    st.markdown("#### 🌐 무료 미디어/BGM 소스 바로가기")
    src_c1, src_c2, src_c3 = st.columns(3)
    with src_c1: st.link_button("🎥 픽사베이 (Pixabay 무료 영상)", "https://pixabay.com/ko/videos/")
    with src_c2: st.link_button("📸 펙셀스 (Pexels 무료 비디오)", "https://www.pexels.com/ko-kr/videos/")
    with src_c3: st.link_button("🎵 픽사베이 무료 음악(BGM)", "https://pixabay.com/ko/music/")

    st.write("---")
    st.markdown("#### 💡 확 끌리는 5가지 쇼츠 제목 & #해시태그 추천")
    if st.button("✨ 조회수 폭발 5가지 쇼츠 제목 & 해시태그 뽑기"):
        with st.spinner("제목 및 태그 분석 중..."):
            prompt = f"설교제목: {st.session_state.sermon_title}, 요약: {st.session_state.full_sermon[:1500]}\n100% 한국어로 쇼츠 클릭률을 높이는 제목 5개와 해시태그 8개를 JSON으로 작성: {{\"titles\": [\"1.제목\", \"2.제목\", \"3.제목\", \"4.제목\", \"5.제목\"], \"hashtags\": [\"#쇼츠\", \"#은혜\"]}}"
            st.session_state.shorts_rec = get_ai_response(prompt, is_json=True)

    selected_title = st.session_state.sermon_title
    if "shorts_rec" in st.session_state and st.session_state.shorts_rec:
        rec = st.session_state.shorts_rec
        t_choice = st.radio("추천 제목 선택", rec.get("titles", []), horizontal=False)
        if t_choice: selected_title = re.sub(r"^\d+\.\s*", "", t_choice)
        st.markdown("**추천 태그:** " + " ".join(rec.get("hashtags", [])))

    st.write("---")
    sh_col1, sh_col2 = st.columns([1.2, 1])
    with sh_col1:
        v_title = st.text_input("쇼츠 제목", value=selected_title)
        v_script = st.text_area("자막 대본 (줄바꿈 구분)", value="내 열심보다 중요한 것은 하나님의 이끄심입니다.\n우리가 멈출 때 비로소 하나님의 역사가 시작됩니다.\n오늘 그분의 인도하심 앞에 온전히 맡기십시오.", height=120)
        
        c_v1, c_v2 = st.columns(2)
        with c_v1: v_ratio = st.radio("비율", ["9:16 (세로 쇼츠)", "16:9 (가로 영상)"])
        with c_v2: v_voice = st.selectbox("보이스", ["인준 (남성)", "선희 (여성)"])

        bg_media = st.file_uploader("배경 동영상/사진 업로드 (최대 300MB)", type=["mp4", "mov", "jpg", "png"])
        bgm_media = st.file_uploader("배경음악 MP3 업로드 (최대 300MB)", type=["mp3", "wav"])

        if st.button("🚀 비디오 렌더링 시작", type="primary"):
            with st.spinner("자막 애니메이션 및 BGM 믹싱 렌더링 중..."):
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

                rendered_out = create_animated_video(v_title, lines, bg_p, bgm_p, ratio_val, voice=voice_id)
                st.session_state.rendered_shorts_out = rendered_out
                st.success("영상 렌더링이 완료되었습니다!")

    with sh_col2:
        st.markdown("### 🎬 완성된 영상")
        if "rendered_shorts_out" in st.session_state and os.path.exists(st.session_state.rendered_shorts_out):
            st.video(st.session_state.rendered_shorts_out)
            with open(st.session_state.rendered_shorts_out, "rb") as vf:
                st.download_button("📥 MP4 비디오 파일 다운로드", data=vf, file_name="sermon_shorts.mp4", mime="video/mp4")

# ==============================================================================
# 5. 📚 설교 서재 (Sermon Library)
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>설교 서재</h1>", unsafe_allow_html=True)
    st.caption(f"설교문 {len(st.session_state.sermon_library)}편 보관 중")

    search_kw = st.text_input("🔍 제목 또는 성경 구절로 검색", placeholder="예: 시편, 로마서, 신앙...")
    for idx, s_item in enumerate(st.session_state.sermon_library):
        if search_kw and (search_kw not in s_item["title"] and search_kw not in s_item["scripture"]):
            continue
        with st.container():
            st.markdown(
                f"""
                <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 12px; padding: 18px; margin-bottom: 12px;">
                    <h3 style="margin: 0 0 6px 0; font-size: 18px; font-weight: bold; color: #f8fafc;">{s_item['title']}</h3>
                    <p style="margin: 0 0 10px 0; font-size: 13px; color: #94a3b8;">{s_item['scripture']} · [{s_item.get('theology', '개혁주의')}] · 등록일: {s_item['date']}</p>
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
                st.success(f"'{s_item['title']}' 설교를 불러왔습니다! [📊 설교 대시보드]로 이동하세요.")
