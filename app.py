import PIL.Image
import PIL.ImageDraw
import PIL.ImageFont
if not hasattr(PIL.Image, 'ANTIALIAS'):
    PIL.Image.ANTIALIAS = getattr(PIL.Image, 'Resampling', PIL.Image).LANCZOS

import streamlit as st
import google.generativeai as genai
import json
import os
import io
import re
import asyncio
import zipfile
import edge_tts
import urllib.parse
import urllib.request
import yt_dlp
import numpy as np
from datetime import datetime
from docx import Document
from docx.shared import Pt as DocxPt, RGBColor as DocxRGB
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from moviepy.editor import VideoFileClip, ColorClip, CompositeVideoClip, ImageClip
from video_engine import create_animated_video, create_pil_text_clip

st.set_page_config(
    page_title="MY 설교 AI 스튜디오 Pro",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 고급 가스모피즘(Glassmorphism) 모던 UI CSS
st.markdown("""
<style>
    .main {
        background-color: #0b1329;
    }
    div[data-testid="column"] button {
        width: 100% !important;
        padding: 5px 8px !important;
        font-size: 12px !important;
        border-radius: 8px !important;
    }
    .content-box {
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid rgba(51, 65, 85, 0.6);
        border-radius: 16px;
        padding: 24px;
        line-height: 1.85;
        color: #f1f5f9;
        font-size: 15px;
        margin-top: 12px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.3);
        backdrop-filter: blur(10px);
    }
    .content-box h3 {
        color: #fde047;
        font-size: 19px;
        margin-top: 18px;
        margin-bottom: 8px;
        font-weight: bold;
    }
    .card-preview-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        margin: 20px 0;
    }
    .card-box-preview {
        width: 480px;
        height: 520px;
        border-radius: 24px;
        padding: 36px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
        box-shadow: 0 14px 36px rgba(0,0,0,0.55);
        background-size: cover;
        background-position: center;
        position: relative;
        overflow: hidden;
    }
    .card-box-preview::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; bottom: 0;
        background: linear-gradient(180deg, rgba(10, 15, 30, 0.68) 0%, rgba(10, 15, 30, 0.88) 100%);
        z-index: 1;
    }
    .card-box-preview > * {
        position: relative;
        z-index: 2;
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

# --- 2. API 키 숨김 처리 & 초고속 Gemini AI 엔진 ---
secret_key = st.secrets.get("GEMINI_API_KEY", "")

with st.sidebar.expander("⚙️ AI 연결 설정 (클릭하여 열기)", expanded=False):
    sidebar_key = st.text_input("🔑 Gemini API Key", value=secret_key, type="password", key="sidebar_api_key_input")

ACTIVE_KEY = sidebar_key.strip() if sidebar_key else secret_key.strip()

def clean_korean_output(text: str) -> str:
    if not text:
        return ""
    
    markers = [
        r"(\[(?:소그룹|주간|가정예배|60초|참고|설교|세대별|리더|신앙).*?\])",
        r"(###?\s*[0-9가-힣])",
        r"(1\.\s*마음\s*열기)",
        r"(1\.\s*[가-힣]{2,})"
    ]
    for marker in markers:
        match = re.search(marker, text)
        if match:
            text = text[match.start():]
            break

    lines = text.split("\n")
    cleaned_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue
            
        if re.match(r'^(Reliable|100%|Check|Ensure|Draft|Concept|Focus|Selection|Idea|Question|Application|Target|Passage|Summary|Matches|Follows)', stripped, re.IGNORECASE):
            continue
        if stripped.startswith("* Check") or stripped.startswith("* Ensure") or stripped.startswith("* Focus:"):
            continue
        if re.search(r'^\*\s*\*(Concept|Drafting|Selection|Draft|Idea \d+|Content|Focus):\*', stripped):
            continue
            
        k_chars = len(re.findall(r'[가-힣]', stripped))
        e_chars = len(re.findall(r'[a-zA-Z]', stripped))
        if e_chars > 8 and k_chars == 0:
            continue
            
        line = re.sub(r'\([A-Za-z0-9\s,\.\?\!\'\":;\-\/]{5,}\)', '', line)
        cleaned_lines.append(line)
        
    result = "\n".join(cleaned_lines).strip()
    result = re.sub(r'(\n\s*[\*\-•]\s*)\n+(\s*)', r'\1 ', result)
    result = re.sub(r'(\n\s*\d+\.\s*)\n+(\s*)', r'\1 ', result)
    result = re.sub(r'(\*\s*)\n+(\s*)', r'\1 ', result)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result

def get_ai_response(prompt: str, is_json: bool = True):
    if not ACTIVE_KEY:
        st.error("🔑 사이드바의 [⚙️ AI 연결 설정] 메뉴에 Gemini API Key를 입력해주세요.")
        return None
    try:
        genai.configure(api_key=ACTIVE_KEY)
    except Exception as e:
        st.error(f"API 키 설정 오류: {str(e)}")
        return None

    valid_models = []
    try:
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                if "gemini-1.0" not in m.name and m.name != "models/gemini-pro":
                    valid_models.append(m.name)
    except Exception:
        pass

    def model_priority(m_name):
        m = m_name.lower()
        if "1.5-flash" in m: return 1
        elif "2.0-flash" in m: return 2
        elif "flash" in m: return 3
        elif "1.5-pro" in m: return 4
        return 10

    if valid_models:
        valid_models.sort(key=model_priority)
    else:
        valid_models = ["models/gemini-1.5-flash", "gemini-1.5-flash", "models/gemini-2.0-flash", "models/gemini-1.5-pro"]

    system_instruction = (
        "당신은 한국 교회의 사역을 돕는 목회 전문 어시스턴트입니다. "
        "영문 생각 과정, 기획 메모(Drafting, Concept, Focus, Checklist 등)나 영어 단어는 일절 작성하지 마십시오. "
        "글머리 기호(*)나 번호(1.) 바로 뒤에 줄바꿈 없이 본문을 이어서 100% 완성된 한국어 사역 문서 본문만 바로 출력하십시오."
    )

    errors = []
    for model_name in valid_models:
        try:
            try:
                model = genai.GenerativeModel(model_name, system_instruction=system_instruction)
            except Exception:
                model = genai.GenerativeModel(model_name)
            
            if is_json:
                res = model.generate_content(
                    prompt,
                    generation_config={"response_mime_type": "application/json", "temperature": 0.2}
                )
                return json.loads(res.text)
            else:
                res = model.generate_content(prompt, generation_config={"temperature": 0.3})
                if res and res.text:
                    return clean_korean_output(res.text)
        except Exception as e:
            errors.append(f"[{model_name}] {str(e)}")
            continue

    st.error(f"AI 호출 실패: {errors[-1] if errors else '모델 응답 오류'}")
    return None

# --- 3. 폰트 캐싱 엔진 ---
PDF_FONT_NAME = "Helvetica"

@st.cache_resource(show_spinner=False)
def init_korean_font():
    paths = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
        "/usr/share/fonts/opentype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/nanum/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        "./NanumGothic.ttf"
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("NanumKorean", p))
                return "NanumKorean"
            except Exception:
                pass

    local_f = "./NanumGothic.ttf"
    if not os.path.exists(local_f):
        try:
            url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
            urllib.request.urlretrieve(url, local_f)
        except Exception:
            pass
    if os.path.exists(local_f):
        try:
            pdfmetrics.registerFont(TTFont("NanumKorean", local_f))
            return "NanumKorean"
        except Exception:
            pass
    return "Helvetica"

PDF_FONT_NAME = init_korean_font()

# --- 4. 고화질 풍경 배경 이미지 큐레이션 목록 ---
CARD_BACKGROUNDS = [
    "https://images.unsplash.com/photo-1506744038136-46273834b3fb?w=1080&q=80",
    "https://images.unsplash.com/photo-1518495973542-4542c06a5843?w=1080&q=80",
    "https://images.unsplash.com/photo-1464822759023-fed622ff2c3b?w=1080&q=80",
    "https://images.unsplash.com/photo-1448375240586-882707db888b?w=1080&q=80",
    "https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=1080&q=80",
    "https://images.unsplash.com/photo-1519681393784-d120267933ba?w=1080&q=80",
    "https://images.unsplash.com/photo-1495616811223-4d98c6e9c869?w=1080&q=80",
    "https://images.unsplash.com/photo-1470252649378-9c29740c9fa8?w=1080&q=80"
]

@st.cache_data(show_spinner=False)
def fetch_image_bytes(url: str):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=4) as response:
            return response.read()
    except Exception:
        return None

# --- 카드뉴스 PNG 이미지 고화질 합성 엔진 ---
def generate_single_card_png(card_item, idx, scripture_str="", church_name=""):
    bg_url = CARD_BACKGROUNDS[idx % len(CARD_BACKGROUNDS)]
    img_b = fetch_image_bytes(bg_url)
    
    if img_b:
        base_img = PIL.Image.open(io.BytesIO(img_b)).convert("RGBA").resize((1080, 1080))
    else:
        base_img = PIL.Image.new("RGBA", (1080, 1080), (15, 23, 42, 255))

    overlay = PIL.Image.new("RGBA", (1080, 1080), (10, 15, 30, 200))
    combined = PIL.Image.alpha_composite(base_img, overlay)
    draw = PIL.ImageDraw.Draw(combined)

    font_b = None
    for f_p in ["/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf", "C:/Windows/Fonts/malgun.ttf"]:
        if os.path.exists(f_p):
            try: font_b = PIL.ImageFont.truetype(f_p, 48); break
            except Exception: pass
    if not font_b: font_b = PIL.ImageFont.load_default()

    font_t = None
    for f_p in ["/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf", "C:/Windows/Fonts/malgun.ttf"]:
        if os.path.exists(f_p):
            try: font_t = PIL.ImageFont.truetype(f_p, 32); break
            except Exception: pass
    if not font_t: font_t = PIL.ImageFont.load_default()

    draw.text((100, 100), f"CARD {card_item.get('card_number', idx+1)}", fill=(99, 102, 241, 255), font=font_t)
    draw.text((100, 200), card_item.get("headline", ""), fill=(253, 224, 71, 255), font=font_b)
    draw.multiline_text((100, 420), card_item.get("body_text", ""), fill=(241, 245, 249, 255), font=font_t, spacing=16)

    if scripture_str:
        draw.text((100, 920), f"「 {scripture_str} 」", fill=(253, 224, 71, 255), font=font_t)
    if church_name:
        draw.text((100, 970), church_name, fill=(147, 197, 253, 255), font=font_t)

    out_buf = io.BytesIO()
    combined.convert("RGB").save(out_buf, format="PNG")
    out_buf.seek(0)
    return out_buf

def generate_cardnews_zip(cards, scripture_str="", church_name=""):
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for i, card in enumerate(cards):
            png_buf = generate_single_card_png(card, i, scripture_str, church_name)
            zf.writestr(f"cardnews_{i+1}.png", png_buf.getvalue())
    zip_buf.seek(0)
    return zip_buf

# --- 5. 유튜브 비디오 다운로드 및 9:16 쇼츠 추출 엔진 (HTTP 403 Forbidden 우회 적용) ---
def extract_youtube_to_shorts(yt_url: str, start_sec: int, duration_sec: int, title: str, subtitle_text: str, church_name: str = ""):
    out_dir = "./outputs"
    os.makedirs(out_dir, exist_ok=True)
    source_template = os.path.join(out_dir, "yt_raw_source.%(ext)s")
    
    # URL 정제 (/live/주소 또는 URL 파라미터 호환 정규화)
    clean_url = re.sub(r'youtube\.com/live/([a-zA-Z0-9_-]+)', r'youtube.com/watch?v=\1', yt_url.strip())
    
    ydl_opts = {
        'format': 'bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best',
        'outtmpl': source_template,
        'overwrites': True,
        'quiet': True,
        'no_warnings': True,
        'nocheckcertificate': True,
        'http_headers': {
            'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Mobile/15E148 Safari/604.1',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        },
        'extractor_args': {
            'youtube': {
                'player_client': ['ios', 'android', 'mweb']
            }
        }
    }
    
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download([clean_url])
        
    src_video = os.path.join(out_dir, "yt_raw_source.mp4")
    if not os.path.exists(src_video):
        for f in os.listdir(out_dir):
            if f.startswith("yt_raw_source"):
                src_video = os.path.join(out_dir, f)
                break
                
    raw_clip = VideoFileClip(src_video)
    max_dur = raw_clip.duration
    start_sec = max(0, min(start_sec, int(max_dur) - 5))
    end_sec = min(start_sec + duration_sec, int(max_dur))
    
    sub_clip = raw_clip.subclip(start_sec, end_sec)
    w, h = sub_clip.size
    
    target_w = int(h * (9 / 16))
    if w > target_w:
        x_center = w / 2
        cropped = sub_clip.crop(x1=x_center - target_w/2, y1=0, x2=x_center + target_w/2, y2=h)
    else:
        cropped = sub_clip
        
    final_video_clip = cropped.resize((1080, 1920))
    clip_dur = final_video_clip.duration
    
    overlays = [final_video_clip]
    
    top_bar = ColorClip(size=(1080, 240), color=(0,0,0), duration=clip_dur).set_opacity(0.45).set_position(('center', 100))
    overlays.append(top_bar)
    
    title_clip = (
        create_pil_text_clip(
            title,
            fontsize=48,
            color="#FDE047",
            stroke_color="black",
            stroke_width=2,
            size=(920, None),
            duration=clip_dur
        )
        .set_position(("center", 140))
    )
    overlays.append(title_clip)
    
    if subtitle_text:
        sub_clip_txt = (
            create_pil_text_clip(
                subtitle_text,
                fontsize=40,
                color="white",
                stroke_color="black",
                stroke_width=2,
                size=(900, None),
                duration=clip_dur
            )
            .set_position(("center", 1400))
        )
        overlays.append(sub_clip_txt)
        
    if church_name:
        church_clip = (
            create_pil_text_clip(
                church_name,
                fontsize=28,
                color="#93C5FD",
                stroke_color="black",
                stroke_width=1,
                size=(800, None),
                duration=clip_dur
            )
            .set_position(("center", 1780))
        )
        overlays.append(church_clip)
        
    comp = CompositeVideoClip(overlays)
    out_file = os.path.join(out_dir, f"yt_shorts_extracted_{int(datetime.now().timestamp())}.mp4")
    comp.write_videofile(out_file, fps=24, codec="libx264", audio_codec="aac", threads=4, preset="ultrafast")
    
    raw_clip.close()
    sub_clip.close()
    comp.close()
    return out_file

# --- 6. 문서 변환 엔진 ---
def create_docx(title: str, content: str) -> io.BytesIO:
    try:
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
    except Exception:
        return io.BytesIO(content.encode("utf-8"))

def create_pdf(title: str, content: str) -> io.BytesIO:
    try:
        font_to_use = init_korean_font()
        bio = io.BytesIO()
        doc = SimpleDocTemplate(
            bio,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36
        )
        
        t_style = ParagraphStyle(
            name="K_Title",
            fontName=font_to_use,
            fontSize=15,
            leading=20,
            textColor="#1e3a8a",
            spaceAfter=8
        )
        m_style = ParagraphStyle(
            name="K_Meta",
            fontName=font_to_use,
            fontSize=8,
            leading=12,
            textColor="#64748b",
            spaceAfter=12
        )
        b_style = ParagraphStyle(
            name="K_Body",
            fontName=font_to_use,
            fontSize=9.5,
            leading=15,
            textColor="#1e293b",
            spaceAfter=5
        )

        story = [
            Paragraph(f"<b>{title}</b>", t_style),
            Paragraph(f"생성일: {datetime.now().strftime('%Y-%m-%d')} | MY 설교 AI 스튜디오", m_style),
            Spacer(1, 8),
        ]

        for line in content.split("\n"):
            clean = line.strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if clean:
                clean = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean)
                story.append(Paragraph(clean, b_style))
            else:
                story.append(Spacer(1, 4))

        doc.build(story)
        bio.seek(0)
        return bio
    except Exception:
        return io.BytesIO(content.encode("utf-8"))

def create_txt(title: str, content: str) -> io.BytesIO:
    text_data = f"[{title}]\n작성일: {datetime.now().strftime('%Y-%m-%d')}\n\n{content}"
    return io.BytesIO(text_data.encode("utf-8"))

def create_document_pptx(title: str, content: str) -> io.BytesIO:
    try:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank_layout = prs.slide_layouts[6]
        
        slide1 = prs.slides.add_slide(blank_layout)
        fill1 = slide1.background.fill
        fill1.solid()
        fill1.fore_color.rgb = RGBColor(15, 23, 42)
        tbox = slide1.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(10.33), Inches(2.5))
        p = tbox.text_frame.paragraphs[0]
        p.text = title
        p.font.size, p.font.bold = Pt(38), True
        p.font.color.rgb, p.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER
        
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
    except Exception:
        return io.BytesIO(content.encode("utf-8"))

def generate_sermon_structure_pptx(title: str, scripture: str, content: str) -> io.BytesIO:
    try:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        def set_bg_with_overlay(slide, img_url=None, color=RGBColor(15, 23, 42)):
            if img_url:
                img_b = fetch_image_bytes(img_url)
                if img_b:
                    slide.shapes.add_picture(io.BytesIO(img_b), 0, 0, width=Inches(13.333), height=Inches(7.5))
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = color

            overlay = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = RGBColor(10, 15, 30)
            overlay.line.fill.background()

        s1 = prs.slides.add_slide(blank_layout)
        set_bg_with_overlay(s1, CARD_BACKGROUNDS[0])
        tb1 = s1.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.8))
        p1 = tb1.text_frame.paragraphs[0]
        p1.text = f"주 일 설 교\n\n{title}\n\n보이지 않는 가장 고귀한 유산\n본문 · {scripture}"
        p1.font.size, p1.font.bold = Pt(38), True
        p1.font.color.rgb, p1.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER

        s2 = prs.slides.add_slide(blank_layout)
        s2.background.fill.solid()
        s2.background.fill.fore_color.rgb = RGBColor(248, 250, 252)
        
        tb2 = s2.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(5.8))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        p2_head = tf2.paragraphs[0]
        p2_head.text = f"들어가며 & 설교의 흐름 ({scripture})"
        p2_head.font.size, p2_head.font.bold = Pt(32), True
        p2_head.font.color.rgb = RGBColor(30, 58, 138)
        
        p2_body = tf2.add_paragraph()
        p2_body.text = f"\n01. 침묵은 곧 삭제입니다\n02. 하나씩 세어가며 전수하라\n03. 부지런히 새기고 가르치라\n04. 가문을 바꾼 한 사람의 결단\n\n{content[:300]}"
        p2_body.font.size = Pt(20)
        p2_body.font.color.rgb = RGBColor(30, 41, 59)

        paragraphs = [p.strip() for p in content.split("\n") if p.strip()]
        chunk_size = max(1, len(paragraphs) // 7)
        
        for idx in range(3, 11):
            s = prs.slides.add_slide(blank_layout)
            if idx % 2 == 1:
                set_bg_with_overlay(s, CARD_BACKGROUNDS[idx % len(CARD_BACKGROUNDS)])
                head_color = RGBColor(253, 224, 71)
                body_color = RGBColor(241, 245, 249)
            else:
                s.background.fill.solid()
                s.background.fill.fore_color.rgb = RGBColor(248, 250, 252)
                head_color = RGBColor(30, 58, 138)
                body_color = RGBColor(30, 41, 59)

            tb_head = s.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(1.0))
            hp = tb_head.text_frame.paragraphs[0]
            hp.text = f"대지 메시지 {idx-2} · {title}"
            hp.font.size, hp.font.bold = Pt(30), True
            hp.font.color.rgb = head_color

            tb_body = s.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.33), Inches(4.8))
            tf_b = tb_body.text_frame
            tf_b.word_wrap = True
            
            p_slice = paragraphs[(idx-3)*chunk_size : (idx-2)*chunk_size]
            slice_text = "\n\n".join(p_slice) if p_slice else f"{title} 핵심 적용 및 축복 선포"
            
            pb = tf_b.paragraphs[0]
            pb.text = slice_text
            pb.font.size = Pt(20)
            pb.font.color.rgb = body_color

        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio
    except Exception:
        return create_document_pptx(title, content)

def generate_cardnews_pptx(slides_data, church_name=""):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(10)
    blank_layout = prs.slide_layouts[6]

    for idx, item in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)
        bg_url = CARD_BACKGROUNDS[idx % len(CARD_BACKGROUNDS)]
        img_bytes = fetch_image_bytes(bg_url)
        if img_bytes:
            img_stream = io.BytesIO(img_bytes)
            slide.shapes.add_picture(img_stream, Inches(0), Inches(0), width=Inches(10), height=Inches(10))
        else:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(15, 23, 42)

        overlay = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(10))
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = RGBColor(10, 15, 30)
        overlay.line.fill.background()
        
        badge_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(2.2), Inches(0.6))
        bp = badge_box.text_frame.paragraphs[0]
        bp.text = f"CARD {item.get('card_number', idx + 1)}"
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

        if church_name:
            cbox = slide.shapes.add_textbox(Inches(0.8), Inches(9.0), Inches(8.4), Inches(0.6))
            cp = cbox.text_frame.paragraphs[0]
            cp.text = church_name
            cp.font.size = Pt(14)
            cp.font.color.rgb, cp.alignment = RGBColor(147, 197, 253), PP_ALIGN.CENTER

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

# --- 7. 모든 섹션 상단 통일 툴바 컴포넌트 [수정 / 복사 / 워드 / PDF / PPT / txt] ---
def render_section_top_toolbar(title: str, content: str, state_key: str):
    col_t, col_btns = st.columns([1.2, 2.8])
    with col_t:
        st.markdown(f"<h3 style='margin: 0; padding: 0; font-size: 20px; font-weight: 800; line-height: 1.3;'>{title}</h3>", unsafe_allow_html=True)
    with col_btns:
        c_edit, c_copy, c_doc, c_pdf, c_ppt, c_txt = st.columns([1, 1, 1.1, 1.1, 1.1, 1])
        with c_edit:
            if st.button("✏️ 수정", key=f"edit_btn_{state_key}"):
                st.session_state[f"edit_mode_{state_key}"] = not st.session_state.get(f"edit_mode_{state_key}", False)
        with c_copy:
            if st.button("📋 복사", key=f"copy_btn_{state_key}"):
                st.session_state[f"show_copy_{state_key}"] = not st.session_state.get(f"show_copy_{state_key}", False)
        with c_doc:
            st.download_button("📥 워드", data=create_docx(title, content if content else "내용 없음"), file_name=f"{title}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"dl_docx_{state_key}")
        with c_pdf:
            st.download_button("📥 PDF", data=create_pdf(title, content if content else "내용 없음"), file_name=f"{title}.pdf", mime="application/pdf", key=f"dl_pdf_{state_key}")
        with c_ppt:
            st.download_button("📥 PPT", data=generate_sermon_structure_pptx(title, st.session_state.get("sermon_scripture", "본문"), content if content else "내용 없음"), file_name=f"{title}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"dl_ppt_struct_{state_key}")
        with c_txt:
            st.download_button("📥 txt", data=create_txt(title, content if content else "내용 없음"), file_name=f"{title}.txt", mime="text/plain", key=f"dl_txt_{state_key}")

    if st.session_state.get(f"show_copy_{state_key}", False):
        st.info("💡 아래 상자의 텍스트를 복사하여 사용하세요:")
        st.code(content, language="text")

# --- 8. 성경 66권 목록 ---
BIBLE_BOOKS = [
    "창세기", "출애굽기", "레위기", "민수기", "신명기", "여호수아", "사사기", "룻기", "사무엘상", "사무엘하",
    "열왕기상", "열왕기하", "역대상", "역대하", "에스라", "느헤미야", "에스더", "욥기", "시편", "잠언",
    "전도서", "아가", "이사야", "예레미야", "예레미야애가", "에스겔", "다니엘", "호세아", "요엘", "아모스",
    "오바댜", "요나", "미가", "나훔", "하박국", "스바냐", "학개", "스가랴", "말라기",
    "마태복음", "마가복음", "누가복음", "요한복음", "사도행전", "로마서", "고린도전서", "고린도후서", "갈라디아서", "에베소서",
    "빌립보서", "골로새서", "데살로니가전서", "데살로니가후서", "디모데전서", "디모데후서", "디도서", "빌레몬서", "히브리서", "야고보서",
    "베드로전서", "베드로후서", "요한일서", "요한이서", "요한삼서", "유다서", "요한계시록"
]

# --- 9. 전역 세션 초기화 ---
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
if "dash_active_view" not in st.session_state:
    st.session_state.dash_active_view = "설교 요약"

# --- 10. 메인 내비게이션 바 ---
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

    with st.expander("💡 설교를 더 풍성하게 — 참고 구절 & 예화", expanded=False):
        if st.button("✨ 참고 성구 및 신학적 예화 생성하기", key="btn_gen_rich"):
            with st.spinner("본문과 연관된 성구와 예화를 분석 중입니다..."):
                prompt = f"""
                성경 본문: {st.session_state.sermon_scripture}
                설교 제목: {st.session_state.sermon_title}
                설교 요약: {st.session_state.full_sermon[:1500]}
                
                [참고 성구 및 예화 자료집: {st.session_state.sermon_title}]
                
                1. 본문 연관 핵심 참고 성구 3가지 및 설교적 연결점
                2. 일상 및 현대적 공감 예화 2가지
                3. 교회사 및 기독교 사상가 명언 2가지
                """
                st.session_state.rich_materials = get_ai_response(prompt, is_json=False)
        
        rich_mat_content = st.session_state.get("rich_materials", "")
        if rich_mat_content:
            render_section_top_toolbar(f"{st.session_state.sermon_title}_참고성구및예화", rich_mat_content, "rich_mat")
            st.markdown(f"<div class='content-box'>{rich_mat_content}</div>", unsafe_allow_html=True)

    with st.expander("🎵 추천 찬양 — 새찬송가 · 복음성가 · CCM (각 5곡 검색 연결)", expanded=False):
        if st.button("🎶 맞춤 찬양 15곡 추천받기", key="btn_gen_praise"):
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

    left_panel, right_panel = st.columns([1, 2.5])
    
    with left_panel:
        st.markdown("<p style='font-size:12px; font-weight:bold; color:#94a3b8; margin-bottom:4px;'>사역 메뉴 선택</p>", unsafe_allow_html=True)
        selected_menu = st.radio(
            "사역 메뉴",
            [
                "설교 요약",
                "소그룹 나눔",
                "QT 5일치",
                "카드뉴스",
                "쇼츠 대본",
                "🏡 세대별 가정예배지",
                "🔍 설교 점검 및 제안",
                "📖 소그룹 리더가이드"
            ],
            index=[
                "설교 요약", "소그룹 나눔", "QT 5일치", "카드뉴스", "쇼츠 대본",
                "🏡 세대별 가정예배지", "🔍 설교 점검 및 제안", "📖 소그룹 리더가이드"
            ].index(st.session_state.dash_active_view) if st.session_state.dash_active_view in [
                "설교 요약", "소그룹 나눔", "QT 5일치", "카드뉴스", "쇼츠 대본",
                "🏡 세대별 가정예배지", "🔍 설교 점검 및 제안", "📖 소그룹 리더가이드"
            ] else 0,
            key="dash_menu_selector"
        )
        st.session_state.dash_active_view = selected_menu

    with right_panel:
        active_view = st.session_state.dash_active_view

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
                st.markdown(f"<div class='content-box'>{st.session_state.full_sermon}</div>", unsafe_allow_html=True)

        elif active_view == "소그룹 나눔":
            grp_txt = st.session_state.get("small_group_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_소그룹나눔지", grp_txt, "sm_grp")
            
            if st.button("✨ 소그룹 나눔 질문 자동 생성", type="primary", key="btn_gen_sm_grp"):
                with st.spinner("소그룹 나눔지 작성 중..."):
                    prompt = f"""
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 요약: {st.session_state.full_sermon[:3500]}
                    
                    [소그룹 나눔지: {st.session_state.sermon_title}]
                    
                    1. 마음 열기 (아이스브레이크)
                    - (일상의 따뜻한 나눔 질문 1가지)
                    
                    2. 말씀 속으로
                    - 1. (본문 말씀 이해 질문)
                    - 2. (설교 핵심 메시지 나눔 질문)
                    
                    3. 삶 속으로
                    - 1. (구체적 실천 방안 질문)
                    - 2. (한 주간의 결단 질문)
                    
                    4. 마침 합심 기도문
                    - (은혜로운 마무리 기도문)
                    """
                    res = get_ai_response(prompt, is_json=False)
                    if res:
                        st.session_state.small_group_text = res
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
                    st.markdown(f"<div class='content-box'>{grp_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 소그룹 나눔지를 생성하세요.")

        elif active_view == "QT 5일치":
            qt_txt = st.session_state.get("qt5_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_주간QT5일치", qt_txt, "qt5")

            if st.button("✨ 5일치 QT 묵상지 자동 생성", type="primary", key="btn_gen_qt5"):
                with st.spinner("주간 5일치 QT 작성 중..."):
                    prompt = f"""
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 요약: {st.session_state.full_sermon[:3000]}
                    
                    [주간 QT 5일치: {st.session_state.sermon_title}]
                    
                    월요일부터 금요일까지 5일치 말씀 묵상지를 작성하세요:
                    각 날짜마다:
                    - 제목:
                    - 본문 구절:
                    - 말씀 묵상 해설:
                    - 삶의 적용 질문:
                    - 오늘의 기도:
                    """
                    res = get_ai_response(prompt, is_json=False)
                    if res:
                        st.session_state.qt5_text = res
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
                    st.markdown(f"<div class='content-box'>{qt_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 5일치 QT를 생성하세요.")

        elif active_view == "카드뉴스":
            card_all_text = "\n\n".join([f"CARD {c['card_number']}. {c['headline']}\n{c['body_text']}" for c in st.session_state.get("card_list", [])]) if "card_list" in st.session_state else ""
            
            c_head1, c_head2 = st.columns([1.2, 1.8])
            with c_head1:
                st.markdown("<h2 style='margin:0; font-size:24px; font-weight:bold;'>카드뉴스</h2>", unsafe_allow_html=True)
            with c_head2:
                col_e, col_d_ppt, col_d_zip = st.columns([1, 1.3, 1.4])
                with col_e:
                    if st.button("✏️ 편집", key="cn_edit_toggle_btn"):
                        st.session_state.cn_edit_mode = not st.session_state.get("cn_edit_mode", False)
                with col_d_ppt:
                    if "card_list" in st.session_state:
                        st.download_button(
                            "📥 PPT 전체",
                            data=generate_cardnews_pptx(st.session_state.card_list, st.session_state.get("cn_church_name", "")),
                            file_name=f"{st.session_state.sermon_title}_카드뉴스.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key="cn_dl_all_ppt_btn"
                        )
                with col_d_zip:
                    if "card_list" in st.session_state:
                        st.download_button(
                            "📦 이미지 (ZIP)",
                            data=generate_cardnews_zip(st.session_state.card_list, st.session_state.sermon_scripture, st.session_state.get("cn_church_name", "")),
                            file_name=f"{st.session_state.sermon_title}_카드뉴스_이미지.zip",
                            mime="application/zip",
                            key="cn_dl_all_zip_btn"
                        )

            st.info("💡 텍스트 수정은 물론, 카드 추가/삭제 및 이미지(PNG/ZIP) / PPT 내려받기가 가능합니다!")

            cn_opt1, cn_opt2, cn_opt3 = st.columns([1.5, 1.5, 2])
            with cn_opt1:
                bg_opt = st.radio("배경", ["사진", "기본", "갤러리", "직접 업로드"], horizontal=True, key="cn_bg_opt")
            with cn_opt2:
                font_opt = st.selectbox("글씨체", ["프리텐다드", "나눔고딕", "본고딕"], key="cn_font_opt")
                font_size = st.number_input("크기", min_value=80, max_value=140, value=100, step=5, key="cn_font_size")
            with cn_opt3:
                church_input = st.text_input("교회명", value=st.session_state.get("cn_church_name", "화광교회"), placeholder="교회 이름 (Enter로 줄바꿈, 최대 2줄)", key="cn_church_input")
                st.session_state.cn_church_name = church_input

            st.write("---")

            if "card_list" not in st.session_state or not st.session_state.card_list:
                if st.button("🎨 카드뉴스 자동 생성하기", type="primary", key="btn_gen_cardnews_init"):
                    with st.spinner("설교 메시지로 카드뉴스 구성 중..."):
                        prompt = f"설교 본문: {st.session_state.sermon_scripture}\n설교문: {st.session_state.full_sermon[:3500]}\n정확히 7장의 카드뉴스 JSON 출력 (100% 한국어): {{\"cards\": [{{\"card_number\": 1, \"headline\": \"제목\", \"body_text\": \"문구\"}}]}}"
                        res = get_ai_response(prompt, is_json=True)
                        if res and "cards" in res:
                            st.session_state.card_list = res["cards"]
                            st.rerun()

            if st.session_state.get("cn_edit_mode", False) and "card_list" in st.session_state:
                st.markdown("#### ✏️ 카드뉴스 텍스트 편집기")
                for c_i, card_item in enumerate(st.session_state.card_list):
                    with st.expander(f"CARD {card_item.get('card_number', c_i+1)} 편집", expanded=(c_i==0)):
                        card_item["headline"] = st.text_input(f"카드 {c_i+1} 헤드라인", value=card_item.get("headline", ""), key=f"cn_h_{c_i}")
                        card_item["body_text"] = st.text_area(f"카드 {c_i+1} 본문", value=card_item.get("body_text", ""), key=f"cn_b_{c_i}")

            if "card_list" in st.session_state and st.session_state.card_list:
                cards = st.session_state.card_list
                if "cn_card_idx" not in st.session_state:
                    st.session_state.cn_card_idx = 0

                total_cards = len(cards)
                curr_idx = st.session_state.cn_card_idx % total_cards
                curr_card = cards[curr_idx]

                car_c1, car_c2, car_c3 = st.columns([1, 4, 1])
                with car_c1:
                    st.markdown("<div style='height: 220px;'></div>", unsafe_allow_html=True)
                    if st.button("❮ 이전", key="cn_prev_btn"):
                        st.session_state.cn_card_idx = (curr_idx - 1) % total_cards
                        st.rerun()
                with car_c3:
                    st.markdown("<div style='height: 220px;'></div>", unsafe_allow_html=True)
                    if st.button("다음 ❯", key="cn_next_btn"):
                        st.session_state.cn_card_idx = (curr_idx + 1) % total_cards
                        st.rerun()

                bg_img_url = CARD_BACKGROUNDS[curr_idx % len(CARD_BACKGROUNDS)]
                
                with car_c2:
                    st.markdown(
                        f"""
                        <div class="card-preview-container">
                            <div class="card-box-preview" style="background-image: url('{bg_img_url}');">
                                <h2 style="color: #ffffff; font-size: {int(26*(font_size/100))}px; font-weight: bold; margin-bottom: 12px; text-shadow: 0 2px 4px rgba(0,0,0,0.8);">{curr_card.get('headline', '')}</h2>
                                <p style="color: #e2e8f0; font-size: {int(16*(font_size/100))}px; line-height: 1.7; text-shadow: 0 1px 3px rgba(0,0,0,0.8); white-space: pre-wrap;">{curr_card.get('body_text', '')}</p>
                                <div style="margin-top: 24px; color: #fde047; font-size: 14px; font-weight: bold;">「 {st.session_state.sermon_scripture} 」</div>
                                <div style="margin-top: 16px; color: #93c5fd; font-size: 12px;">{church_input}</div>
                            </div>
                            <div style="color: #94a3b8; font-size: 13px; margin-top: 14px;">
                                <strong>{curr_idx + 1} / {total_cards} 슬라이드</strong> (← → 키로 이동)
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                    
                    single_png_bytes = generate_single_card_png(curr_card, curr_idx, st.session_state.sermon_scripture, church_input)
                    st.download_button(
                        f"🖼️ CARD {curr_idx + 1} 개별 PNG 이미지 다운로드",
                        data=single_png_bytes,
                        file_name=f"{st.session_state.sermon_title}_card_{curr_idx + 1}.png",
                        mime="image/png",
                        key=f"dl_single_png_{curr_idx}"
                    )

                st.write("---")

                st.markdown("#### 인스타그램 캡션")
                insta_c1, insta_c2 = st.columns([4, 1])
                
                insta_text = f"집은 물려주려고 평생을 아끼고 통장은 자녀 이름으로 만들어 두면서, 정작 가장 귀한 하나님 이야기는 언제 마지막으로 들려주셨나요?\n\n오늘 선포된 [{st.session_state.sermon_title}] 말씀을 통해 믿음의 유산을 전하는 가정이 되기를 축복합니다."
                insta_tags = "#주일설교 #신앙전수 #말씀묵상 #가정예배 #시편78편 #크리스천"

                with insta_c1:
                    st.info(insta_text)
                    st.code(insta_tags, language="text")
                with insta_c2:
                    if st.button("📋 전체 복사", key="btn_copy_insta"):
                        st.success("인스타그램 캡션이 복사되었습니다!")

        elif active_view == "쇼츠 대본":
            sh_txt = st.session_state.get("shorts_script_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_쇼츠대본", sh_txt, "sh_script")

            if st.button("🎬 바이럴 쇼츠 대본 3종 생성", type="primary", key="btn_gen_shorts_script"):
                with st.spinner("쇼츠 대본 작성 중..."):
                    prompt = f"""
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 요약: {st.session_state.full_sermon[:3000]}
                    
                    [60초 세로 쇼츠 대본 3종: {st.session_state.sermon_title}]
                    
                    1. 감동 및 위로형 대본
                    - [0~5초 후킹 멘트]:
                    - [5~45초 본론 메시지]:
                    - [45~60초 결단 및 축복]:
                    
                    2. 질문 및 호기심 자극형 대본
                    - [0~5초 후킹 멘트]:
                    - [5~45초 본론 메시지]:
                    - [45~60초 결단 및 축복]:
                    
                    3. 강한 결단 선포형 대본
                    - [0~5초 후킹 멘트]:
                    - [5~45초 본론 메시지]:
                    - [45~60초 결단 및 축복]:
                    """
                    res = get_ai_response(prompt, is_json=False)
                    if res:
                        st.session_state.shorts_script_text = res
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
                    st.markdown(f"<div class='content-box'>{sh_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 쇼츠 대본을 생성하세요.")

        elif active_view == "🏡 세대별 가정예배지":
            age_group = st.selectbox("예배 대상 선택", ["👶 영유아용", "🧒 어린이용", "🧑 청소년용", "👨‍👩‍👧 청장년용"], key="sel_age_group")
            fam_txt = st.session_state.get(f"family_worship_{age_group}", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_가정예배지_{age_group}", fam_txt, f"fam_{age_group}")

            if st.button(f"✨ {age_group} 맞춤 가정예배지 생성", type="primary", key="btn_gen_fam"):
                with st.spinner(f"{age_group} 가정예배지 작성 중..."):
                    prompt = f"""
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 요약: {st.session_state.full_sermon[:3000]}
                    대상: {age_group}
                    
                    [가정예배 순서지 ({age_group}): {st.session_state.sermon_title}]
                    
                    1. 찬양 및 신앙고백
                    2. 함께 읽는 성경 말씀
                    3. {age_group} 눈높이에 맞춘 3분 가족 메시지
                    4. 온 가족 나눔 질문 2가지
                    5. 가정을 축복하는 마무리 기도문
                    """
                    res = get_ai_response(prompt, is_json=False)
                    if res:
                        st.session_state[f"family_worship_{age_group}"] = res
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
                    st.markdown(f"<div class='content-box'>{fam_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption(f"위 버튼을 눌러 {age_group} 맞춤 가정예배지를 생성하세요.")

        elif active_view == "🔍 설교 점검 및 제안":
            audit_txt = st.session_state.get("sermon_audit_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교점검및제안", audit_txt, "sermon_audit")

            if st.button("🔍 설교 전달력 및 신학적 완성도 정밀 점검", type="primary", key="btn_gen_audit"):
                with st.spinner("설교 분석 리포트 작성 중..."):
                    prompt = f"""
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 제목: {st.session_state.sermon_title}
                    설교 원고: {st.session_state.full_sermon[:4000]}
                    
                    [설교 전문 피드백 리포트: {st.session_state.sermon_title}]
                    
                    1. 본문 주해의 정확성 및 성경 중심성 평가
                    2. 논리적 대지 전개 및 설교 구조 분석
                    3. 청중 공감 예화 및 삶의 적용 적절성
                    4. 스피치 전달력 및 표현 개선 제안
                    5. 총평 및 3가지 핵심 권고사항
                    """
                    res = get_ai_response(prompt, is_json=False)
                    if res:
                        st.session_state.sermon_audit_text = res
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
                    st.markdown(f"<div class='content-box'>{audit_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 설교 점검 리포트를 생성하세요.")

        elif active_view == "📖 소그룹 리더가이드":
            ldr_txt = st.session_state.get("leader_guide_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_소그룹리더가이드", ldr_txt, "ldr_guide")

            if st.button("📖 구역장/순장용 소그룹 리더가이드 생성", type="primary", key="btn_gen_ldr_guide"):
                with st.spinner("소그룹 인도자 가이드 작성 중..."):
                    prompt = f"""
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 요약: {st.session_state.full_sermon[:3500]}
                    
                    [소그룹 리더(구역장/셀리더/순장) 심화 가이드: {st.session_state.sermon_title}]
                    
                    1. 이번 주 모임의 핵심 목표 및 주제 방향
                    2. 본문 배경 및 신학적 핵심 해설 (리더용 심화 자료)
                    3. 나눔 질문별 성도들의 예상 답변 및 리더 피드백 팁
                    4. 모임 중 침묵 또는 돌발 상황 대처 요령
                    5. 소그룹을 위한 맞춤 중보기도 제목 3가지
                    """
                    res = get_ai_response(prompt, is_json=False)
                    if res:
                        st.session_state.leader_guide_text = res
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
                    st.markdown(f"<div class='content-box'>{ldr_txt}</div>", unsafe_allow_html=True)
            else:
                st.caption("위 버튼을 눌러 소그룹 리더가이드를 생성하세요.")

# ==============================================================================
# 2. 📤 새 설교 등록/원고작성
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
                st.session_state.dash_active_view = "설교 요약"
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
            st.session_state.dash_active_view = "설교 요약"
            st.success("파일 등록이 완료되었습니다! [📊 설교 대시보드]로 이동하세요.")

    with tab_ai:
        st.markdown("#### 📖 성경 본문 선택 기반 정통 강해설교문 자동 생성")
        ai_c1, ai_c2, ai_c3 = st.columns([1.2, 1.5, 1.3])
        with ai_c1: sel_book = st.selectbox("성경 66권 선택", BIBLE_BOOKS, index=44, key="sel_ai_book")
        with ai_c2: sel_chap_verse = st.text_input("장 및 절 입력", value="8장 28절~39절", key="sel_ai_cv")
        with ai_c3:
            theology_choice = st.selectbox(
                "신학적 관점 선택",
                [
                    "개혁주의 (Reformed - 칼빈주의/하나님 주권)",
                    "장로교 정통 (Presbyterian - 웨스트민스터/구속사)",
                    "복음주의 (Evangelical - 십자가/은혜/복음선포)"
                ],
                key="sel_ai_theology"
            )

        ai_t1, ai_t2 = st.columns([2, 1])
        with ai_t1: ai_sermon_topic = st.text_input("설교 주제 / 강조 포인트 (선택)", value="고난 속에서도 흔들리지 않는 하나님의 영원한 사랑과 구원의 확신", key="sel_ai_topic")
        with ai_t2: sermon_style = st.selectbox("설교 형태", ["3대지 본문중심 강해설교", "구속사적 복음설교", "원어 주해 중심 강해설교"], key="sel_ai_style")

        full_scripture_str = f"{sel_book} {sel_chap_verse}"

        if st.button("🚀 정통 강해설교문 전문 자동 작성 시작 (25~30분 분량)", type="primary", key="btn_gen_ai_sermon"):
            with st.spinner(f"[{theology_choice.split(' ')[0]}] 관점으로 강해설교문 작성 중..."):
                prompt = f"""
                성경 본문: {full_scripture_str}
                설교 주제: {ai_sermon_topic}
                신학적 관점: {theology_choice}
                설교 형태: {sermon_style}
                
                [강해설교문 전문: {sel_book} 강해 - {ai_sermon_topic}]
                
                서론(도입부) - 본론(제1대지, 제2대지, 제3대지) - 결론 및 결단의 기도 구조로 강단 선포용 25~30분 완성 원고를 작성하세요.
                각 대지마다 본문 주해와 성도들의 삶에 와닿는 구체적 예화, 실천 적용 방안을 풍성히 포함하세요.
                """
                generated_sermon = get_ai_response(prompt, is_json=False)
                if generated_sermon:
                    st.session_state.temp_generated_sermon = generated_sermon
                    st.session_state.temp_ai_title = f"{sel_book} 강해: {ai_sermon_topic}"
                    st.session_state.temp_ai_scrip = full_scripture_str
                    st.success("강해설교문 전문이 완성되었습니다!")

        if "temp_generated_sermon" in st.session_state and st.session_state.temp_generated_sermon:
            st.write("---")
            render_section_top_toolbar(st.session_state.temp_ai_title, st.session_state.temp_generated_sermon, "ai_gen_sermon")
            st.session_state.temp_generated_sermon = st.text_area("작성된 강해설교문 검토 및 수정", value=st.session_state.temp_generated_sermon, height=450, key="edit_ai_sermon_area")
            
            if st.button("✅ 이 설교문을 내 서재와 대시보드에 최종 등록하기", type="primary", key="btn_save_ai_sermon"):
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
                st.session_state.dash_active_view = "설교 요약"
                st.success("설교문이 등록되었습니다! [📊 설교 대시보드] 메뉴에서 확인하세요.")

# ==============================================================================
# 3. 🎙️ AI 보이스오버 스튜디오
# ==============================================================================
elif app_mode == "🎙️ AI 보이스오버 스튜디오":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>🎙️ AI 보이스오버 스튜디오</h1>", unsafe_allow_html=True)
    st.caption("설교문이나 요약문을 신경망 음성으로 합성하여 실시간 청취 및 고음질 MP3로 다운로드합니다.")

    vo_col1, vo_col2 = st.columns([1.5, 1])
    with vo_col1:
        vo_text = st.text_area("음성 변환 텍스트", value=st.session_state.full_sermon, height=350, key="vo_text_area")
        vo_voice = st.selectbox("성우 보이스 선택", ["인준 (남성 - 차분하고 신뢰감 있는 톤)", "선희 (여성 - 맑고 또렷한 톤)"], key="vo_voice_sel")
        voice_id = "ko-KR-InJoonNeural" if "인준" in vo_voice else "ko-KR-SunHiNeural"
        
        if st.button("🎙️ 고음질 음성(TTS) 즉시 생성하기", type="primary", key="btn_gen_tts"):
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
                st.download_button("📥 MP3 오디오 파일 다운로드", data=af, file_name=f"{st.session_state.sermon_title}_voice.mp3", mime="audio/mp3", key="dl_vo_mp3")
        else:
            st.info("왼쪽에서 버튼을 누르면 이곳에 재생 플레이어가 나타납니다.")

# ==============================================================================
# 4. 🎬 쇼츠 만들기 (스튜디오)
# ==============================================================================
elif app_mode == "🎬 쇼츠 만들기 (스튜디오)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>▶️ 쇼츠 만들기 스튜디오</h1>", unsafe_allow_html=True)
    
    tab_yt_extract, tab_ai_compose = st.tabs([
        "🔗 내 영상(유튜브 링크)으로 숏츠 추출",
        "🎨 AI 나레이션 & 템플릿 숏츠 제작"
    ])

    with tab_yt_extract:
        st.markdown("#### 📺 유튜브 예배/설교 영상에서 세로 숏츠 자동 추출")
        st.caption("유튜브 영상 링크를 입력하고 원하는 시작 시간과 길이를 설정하면 9:16 고화질 숏츠를 즉시 추출합니다.")
        
        yt_url_input = st.text_input("유튜브 영상 링크 (URL)", placeholder="https://www.youtube.com/watch?v=... 또는 https://youtu.be/...", key="yt_url_input")
        
        yt_c1, yt_c2, yt_c3 = st.columns(3)
        with yt_c1:
            yt_start_min = st.number_input("시작 분 (Minute)", min_value=0, max_value=300, value=12, step=1, key="yt_s_min")
        with yt_c2:
            yt_start_sec = st.number_input("시작 초 (Second)", min_value=0, max_value=59, value=30, step=1, key="yt_s_sec")
        with yt_c3:
            yt_duration = st.slider("숏츠 길이 (초)", min_value=15, max_value=60, value=45, key="yt_dur_slider")
            
        yt_shorts_title = st.text_input("숏츠 상단 헤드라인 제목", value=f"{st.session_state.sermon_title}", key="yt_shorts_title")
        yt_shorts_sub = st.text_input("강조 자막 텍스트 (선택)", value="내 열심보다 중요한 하나님의 은혜와 인도하심", key="yt_shorts_sub")
        yt_church = st.text_input("하단 교회명 워터마크 배지", value="화광교회", key="yt_church")

        total_start_seconds = (yt_start_min * 60) + yt_start_sec

        if st.button("🚀 유튜브 영상에서 9:16 세로 숏츠 즉시 추출하기", type="primary", key="btn_yt_extract"):
            if not yt_url_input.strip():
                st.warning("유튜브 영상 링크를 입력해주세요.")
            else:
                with st.spinner("유튜브 영상 다운로드, 9:16 화면 크롭 및 자막 합성 중... (약 20~40초)"):
                    try:
                        extracted_file = extract_youtube_to_shorts(
                            yt_url=yt_url_input.strip(),
                            start_sec=total_start_seconds,
                            duration_sec=yt_duration,
                            title=yt_shorts_title,
                            subtitle_text=yt_shorts_sub,
                            church_name=yt_church
                        )
                        st.session_state.yt_extracted_result = extracted_file
                        st.success("유튜브 영상 숏츠 추출이 완료되었습니다!")
                    except Exception as e:
                        st.error(f"유튜브 추출 오류: {str(e)}")

        if "yt_extracted_result" in st.session_state and os.path.exists(st.session_state.yt_extracted_result):
            st.write("---")
            st.markdown("### 🎬 추출 완료된 9:16 쇼츠 미리보기")
            res_c1, res_c2 = st.columns([1, 1])
            with res_c1:
                st.video(st.session_state.yt_extracted_result)
            with res_c2:
                with open(st.session_state.yt_extracted_result, "rb") as yf:
                    st.download_button(
                        "📥 추출된 MP4 세로 숏츠 다운로드",
                        data=yf,
                        file_name=f"{yt_shorts_title}_shorts.mp4",
                        mime="video/mp4",
                        key="dl_yt_shorts_btn"
                    )

    with tab_ai_compose:
        st.markdown("#### 🌐 무료 미디어/BGM 소스 바로가기")
        src_c1, src_c2, src_c3 = st.columns(3)
        with src_c1: st.link_button("🎥 픽사베이 (Pixabay 무료 영상)", "https://pixabay.com/ko/videos/")
        with src_c2: st.link_button("📸 펙셀스 (Pexels 무료 비디오)", "https://www.pexels.com/ko-kr/videos/")
        with src_c3: st.link_button("🎵 픽사베이 무료 음악(BGM)", "https://pixabay.com/ko/music/")

        st.write("---")
        st.markdown("#### 💡 확 끌리는 5가지 쇼츠 제목 & #해시태그 추천")
        if st.button("✨ 조회수 폭발 5가지 쇼츠 제목 & 해시태그 뽑기", key="btn_gen_shorts_meta"):
            with st.spinner("제목 및 태그 분석 중..."):
                prompt = f"설교제목: {st.session_state.sermon_title}, 요약: {st.session_state.full_sermon[:1500]}\n100% 한국어로 쇼츠 클릭률을 높이는 제목 5개와 해시태그 8개를 JSON으로 작성: {{\"titles\": [\"1.제목\", \"2.제목\", \"3.제목\", \"4.제목\", \"5.제목\"], \"hashtags\": [\"#쇼츠\", \"#은혜\"]}}"
                st.session_state.shorts_rec = get_ai_response(prompt, is_json=True)

        selected_title = st.session_state.sermon_title
        if "shorts_rec" in st.session_state and st.session_state.shorts_rec:
            rec = st.session_state.shorts_rec
            t_choice = st.radio("추천 제목 선택", rec.get("titles", []), horizontal=False, key="rad_shorts_title")
            if t_choice: selected_title = re.sub(r"^\d+\.\s*", "", t_choice)
            st.markdown("**추천 태그:** " + " ".join(rec.get("hashtags", [])))

        st.write("---")
        sh_col1, sh_col2 = st.columns([1.2, 1])
        with sh_col1:
            v_title = st.text_input("쇼츠 제목", value=selected_title, key="in_shorts_title")
            v_script = st.text_area("자막 대본 (줄바꿈 구분)", value="내 열심보다 중요한 것은 하나님의 이끄심입니다.\n우리가 멈출 때 비로소 하나님의 역사가 시작됩니다.\n오늘 그분의 인도하심 앞에 온전히 맡기십시오.", height=120, key="in_shorts_script")
            
            c_v1, c_v2 = st.columns(2)
            with c_v1: v_ratio = st.radio("비율", ["9:16 (세로 쇼츠)", "16:9 (가로 영상)"], key="rad_shorts_ratio")
            with c_v2: v_voice = st.selectbox("보이스", ["인준 (남성)", "선희 (여성)"], key="sel_shorts_voice")

            with st.expander("🎨 폰트 크기 및 자막 위치 정밀 편집", expanded=True):
                font_c1, font_c2 = st.columns(2)
                with font_c1:
                    t_fsize = st.slider("제목 폰트 크기 (pt)", min_value=32, max_value=72, value=48, step=2, key="sh_t_fsize")
                    t_ypos = st.slider("제목 Y 위치 (높이)", min_value=80, max_value=500, value=180, step=10, key="sh_t_ypos")
                with font_c2:
                    s_fsize = st.slider("자막/본문 폰트 크기 (pt)", min_value=24, max_value=60, value=42, step=2, key="sh_s_fsize")
                    s_ypos = st.slider("자막 Y 위치 (높이)", min_value=800, max_value=1700, value=1400, step=20, key="sh_s_ypos")

            bg_media = st.file_uploader("배경 동영상/사진 업로드 (최대 300MB)", type=["mp4", "mov", "jpg", "png"], key="up_shorts_bg")
            bgm_media = st.file_uploader("배경음악 MP3 업로드 (최대 300MB)", type=["mp3", "wav"], key="up_shorts_bgm")

            if st.button("🚀 비디오 렌더링 시작", type="primary", key="btn_render_video"):
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

                    rendered_out = create_animated_video(
                        title=v_title,
                        script_paragraphs=lines,
                        bg_media_path=bg_p,
                        bgm_path=bgm_p,
                        aspect_ratio=ratio_val,
                        voice=voice_id,
                        title_fontsize=t_fsize,
                        sub_fontsize=s_fsize,
                        title_y=t_ypos,
                        sub_y=s_ypos
                    )
                    st.session_state.rendered_shorts_out = rendered_out
                    st.success("영상 렌더링이 완료되었습니다!")

    with sh_col2:
        st.markdown("### 🎬 완성된 영상")
        if "rendered_shorts_out" in st.session_state and os.path.exists(st.session_state.rendered_shorts_out):
            st.video(st.session_state.rendered_shorts_out)
            with open(st.session_state.rendered_shorts_out, "rb") as vf:
                st.download_button("📥 MP4 비디오 파일 다운로드", data=vf, file_name="sermon_shorts.mp4", mime="video/mp4", key="dl_shorts_mp4")

# ==============================================================================
# 5. 📚 설교 서재 (Sermon Library)
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>설교 서재</h1>", unsafe_allow_html=True)
    st.caption(f"설교문 {len(st.session_state.sermon_library)}편 보관 중")

    search_kw = st.text_input("🔍 제목 또는 성경 구절로 검색", placeholder="예: 시편, 로마서, 신앙...", key="in_lib_search")
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
                st.session_state.dash_active_view = "설교 요약"
                st.success(f"'{s_item['title']}' 설교를 불러왔습니다! [📊 설교 대시보드]로 이동하세요.")
