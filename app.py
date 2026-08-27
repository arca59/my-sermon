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

# --- 성경 66권 구약/신약 자동 분류 모듈 ---
OLD_TESTAMENT_BOOKS = [
    "창세기", "출애굽기", "레위기", "민수기", "신명기", "여호수아", "사사기", "룻기", 
    "사무엘상", "사무엘하", "열왕기상", "열왕기하", "역대상", "역대하", "에스라", "느헤미야", 
    "에스더", "욥기", "시편", "잠언", "전도서", "아가", "이사야", "예레미야", "예레미야애가", 
    "에스겔", "다니엘", "호세아", "요엘", "아모스", "오바댜", "요나", "미가", "나훔", 
    "하박국", "스바냐", "학개", "스가랴", "말라기"
]

NEW_TESTAMENT_BOOKS = [
    "마태복음", "마가복음", "누가복음", "요한복음", "사도행전", "로마서", "고린도전서", "고린도후서", 
    "갈라디아서", "에베소서", "빌립보서", "골로새서", "데살로니가전서", "데살로니가후서", "디모데전서", 
    "디모데후서", "디도서", "빌레몬서", "히브리서", "야고보서", "베드로전서", "베드로후서", 
    "요한일서", "요한이서", "요한삼서", "유다서", "요한계시록"
]

BIBLE_BOOKS = OLD_TESTAMENT_BOOKS + NEW_TESTAMENT_BOOKS

def classify_scripture(scripture_text: str):
    if not scripture_text:
        return "기타", "성경전체"
    
    for b in OLD_TESTAMENT_BOOKS:
        if b in scripture_text:
            return "구약", b
            
    for b in NEW_TESTAMENT_BOOKS:
        if b in scripture_text:
            return "신약", b
            
    return "기타", "성경전체"

# --- 설교 영구 보관소 파일 데이터베이스 엔진 ---
SERMON_DB_PATH = "./outputs/sermons_db.json"

def load_sermons_from_db():
    os.makedirs("./outputs", exist_ok=True)
    if os.path.exists(SERMON_DB_PATH):
        try:
            with open(SERMON_DB_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass

    default_data = [
        {
            "id": 1,
            "title": "눈동자처럼 은혜 가운데",
            "scripture": "시편 17:8",
            "testament": "구약",
            "book": "시편",
            "topic": "보호와 은혜",
            "theology": "개혁주의/장로교",
            "date": "2026-08-27",
            "tags": ["보호", "은혜", "기도", "시편"],
            "summary": """🎯 설교 핵심 명제:
나를 눈동자 같이 지키시고 주의 날개 그늘 아래에 감추시는 하나님의 신실한 은혜를 온전히 신뢰하며 나아갑니다.

📌 설교 3대지 핵심 요약:
1. 눈동자처럼 우리를 보호하시는 하나님 (시편 17:8)
- 하나님께서는 우리의 모든 순간을 눈동자처럼 아끼며 지켜주십니다.

2. 주의 날개 그늘 아래 참된 안식과 평안
- 거친 풍랑 속에서도 주의 날개 아래 피할 때 참된 평안과 안식을 누립니다.

3. 고난 속에서도 흔들리지 않는 기도와 선포
- 기도로 나아가면 주님께서 우리의 부르짖음에 응답하시며 담대함을 주십니다.

💡 성도를 위한 구체적 실천 적용:
- 매일의 삶 속에서 나를 지키시는 하나님의 은혜를 묵상하고 감사하기
- 두려움과 두 갈래 길에 설 때 세상이 아닌 주의 날개 아래로 먼저 피하기
- 만나는 사람들에게 주의 신실하신 사랑과 간증을 기쁨으로 나누기""",
            "text": """1. 눈동자처럼 지키시는 사랑입니다.
하나님께서는 우리의 모든 순간을 시선을 떼지 않으시고 눈동자처럼 아끼며 지켜주십니다. 세상의 고난과 위협 속에서도 우리를 결코 혼자 두지 않으시는 은혜를 믿으십시오.

2. 날개 그늘 아래의 참된 피난처입니다.
거친 풍랑 속에서도 주의 날개 아래 피할 때 참된 평안과 안식을 누립니다. 세상의 안전지대가 아닌 오직 하나님의 품만이 우리의 영원한 피난처가 됩니다.

3. 고난 속에서도 흔들리지 않는 기도입니다.
우리가 절망 가운데 있을 때 기도로 나아가면 주님께서 우리의 부르짖음에 응답하십니다. 어떠한 상황에서도 소망을 주님께 두고 온전히 선포하십시오."""
        }
    ]
    save_sermons_to_db(default_data)
    return default_data

def save_sermons_to_db(sermons_list):
    os.makedirs("./outputs", exist_ok=True)
    try:
        with open(SERMON_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(sermons_list, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"서재 파일 저장 오류: {str(e)}")

# --- 모든 메뉴 완벽 동기화 및 캐시 초기화 핵심 함수 ---
def load_sermon_to_workspace(sermon_item, idx=0):
    st.session_state.current_sermon_idx = idx
    st.session_state.sermon_title = sermon_item.get("title", "")
    st.session_state.sermon_scripture = sermon_item.get("scripture", "")
    st.session_state.full_sermon = sermon_item.get("text", "")
    st.session_state.sermon_summary_text = sermon_item.get("summary", "")
    
    keys_to_clear = [
        "small_group_text", "qt5_text", "card_list", "shorts_script_text",
        "sermon_audit_text", "leader_guide_text", "rich_materials", 
        "praise_list", "shorts_rec", "yt_extracted_result", 
        "rendered_shorts_out", "vo_audio_path", "verse_card_img"
    ]
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
            
    for k in list(st.session_state.keys()):
        if k.startswith("family_worship_"):
            del st.session_state[k]

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

# --- PIL 텍스트 줄바꿈 헬퍼 함수 ---
def wrap_korean_text(text: str, font, max_width: int, draw: PIL.ImageDraw.ImageDraw) -> str:
    if not text:
        return ""
    
    wrapped_lines = []
    paragraphs = str(text).split('\n')
    for p in paragraphs:
        words = p.split(' ')
        curr_line = ""
        for w in words:
            test_line = f"{curr_line} {w}".strip()
            bbox = draw.textbbox((0, 0), test_line, font=font)
            line_w = bbox[2] - bbox[0]
            if line_w > max_width and curr_line:
                wrapped_lines.append(curr_line)
                curr_line = w
            else:
                curr_line = test_line
        if curr_line:
            wrapped_lines.append(curr_line)
            
    return "\n".join(wrapped_lines)

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
            try: font_b = PIL.ImageFont.truetype(f_p, 44); break
            except Exception: pass
    if not font_b: font_b = PIL.ImageFont.load_default()

    font_t = None
    for f_p in ["/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf", "C:/Windows/Fonts/malgun.ttf"]:
        if os.path.exists(f_p):
            try: font_t = PIL.ImageFont.truetype(f_p, 30); break
            except Exception: pass
    if not font_t: font_t = PIL.ImageFont.load_default()

    draw.text((100, 90), f"CARD {card_item.get('card_number', idx+1)}", fill=(99, 102, 241, 255), font=font_t)
    
    headline_raw = card_item.get("headline", "")
    headline_wrapped = wrap_korean_text(headline_raw, font_b, 880, draw)
    draw.multiline_text((100, 170), headline_wrapped, fill=(253, 224, 71, 255), font=font_b, spacing=12)

    body_raw = card_item.get("body_text", "")
    body_wrapped = wrap_korean_text(body_raw, font_t, 880, draw)
    draw.multiline_text((100, 380), body_wrapped, fill=(241, 245, 249, 255), font=font_t, spacing=16)

    if scripture_str:
        draw.text((100, 910), f"「 {scripture_str} 」", fill=(253, 224, 71, 255), font=font_t)
    if church_name:
        draw.text((100, 965), church_name, fill=(147, 197, 253, 255), font=font_t)

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

# --- 5. 유튜브 비디오 다운로드 및 9:16 쇼츠 추출 엔진 ---
def extract_youtube_to_shorts(yt_url: str, start_sec: int, duration_sec: int, title: str, subtitle_text: str, church_name: str = ""):
    out_dir = "./outputs"
    os.makedirs(out_dir, exist_ok=True)
    src_video = os.path.join(out_dir, "yt_raw_source.mp4")
    
    clean_input = yt_url.strip()
    vid_match = re.search(r'(?:v=|\/live\/|\/shorts\/|youtu\.be\/)([a-zA-Z0-9_-]{11})', clean_input)
    video_id = vid_match.group(1) if vid_match else None
    clean_url = f"https://www.youtube.com/watch?v={video_id}" if video_id else clean_input

    download_success = False
    last_err = ""

    cobalt_endpoints = [
        "https://api.cobalt.tools/api/json",
        "https://co.wuk.sh/api/json",
        "https://cobalt-api.koyeb.app/api/json"
    ]
    for ep in cobalt_endpoints:
        try:
            payload = json.dumps({"url": clean_url, "videoQuality": "720"}).encode('utf-8')
            req = urllib.request.Request(
                ep,
                data=payload,
                headers={
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                }
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                resp_bytes = resp.read()
                try:
                    res_json = json.loads(resp_bytes.decode('utf-8'))
                    dl_link = res_json.get("url")
                    if dl_link:
                        urllib.request.urlretrieve(dl_link, src_video)
                        if os.path.exists(src_video) and os.path.getsize(src_video) > 100000:
                            download_success = True
                            break
                except Exception:
                    continue
        except Exception as ex:
            last_err = str(ex)
            continue

    if not download_success and video_id:
        piped_apis = [
            f"https://pipedapi.kavin.rocks/streams/{video_id}",
            f"https://pipedapi.adminforge.de/streams/{video_id}",
            f"https://piped-api.lunar.icu/streams/{video_id}"
        ]
        for p_api in piped_apis:
            try:
                req = urllib.request.Request(p_api, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=8) as resp:
                    resp_bytes = resp.read()
                    try:
                        p_data = json.loads(resp_bytes.decode('utf-8'))
                        v_streams = p_data.get("videoStreams", [])
                        mp4_url = None
                        for vs in v_streams:
                            if vs.get("mimeType", "").startswith("video/mp4") and not vs.get("videoOnly", True):
                                mp4_url = vs.get("url")
                                break
                        if not mp4_url:
                            for vs in v_streams:
                                if vs.get("mimeType", "").startswith("video/mp4"):
                                    mp4_url = vs.get("url")
                                    break
                        if mp4_url:
                            urllib.request.urlretrieve(mp4_url, src_video)
                            if os.path.exists(src_video) and os.path.getsize(src_video) > 100000:
                                download_success = True
                                break
                    except Exception:
                        continue
            except Exception as ex:
                last_err = str(ex)
                continue

    if not download_success:
        client_strategies = [['ios'], ['android_creator'], ['mweb'], ['tv_embedded']]
        for clients in client_strategies:
            if os.path.exists(src_video):
                try: os.remove(src_video)
                except Exception: pass

            ydl_opts = {
                'format': 'best[ext=mp4]/bestvideo[ext=mp4]+bestaudio[ext=m4a]/best',
                'outtmpl': src_video,
                'overwrites': True,
                'quiet': True,
                'no_warnings': True,
                'nocheckcertificate': True,
                'geo_bypass': True,
                'http_headers': {
                    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_4_1 like Mac OS X) AppleWebKit/605.1.15',
                },
                'extractor_args': {
                    'youtube': {
                        'player_client': clients
                    }
                }
            }

            try:
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    ydl.download([clean_url])
                if os.path.exists(src_video) and os.path.getsize(src_video) > 100000:
                    download_success = True
                    break
            except Exception as e:
                last_err = str(e)
                continue

    if not download_success or not os.path.exists(src_video):
        raise Exception(f"클라우드 차단으로 자동 우회 다운로드가 어려울 경우, [🎨 AI 나레이션 & 템플릿 숏츠 제작] 탭에서 설교 동영상을 직접 업로드하여 9:16 숏츠를 바로 렌더링하실 수 있습니다.")

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

# --- 100% 동적 설교 파싱 엔진 ---
def parse_sermon_content(title, scripture, summary_content, full_sermon=""):
    text = summary_content if summary_content and len(summary_content) > 50 else full_sermon
    
    prop = ""
    prop_match = re.search(r'(?:🎯|핵심\s*명제|명제|개요)[:\s]*(.*?)(?=\n📌|\n💡|\n🙏|\n\d+\.|\n[1-9]대지|$)', text, re.DOTALL)
    if prop_match:
        prop = prop_match.group(1).strip()
    else:
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        prop = "\n".join(lines[:2]) if lines else title

    points = []
    point_matches = re.findall(r'(?:^|\n)\s*(?:[1-4]\.|\d+[\.\)]|제\s*[1-4]\s*대지|•|-)\s*(.*?)(?=\n\s*(?:[1-4]\.|\d+[\.\)]|제\s*[1-4]\s*대지|•|-|💡|🙏|🎯|📌)|$)', text, re.DOTALL)
    
    for pm in point_matches:
        pm_clean = pm.strip()
        if pm_clean and len(pm_clean) > 3:
            points.append(pm_clean)
            
    if not points:
        paragraphs = [p.strip() for p in text.split('\n') if p.strip()]
        chunk_size = max(1, len(paragraphs) // 4)
        for i in range(4):
            sub_p = paragraphs[i*chunk_size:(i+1)*chunk_size]
            if sub_p:
                points.append("\n".join(sub_p))

    while len(points) < 4:
        points.append(f"{title} - 하나님의 은혜와 인도하심에 감사하며 말씀 중심의 삶을 살아가십시오.")

    app_text = ""
    app_match = re.search(r'(?:💡|실천\s*적용|삶의\s*적용|적용)[:\s]*(.*?)(?=\n🙏|\n기도|$)', text, re.DOTALL)
    if app_match:
        app_text = app_match.group(1).strip()
    else:
        app_text = f"1. 매일 일상 속에서 {scripture} 말씀을 생각하며 마음에 새깁니다.\n2. 세상의 자랑이 아닌 오직 하나님의 신실하신 약속을 선포합니다.\n3. 주신 말씀에 순종하여 담대함과 기도로 승리하는 삶을 살아갑니다."

    prayer_text = ""
    prayer_match = re.search(r'(?:🙏|기도문|결단\s*및\s*축복|마침\s*기도|기도)[:\s]*(.*?)$', text, re.DOTALL)
    if prayer_match:
        prayer_text = prayer_match.group(1).strip()
    else:
        prayer_text = f"살아계신 하나님, 오늘 선포된 [{title}] 말씀을 통해 주님의 거룩하신 뜻을 깨닫게 하시니 감사합니다. 주신 은혜의 구절({scripture})을 심비에 새기고, 날마다 믿음으로 승리하는 복된 성도가 되게 하옵소서. 예수님의 이름으로 기도드립니다. 아멘."

    return {
        "prop": prop,
        "points": points,
        "app": app_text,
        "prayer": prayer_text
    }

# --- 동적 10-슬라이드 구조화 PPTX 생성기 ---
def generate_sermon_structure_pptx(title: str, scripture: str, summary_content: str, full_sermon: str = "") -> io.BytesIO:
    try:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        def set_dark_slide(slide, img_url=None):
            if img_url:
                img_b = fetch_image_bytes(img_url)
                if img_b:
                    slide.shapes.add_picture(io.BytesIO(img_b), 0, 0, width=Inches(13.333), height=Inches(7.5))
            else:
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(15, 23, 42)

            overlay = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
            overlay.fill.solid()
            overlay.fill.fore_color.rgb = RGBColor(10, 15, 30)
            overlay.line.fill.background()

        def set_light_slide(slide):
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(248, 250, 252)

        parsed = parse_sermon_content(title, scripture, summary_content, full_sermon)
        prop_text = parsed["prop"]
        points = parsed["points"]
        app_text = parsed["app"]
        prayer_text = parsed["prayer"]

        # [Slide 1: 표지]
        s1 = prs.slides.add_slide(blank_layout)
        set_dark_slide(s1, CARD_BACKGROUNDS[0])
        tb1 = s1.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.8))
        p1 = tb1.text_frame.paragraphs[0]
        p1.text = f"주 일 설 교\n\n{title}\n\n본문 · {scripture}"
        p1.font.size, p1.font.bold = Pt(38), True
        p1.font.color.rgb, p1.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER

        # [Slide 2: 들어가며]
        s2 = prs.slides.add_slide(blank_layout)
        set_light_slide(s2)
        tb2 = s2.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(5.8))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        p2_head = tf2.paragraphs[0]
        p2_head.text = f"들어가며 · 핵심 메시지 ({scripture})"
        p2_head.font.size, p2_head.font.bold = Pt(32), True
        p2_head.font.color.rgb = RGBColor(30, 58, 138)
        
        p2_body = tf2.add_paragraph()
        p2_body.text = f"\n{prop_text}"
        p2_body.font.size = Pt(20)
        p2_body.font.color.rgb = RGBColor(30, 41, 59)

        # [Slide 3: 설교의 전체 흐름]
        s3 = prs.slides.add_slide(blank_layout)
        set_dark_slide(s3)
        tb3_h = s3.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(1.0))
        hp3 = tb3_h.text_frame.paragraphs[0]
        hp3.text = f"설교의 흐름 (Sermon Outline)"
        hp3.font.size, hp3.font.bold = Pt(32), True
        hp3.font.color.rgb = RGBColor(253, 224, 71)

        for card_i in range(min(4, len(points))):
            top_pos = Inches(2.2 + (card_i * 1.1))
            shape = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), top_pos, Inches(11.33), Inches(0.9))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(30, 41, 59)
            shape.line.color.rgb = RGBColor(51, 65, 85)
            tf_card = shape.text_frame
            p_c = tf_card.paragraphs[0]
            first_line = points[card_i].split('\n')[0][:50]
            p_c.text = f"  0{card_i+1}   {first_line}"
            p_c.font.size, p_c.font.bold = Pt(20), True
            p_c.font.color.rgb = RGBColor(241, 245, 249)

        # [Slide 4: 본문 핵심 성구]
        s4 = prs.slides.add_slide(blank_layout)
        set_dark_slide(s4, CARD_BACKGROUNDS[1])
        tb4 = s4.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.93), Inches(5.0))
        tf4 = tb4.text_frame
        tf4.word_wrap = True
        p4_h = tf4.paragraphs[0]
        p4_h.text = f"본문 말씀  ·  {scripture}\n"
        p4_h.font.size, p4_h.font.bold = Pt(28), True
        p4_h.font.color.rgb = RGBColor(147, 197, 253)
        
        p4_b = tf4.add_paragraph()
        p4_b.text = f"[{title}]\n\n{scripture} 본문 말씀 중심 선포"
        p4_b.font.size = Pt(20)
        p4_b.font.color.rgb = RGBColor(241, 245, 249)

        # [Slide 5: 제1대지]
        s5 = prs.slides.add_slide(blank_layout)
        set_light_slide(s5)
        tb5 = s5.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(5.8))
        tf5 = tb5.text_frame
        tf5.word_wrap = True
        p5_h = tf5.paragraphs[0]
        p5_h.text = f"01. 첫 번째 메시지"
        p5_h.font.size, p5_h.font.bold = Pt(30), True
        p5_h.font.color.rgb = RGBColor(30, 58, 138)
        p5_b = tf5.add_paragraph()
        p5_b.text = f"\n{points[0] if len(points) > 0 else title}"
        p5_b.font.size = Pt(20)
        p5_b.font.color.rgb = RGBColor(30, 41, 59)

        # [Slide 6: 제2대지]
        s6 = prs.slides.add_slide(blank_layout)
        set_dark_slide(s6)
        tb6 = s6.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(5.8))
        tf6 = tb6.text_frame
        tf6.word_wrap = True
        p6_h = tf6.paragraphs[0]
        p6_h.text = f"02. 두 번째 메시지"
        p6_h.font.size, p6_h.font.bold = Pt(30), True
        p6_h.font.color.rgb = RGBColor(253, 224, 71)
        p6_b = tf6.add_paragraph()
        p6_b.text = f"\n{points[1] if len(points) > 1 else title}"
        p6_b.font.size = Pt(20)
        p6_b.font.color.rgb = RGBColor(241, 245, 249)

        # [Slide 7: 제3대지]
        s7 = prs.slides.add_slide(blank_layout)
        set_light_slide(s7)
        tb7 = s7.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(5.8))
        tf7 = tb7.text_frame
        tf7.word_wrap = True
        p7_h = tf7.paragraphs[0]
        p7_h.text = f"03. 세 번째 메시지"
        p7_h.font.size, p7_h.font.bold = Pt(30), True
        p7_h.font.color.rgb = RGBColor(30, 58, 138)
        p7_b = tf7.add_paragraph()
        p7_b.text = f"\n{points[2] if len(points) > 2 else title}"
        p7_b.font.size = Pt(20)
        p7_b.font.color.rgb = RGBColor(30, 41, 59)

        # [Slide 8: 핵심 메시지]
        s8 = prs.slides.add_slide(blank_layout)
        set_dark_slide(s8, CARD_BACKGROUNDS[2])
        tb8 = s8.shapes.add_textbox(Inches(1.0), Inches(1.0), Inches(11.33), Inches(5.5))
        tf8 = tb8.text_frame
        tf8.word_wrap = True
        p8_h = tf8.paragraphs[0]
        p8_h.text = f"04. 깊은 묵상과 선포"
        p8_h.font.size, p8_h.font.bold = Pt(30), True
        p8_h.font.color.rgb = RGBColor(253, 224, 71)
        p8_b = tf8.add_paragraph()
        p8_b.text = f"\n{points[3] if len(points) > 3 else title}"
        p8_b.font.size = Pt(20)
        p8_b.font.color.rgb = RGBColor(241, 245, 249)

        # [Slide 9: 삶의 적용]
        s9 = prs.slides.add_slide(blank_layout)
        set_light_slide(s9)
        tb9 = s9.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(5.8))
        tf9 = tb9.text_frame
        tf9.word_wrap = True
        p9_h = tf9.paragraphs[0]
        p9_h.text = f"삶의 적용 · 이렇게 살아갑시다"
        p9_h.font.size, p9_h.font.bold = Pt(30), True
        p9_h.font.color.rgb = RGBColor(30, 58, 138)
        p9_b = tf9.add_paragraph()
        p9_b.text = f"\n{app_text}"
        p9_b.font.size = Pt(19)
        p9_b.font.color.rgb = RGBColor(30, 41, 59)

        # [Slide 10: 결단 및 기도]
        s10 = prs.slides.add_slide(blank_layout)
        set_dark_slide(s10, CARD_BACKGROUNDS[0])
        tb10 = s10.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.93), Inches(5.2))
        tf10 = tb10.text_frame
        tf10.word_wrap = True
        p10_h = tf10.paragraphs[0]
        p10_h.text = f"결단과 마침 기도문\n"
        p10_h.font.size, p10_h.font.bold = Pt(28), True
        p10_h.font.color.rgb = RGBColor(253, 224, 71)
        p10_b = tf10.add_paragraph()
        p10_b.text = f"\n{prayer_text}"
        p10_b.font.size = Pt(18)
        p10_b.font.color.rgb = RGBColor(241, 245, 249)

        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        return bio
    except Exception:
        return create_document_pptx(title, summary_content)

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

# --- 말씀카드 고화질 PNG 합성 엔진 ---
def generate_verse_card_png(text_str, scripture_str, bg_option="사진", custom_bg_file=None, font_size=42, line_spacing=18, font_color="#FDE047", stroke_color="#000000", overlay_opacity=0.6, church_name=""):
    canvas_w, canvas_h = 1080, 1080
    
    if bg_option == "직접 업로드" and custom_bg_file:
        try:
            base_img = PIL.Image.open(custom_bg_file).convert("RGBA").resize((canvas_w, canvas_h))
        except Exception:
            base_img = PIL.Image.new("RGBA", (canvas_w, canvas_h), (15, 23, 42, 255))
    elif bg_option == "기본":
        base_img = PIL.Image.new("RGBA", (canvas_w, canvas_h), (15, 23, 42, 255))
    else:
        bg_url = CARD_BACKGROUNDS[0]
        img_b = fetch_image_bytes(bg_url)
        if img_b:
            base_img = PIL.Image.open(io.BytesIO(img_b)).convert("RGBA").resize((canvas_w, canvas_h))
        else:
            base_img = PIL.Image.new("RGBA", (canvas_w, canvas_h), (15, 23, 42, 255))

    alpha_val = int(255 * overlay_opacity)
    overlay = PIL.Image.new("RGBA", (canvas_w, canvas_h), (10, 15, 30, alpha_val))
    combined = PIL.Image.alpha_composite(base_img, overlay)
    draw = PIL.ImageDraw.Draw(combined)

    font_main = None
    for f_p in ["/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf", "C:/Windows/Fonts/malgun.ttf"]:
        if os.path.exists(f_p):
            try: font_main = PIL.ImageFont.truetype(f_p, font_size); break
            except Exception: pass
    if not font_main: font_main = PIL.ImageFont.load_default()

    font_sub = None
    for f_p in ["/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf", "C:/Windows/Fonts/malgun.ttf"]:
        if os.path.exists(f_p):
            try: font_sub = PIL.ImageFont.truetype(f_p, 30); break
            except Exception: pass
    if not font_sub: font_sub = PIL.ImageFont.load_default()

    def parse_hex(c):
        if c.startswith('#'):
            hex_c = c.lstrip('#')
            return tuple(int(hex_c[i:i+2], 16) for i in (0, 2, 4)) + (255,)
        return (255, 255, 255, 255)

    f_color = parse_hex(font_color)
    s_color = parse_hex(stroke_color) if stroke_color else None

    max_w = 880
    wrapped_text = wrap_korean_text(text_str, font_main, max_w, draw)

    bbox = draw.multiline_textbbox((0, 0), wrapped_text, font=font_main, spacing=line_spacing)
    t_w = bbox[2] - bbox[0]
    t_h = bbox[3] - bbox[1]
    
    x_pos = (canvas_w - t_w) // 2
    y_pos = (canvas_h - t_h) // 2 - 40

    if s_color:
        for dx in range(-2, 3):
            for dy in range(-2, 3):
                if dx != 0 or dy != 0:
                    draw.multiline_text((x_pos + dx, y_pos + dy), wrapped_text, font=font_main, fill=s_color, align="center", spacing=line_spacing)

    draw.multiline_text((x_pos, y_pos), wrapped_text, font=font_main, fill=f_color, align="center", spacing=line_spacing)

    if scripture_str:
        scrip_text = f"「 {scripture_str} 」"
        s_bbox = draw.textbbox((0, 0), scrip_text, font=font_sub)
        sx = (canvas_w - (s_bbox[2] - s_bbox[0])) // 2
        sy = y_pos + t_h + 50
        draw.text((sx, sy), scrip_text, fill=(253, 224, 71, 255), font=font_sub)

    if church_name:
        c_bbox = draw.textbbox((0, 0), church_name, font=font_sub)
        cx = (canvas_w - (c_bbox[2] - c_bbox[0])) // 2
        cy = canvas_h - 100
        draw.text((cx, cy), church_name, fill=(147, 197, 253, 255), font=font_sub)

    out_buf = io.BytesIO()
    combined.convert("RGB").save(out_buf, format="PNG")
    out_buf.seek(0)
    return out_buf

async def generate_voiceover_audio(text: str, voice: str = "ko-KR-InJoonNeural") -> str:
    out_path = "./outputs/voiceover_temp.mp3"
    os.makedirs("./outputs", exist_ok=True)
    communicate = edge_tts.Communicate(text, voice)
    await communicate.save(out_path)
    return out_path

# --- 7. 모든 섹션 상단 통일 툴바 컴포넌트 ---
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
            st.download_button(
                "📥 PPT", 
                data=generate_sermon_structure_pptx(
                    title, 
                    st.session_state.get("sermon_scripture", "본문"), 
                    content if content else "내용 없음",
                    st.session_state.get("full_sermon", "")
                ), 
                file_name=f"{title}.pptx", 
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", 
                key=f"dl_ppt_struct_{state_key}"
            )
        with c_txt:
            st.download_button("📥 txt", data=create_txt(title, content if content else "내용 없음"), file_name=f"{title}.txt", mime="text/plain", key=f"dl_txt_{state_key}")

    if st.session_state.get(f"show_copy_{state_key}", False):
        st.info("💡 아래 상자의 텍스트를 복사하여 사용하세요:")
        st.code(content, language="text")

# --- 8. 전역 세션 및 영구 서재 DB 초기화 ---
if "sermon_library" not in st.session_state:
    st.session_state.sermon_library = load_sermons_from_db()

if "current_sermon_idx" not in st.session_state:
    st.session_state.current_sermon_idx = 0

current_s = st.session_state.sermon_library[st.session_state.current_sermon_idx]
if "full_sermon" not in st.session_state or not st.session_state.full_sermon:
    st.session_state.full_sermon = current_s["text"]
if "sermon_summary_text" not in st.session_state or not st.session_state.sermon_summary_text:
    st.session_state.sermon_summary_text = current_s.get("summary", "")
if "sermon_title" not in st.session_state:
    st.session_state.sermon_title = current_s["title"]
if "sermon_scripture" not in st.session_state:
    st.session_state.sermon_scripture = current_s["scripture"]
if "preacher_name" not in st.session_state:
    st.session_state.preacher_name = "김세훈목사"
if "dash_active_view" not in st.session_state:
    st.session_state.dash_active_view = "설교 요약"

# --- 9. 메인 내비게이션 바 ---
app_mode = st.sidebar.radio(
    "🕊️ 플랫폼 대메뉴",
    [
        "📊 설교 대시보드 (메인 작업실)",
        "📤 새 설교 등록/원고작성",
        "🎙️ AI 보이스오버 스튜디오",
        "🎬 쇼츠 만들기 (스튜디오)",
        "📷 말씀카드 이미지",
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

        # 1. 설교 요약
        if active_view == "설교 요약":
            summary_val = st.session_state.get("sermon_summary_text", "")
            
            if not summary_val or summary_val == st.session_state.full_sermon:
                with st.spinner("AI가 설교 원고에서 핵심 요약을 추출 중입니다..."):
                    prompt = f"""
                    설교 제목: {st.session_state.sermon_title}
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 전문: {st.session_state.full_sermon[:4000]}

                    [설교 핵심 요약 작성]
                    위 설교 전문을 바탕으로 강단 선포용 핵심 요약을 100% 한국어로만 작성하세요:
                    🎯 설교 핵심 명제 (1문장)
                    📌 설교 3대지 핵심 요약 (제1대지, 제2대지, 제3대지별 핵심 구절과 2줄 설명)
                    💡 성도를 위한 구체적 실천 적용 3가지
                    🙏 결단 및 축복 기도문 (2줄)
                    """
                    gen_sum = get_ai_response(prompt, is_json=False)
                    if gen_sum:
                        st.session_state.sermon_summary_text = gen_sum
                        summary_val = gen_sum

            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교요약", summary_val, "sermon_sum")
            
            if st.session_state.get("edit_mode_sermon_sum", False):
                s_edit = st.text_area("설교 요약 내용 편집", value=summary_val, height=380, key="edit_sum_area")
                if st.button("💾 저장", key="save_full_sermon"):
                    st.session_state.sermon_summary_text = s_edit
                    st.session_state.edit_mode_sermon_sum = False
                    
                    if "sermon_library" in st.session_state and len(st.session_state.sermon_library) > st.session_state.current_sermon_idx:
                        st.session_state.sermon_library[st.session_state.current_sermon_idx]["summary"] = s_edit
                        save_sermons_to_db(st.session_state.sermon_library)
                        
                    st.success("요약 내용이 저장 및 동기화되었습니다.")
                    st.rerun()
            else:
                st.markdown(f"<div class='content-box'>{summary_val}</div>", unsafe_allow_html=True)

            with st.expander("📜 설교문 원고 전문 보기", expanded=False):
                st.write(st.session_state.full_sermon)

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
                            "📦 전체 PNG (ZIP)",
                            data=generate_cardnews_zip(st.session_state.card_list, st.session_state.sermon_scripture, st.session_state.get("cn_church_name", "")),
                            file_name=f"{st.session_state.sermon_title}_카드뉴스_이미지.zip",
                            mime="application/zip",
                            key="cn_dl_all_zip_btn_top"
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
                    
                    dl_c1, dl_c2 = st.columns(2)
                    with dl_c1:
                        single_png_bytes = generate_single_card_png(curr_card, curr_idx, st.session_state.sermon_scripture, church_input)
                        st.download_button(
                            f"🖼️ CARD {curr_idx + 1} 개별 PNG 다운로드",
                            data=single_png_bytes,
                            file_name=f"{st.session_state.sermon_title}_card_{curr_idx + 1}.png",
                            mime="image/png",
                            key=f"dl_single_png_{curr_idx}"
                        )
                    with dl_c2:
                        st.download_button(
                            "📦 전체 카드뉴스 PNG (ZIP) 다운로드",
                            data=generate_cardnews_zip(cards, st.session_state.sermon_scripture, church_input),
                            file_name=f"{st.session_state.sermon_title}_전체_카드뉴스_PNG.zip",
                            mime="application/zip",
                            key=f"dl_all_zip_under_preview_{curr_idx}"
                        )

                st.write("---")

                st.markdown("#### 인스타그램 캡션")
                insta_c1, insta_c2 = st.columns([4, 1])
                
                insta_text = f"오늘 선포된 [{st.session_state.sermon_title}] ({st.session_state.sermon_scripture}) 말씀을 나눕니다.\n\n주신 은혜를 기억하고 마음에 깊이 새기며 삶 속에서 선한 능력으로 승리하시기를 축복합니다."
                insta_tags = f"#주일설교 #{st.session_state.sermon_title.replace(' ', '')} #말씀묵상 #가정예배 #크리스천"

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
                testament, book = classify_scripture(t_scripture.strip())
                new_entry = {
                    "id": len(st.session_state.sermon_library) + 1,
                    "title": t_title.strip(),
                    "scripture": t_scripture.strip(),
                    "testament": testament,
                    "book": book,
                    "topic": t_tags.split(",")[0].strip() if t_tags else "일반설교",
                    "theology": "직접작성",
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "tags": [t.strip() for t in t_tags.split(",") if t.strip()],
                    "summary": "",
                    "text": t_content.strip()
                }
                st.session_state.sermon_library.append(new_entry)
                save_sermons_to_db(st.session_state.sermon_library)
                
                load_sermon_to_workspace(new_entry, idx=len(st.session_state.sermon_library) - 1)
                st.session_state.dash_active_view = "설교 요약"
                st.success(f"'{t_title}' 설교가 등록되고 모든 메뉴가 동기화되었습니다!")

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

            testament, book = classify_scripture(f_scripture.strip())
            new_entry = {
                "id": len(st.session_state.sermon_library) + 1,
                "title": f_title,
                "scripture": f_scripture,
                "testament": testament,
                "book": book,
                "topic": "파일등록",
                "theology": "파일업로드",
                "date": datetime.now().strftime("%Y-%m-%d"),
                "tags": ["파일등록"],
                "summary": "",
                "text": text
            }
            st.session_state.sermon_library.append(new_entry)
            save_sermons_to_db(st.session_state.sermon_library)

            load_sermon_to_workspace(new_entry, idx=len(st.session_state.sermon_library) - 1)
            st.session_state.dash_active_view = "설교 요약"
            st.success("파일 등록 및 서재 동기화가 완료되었습니다!")

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
                testament, book = classify_scripture(st.session_state.temp_ai_scrip)
                new_entry = {
                    "id": len(st.session_state.sermon_library) + 1,
                    "title": st.session_state.temp_ai_title,
                    "scripture": st.session_state.temp_ai_scrip,
                    "testament": testament,
                    "book": book,
                    "topic": sel_book,
                    "theology": theology_choice.split(' ')[0],
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "tags": [sel_book, theology_choice.split(' ')[0], "강해설교"],
                    "summary": "",
                    "text": st.session_state.temp_generated_sermon
                }
                st.session_state.sermon_library.append(new_entry)
                save_sermons_to_db(st.session_state.sermon_library)

                load_sermon_to_workspace(new_entry, idx=len(st.session_state.sermon_library) - 1)
                st.session_state.dash_active_view = "설교 요약"
                st.success("설교문이 영구 서재에 등록되고 플랫폼 전체가 동기화되었습니다!")

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
                with st.spinner("유튜브 차단 우회망 가동 및 영상 추출 중... (약 20~40초)"):
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
                        st.success("유튜브 영상 숏츠 추출이 완벽하게 완료되었습니다!")
                    except Exception as e:
                        st.warning(f"{str(e)}")

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
# 5. 📷 말씀카드 이미지 스튜디오
# ==============================================================================
elif app_mode == "📷 말씀카드 이미지":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>📷 말씀카드 이미지 스튜디오</h1>", unsafe_allow_html=True)
    st.caption("설교 핵심 구절과 메시지를 바탕으로 무한 고화질 말씀카드 이미지를 제작하고 내려받습니다.")

    vc_c1, vc_c2 = st.columns([1.3, 1.2])

    default_v_text = st.session_state.get("sermon_summary_text", "")
    if default_v_text:
        first_lines = [l.strip() for l in default_v_text.split('\n') if l.strip() and not l.startswith('🎯') and not l.startswith('📌')][:3]
        v_init_str = " ".join(first_lines) if first_lines else st.session_state.sermon_title
    else:
        v_init_str = st.session_state.full_sermon[:120]

    with vc_c1:
        st.markdown("#### ✏️ 말씀 문구 & 디자인 커스텀")
        v_text_input = st.text_area("말씀문구 / 메시지", value=v_init_str, height=120, key="vc_text_in")
        v_scrip_input = st.text_input("성경 구절", value=st.session_state.sermon_scripture, key="vc_scrip_in")
        v_church_input = st.text_input("교회명 배지", value=st.session_state.get("cn_church_name", "화광교회"), key="vc_church_in")

        st.markdown("#### 🎨 배경 & 폰트 스타일링")
        bg_col1, bg_col2 = st.columns(2)
        with bg_col1:
            vc_bg_opt = st.radio("배경 방식", ["사진", "기본", "직접 업로드"], key="vc_bg_radio")
        with bg_col2:
            vc_up_file = None
            if vc_bg_opt == "직접 업로드":
                vc_up_file = st.file_uploader("배경 이미지 업로드", type=["jpg", "png", "jpeg"], key="vc_up_file_widget")

        st.markdown("#### 📐 세부 텍스트 & 오버레이 세팅")
        s_col1, s_col2 = st.columns(2)
        with s_col1:
            vc_fsize = st.slider("폰트 크기 (pt)", min_value=28, max_value=68, value=42, step=2, key="vc_fsize_sl")
            vc_lspace = st.slider("줄간격 (행간)", min_value=10, max_value=40, value=18, step=2, key="vc_lspace_sl")
        with s_col2:
            vc_fcolor = st.color_picker("글자 색상", value="#FDE047", key="vc_fcolor_pk")
            vc_scolor = st.color_picker("테두리 색상", value="#000000", key="vc_scolor_pk")
            vc_opacity = st.slider("배경 어둡기 (투명도)", min_value=0.2, max_value=0.9, value=0.6, step=0.05, key="vc_op_sl")

    with vc_c2:
        st.markdown("### 🖼️ 말씀카드 미리보기 & 다운로드")
        
        card_png_bytes = generate_verse_card_png(
            text_str=v_text_input,
            scripture_str=v_scrip_input,
            bg_option=vc_bg_opt,
            custom_bg_file=vc_up_file,
            font_size=vc_fsize,
            line_spacing=vc_lspace,
            font_color=vc_fcolor,
            stroke_color=vc_scolor,
            overlay_opacity=vc_opacity,
            church_name=v_church_input
        )
        
        # Streamlit 1.38+ 호환성 적용 (use_container_width=True)
        st.image(card_png_bytes.getvalue(), caption="1:1 정사각형 고화질 말씀카드", use_container_width=True)
        
        st.download_button(
            "📥 말씀카드 PNG 고화질 다운로드",
            data=card_png_bytes,
            file_name=f"{st.session_state.sermon_title}_말씀카드.png",
            mime="image/png",
            key="dl_verse_card_png_btn"
        )

# ==============================================================================
# 6. 📚 설교 서재 (Sermon Library) - 영구 보관 및 다차원 정밀 필터링
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>📚 설교 서재 (영구 기록보관소)</h1>", unsafe_allow_html=True)
    
    sermons_db = load_sermons_from_db()
    st.caption(f"총 {len(sermons_db)}편의 설교 원고가 영구 보관 중입니다.")

    st.markdown("#### 🔍 다차원 정밀 검색 및 분류 필터")
    f_c1, f_c2, f_c3, f_c4 = st.columns([1.5, 1.2, 1.5, 1.2])
    
    with f_c1:
        search_kw = st.text_input("검색어 (제목/키워드)", placeholder="예: 신앙전수, 은혜, 고난...", key="lib_search_kw")
    with f_c2:
        testament_filter = st.selectbox("구약/신약 구분", ["전체", "구약", "신약"], key="lib_testament_filter")
    with f_c3:
        book_filter = st.selectbox("성경 66권 책별", ["전체"] + BIBLE_BOOKS, key="lib_book_filter")
    with f_c4:
        sort_order = st.selectbox("정렬 순서", ["최신순 (등록일)", "오래된순"], key="lib_sort_order")

    st.write("---")

    filtered_sermons = []
    for s_item in sermons_db:
        if "testament" not in s_item or "book" not in s_item:
            testament, book = classify_scripture(s_item.get("scripture", ""))
            s_item["testament"] = testament
            s_item["book"] = book

        if search_kw:
            kw_match = (
                search_kw.lower() in s_item.get("title", "").lower() or
                search_kw.lower() in s_item.get("scripture", "").lower() or
                search_kw.lower() in s_item.get("text", "").lower() or
                any(search_kw.lower() in t.lower() for t in s_item.get("tags", []))
            )
            if not kw_match:
                continue

        if testament_filter != "전체" and s_item.get("testament") != testament_filter:
            continue

        if book_filter != "전체" and s_item.get("book") != book_filter:
            continue

        filtered_sermons.append(s_item)

    if sort_order == "오래된순":
        filtered_sermons.sort(key=lambda x: x.get("id", 0))
    else:
        filtered_sermons.sort(key=lambda x: x.get("id", 0), reverse=True)

    st.markdown(f"**검색 결과:** `{len(filtered_sermons)}편` 검색됨")

    if not filtered_sermons:
        st.info("조건에 해당하는 설교문이 서재에 없습니다.")
    else:
        for item_i, s_item in enumerate(filtered_sermons):
            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 14px; padding: 20px; margin-bottom: 14px;">
                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px;">
                            <h3 style="margin: 0; font-size: 20px; font-weight: bold; color: #f8fafc;">{s_item.get('title')}</h3>
                            <span style="background-color: #1e3a8a; color: #fde047; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: bold;">{s_item.get('testament', '구약')} · {s_item.get('book', '성경')}</span>
                        </div>
                        <p style="margin: 0 0 10px 0; font-size: 14px; color: #94a3b8;">
                            📖 <strong>{s_item.get('scripture')}</strong> · [{s_item.get('theology', '개혁주의')}] · 📅 등록일: {s_item.get('date', '2026-08-27')}
                        </p>
                        <div style="margin-bottom: 14px;">
                            {' '.join([f'<span style=\"background:#1e293b; color:#38bdf8; padding:3px 10px; border-radius:6px; font-size:12px; margin-right:4px;\">#{t}</span>' for t in s_item.get('tags', [])])}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                
                b_col1, b_col2 = st.columns([3, 1])
                with b_col1:
                    if st.button("📖 이 설교를 대시보드로 불러와서 작업하기", key=f"lib_load_btn_{s_item.get('id')}_{item_i}"):
                        load_sermon_to_workspace(s_item, idx=st.session_state.sermon_library.index(s_item) if s_item in st.session_state.sermon_library else 0)
                        st.session_state.dash_active_view = "설교 요약"
                        st.success(f"'{s_item.get('title')}' 설교로 모든 플랫폼 메뉴가 완벽히 동기화되었습니다!")
                        st.rerun()
                with b_col2:
                    if st.button("🗑️ 서재에서 삭제", key=f"lib_del_btn_{s_item.get('id')}_{item_i}"):
                        st.session_state.sermon_library = [x for x in st.session_state.sermon_library if x.get('id') != s_item.get('id')]
                        save_sermons_to_db(st.session_state.sermon_library)
                        st.success("설교가 서재에서 삭제되었습니다.")
                        st.rerun()
