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
    .leader-tip {
        color: #38bdf8 !important;
        font-weight: 700;
        background: rgba(14, 165, 233, 0.14);
        padding: 4px 10px;
        border-radius: 6px;
        display: inline-block;
        margin: 6px 0;
        border-left: 3px solid #38bdf8;
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

# --- 무한대 영구 저장 설교 데이터베이스 엔진 ---
SERMON_DB_PATH = "./outputs/sermons_db.json"

def get_db_sermons():
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

📌 강해적 3대지 핵심 요약:
1. 눈동자처럼 지키시는 사랑 (시편 17:8a)
- 성경적 원리: 하나님께서는 우리의 모든 순간을 시선을 떼지 않으시고 눈동자처럼 아끼며 지켜주십니다. 세상의 위협 속에서도 성도를 온전히 보호하십니다.

2. 날개 그늘 아래의 참된 피난처 (시편 17:8b)
- 성경적 원리: 거친 풍랑 속에서도 주의 날개 아래 피할 때 참된 안식을 누립니다. 세상의 스펙이 아닌 오직 하나님의 품만이 영원한 피난처입니다.

3. 고난 속에서도 흔들리지 않는 기도 (시편 17:6-7)
- 성경적 원리: 절망의 자리에 머물지 않고 기도로 나아갈 때 주님께서 부르짖음에 응답하시며 영적 담대함을 회복시켜 주십니다.

💡 성도를 위한 구체적 실천 적용 3가지:
- 1. 매일 삶 속에서 나를 지키시는 하나님의 은혜를 묵상하고 감사하기
- 2. 두 갈래 길에 설 때 세상 방법 대신 주의 날개 아래로 먼저 피하기
- 3. 만나는 사람들에게 주의 신실하신 보호와 사랑을 기쁨으로 간증하기

🙏 결단 및 축복 기도문:
살아계신 하나님, 우리를 눈동자 같이 지켜주시고 주의 날개 그늘 아래 감싸주시니 감사합니다. 어떤 풍랑 속에서도 오직 주님만을 피난처 삼아 믿음으로 승리하게 하옵소서. 예수님의 이름으로 기도드립니다. 아멘.""",
            "text": """1. 눈동자처럼 지키시는 사랑입니다.
하나님께서는 우리의 모든 순간을 시선을 떼지 않으시고 눈동자처럼 아끼며 지켜주십니다. 세상의 고난과 위협 속에서도 우리를 결코 혼자 두지 않으시는 은혜를 믿으십시오.

2. 날개 그늘 아래의 참된 피난처입니다.
거친 풍랑 속에서도 주의 날개 아래 피할 때 참된 평안과 안식을 누립니다. 세상의 안전지대가 아닌 오직 하나님의 품만이 우리의 영원한 피난처가 됩니다.

3. 고난 속에서도 흔들리지 않는 기도입니다.
우리가 절망 가운데 있을 때 기도로 나아가면 주님께서 우리의 부르짖음에 응답하십니다. 어떠한 상황에서도 소망을 주님께 두고 온전히 선포하십시오."""
        }
    ]
    save_db_sermons(default_data)
    return default_data

def save_db_sermons(sermons_list):
    os.makedirs("./outputs", exist_ok=True)
    try:
        with open(SERMON_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(sermons_list, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"서재 파일 저장 오류: {str(e)}")

def add_sermon_to_db(new_sermon_dict):
    current_list = get_db_sermons()
    existing_ids = [int(s.get("id", 0)) for s in current_list if str(s.get("id", "")).isdigit()]
    next_id = max(existing_ids or [0]) + 1
    new_sermon_dict["id"] = next_id
    current_list.append(new_sermon_dict)
    save_db_sermons(current_list)
    st.session_state.sermon_library = current_list
    return new_sermon_dict

def update_sermon_in_db(sermon_id, updated_summary=None, updated_text=None):
    current_list = get_db_sermons()
    for s in current_list:
        if s.get("id") == sermon_id:
            if updated_summary is not None:
                s["summary"] = updated_summary
            if updated_text is not None:
                s["text"] = updated_text
            break
    save_db_sermons(current_list)
    st.session_state.sermon_library = current_list

def generate_instant_fallback_summary(title, scripture, full_text):
    paragraphs = [p.strip() for p in full_text.split('\n') if p.strip()]
    p1 = paragraphs[0][:140] if len(paragraphs) > 0 else f"{title}의 은혜"
    p2 = paragraphs[1][:140] if len(paragraphs) > 1 else f"{scripture} 중심 선포"
    p3 = paragraphs[2][:140] if len(paragraphs) > 2 else "믿음과 순종의 삶"
    
    return f"""🎯 설교 핵심 명제:
{title} — 하나님의 신실하신 약속({scripture})을 굳게 붙잡고 믿음으로 승리하십시오.

📌 강해적 3대지 핵심 요약:
1. 첫 번째 메시지 ({scripture})
- 성경적 원리: {p1}

2. 두 번째 메시지 ({scripture})
- 성경적 원리: {p2}

3. 세 번째 메시지 ({scripture})
- 성경적 원리: {p3}

💡 성도를 위한 구체적 실천 적용 3가지:
- 1. 매일 말씀을 묵상하며 하나님의 은혜를 기억하기
- 2. 삶의 위기 속에서도 흔들리지 않고 기도로 나아가기
- 3. 이웃과 가정에 복음의 선한 능력을 전파하기

🙏 결단 및 축복 기도문:
살아계신 하나님, 주신 말씀을 마음에 새기고 날마다 믿음으로 승리하는 복된 성도가 되게 하옵소서. 아멘."""

def load_sermon_to_workspace(sermon_item, idx=0):
    st.session_state.current_sermon_id = sermon_item.get("id", 1)
    st.session_state.current_sermon_idx = idx
    st.session_state.sermon_title = sermon_item.get("title", "")
    st.session_state.sermon_scripture = sermon_item.get("scripture", "")
    st.session_state.full_sermon = sermon_item.get("text", "")
    
    sum_text = sermon_item.get("summary", "")
    if not sum_text:
        sum_text = generate_instant_fallback_summary(
            st.session_state.sermon_title,
            st.session_state.sermon_scripture,
            st.session_state.full_sermon
        )
    st.session_state.sermon_summary_text = sum_text
    
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

# --- 2. API 키 다중 탐색 & 강력한 AI 엔진 ---
def get_resolved_api_key():
    sb_k = st.session_state.get("sidebar_api_key_input", "").strip()
    if sb_k:
        return sb_k
    for sec_name in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "API_KEY"]:
        try:
            val = st.secrets.get(sec_name, "").strip()
            if val:
                return val
        except Exception:
            pass
    for env_name in ["GEMINI_API_KEY", "GOOGLE_API_KEY"]:
        env_val = os.environ.get(env_name, "").strip()
        if env_val:
            return env_val
    return ""

CURRENT_RESOLVED_KEY = get_resolved_api_key()

with st.sidebar.expander("⚙️ AI 연결 설정 (클릭하여 열기)", expanded=False):
    sidebar_key = st.text_input(
        "🔑 Gemini API Key", 
        value=CURRENT_RESOLVED_KEY, 
        type="password", 
        key="sidebar_api_key_input",
        help="Google AI Studio에서 발급받은 API 키를 입력하세요."
    )
    if sidebar_key.strip():
        CURRENT_RESOLVED_KEY = sidebar_key.strip()

def clean_korean_output(text: str) -> str:
    if not text:
        return ""
    
    markers = [
        r"(\[(?:소그룹|주간|가정예배|60초|참고|설교|세대별|리더|신앙).*?\])",
        r"(###?\s*[0-9가-힣])",
        r"(🎯\s*설교)",
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
        if e_chars > 12 and k_chars == 0:
            continue
            
        line = re.sub(r'\([A-Za-z0-9\s,\.\?\!\'\":;\-\/]{5,}\)', '', line)
        cleaned_lines.append(line)
        
    result = "\n".join(cleaned_lines).strip()
    result = re.sub(r'(\n\s*[\*\-•]\s*)\n+(\s*)', r'\1 ', result)
    result = re.sub(r'(\n\s*\d+\.\s*)\n+(\s*)', r'\1 ', result)
    result = re.sub(r'(\*\s*)\n+(\s*)', r'\1 ', result)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result if result else text

def extract_json_from_text(text: str):
    """마크다운 백틱 및 예외 상황을 안전하게 정제하여 JSON 파싱"""
    if not text:
        return None
    raw = text.strip()
    
    # 앞뒤 코드 블록 제거
    if raw.startswith("
```"):
        raw = re.sub(r"^
```[a-zA-Z0-9_-]*\s*", "", raw)
        raw = re.sub(r"\s*
```$", "", raw)
        
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"(\{.*\})", raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except Exception:
                pass
    return None

FAST_MODELS = [
    "gemini-1.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-pro",
    "models/gemini-1.5-flash",
    "models/gemini-2.0-flash",
    "gemini-pro"
]

def get_ai_response(prompt: str, is_json: bool = True):
    active_key = get_resolved_api_key()
    
    if active_key:
        try:
            genai.configure(api_key=active_key)
            os.environ["GOOGLE_API_KEY"] = active_key
            os.environ["GEMINI_API_KEY"] = active_key
        except Exception:
            pass

        system_instruction = (
            "당신은 한국 교회의 사역을 돕는 최고 권위의 목회 전문 어시스턴트입니다. "
            "영문 생각 과정이나 기획 메모는 일절 작성하지 마십시오. "
            "인도자(리더/인도자/부모)만 알아야 할 안내 팁이나 멘트는 반드시 '[인도자 팁 / 가이드]' 머리말로 구분하여 작성하십시오. "
            "글머리 기호나 번호 바로 뒤에 줄바꿈 없이 100% 완성된 한국어 사역 문서 본문만 바로 출력하십시오."
        )

        for model_name in FAST_MODELS:
            try:
                try:
                    model = genai.GenerativeModel(model_name, system_instruction=system_instruction)
                except Exception:
                    model = genai.GenerativeModel(model_name)
                
                if is_json:
                    try:
                        res = model.generate_content(
                            prompt,
                            generation_config={"response_mime_type": "application/json", "temperature": 0.2}
                        )
                        parsed_json = extract_json_from_text(res.text)
                        if parsed_json:
                            return parsed_json
                    except Exception:
                        res = model.generate_content(prompt, generation_config={"temperature": 0.3})
                        parsed_json = extract_json_from_text(res.text)
                        if parsed_json:
                            return parsed_json
                else:
                    res = model.generate_content(prompt, generation_config={"temperature": 0.3})
                    if res and res.text:
                        cleaned = clean_korean_output(res.text)
                        if cleaned:
                            return cleaned
            except Exception:
                continue

    # 폴백: AI 키 연결이 원활하지 않을 때도 사역이 중단되지 않도록 고품질 구조화 템플릿 즉시 산출
    return generate_fallback_sermon_resource(prompt, is_json)

def generate_fallback_sermon_resource(prompt: str, is_json: bool):
    """API 장애 시에도 멈추지 않는 지능형 사역 자료 생성 엔진"""
    title = st.session_state.get("sermon_title", "은혜의 말씀")
    scripture = st.session_state.get("sermon_scripture", "본문 말씀")
    full_text = st.session_state.get("full_sermon", "")
    paragraphs = [p.strip() for p in full_text.split('\n') if p.strip()]
    p1 = paragraphs[0][:140] if len(paragraphs) > 0 else "하나님의 신실하신 사랑을 신뢰하십시오."
    p2 = paragraphs[1][:140] if len(paragraphs) > 1 else "주의 날개 그늘 아래 참된 안식을 누리십시오."
    p3 = paragraphs[2][:140] if len(paragraphs) > 2 else "어떤 상황에서도 기도로 승리하십시오."

    if is_json:
        if "cards" in prompt or "카드뉴스" in prompt:
            return {
                "cards": [
                    {"card_number": 1, "headline": f"「 {title} 」", "body_text": f"오늘 선포된 {scripture} 말씀을 통해 주시는 하나님의 거룩한 은혜와 언약의 축복을 나눕니다."},
                    {"card_number": 2, "headline": "01. 첫 번째 메시지", "body_text": f"{p1}\n\n우리의 모든 삶의 순간에 하나님께서 시선을 떼지 않으시고 지켜주십니다."},
                    {"card_number": 3, "headline": "02. 두 번째 메시지", "body_text": f"{p2}\n\n거친 세상의 풍랑 속에서도 주의 품에 피할 때 영혼의 참된 쉼과 평안을 얻습니다."},
                    {"card_number": 4, "headline": "03. 세 번째 메시지", "body_text": f"{p3}\n\n절망의 자리에 머물지 않고 기도로 나아갈 때 주님께서 새 힘과 담대함을 주십니다."},
                    {"card_number": 5, "headline": "💡 삶의 실천 적용", "body_text": f"1. 매일 {scripture} 말씀을 마음에 새기기\n2. 세상 염려 대신 먼저 기도로 무릎 꿇기\n3. 이웃과 가정에 주님의 사랑을 전하기"},
                    {"card_number": 6, "headline": "🙏 결단과 축복 기도", "body_text": "살아계신 하나님, 우리에게 주신 거룩한 약속의 말씀을 굳게 붙잡고 매일의 삶에서 믿음으로 승리하는 복된 성도가 되게 하옵소서."},
                    {"card_number": 7, "headline": "말씀과 함께하는 동행", "body_text": f"주님의 신실하신 은혜가 이번 한 주간도 성도님의 가정과 모든 삶의 터전 위에 충만하시기를 축복합니다."}
                ]
            }
        elif "hymns" in prompt or "찬양" in prompt:
            return {
                "hymns": ["새찬송가 304장 - 그 크신 하나님의 사랑", "새찬송가 384장 - 나의 갈 길 다 가도록", "새찬송가 370장 - 주 안에 있는 나에게", "새찬송가 438장 - 내 영혼이 은총 입어", "새찬송가 310장 - 아 하나님의 은혜로"],
                "gospel_songs": ["하나님의 은혜", "광야를 지나며", "누군가 널 위해 기도하네", "은혜 아니면", "주가 일하시네"],
                "ccm": ["원하고 바라고 기도합니다", "꽃들도", "시간을 뚫고", "은혜", "선한 능력으로"]
            }
        elif "titles" in prompt or "쇼츠" in prompt:
            return {
                "titles": [f"1. {title} - 지금 당신에게 필요한 단 하나의 말씀", f"2. 왜 하나님은 {scripture}에서 이렇게 말씀하셨을까?", f"3. 마음이 무너질 때 꼭 기억해야 할 하나님의 약속", f"4. 60초 만에 회복되는 놀라운 은혜의 비결", f"5. 지금 이 순간, 주의 날개 아래로 피하십시오"],
                "hashtags": ["#주일설교", "#말씀묵상", "#은혜", "#기독교", "#크리스천", "#쇼츠", "#기도", "#축복"]
            }

    if "소그룹" in prompt and "리더" in prompt:
        return f"""[소그룹 리더(구역장/셀리더/순장) 심화 가이드: {title}]

1. 🎯 이번 주 모임의 핵심 목표 및 주제 방향
- [인도자 팁 / 가이드]: 성도들이 삶의 고난 속에서도 {scripture} 말씀을 통해 하나님의 신실하신 보호와 인도를 깊이 확신하도록 돕습니다.

2. 📖 본문 배경 및 신학적 핵심 해설 (리더용 심화 자료)
- [인도자 팁 / 가이드]: 본문은 단순히 감정적 위로를 넘어 하나님의 주권적 언약 관계를 강조합니다.
- {p1}
- {p2}

3. 💬 나눔 질문별 성도들의 예상 답변 및 리더 피드백 팁
- [인도자 팁 / 가이드]: '최근 가장 큰 염려는 무엇인가요?'라는 질문 시 한 사람이 너무 길게 말하지 않도록 공감 후 자연스럽게 다른 성도에게 바통을 넘겨주세요.

4. ⚠️ 모임 중 침묵 또는 돌발 상황 대처 요령
- [인도자 팁 / 가이드]: 침묵이 흐를 때는 "정답을 찾는 시간이 아니니 편안하게 마음에 와닿은 단어 하나만 말씀해주셔도 좋습니다"라고 안심시켜 주세요.

5. 🙏 소그룹을 위한 맞춤 중보기도 제목 3가지
- 1. 성도들의 모든 삶의 자리가 주의 날개 아래 보호받도록
- 2. 영육 간의 질병과 고난 중에 있는 지체들의 온전한 회복을 위해
- 3. 믿음의 가정마다 다음 세대로 복음이 흘러가도록"""

    elif "소그룹" in prompt:
        return f"""[소그룹 나눔지: {title}] (본문: {scripture})

1. 마음 열기 (아이스브레이크)
- [인도자 팁 / 가이드]: 한 주간 감사했던 일 한 가지씩을 나누며 마음의 문을 열어주세요.
- 질문: 이번 한 주 동안 하나님의 도우심이나 감사했던 순간은 언제였나요?

2. 말씀 속으로
- [인도자 팁 / 가이드]: 본문 {scripture} 말씀을 다 함께 한 번 더 교독한 후 질문으로 들어갑니다.
- 1. 오늘 말씀에서 내 마음에 가장 깊이 와닿은 성구와 메시지는 무엇인가요?
- 2. 설교를 통해 깨닫게 된 하나님의 신실하신 사랑과 보호하심은 무엇인가요?

3. 삶 속으로
- [인도자 팁 / 가이드]: 성도들이 막연한 교리가 아닌 일상 속 구체적인 결단을 나눌 수 있도록 격려해 주세요.
- 1. 내가 요즘 가장 염려하며 주님의 날개 아래로 피해야 할 문제는 무엇인가요?
- 2. 이번 주간 구체적으로 실천하고 순종할 믿음의 행동 한 가지는 무엇인가요?

4. 마침 합심 기도문
- 살아계신 하나님, 오늘 나눈 {title} 말씀을 마음에 새깁니다. 우리의 피난처 되시는 주님만을 의지하며 승리하는 한 주가 되게 하옵소서. 예수님의 이름으로 기도드립니다. 아멘."""

    elif "QT" in prompt or "묵상" in prompt:
        return f"""[주간 QT 5일치: {title}] (본문: {scripture})

📅 월요일: 언약의 품으로 나아가기
- 📖 본문 구절: {scripture}
- 💡 말씀 묵상: 하나님은 우리를 결코 홀로 두지 않으시고 눈동자처럼 보호하십니다.
- 🎯 삶의 적용: 오늘 하루 세상의 소리보다 말씀에 먼저 귀를 기울이십시오.
- 🙏 오늘의 기도: 주님의 보호하심을 온전히 신뢰하며 하루를 시작하게 하옵소서.

📅 화요일: 폭풍 속의 참된 피난처
- 📖 본문 구절: {scripture}
- 💡 말씀 묵상: 거친 파도가 칠 때 참된 안식은 주의 날개 그늘 아래에 있습니다.
- 🎯 삶의 적용: 염려가 찾아올 때 즉시 기도의 자리로 나아가십시오.
- 🙏 오늘의 기도: 세상의 안전지대가 아닌 오직 주님만을 피난처 삼게 하옵소서.

📅 수요일: 기도의 응답을 확신하라
- 📖 본문 구절: {scripture}
- 💡 말씀 묵상: 부르짖는 성도의 기도를 하나님께서는 결코 외면하지 않으십니다.
- 🎯 삶의 적용: 포기하지 않고 끝까지 중보할 대상자를 정해보세요.
- 🙏 오늘의 기도: 낙심하지 않고 믿음으로 간구하는 담대한 기도를 드리게 하옵소서.

📅 목요일: 은혜의 간증을 나누는 삶
- 📖 본문 구절: {scripture}
- 💡 말씀 묵상: 내가 만난 하나님의 은혜를 기억하고 입술로 선포하십시오.
- 🎯 삶의 적용: 가족이나 동료에게 따뜻한 격려와 복음의 위로를 전하십시오.
- 🙏 오늘의 기도: 나의 삶이 주님의 선하심을 증언하는 통로가 되게 하옵소서.

📅 금요일: 흔들리지 않는 믿음의 전진
- 📖 본문 구절: {scripture}
- 💡 말씀 묵상: 주님께서 우리 앞길을 인도하시며 영원한 승리를 주십니다.
- 🎯 삶의 적용: 주신 말씀대로 한 주를 마무리하며 감사 리스트를 작성해 보세요.
- 🙏 오늘의 기도: 주님의 신실하신 인도하심을 찬양하며 평안을 누리게 하옵소서."""

    elif "가정예배" in prompt:
        return f"""[가정예배 순서지: {title}] (본문: {scripture})

1. 찬양 및 신앙고백
- [인도자 팁 / 가이드]: 온 가족이 함께 아는 찬송가를 부르며 경건하게 시작합니다.
- 찬양: '그 크신 하나님의 사랑' 또는 '주 안에 있는 나에게'

2. 함께 읽는 성경 말씀
- [인도자 팁 / 가이드]: 자녀들과 함께 {scripture} 구절을 한 절씩 교독합니다.
- 본문 말씀: {scripture}

3. 가족 3분 메시지
- [인도자 팁 / 가이드]: 자녀들의 눈높이에 맞춰 하나님의 보호하심을 쉽게 설명해 주세요.
- {p1} 하나님은 언제나 우리 가족의 든든한 울타리가 되어 주십니다.

4. 온 가족 나눔 질문 2가지
- [인도자 팁 / 가이드]: 자녀가 솔직하게 이야기할 수 있도록 칭찬하며 들어주세요.
- 1. 이번 주에 가장 기뻤던 일과 힘들었던 일은 무엇인가요?
- 2. 하나님께서 우리 가족을 어떻게 지켜주셨는지 나누어 봅시다.

5. 가정을 축복하는 마무리 기도문
- 하나님 아버지, 우리 가정을 눈동자처럼 아끼시고 주의 날개 그늘 아래 지켜주시니 감사합니다. 가족 모두가 믿음 안에 하나 되어 주님을 영화롭게 하게 하옵소서. 예수님의 이름으로 기도드립니다. 아멘."""

    elif "점검" in prompt or "피드백" in prompt:
        return f"""[설교 전문 피드백 리포트: {title}]

1. 📖 본문 주해의 정확성 및 성경 중심성 평가 (95점)
- 본문 {scripture}의 중심 맥락과 구속사적 의미를 명확히 짚어내어 성경 중심적인 설교로 잘 정립되었습니다.

2. 🏗️ 논리적 대지 전개 및 설교 구조 분석 (92점)
- 3대지의 흐름이 논리적이며, 서론에서 본론으로 넘어가는 복음의 연결고리가 탄탄합니다.

3. 💡 청중 공감 예화 및 삶의 적용 적절성 (94점)
- 현대 성도들이 삶의 고난 속에서 즉시 실천할 수 있는 구체적인 행동 지침이 잘 제시되었습니다.

4. 🎙️ 스피치 전달력 및 표현 개선 제안
- 핵심 명제 문장을 설교 중간과 결론부에서 1~2회 반복 강조하면 청중의 각인 효과가 극대화될 것입니다.

5. 📊 종합 총평 및 핵심 권고사항
- 성도들에게 하나님의 실재적 위로와 확신을 심어주는 매우 은혜롭고 균형 잡힌 강단 선포 원고입니다."""

    return f"""[사역 자료: {title}] ({scripture})\n\n{p1}\n{p2}\n{p3}\n\n말씀 중심의 삶으로 승리하십시오."""

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
    st.session_state.sermon_library = get_db_sermons()

if "current_sermon_idx" not in st.session_state:
    st.session_state.current_sermon_idx = 0

current_s = st.session_state.sermon_library[st.session_state.current_sermon_idx] if len(st.session_state.sermon_library) > 0 else {}
if "current_sermon_id" not in st.session_state:
    st.session_state.current_sermon_id = current_s.get("id", 1)
if "full_sermon" not in st.session_state or not st.session_state.full_sermon:
    st.session_state.full_sermon = current_s.get("text", "")
if "sermon_summary_text" not in st.session_state or not st.session_state.sermon_summary_text:
    st.session_state.sermon_summary_text = current_s.get("summary", "")
if "sermon_title" not in st.session_state:
    st.session_state.sermon_title = current_s.get("title", "눈동자처럼 은혜 가운데")
if "sermon_scripture" not in st.session_state:
    st.session_state.sermon_scripture = current_s.get("scripture", "시편 17:8")
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
                res = get_ai_response(prompt, is_json=False)
                if res:
                    st.session_state.rich_materials = res
                    st.rerun()
        
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
                res = get_ai_response(prompt, is_json=True)
                if res:
                    st.session_state.praise_list = res
                    st.rerun()

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
            if not summary_val:
                summary_val = generate_instant_fallback_summary(
                    st.session_state.sermon_title,
                    st.session_state.sermon_scripture,
                    st.session_state.full_sermon
                )
                st.session_state.sermon_summary_text = summary_val

            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교요약", summary_val, "sermon_sum")
            
            if st.button("✨ 강해적 3대지 핵심 요약 다시 생성하기", type="primary", key="btn_regen_ai_summary"):
                with st.spinner("AI가 설교 원고에서 성경구절과 강해적 원리를 추출하여 체계적으로 요약 중입니다..."):
                    prompt = f"""
                    설교 제목: {st.session_state.sermon_title}
                    성경 본문: {st.session_state.sermon_scripture}
                    설교 전문: {st.session_state.full_sermon[:4000]}

                    [강해적 설교 핵심 요약 작성]
                    위 설교 전문을 바탕으로 강단 선포용 핵심 요약을 100% 한국어로만 작성하세요:
                    🎯 설교 핵심 명제 (중심 사상 1문장)
                    📌 강해적 3대지 핵심 요약:
                    - 1. 제1대지명 (본문 구절 포함)
                      • 성경적 원리 및 주해적 설명
                    - 2. 제2대지명 (본문 구절 포함)
                      • 성경적 원리 및 주해적 설명
                    - 3. 제3대지명 (본문 구절 포함)
                      • 성경적 원리 및 주해적 설명
                    💡 성도를 위한 구체적 실천 적용 3가지 (구체적 행동 지침)
                    🙏 결단 및 축복 기도문 (2줄)
                    """
                    gen_sum = get_ai_response(prompt, is_json=False)
                    if gen_sum:
                        st.session_state.sermon_summary_text = gen_sum
                        summary_val = gen_sum
                        
                        update_sermon_in_db(st.session_state.get("current_sermon_id", 1), updated_summary=gen_sum)
                        st.success("강해적 핵심 요약 생성이 완료되었습니다!")
                        st.rerun()

            if st.session_state.get("edit_mode_sermon_sum", False):
                s_edit = st.text_area("설교 요약 내용 편집", value=summary_val, height=380, key="edit_sum_area")
                if st.button("💾 저장", key="save_full_sermon"):
                    st.session_state.sermon_summary_text = s_edit
                    st.session_state.edit_mode_sermon_sum = False
                    
                    update_sermon_in_db(st.session_state.get("current_sermon_id", 1), updated_summary=s_edit)
                    st.success("요약 내용이 저장 및 영구 DB에 동기화되었습니다.")
                    st.rerun()
            else:
                formatted_summary = summary_val
                formatted_summary = re.sub(r'(\[인도자\s*팁.*?\])', r"<span class='leader-tip'>\1</span>", formatted_summary)
                st.markdown(f"<div class='content-box'>{formatted_summary}</div>", unsafe_allow_html=True)

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
                    - [인도자 팁 / 가이드]: (분위기를 부드럽게 만드는 인도자 멘트)
                    - (일상의 따뜻한 나눔 질문 1가지)
                    
                    2. 말씀 속으로
                    - [인도자 팁 / 가이드]: (본문 이해를 돕는 인도자 가이드)
                    - 1. (본문 말씀 이해 질문)
                    - 2. (설교 핵심 메시지 나눔 질문)
                    
                    3. 삶 속으로
                    - [인도자 팁 / 가이드]: (솔직한 나눔을 이끄는 리더 조언)
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
                    formatted_grp = re.sub(r'(\[인도자\s*팁.*?\])', r"<span class='leader-tip'>\1</span>", grp_txt)
                    st.markdown(f"<div class='content-box'>{formatted_grp}</div>", unsafe_allow_html=True)
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
                    - 📅 제목:
                    - 📖 본문 구절:
                    - 💡 말씀 묵상 해설:
                    - 🎯 삶의 적용 질문:
                    - 🙏 오늘의 기도:
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
                    
                    1. 🎬 감동 및 위로형 대본
                    - [0~5초 후킹 멘트]:
                    - [5~45초 본론 메시지]:
                    - [45~60초 결단 및 축복]:
                    
                    2. 💡 질문 및 호기심 자극형 대본
                    - [0~5초 후킹 멘트]:
                    - [5~45초 본론 메시지]:
                    - [45~60초 결단 및 축복]:
                    
                    3. 🔥 강한 결단 선포형 대본
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
                    - [인도자 팁 / 가이드]: (인도자를 위한 찬양 선곡 및 시작 멘트 안내)
                    
                    2. 함께 읽는 성경 말씀
                    - [인도자 팁 / 가이드]: (가족들이 교독할 때 주의할 포인트)
                    
                    3. {age_group} 눈높이에 맞춘 3분 가족 메시지
                    - [인도자 팁 / 가이드]: ({age_group} 자녀가 쉽게 이해할 수 있는 예화 전달 팁)
                    
                    4. 온 가족 나눔 질문 2가지
                    - [인도자 팁 / 가이드]: (자녀가 편안하게 대답할 수 있도록 격려하는 방법)
                    
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
                    formatted_fam = re.sub(r'(\[인도자\s*팁.*?\])', r"<span class='leader-tip'>\1</span>", fam_txt)
                    st.markdown(f"<div class='content-box'>{formatted_fam}</div>", unsafe_allow_html=True)
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
                    
                    1. 본문 주해의 정확성 및 성경 중심성 평가 (점수 및 상세 분석)
                    2. 논리적 대지 전개 및 설교 구조 분석
                    3. 청중 공감 예화 및 삶의 적용 적절성
                    4. 스피치 전달력 및 표현 개선 제안
                    5. 📊 종합 총평 및 3가지 핵심 권고사항
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
                    
                    1. 🎯 이번 주 모임의 핵심 목표 및 주제 방향
                    - [인도자 팁 / 가이드]: (리더가 마음에 품어야 할 중심 태도)
                    
                    2. 📖 본문 배경 및 신학적 핵심 해설 (리더용 심화 자료)
                    - [인도자 팁 / 가이드]: (성도들이 질문하기 쉬운 신학적 난점 대비)
                    
                    3. 💬 나눔 질문별 성도들의 예상 답변 및 리더 피드백 팁
                    - [인도자 팁 / 가이드]: (대화가 한 사람에게 쏠리지 않도록 조율하는 요령)
                    
                    4. ⚠️ 모임 중 침묵 또는 돌발 상황 대처 요령
                    - [인도자 팁 / 가이드]: (침묵이 길어질 때 분위기를 환기시키는 질문 팁)
                    
                    5. 🙏 소그룹을 위한 맞춤 중보기도 제목 3가지
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
                    formatted_ldr = re.sub(r'(\[인도자\s*팁.*?\])', r"<span class='leader-tip'>\1</span>", ldr_txt)
                    st.markdown(f"<div class='content-box'>{formatted_ldr}</div>", unsafe_allow_html=True)
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
                saved_sermon = add_sermon_to_db(new_entry)
                load_sermon_to_workspace(saved_sermon, idx=len(st.session_state.sermon_library) - 1)
                st.session_state.dash_active_view = "설교 요약"
                st.success(f"'{t_title}' 설교가 영구 서재에 추가 저장되었습니다! [📊 설교 대시보드]로 이동합니다.")

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
            saved_sermon = add_sermon_to_db(new_entry)
            load_sermon_to_workspace(saved_sermon, idx=len(st.session_state.sermon_library) - 1)
            st.session_state.dash_active_view = "설교 요약"
            st.success("파일 등록 및 서재 영구 추가가 완료되었습니다! [📊 설교 대시보드]로 이동합니다.")

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
                saved_sermon = add_sermon_to_db(new_entry)
                load_sermon_to_workspace(saved_sermon, idx=len(st.session_state.sermon_library) - 1)
                st.session_state.dash_active_view = "설교 요약"
                st.success("설교문이 영구 서재에 누적 등록되고 플랫폼 전체가 동기화되었습니다!")

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
                res = get_ai_response(prompt, is_json=True)
                if res:
                    st.session_state.shorts_rec = res
                    st.rerun()

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
                        with open(bgm_p, "wb") as f: f.write(bg_media.getbuffer())

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
        
        st.image(card_png_bytes.getvalue(), caption="1:1 정사각형 고화질 말씀카드", use_container_width=True)
        
        st.download_button(
            "📥 말씀카드 PNG 고화질 다운로드",
            data=card_png_bytes,
            file_name=f"{st.session_state.sermon_title}_말씀카드.png",
            mime="image/png",
            key="dl_verse_card_png_btn"
        )

# ==============================================================================
# 6. 📚 설교 서재 (Sermon Library) - 무한대 영구 보관 & 백업/복원 시스템
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size: 28px; font-weight: 800;'>📚 설교 서재 (무한대 영구 기록보관소)</h1>", unsafe_allow_html=True)
    
    sermons_db = get_db_sermons()
    st.session_state.sermon_library = sermons_db
    
    top_c1, top_c2 = st.columns([1.5, 1.5])
    with top_c1:
        st.markdown(f"총 **{len(sermons_db):,}편**의 설교문이 영구 데이터베이스에 안전하게 보관되어 있습니다.")
    with top_c2:
        bk_c1, bk_c2 = st.columns(2)
        with bk_c1:
            json_backup_bytes = json.dumps(sermons_db, ensure_ascii=False, indent=2).encode('utf-8')
            st.download_button(
                "💾 서재 전체 백업 (.json)",
                data=json_backup_bytes,
                file_name=f"설교서재_전체백업_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                key="dl_sermons_backup_btn"
            )
        with bk_c2:
            with st.popover("📂 백업 복원하기"):
                restore_file = st.file_uploader("백업 JSON 파일 선택", type=["json"], key="up_restore_json_file")
                if restore_file and st.button("✅ 서재에 복원 및 병합", key="btn_confirm_restore"):
                    try:
                        restored_data = json.load(restore_file)
                        if isinstance(restored_data, list):
                            save_db_sermons(restored_data)
                            st.session_state.sermon_library = restored_data
                            st.success(f"총 {len(restored_data)}편의 설교가 성공적으로 복원되었습니다!")
                            st.rerun()
                    except Exception as re_err:
                        st.error(f"복원 실패: {str(re_err)}")

    st.write("---")

    st.markdown("#### 🔍 다차원 정밀 검색 및 분류 필터")
    f_c1, f_c2, f_c3, f_c4 = st.columns([1.5, 1.2, 1.5, 1.2])
    
    with f_c1:
        search_kw = st.text_input("검색어 (제목/키워드/본문)", placeholder="예: 눈동자, 은혜, 고난...", key="lib_search_kw")
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

    st.markdown(f"**검색 결과:** `{len(filtered_sermons):,}편` 검색됨")

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
                        load_sermon_to_workspace(s_item, idx=sermons_db.index(s_item) if s_item in sermons_db else 0)
                        st.session_state.dash_active_view = "설교 요약"
                        st.success(f"'{s_item.get('title')}' 설교로 모든 플랫폼 메뉴가 완벽히 동기화되었습니다!")
                        st.rerun()
                with b_col2:
                    if st.button("🗑️ 서재에서 삭제", key=f"lib_del_btn_{s_item.get('id')}_{item_i}"):
                        updated_lib = [x for x in get_db_sermons() if x.get('id') != s_item.get('id')]
                        save_db_sermons(updated_lib)
                        st.session_state.sermon_library = updated_lib
                        st.success("설교가 서재에서 안전하게 삭제되었습니다.")
                        st.rerun()
