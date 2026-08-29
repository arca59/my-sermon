# -*- coding: utf-8 -*-
"""
MY 설교 AI 스튜디오 Pro  (v3.0 - 원고 기반 정밀 분석 엔진)

[v3.0 핵심 개선]
 1) 원고 기반 근거(Grounding) 엔진 도입
    - 어떤 본문/제목을 넣어도 '그 원고에서 실제로 뽑아낸' 요약/카드뉴스가 나옵니다.
    - AI 프롬프트에 '원고에 없는 내용 창작 금지 + 근거 문장 인용 의무' 규칙을 강제합니다.
 2) AI 연결 실패를 숨기지 않음
    - 예전에는 API가 죽으면 몰래 '템플릿 문구'가 나와서 엉망인 결과처럼 보였습니다.
    - 이제는 상단에 경고를 띄우고, 대신 원고에서 문장을 직접 추출한 실제 요약을 보여줍니다.
 3) 최신 Gemini 모델 목록 + 자동 탐색
    - 구형 모델명(gemini-1.5-*)만 호출하다 전부 실패하던 문제 해결.
 4) 설교 전환 시 세션/캐시 완전 격리
 5) PPTX 배경 이미지가 검은 사각형에 완전히 가려지던 버그 수정(투명도 적용)
 6) 다운로드 버튼이 매 rerun마다 문서를 새로 만들던 성능 폭탄 제거(캐시)
"""

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
import hashlib
import asyncio
import zipfile
import urllib.parse
import urllib.request
from collections import Counter
from datetime import datetime

from docx import Document
from docx.shared import Pt as DocxPt, RGBColor as DocxRGB
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 선택적 의존성 (없어도 앱 전체가 죽지 않도록 보호)
try:
    import edge_tts
    HAS_TTS = True
except Exception:
    HAS_TTS = False

try:
    import yt_dlp
    HAS_YTDLP = True
except Exception:
    HAS_YTDLP = False

# moviepy 1.x 는 최신 파이썬에서 설치가 실패하므로 2.x 를 우선 사용하고,
# 1.x 스타일 메서드 이름(set_position / subclip / resize / crop ...)은 별칭으로 되살린다.
HAS_MOVIEPY = False
try:
    from moviepy import VideoFileClip, ColorClip, CompositeVideoClip   # moviepy 2.x
    HAS_MOVIEPY = True
except Exception:
    try:
        from moviepy.editor import VideoFileClip, ColorClip, CompositeVideoClip   # moviepy 1.x
        HAS_MOVIEPY = True
    except Exception:
        HAS_MOVIEPY = False

if HAS_MOVIEPY:
    for _cls in (VideoFileClip, ColorClip, CompositeVideoClip):
        for _old, _new in (("set_position", "with_position"), ("set_duration", "with_duration"),
                           ("set_start", "with_start"), ("set_opacity", "with_opacity"),
                           ("set_audio", "with_audio"), ("subclip", "subclipped"),
                           ("resize", "resized"), ("crop", "cropped")):
            if not hasattr(_cls, _old) and hasattr(_cls, _new):
                try:
                    setattr(_cls, _old, getattr(_cls, _new))
                except Exception:
                    pass

try:
    from video_engine import create_animated_video, create_pil_text_clip
    HAS_VIDEO_ENGINE = True
except Exception:
    HAS_VIDEO_ENGINE = False

    def create_animated_video(*args, **kwargs):
        raise RuntimeError("video_engine.py 모듈을 찾을 수 없습니다.")

    def create_pil_text_clip(*args, **kwargs):
        raise RuntimeError("video_engine.py 모듈을 찾을 수 없습니다.")


st.set_page_config(
    page_title="MY 설교 AI 스튜디오 Pro",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main { background-color: #0b1329; }
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
        white-space: pre-wrap;
    }
    .content-box h3 { color: #fde047; font-size: 19px; margin-top: 18px; font-weight: bold; }
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
    .ground-quote {
        color: #a5b4fc !important;
        background: rgba(99, 102, 241, 0.12);
        padding: 3px 8px;
        border-radius: 5px;
        display: inline-block;
        border-left: 3px solid #6366f1;
        font-size: 13.5px;
    }
    .badge-ok  { background:#065f46; color:#d1fae5; padding:3px 10px; border-radius:8px; font-size:12px; }
    .badge-bad { background:#7f1d1d; color:#fee2e2; padding:3px 10px; border-radius:8px; font-size:12px; }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 성경 66권 · 약어 · 분류
# ==============================================================================
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

# 약어 → 정식명 (설교문에 '롬 8:28' 처럼 쓰는 경우 대응)
BIBLE_ABBREV = {
    "창": "창세기", "출": "출애굽기", "레": "레위기", "민": "민수기", "신": "신명기",
    "수": "여호수아", "삿": "사사기", "룻": "룻기", "삼상": "사무엘상", "삼하": "사무엘하",
    "왕상": "열왕기상", "왕하": "열왕기하", "대상": "역대상", "대하": "역대하",
    "스": "에스라", "느": "느헤미야", "에": "에스더", "욥": "욥기", "시": "시편",
    "잠": "잠언", "전": "전도서", "아": "아가", "사": "이사야", "렘": "예레미야",
    "애": "예레미야애가", "겔": "에스겔", "단": "다니엘", "호": "호세아", "욜": "요엘",
    "암": "아모스", "옵": "오바댜", "욘": "요나", "미": "미가", "나": "나훔",
    "합": "하박국", "습": "스바냐", "학": "학개", "슥": "스가랴", "말": "말라기",
    "마": "마태복음", "막": "마가복음", "눅": "누가복음", "요": "요한복음", "행": "사도행전",
    "롬": "로마서", "고전": "고린도전서", "고후": "고린도후서", "갈": "갈라디아서",
    "엡": "에베소서", "빌": "빌립보서", "골": "골로새서", "살전": "데살로니가전서",
    "살후": "데살로니가후서", "딤전": "디모데전서", "딤후": "디모데후서", "딛": "디도서",
    "몬": "빌레몬서", "히": "히브리서", "약": "야고보서", "벧전": "베드로전서",
    "벧후": "베드로후서", "요일": "요한일서", "요이": "요한이서", "요삼": "요한삼서",
    "유": "유다서", "계": "요한계시록",
}


def classify_scripture(scripture_text: str):
    """본문 문자열에서 구약/신약 및 책 이름을 판별한다. (긴 이름 우선 매칭)"""
    if not scripture_text:
        return "기타", "성경전체"
    txt = scripture_text.strip()

    for b in sorted(BIBLE_BOOKS, key=len, reverse=True):
        if b in txt:
            return ("구약" if b in OLD_TESTAMENT_BOOKS else "신약"), b

    for ab in sorted(BIBLE_ABBREV.keys(), key=len, reverse=True):
        if re.search(rf'(^|[^가-힣]){re.escape(ab)}\s*\d', txt):
            full = BIBLE_ABBREV[ab]
            return ("구약" if full in OLD_TESTAMENT_BOOKS else "신약"), full

    return "기타", "성경전체"


# ==============================================================================
# ★★★ 원고 정밀 분석 엔진 (본 앱 품질의 심장) ★★★
#   - 지어내지 않는다. 원고에 실제로 있는 문장/단어만 뽑는다.
# ==============================================================================

KOR_STOPWORDS = set("""
그리고 그러나 그런데 하지만 그래서 그러므로 우리 여러분 오늘 이것 그것 저것 때문 통해 위해 대해
정말 참으로 지금 다시 모든 어떤 무엇 이런 저런 그런 이렇게 그렇게 저렇게 사람 사람들 이야기 경우
때문에 그때 지금은 여기 거기 저기 하나 둘 셋 자신 자기 서로 함께 아주 매우 너무 조금 많이 항상
말씀 하나님 예수님 예수 주님 아멘 성도 성도님 교회 오늘날 우리가 우리는 우리를 저는 제가 여러분이
그분 그것이 있습니다 없습니다 합니다 됩니다 입니다 것입니다 것이다 라고 라는 이라는 하는 되는 있는
없는 같은 위한 통한 한번 다시한번 사실 물론 결국 특히 바로 아마 혹시 만약 그럼 자 이제
안에 밖에 위에 아래 아니라 아니고 말고 자가 나도 너도 여기에 거기에 그때에 오늘은 본문 본문은
읽은 읽는 보면 보니 보십시오 하십시오 합시다 첫째 둘째 셋째 넷째 먼저 다음으로 끝으로 마지막으로
사랑하는 존경하는 축복 은혜 믿음 기도 신앙 삶의 인생 하나님의 우리의 저희 여러분들
""".split())

JOSA_SUFFIX = (
    "으로써", "에게서", "께서는", "에서는", "으로는", "이라는", "에게는", "까지도", "부터는", "라고도",
    "께서", "에게", "에서", "으로", "이라", "라는", "처럼", "까지", "부터", "보다", "조차", "마저",
    "한테", "이나", "이며", "이고", "라도", "든지",
    "은", "는", "이", "가", "을", "를", "에", "의", "도", "와", "과", "로", "만", "랑",
)

VERB_TAIL_PAT = re.compile(
    r'(니다|습니다|십시오|시오|세요|읍시다|했다|한다|하다|하는|하며|하고|해서|해도|'
    r'되어|되고|되는|겠다|것이|것을|것은|것도|이라|라고|면서|는데|지만|어서|아서)$'
)

# 원고 안에서 성경 인용을 찾아내는 정규식
_BOOK_ALT = "|".join(sorted(BIBLE_BOOKS, key=len, reverse=True) +
                     sorted(BIBLE_ABBREV.keys(), key=len, reverse=True))
SCRIPTURE_REF_RE = re.compile(
    rf'((?:{_BOOK_ALT}))\s*(\d{{1,3}})\s*(?:장|[:：])\s*(\d{{1,3}})?\s*(?:[-~–]\s*(\d{{1,3}}))?\s*절?'
)

# 적용/권면 신호 어미 (한국어 설교의 권면 어미를 폭넓게 포착)
APPLY_TAIL_RE = re.compile(
    r'(십시오|ㅂ시다|읍시다|합시다|시기\s*바랍니다|시길\s*바랍니다|기를\s*바랍니다|'
    r'해야\s*합니다|해야만\s*합니다|되시기\s*바랍니다|되기를\s*축복합니다|'
    r'하시기\s*축원합니다|되시기\s*축원합니다|합시다\.|드립시다|삽시다|나아갑시다)'
)

# 대지(포인트) 신호 — 한국 설교의 전형적 구조 표지
POINT_MARKER_RE = re.compile(
    r'^\s*(첫째|둘째|셋째|넷째|다섯째|먼저|다음으로|끝으로|마지막으로|'
    r'제\s*[1-5]\s*대지|[1-5]\s*대지|[1-5][\.\)])[,\s:：]'
)

# 서론·봉독 등 대지 제목으로 부적합한 상투 문장
INTRO_NOISE_RE = re.compile(
    r'(본문은|본문\s*말씀은|함께\s*읽은|봉독|읽어\s*드린|말씀을\s*읽었|오늘\s*우리가\s*함께|'
    r'설교\s*제목은|사랑하는\s*성도\s*여러분|기도하겠습니다)'
)

PRAYER_HINT_RE = re.compile(r'(기도(를)?\s*(드립니다|합니다|드리겠습니다)|예수님의\s*이름으로|아멘)')


_JOSA_SORTED = sorted(JOSA_SUFFIX, key=len, reverse=True)


def _strip_josa(word: str):
    """가장 긴 조사 1개만 제거한 후보를 돌려준다. (없으면 None)"""
    for j in _JOSA_SORTED:
        if len(word) > len(j) + 1 and word.endswith(j):
            return word[:-len(j)]
    return None


def build_stem_map(words):
    """
    조사를 무조건 떼면 '십자가'가 '십자'로 잘려 엉뚱한 키워드가 된다.
    → 어간으로 인정하는 조건:
       (a) 그 어간이 원고에 단독으로도 등장하거나
       (b) 서로 다른 조사 형태가 2개 이상 붙어 나타날 때
    """
    vocab = set(words)
    forms = {}
    for w in vocab:
        cand = _strip_josa(w)
        if cand:
            forms.setdefault(cand, set()).add(w)
    accepted = {s for s, fs in forms.items() if len(fs) >= 2 or s in vocab}
    mapping = {}
    for w in vocab:
        cand = _strip_josa(w)
        if cand and cand in accepted:
            mapping[w] = cand
    return mapping


def _stem(word: str, stem_map=None) -> str:
    if stem_map is not None:
        return stem_map.get(word, word)
    cand = _strip_josa(word)
    return cand if cand else word


def split_sentences(text: str):
    """한국어 문장 분리 (마침표/물음표/느낌표 + '~다.' 종결 고려)"""
    if not text:
        return []
    t = re.sub(r'\s+', ' ', text.replace('\n', ' '))
    raw = re.split(r'(?<=[.!?…])\s+|(?<=다\.)\s*|(?<=요\.)\s*|(?<=까\?)\s*', t)
    out = []
    for s in raw:
        s = s.strip()
        if len(s) >= 8:
            out.append(s)
    return out


def split_paragraphs(text: str):
    if not text:
        return []
    parts = [p.strip() for p in re.split(r'\n\s*\n|\n', text)]
    return [p for p in parts if len(p) >= 15]


def extract_scripture_refs(text: str):
    """원고에 실제로 인용된 성경 구절 목록 (등장 순, 중복 제거)"""
    refs, seen = [], set()
    for m in SCRIPTURE_REF_RE.finditer(text or ""):
        book = BIBLE_ABBREV.get(m.group(1), m.group(1))
        ch = m.group(2)
        v1 = m.group(3)
        v2 = m.group(4)
        if v1 and v2:
            label = f"{book} {ch}:{v1}-{v2}"
        elif v1:
            label = f"{book} {ch}:{v1}"
        else:
            label = f"{book} {ch}장"
        if label not in seen:
            seen.add(label)
            refs.append(label)
    return refs


NOISE_TOKEN_RE = re.compile(r'^(절부터|절까지|장에서|절에서|이라는|말씀은|본문의|장의|절의)$')


def extract_keywords(text: str, top_n: int = 20, stem_map=None):
    """원고에서 '이 설교만의' 특징 키워드 추출"""
    words = re.findall(r'[가-힣]{2,}', text or "")
    if stem_map is None:
        stem_map = build_stem_map(words)
    counter = Counter()
    for w in words:
        s = _stem(w, stem_map)
        if len(s) < 2:
            continue
        if s in KOR_STOPWORDS or w in KOR_STOPWORDS:
            continue
        if VERB_TAIL_PAT.search(s) or NOISE_TOKEN_RE.match(s):
            continue
        if re.match(r'^\d+$', s):
            continue
        counter[s] += 1
    # 1회만 나온 단어는 특징으로 보지 않음 (원고가 아주 짧으면 예외)
    min_cnt = 2 if len(words) > 400 else 1
    return [w for w, c in counter.most_common(top_n * 3) if c >= min_cnt][:top_n]


def _score_sentence(sent: str, kw_rank: dict, stem_map=None) -> float:
    score = 0.0
    for w in re.findall(r'[가-힣]{2,}', sent):
        s = _stem(w, stem_map)
        if s in kw_rank:
            score += kw_rank[s]
    if SCRIPTURE_REF_RE.search(sent):
        score += 3.0
    if '"' in sent or '“' in sent or "'" in sent:
        score += 1.0
    n = len(sent)
    if n < 20:
        score *= 0.5
    elif n > 160:
        score *= 0.8
    return score


@st.cache_data(show_spinner=False, max_entries=64)
def analyze_manuscript(text: str) -> dict:
    """
    설교 원고를 실제로 읽어 구조를 뽑아낸다.
    반환: keywords / refs / paragraphs / sentences / key_sentences / applications /
          prayer / sections(3분할 대표 문단)
    """
    text = (text or "").strip()
    paragraphs = split_paragraphs(text)
    sentences = split_sentences(text)
    stem_map = build_stem_map(re.findall(r'[가-힣]{2,}', text))
    keywords = extract_keywords(text, stem_map=stem_map)
    refs = extract_scripture_refs(text)

    kw_rank = {w: (len(keywords) - i) / len(keywords) * 2.0 for i, w in enumerate(keywords)} if keywords else {}

    def sc(s):
        return _score_sentence(s, kw_rank, stem_map)

    scored = sorted([(s, sc(s)) for s in sentences], key=lambda x: x[1], reverse=True)
    key_sentences = [s for s, _ in scored[:12]]

    # 적용/권면 문장 (원고에 실제로 있는 것만)
    applications = []
    for s in sentences:
        if APPLY_TAIL_RE.search(s) and 15 <= len(s) <= 160:
            applications.append(s)
    applications = sorted(applications, key=sc, reverse=True)[:8]

    # 기도문: 원고 뒤쪽 30% 안에서 기도 신호가 있는 문단
    prayer = ""
    tail_start = int(len(paragraphs) * 0.65)
    for p in reversed(paragraphs[tail_start:] or paragraphs[-3:] if paragraphs else []):
        if PRAYER_HINT_RE.search(p):
            prayer = p
            break

    # ── 대지(sections) 추출 ────────────────────────────────────────────────
    # 1순위: 설교자가 직접 표시한 '첫째/둘째/셋째, 먼저/다음으로/끝으로' 구조를 그대로 사용
    sections = []
    marked = [p for p in paragraphs if POINT_MARKER_RE.match(p)]
    if len(marked) >= 2:
        for p in marked[:4]:
            sents = split_sentences(p)
            head = sents[0] if sents else p
            sections.append({"headline": head[:90], "evidence": p[:420], "chunk": [p]})
    else:
        # 2순위: 원고를 3등분해 각 구간의 대표 문단을 뽑되, 서론·봉독 상투구는 제외
        if paragraphs:
            n = len(paragraphs)
            bounds = [(0, max(1, n // 3)), (max(1, n // 3), max(2, (2 * n) // 3)),
                      (max(2, (2 * n) // 3), n)]
            used = set()
            for a, b in bounds:
                chunk = paragraphs[a:b] or paragraphs[-1:]
                ranked = sorted(chunk, key=sc, reverse=True)
                best = next((p for p in ranked if p not in used and not INTRO_NOISE_RE.search(p[:60])),
                            ranked[0])
                used.add(best)
                cand = [s for s in split_sentences(best) if not INTRO_NOISE_RE.search(s)]
                head = max(cand, key=sc) if cand else best
                sections.append({"headline": head[:90], "evidence": best[:420], "chunk": chunk})

    # 명제로 쓸 문장: 서론·봉독 문장은 제외
    prop_pool = [s for s in key_sentences if not INTRO_NOISE_RE.search(s)]
    key_sentences = prop_pool + [s for s in key_sentences if s not in prop_pool]

    return {
        "keywords": keywords,
        "refs": refs,
        "paragraphs": paragraphs,
        "sentences": sentences,
        "key_sentences": key_sentences,
        "applications": applications,
        "prayer": prayer,
        "sections": sections,
        "char_count": len(text),
    }


def _trim_title(sent: str, limit: int = 42) -> str:
    """문장을 대지 제목처럼 다듬는다 (원문 단어는 유지)"""
    s = re.sub(r'^[0-9]+[\.\)]\s*', '', sent).strip()
    s = re.sub(r'^(첫째|둘째|셋째|넷째|먼저|다음으로|끝으로)[,\s]*', '', s).strip()
    s = re.sub(r'\s+', ' ', s)
    if len(s) <= limit:
        return s.rstrip('.。')
    cut = s[:limit]
    if ' ' in cut:
        cut = cut[:cut.rfind(' ')]
    return cut + "…"


def build_local_summary(title: str, scripture: str, full_text: str) -> str:
    """
    AI 미연결/실패 시에도 '원고에서 실제로 뽑아낸' 요약을 만든다.
    (예전처럼 아무 설교에나 붙는 템플릿 문구를 쓰지 않는다.)
    """
    a = analyze_manuscript(full_text)
    if a["char_count"] < 80:
        return ("⚠️ 설교 원고가 비어 있거나 너무 짧습니다.\n"
                "[📤 새 설교 등록/원고작성] 메뉴에서 설교문 전문을 먼저 등록해 주세요.")

    kw = a["keywords"][:8]
    refs = a["refs"][:8]
    secs = a["sections"]
    apps = a["applications"]

    # 명제는 대지 제목과 중복되지 않는 문장으로 고른다
    sec_heads = {s["headline"] for s in secs}
    prop_src = next((s for s in a["key_sentences"] if s[:90] not in sec_heads),
                    (a["key_sentences"][0] if a["key_sentences"]
                     else (a["paragraphs"][0] if a["paragraphs"] else title)))

    lines = []
    lines.append(f"🧭 [원고 자동 추출 요약]  ·  {title}  ·  {scripture}")
    lines.append(f"※ AI 서버 연결 없이, 업로드하신 설교 원고 본문에서 직접 추출한 결과입니다. (원고 {a['char_count']:,}자 분석)")
    lines.append("")
    lines.append("🎯 설교 핵심 명제")
    lines.append(f"{prop_src}")
    lines.append("")
    lines.append(f"🔑 이 설교의 핵심 키워드: {', '.join(kw) if kw else '(추출 불가)'}")
    if refs:
        lines.append(f"📖 원고에 인용된 성경 구절: {', '.join(refs)}")
    lines.append("")
    lines.append("📌 원고 흐름에 따른 3대지")
    for i, sec in enumerate(secs[:3], start=1):
        lines.append(f"{i}. {_trim_title(sec['headline'])}")
        lines.append(f"   ▸ 원고 근거: \"{sec['evidence'][:220]}{'…' if len(sec['evidence']) > 220 else ''}\"")
        lines.append("")

    lines.append("💡 원고에서 뽑은 실천 적용")
    if apps:
        for i, s in enumerate(apps[:3], start=1):
            lines.append(f"- {i}. {s}")
    else:
        lines.append("- (원고에 직접적인 권면·적용 문장이 없어 추출하지 못했습니다. AI 요약 버튼을 눌러주세요.)")
    lines.append("")

    lines.append("🙏 마침 기도")
    if a["prayer"]:
        lines.append(a["prayer"][:400])
    else:
        lines.append(f"(원고에 기도문 단락이 없습니다. 본문 {scripture}과 위 적용을 근거로 기도문을 작성해 주세요.)")

    return "\n".join(lines)


# 기존 함수명 호환 유지
def analyze_expository_sermon(title: str, scripture: str, full_text: str) -> str:
    return build_local_summary(title, scripture, full_text)


def build_sermon_context(text: str, max_chars: int = 9000) -> str:
    """
    긴 원고를 프롬프트에 넣을 때 앞부분만 잘라내면 결론/적용이 통째로 날아간다.
    앞 45% / 중간 25% / 뒤 30% 를 균형 있게 샘플링해서 전체 흐름을 살린다.
    """
    text = (text or "").strip()
    if len(text) <= max_chars:
        return text
    head_n = int(max_chars * 0.45)
    mid_n = int(max_chars * 0.25)
    tail_n = max_chars - head_n - mid_n
    mid_start = (len(text) - mid_n) // 2
    return (
        text[:head_n]
        + "\n\n…(중략)…\n\n"
        + text[mid_start:mid_start + mid_n]
        + "\n\n…(중략)…\n\n"
        + text[-tail_n:]
    )


# ==============================================================================
# 영구 저장 설교 DB
# ==============================================================================
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
    default_data = [{
        "id": 1,
        "title": "예배와 선교",
        "scripture": "이사야 59:21",
        "testament": "구약",
        "book": "이사야",
        "topic": "예배와 선교",
        "theology": "개혁주의/장로교",
        "date": "2026-08-27",
        "tags": ["선교", "예배", "이사야", "복음"],
        "summary": "",
        "text": """선교전략 용어 중에 1040창(10/40 Window)을 들어보셨을 것입니다. 북위 10도와 40도 사이의 아시아, 북아프리카, 중동 지역을 일컫는 말입니다. 미전도 종족과 빈곤율이 가장 높고 영적 어둠이 짙은 곳입니다.

오늘 본문 이사야 59장 21절은 "내 영과 내 말이 네 입에서 영원토록 떠나지 아니하리라" 말씀하십니다. 참된 예배를 회복할 때 비로소 선교의 문이 열립니다.

우리가 말씀과 성령으로 충만하여 대대손손 복음의 유산을 전수하고 땅끝까지 증인 되는 삶을 살아갑시다."""
    }]
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
    new_sermon_dict["id"] = max(existing_ids or [0]) + 1
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


# ------------------------------------------------------------------------------
# 설교 전환 시 세션 완전 격리
# ------------------------------------------------------------------------------
DERIVED_KEYS = [
    "small_group_text", "qt5_text", "card_list", "shorts_script_text",
    "sermon_audit_text", "leader_guide_text", "rich_materials",
    "praise_list", "shorts_rec", "yt_extracted_result",
    "rendered_shorts_out", "vo_audio_path", "verse_card_img",
    "cn_card_idx", "cn_edit_mode", "cn_church_name",
    "temp_generated_sermon", "temp_ai_title", "temp_ai_scrip",
    "ai_fallback_used", "ai_last_error",
]
DERIVED_PREFIXES = (
    "family_worship_", "edit_mode_", "show_copy_", "cn_h_", "cn_b_",
    "edit_sum_area", "edit_grp_area", "edit_qt_area", "edit_sh_area",
    "edit_audit_area", "edit_ldr_area", "edit_fam_", "vc_text_in",
)


def load_sermon_to_workspace(sermon_item, idx=0):
    """서재/등록에서 설교를 불러올 때, 이전 설교의 결과물 캐시를 100% 소각한다."""
    for k in DERIVED_KEYS:
        st.session_state.pop(k, None)
    for k in list(st.session_state.keys()):
        if any(k.startswith(p) for p in DERIVED_PREFIXES):
            st.session_state.pop(k, None)

    st.session_state.current_sermon_id = sermon_item.get("id", 1)
    st.session_state.current_sermon_idx = idx
    st.session_state.sermon_title = sermon_item.get("title", "")
    st.session_state.sermon_scripture = sermon_item.get("scripture", "")
    st.session_state.full_sermon = sermon_item.get("text", "")

    sum_text = sermon_item.get("summary", "")
    if not sum_text or len(sum_text.strip()) < 50:
        sum_text = build_local_summary(
            st.session_state.sermon_title,
            st.session_state.sermon_scripture,
            st.session_state.full_sermon
        )
    st.session_state.sermon_summary_text = sum_text
    st.session_state.dash_active_view = "설교 요약"


# ==============================================================================
# 보안 접속
# ==============================================================================
def _get_secret(name: str, default: str = "") -> str:
    """secrets.toml 이 아예 없는 환경에서도 죽지 않도록 감싼다."""
    try:
        v = st.secrets.get(name, default)
        return str(v) if v is not None else default
    except Exception:
        return default


USER_PIN = _get_secret("APP_PIN", "7777") or "7777"
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


# ==============================================================================
# AI 엔진 (모델 최신화 + 근거 강제 프롬프트 + 실패 투명화)
# ==============================================================================
def get_resolved_api_key():
    sb_k = st.session_state.get("sidebar_api_key_input", "")
    if sb_k and sb_k.strip():
        return sb_k.strip()
    for sec_name in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "API_KEY"]:
        try:
            val = st.secrets.get(sec_name, "")
            if val and str(val).strip():
                return str(val).strip()
        except Exception:
            pass
    for env_name in ["GEMINI_API_KEY", "GOOGLE_API_KEY"]:
        env_val = os.environ.get(env_name, "").strip()
        if env_val:
            return env_val
    return ""


# 2026년 기준 사용 가능한 모델 우선순위 (구형 1.5 계열은 마지막 예비용)
PREFERRED_MODELS = [
    "gemini-2.5-flash",
    "gemini-flash-latest",
    "gemini-2.0-flash",
    "gemini-2.5-pro",
    "gemini-pro-latest",
    "gemini-2.0-flash-001",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]


@st.cache_data(show_spinner=False, ttl=1800)
def discover_available_models(key_fingerprint: str):
    """계정에서 실제로 호출 가능한 모델을 조회해 우선순위와 교차한다."""
    try:
        names = []
        for m in genai.list_models():
            methods = getattr(m, "supported_generation_methods", []) or []
            if "generateContent" in methods:
                names.append(m.name.replace("models/", ""))
        if not names:
            return PREFERRED_MODELS
        ordered = [m for m in PREFERRED_MODELS if m in names]
        extras = [n for n in names if n.startswith("gemini") and n not in ordered
                  and "vision" not in n and "embedding" not in n]
        return (ordered + extras)[:8] or PREFERRED_MODELS
    except Exception:
        return PREFERRED_MODELS


SYSTEM_INSTRUCTION = (
    "당신은 한국 교회 강단 사역을 돕는 최고 수준의 설교 분석 전문가입니다.\n"
    "반드시 지킬 것:\n"
    "1) 사용자가 제공한 <설교원고> 안에 실제로 존재하는 내용만 사용합니다. "
    "원고에 없는 예화·인물·통계·성경구절을 절대 새로 만들어내지 않습니다.\n"
    "2) 어떤 설교에나 붙일 수 있는 상투적 표현(예: '말씀의 깊은 은혜', '믿음의 결단', "
    "'주님의 신실하신 은혜가 충만하기를')만으로 항목을 채우지 않습니다. "
    "반드시 이 원고에만 등장하는 고유한 단어·지명·인물·숫자·예화를 포함시킵니다.\n"
    "3) 영어 사고 과정, 기획 메모, 인사말, 서두 설명 없이 요청된 결과물 본문만 출력합니다.\n"
    "4) 인도자(리더/구역장/셀리더/부모)용 안내는 '[인도자 팁 / 가이드]: ...' 형식으로 씁니다.\n"
    "5) 100% 한국어로 작성합니다."
)


def build_grounded_prompt(task_block: str, ctx_chars: int = 9000, extra: str = "") -> str:
    """모든 생성 작업에 공통으로 붙는 '근거 강제' 프롬프트 골격"""
    title = st.session_state.get("sermon_title", "")
    scripture = st.session_state.get("sermon_scripture", "")
    raw = st.session_state.get("full_sermon", "")
    a = analyze_manuscript(raw)
    body = build_sermon_context(raw, ctx_chars)

    kw = ", ".join(a["keywords"][:15]) or "(없음)"
    refs = ", ".join(a["refs"][:15]) or "(원고 내 명시적 인용 없음)"

    return f"""[작업 대상 설교]
제목: {title}
설교 본문(대표 성구): {scripture}

[이 원고에서 자동 추출된 고유 키워드]
{kw}

[이 원고에 실제로 인용된 성경 구절]
{refs}

<설교원고>
{body}
</설교원고>

[절대 준수 규칙]
- 위 <설교원고>에 실제로 나오는 내용만 근거로 삼습니다. 원고에 없는 예화/인물/사건/통계/성구를 지어내지 마십시오.
- 위 '고유 키워드' 중 최소 5개 이상을 결과물에 자연스럽게 반영하십시오.
- 모든 항목은 이 설교에만 해당되어야 합니다. 다른 설교에도 그대로 쓸 수 있는 일반론 문장은 금지입니다.
- 대표 성구는 {scripture} 입니다. 원고에 인용되지 않은 다른 성경 구절을 임의로 끌어오지 마십시오.
- 100% 한국어. 영어 메모·머리말·마무리 인사 금지. 아래 형식 그대로만 출력.

{task_block}
{extra}"""


def clean_korean_output(text: str) -> str:
    if not text:
        return ""
    lines = text.split("\n")
    cleaned = []
    for line in lines:
        s = line.strip()
        if not s:
            cleaned.append("")
            continue
        # 순수 영어 메모 줄만 제거 (한글이 섞인 줄은 절대 건드리지 않음)
        k = len(re.findall(r'[가-힣]', s))
        e = len(re.findall(r'[a-zA-Z]', s))
        if k == 0 and e > 20:
            continue
        if re.match(r'^(Okay|Sure|Certainly|Here is|Here\'s|I will|Let me|Note:)\b', s, re.IGNORECASE):
            continue
        cleaned.append(line.rstrip())
    result = re.sub(r'\n{3,}', '\n\n', "\n".join(cleaned).strip())
    return result if result else text


def extract_json_from_text(text):
    if not text:
        return None
    raw = str(text).strip()
    raw = re.sub(r"^`{1,3}[a-zA-Z0-9_-]*\s*", "", raw)
    raw = re.sub(r"\s*`{1,3}$", "", raw).strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"(\{[\s\S]*\})", raw)
        if m:
            try:
                return json.loads(m.group(1))
            except Exception:
                pass
    return None


def get_ai_response(prompt: str, is_json: bool = True, temperature: float = 0.35,
                    kind: str = "summary", card_count: int = 7):
    """
    AI 호출. 실패하면 조용히 템플릿을 뱉지 않고,
    세션에 실패 사유를 기록한 뒤 '원고 기반' 대체 결과를 돌려준다.
    """
    st.session_state.ai_fallback_used = False
    st.session_state.ai_last_error = ""

    active_key = get_resolved_api_key()
    if not active_key:
        st.session_state.ai_fallback_used = True
        st.session_state.ai_last_error = "Gemini API 키가 설정되지 않았습니다. (사이드바 ⚙️ AI 연결 설정)"
        return grounded_fallback(kind, is_json, card_count)

    try:
        genai.configure(api_key=active_key)
        os.environ["GOOGLE_API_KEY"] = active_key
        os.environ["GEMINI_API_KEY"] = active_key
    except Exception as e:
        st.session_state.ai_fallback_used = True
        st.session_state.ai_last_error = f"API 키 설정 오류: {e}"
        return grounded_fallback(kind, is_json, card_count)

    fingerprint = hashlib.sha256(active_key.encode()).hexdigest()[:16]
    models = discover_available_models(fingerprint)

    errors = []
    for model_name in models:
        try:
            try:
                model = genai.GenerativeModel(model_name, system_instruction=SYSTEM_INSTRUCTION)
            except Exception:
                model = genai.GenerativeModel(model_name)

            if is_json:
                cfg = {"response_mime_type": "application/json", "temperature": temperature}
                try:
                    res = model.generate_content(prompt, generation_config=cfg)
                except Exception:
                    res = model.generate_content(prompt, generation_config={"temperature": temperature})
                parsed = extract_json_from_text(getattr(res, "text", ""))
                if parsed:
                    st.session_state.ai_model_used = model_name
                    return parsed
                errors.append(f"{model_name}: JSON 파싱 실패")
            else:
                res = model.generate_content(
                    prompt,
                    generation_config={"temperature": temperature, "max_output_tokens": 8192}
                )
                txt = getattr(res, "text", "") or ""
                cleaned = clean_korean_output(txt)
                if cleaned and len(cleaned.strip()) > 60:
                    st.session_state.ai_model_used = model_name
                    return cleaned
                errors.append(f"{model_name}: 응답이 너무 짧음")
        except Exception as e:
            errors.append(f"{model_name}: {str(e)[:120]}")
            continue

    st.session_state.ai_fallback_used = True
    st.session_state.ai_last_error = " / ".join(errors[:3]) or "알 수 없는 오류"
    return grounded_fallback(kind, is_json, card_count)


# ------------------------------------------------------------------------------
# 원고 기반 대체 생성기 (AI 실패 시에도 '이 설교의' 내용이 나오도록)
# ------------------------------------------------------------------------------
def grounded_fallback(kind: str, is_json: bool, card_count: int = 7):
    title = st.session_state.get("sermon_title", "은혜의 말씀")
    scripture = st.session_state.get("sermon_scripture", "본문 말씀")
    raw = st.session_state.get("full_sermon", "")
    a = analyze_manuscript(raw)
    keys = a["keywords"]
    sents = a["key_sentences"]
    secs = a["sections"]
    apps = a["applications"]

    def pick(i, default=""):
        return sents[i] if len(sents) > i else default

    if is_json:
        if kind == "cards":
            n = max(3, int(card_count))
            cards = [{
                "card_number": 1,
                "headline": f"「 {title} 」",
                "body_text": f"{scripture}\n\n{pick(0, title)}"
            }]
            for i, sec in enumerate(secs[:3], start=2):
                cards.append({
                    "card_number": i,
                    "headline": f"0{i-1}. {_trim_title(sec['headline'], 28)}",
                    "body_text": sec["evidence"][:180]
                })
            idx = len(cards) + 1
            for s in sents[3:3 + max(0, n - len(cards) - 2)]:
                cards.append({"card_number": idx, "headline": f"0{idx-1}. 원고 속 핵심 문장",
                              "body_text": s[:180]})
                idx += 1
            if apps:
                cards.append({
                    "card_number": idx,
                    "headline": "💡 삶의 적용",
                    "body_text": "\n".join(f"{k}. {s[:90]}" for k, s in enumerate(apps[:3], start=1))
                })
                idx += 1
            cards.append({
                "card_number": idx,
                "headline": "🙏 함께 드리는 기도",
                "body_text": (a["prayer"][:180] if a["prayer"]
                              else f"{scripture} 말씀을 붙들고 {', '.join(keys[:3])}의 은혜 안에서 살게 하옵소서.")
            })
            for c in cards:
                pass
            cards = cards[:n]
            for i, c in enumerate(cards, start=1):
                c["card_number"] = i
            while len(cards) < n:
                extra_i = len(cards) + 1
                src = sents[(extra_i + 3) % max(1, len(sents))] if sents else title
                cards.append({"card_number": extra_i, "headline": f"0{extra_i-1}. 말씀 되새김",
                              "body_text": src[:180]})
            return {"cards": cards}

        if kind == "praise":
            topic = ", ".join(keys[:3]) or title
            return {
                "hymns": [f"(AI 미연결) '{topic}' 주제 찬송 검색 필요 - 새찬송가 색인 참조"],
                "gospel_songs": [f"(AI 미연결) '{topic}' 주제 복음성가 검색 필요"],
                "ccm": [f"(AI 미연결) '{topic}' 주제 CCM 검색 필요"],
            }

        if kind == "shorts_meta":
            base = [s for s in sents[:5]] or [title]
            return {
                "titles": [f"{i+1}. {_trim_title(s, 34)}" for i, s in enumerate(base)],
                "hashtags": ["#주일설교", "#말씀묵상"] + [f"#{k}" for k in keys[:5]],
            }
        return {}

    # ---- 텍스트 계열 ----
    kw_line = ", ".join(keys[:8]) or "(추출 불가)"
    warn = ("⚠️ AI 서버에 연결되지 않아, 업로드하신 설교 원고에서 직접 추출한 결과를 표시합니다.\n"
            "   사이드바 [⚙️ AI 연결 설정]에서 Gemini API 키를 확인해 주세요.\n")

    if kind == "rich":
        body = [warn, f"[원고 기반 자료 추출: {title} / {scripture}]", "",
                f"🔑 핵심 키워드: {kw_line}",
                f"📖 원고에 인용된 성구: {', '.join(a['refs'][:10]) or '(없음)'}", "",
                "📝 원고에서 뽑은 인용 가치 있는 문장"]
        for i, s in enumerate(sents[:6], start=1):
            body.append(f"- {i}. {s}")
        return "\n".join(body)

    if kind == "leader":
        out = [warn, f"[소그룹 리더 가이드(원고 추출): {title} / {scripture}]", "",
               "1. 🎯 이번 주 모임의 핵심 방향",
               f"- [인도자 팁 / 가이드]: 원고의 중심 문장 — \"{pick(0, title)}\"", "",
               "2. 📖 원고 흐름 요약(리더 숙지용)"]
        for i, sec in enumerate(secs[:3], start=1):
            out.append(f"- {i}. {_trim_title(sec['headline'])}")
            out.append(f"  ▸ 근거: \"{sec['evidence'][:160]}…\"")
        out += ["", "3. 💬 나눔으로 이어갈 원고 속 권면"]
        for s in (apps[:3] or ["(원고에 권면 문장이 없습니다)"]):
            out.append(f"- {s}")
        return "\n".join(out)

    if kind == "smallgroup":
        out = [warn, f"[소그룹 나눔지(원고 추출): {title} / {scripture}]", "",
               "1. 마음 열기",
               f"- [인도자 팁 / 가이드]: 이번 설교의 키워드 '{keys[0] if keys else title}'로 시작하세요.",
               "", "2. 말씀 속으로"]
        for i, sec in enumerate(secs[:3], start=1):
            out.append(f"- {i}. \"{_trim_title(sec['headline'], 60)}\" 이 대목에서 마음에 남은 것은 무엇인가요?")
        out += ["", "3. 삶 속으로"]
        for i, s in enumerate(apps[:2] or [f"{scripture} 말씀을 이번 주 어디에 적용하시겠습니까?"], start=1):
            out.append(f"- {i}. \"{s}\" — 나에게는 어떻게 적용됩니까?")
        out += ["", "4. 마침 기도", a["prayer"][:300] if a["prayer"] else f"{title} 말씀대로 살게 하옵소서. 아멘."]
        return "\n".join(out)

    if kind == "qt":
        days = ["월요일", "화요일", "수요일", "목요일", "금요일"]
        out = [warn, f"[주간 QT 5일치(원고 추출): {title} / {scripture}]", ""]
        pool = sents if sents else [title] * 5
        for i, d in enumerate(days):
            src = pool[i % len(pool)]
            ref = a["refs"][i % len(a["refs"])] if a["refs"] else scripture
            out += [f"📅 {d}",
                    f"- 📖 본문 구절: {ref}",
                    f"- 💡 말씀 묵상: {src}",
                    f"- 🎯 삶의 적용: {(apps[i % len(apps)] if apps else '오늘 이 말씀을 한 가지 행동으로 옮겨 보십시오.')}",
                    f"- 🙏 오늘의 기도: 주님, 이 말씀을 붙들게 하옵소서.", ""]
        return "\n".join(out)

    if kind == "family":
        out = [warn, f"[가정예배 순서지(원고 추출): {title} / {scripture}]", "",
               "1. 찬양 및 신앙고백",
               "- [인도자 팁 / 가이드]: 가족이 함께 아는 찬송으로 시작하세요.", "",
               "2. 함께 읽는 말씀", f"- 본문: {scripture}", "",
               "3. 3분 가족 메시지(원고 근거)"]
        for sec in secs[:2]:
            out.append(f"- {_trim_title(sec['headline'], 60)}")
        out += ["", "4. 나눔 질문",
                f"- 1. 오늘 말씀에서 '{keys[0] if keys else '은혜'}'는 우리 가정에 어떤 의미일까요?",
                "- 2. 이번 주 우리 가족이 실천할 한 가지는 무엇인가요?", "",
                "5. 마무리 기도", a["prayer"][:250] if a["prayer"] else "우리 가정을 지켜주옵소서. 아멘."]
        return "\n".join(out)

    if kind == "audit":
        out = [warn, f"[원고 구조 자동 점검: {title} / {scripture}]", "",
               f"- 원고 분량: {a['char_count']:,}자 (약 {max(1, a['char_count']//350)}분 분량 추정)",
               f"- 문단 수: {len(a['paragraphs'])} / 문장 수: {len(a['sentences'])}",
               f"- 인용 성구: {len(a['refs'])}개 — {', '.join(a['refs'][:8]) or '없음'}",
               f"- 권면·적용 문장: {len(a['applications'])}개",
               f"- 핵심 키워드: {kw_line}", "",
               "🔎 구조 관찰"]
        if len(a['refs']) < 2:
            out.append("- 본문 인용이 적습니다. 대지마다 근거 구절을 명시하면 강해적 설득력이 올라갑니다.")
        if len(a['applications']) < 3:
            out.append("- 구체적 적용 문장이 부족합니다. '이번 주에 ~하십시오' 형태의 실행 문장을 보강해 보십시오.")
        if not a['prayer']:
            out.append("- 마무리 기도 단락이 확인되지 않습니다.")
        if len(out) == 8:
            out.append("- 구조상 큰 결손은 발견되지 않았습니다.")
        return "\n".join(out)

    if kind == "shorts_script":
        pool = sents or [title]
        out = [warn, f"[쇼츠 대본(원고 추출): {title} / {scripture}]", ""]
        labels = ["🎬 1) 감동·위로형", "💡 2) 질문·호기심형", "🔥 3) 결단 선포형"]
        for i, lab in enumerate(labels):
            hook = pool[i % len(pool)]
            mid = pool[(i + 3) % len(pool)]
            end = (apps[i % len(apps)] if apps else f"{scripture} 말씀을 오늘 붙드십시오.")
            out += [lab,
                    f"- 후킹(0~5초): {hook[:60]}",
                    f"- 본론(5~45초): {mid}",
                    f"- 결단(45~60초): {end}",
                    f"- 자막 키워드: {', '.join(keys[i*2:i*2+5]) or ', '.join(keys[:5])}", ""]
        return "\n".join(out)

    if kind == "sermon_write":
        return ("⚠️ 강해설교문 자동 작성은 AI 전용 기능입니다.\n"
                "사이드바 [⚙️ AI 연결 설정]에서 Gemini API 키를 등록한 뒤 다시 시도해 주세요.\n"
                f"(사유: {st.session_state.get('ai_last_error', '연결 실패')})")

    return build_local_summary(title, scripture, raw)


# ==============================================================================
# 폰트 (캐싱)
# ==============================================================================
# 시스템에 설치돼 있을 만한 한글 폰트 경로 (우선순위 순)
SYSTEM_FONTS_REGULAR = [
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
    "/usr/share/fonts/opentype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "C:/Windows/Fonts/malgun.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
]
SYSTEM_FONTS_BOLD = [
    "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
    "/usr/share/fonts/opentype/nanum/NanumGothicBold.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
    "C:/Windows/Fonts/malgunbd.ttf",
] + SYSTEM_FONTS_REGULAR

# 폰트가 없을 때 내려받을 미러 (User-Agent 없으면 403 이 나므로 반드시 헤더를 붙인다)
FONT_MIRRORS = {
    "regular": [
        "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/nanumgothic/NanumGothic-Regular.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/nanumgothic/NanumGothic-Regular.ttf",
        "https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts@master/NanumGothic.ttf",
    ],
    "bold": [
        "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/nanumgothic/NanumGothic-Bold.ttf",
        "https://raw.githubusercontent.com/google/fonts/main/ofl/nanumgothic/NanumGothic-Bold.ttf",
    ],
}
FONT_DIR = "./fonts"


def _download_font(kind: str, dest: str) -> bool:
    for url in FONT_MIRRORS.get(kind, []):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=20) as r:
                data = r.read()
            if len(data) > 200000:                       # 정상 TTF 인지 최소 검증
                os.makedirs(FONT_DIR, exist_ok=True)
                with open(dest, "wb") as f:
                    f.write(data)
                return True
        except Exception:
            continue
    return False


@st.cache_resource(show_spinner=False)
def ensure_korean_fonts():
    """(regular_path, bold_path) 반환. 확보 실패 시 (None, None)."""
    reg = next((p for p in SYSTEM_FONTS_REGULAR if os.path.exists(p)), None)
    bold = next((p for p in SYSTEM_FONTS_BOLD if os.path.exists(p)), None)

    if not reg:
        cand = os.path.join(FONT_DIR, "NanumGothic-Regular.ttf")
        if os.path.exists(cand) or _download_font("regular", cand):
            reg = cand
    if not bold:
        cand = os.path.join(FONT_DIR, "NanumGothic-Bold.ttf")
        if os.path.exists(cand) or _download_font("bold", cand):
            bold = cand
    return reg, (bold or reg)


@st.cache_resource(show_spinner=False)
def init_korean_font():
    """
    reportlab(PDF)용 한글 폰트 등록.
    reportlab은 .ttc(폰트 컬렉션)나 CFF 계열을 제대로 못 읽는 경우가 많아
    반드시 순수 .ttf 를 우선 시도하고, 없으면 NanumGothic 을 내려받는다.
    """
    reg, _ = ensure_korean_fonts()

    tries = []
    for p in SYSTEM_FONTS_REGULAR:
        if os.path.exists(p) and p.lower().endswith(".ttf"):
            tries.append((p, None))

    dl = os.path.join(FONT_DIR, "NanumGothic-Regular.ttf")
    if os.path.exists(dl) or _download_font("regular", dl):
        tries.append((dl, None))

    if reg and reg.lower().endswith((".ttc", ".otc")):
        tries.append((reg, 0))          # 컬렉션은 첫 번째 서브폰트로 시도

    for path, sub in tries:
        try:
            if sub is None:
                pdfmetrics.registerFont(TTFont("NanumKorean", path))
            else:
                pdfmetrics.registerFont(TTFont("NanumKorean", path, subfontIndex=sub))
            return "NanumKorean"
        except Exception:
            continue
    return "Helvetica"


PDF_FONT_NAME = init_korean_font()
KOREAN_FONT_OK = (PDF_FONT_NAME != "Helvetica")


@st.cache_resource(show_spinner=False)
def get_pil_font(size: int):
    _, bold = ensure_korean_fonts()
    if bold:
        try:
            return PIL.ImageFont.truetype(bold, size)
        except Exception:
            pass
    return PIL.ImageFont.load_default()


# ==============================================================================
# 배경 이미지
# ==============================================================================
CARD_BACKGROUNDS = [
    "https://images.unsplash.com/photo-1506744038136-46273834b3fb?w=1080&q=80",
    "https://images.unsplash.com/photo-1518495973542-4542c06a5843?w=1080&q=80",
    "https://images.unsplash.com/photo-1464822759023-fed622ff2c3b?w=1080&q=80",
    "https://images.unsplash.com/photo-1448375240586-882707db888b?w=1080&q=80",
    "https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=1080&q=80",
    "https://images.unsplash.com/photo-1519681393784-d120267933ba?w=1080&q=80",
    "https://images.unsplash.com/photo-1495616811223-4d98c6e9c869?w=1080&q=80",
    "https://images.unsplash.com/photo-1470252649378-9c29740c9fa8?w=1080&q=80",
]


@st.cache_data(show_spinner=False, ttl=86400)
def fetch_image_bytes(url: str):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=4) as response:
            return response.read()
    except Exception:
        return None


EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FAFF\U00002600-\U000027BF\U0001F1E6-\U0001F1FF⬀-⯿️‍]+"
)


def strip_emoji(text: str) -> str:
    """한글 폰트에는 이모지 글리프가 없어 □(두부)로 찍히므로 이미지에서는 제거한다."""
    return re.sub(r'\s{2,}', ' ', EMOJI_RE.sub('', text or "")).strip()


def wrap_korean_text(text: str, font, max_width: int, draw) -> str:
    if not text:
        return ""
    wrapped = []
    for p in str(text).split('\n'):
        words = p.split(' ')
        curr = ""
        for w in words:
            test = f"{curr} {w}".strip()
            bbox = draw.textbbox((0, 0), test, font=font)
            if (bbox[2] - bbox[0]) > max_width and curr:
                wrapped.append(curr)
                curr = w
            else:
                curr = test
        wrapped.append(curr)
    return "\n".join(wrapped)


# ==============================================================================
# PPTX 투명도 헬퍼 (배경 이미지가 검게 덮이던 버그 수정)
# ==============================================================================
def set_shape_fill_alpha(shape, alpha: float):
    """alpha 0.0(투명) ~ 1.0(불투명)"""
    try:
        spPr = shape.fill._xPr
        solid = spPr.find(qn('a:solidFill'))
        if solid is None:
            return
        clr = solid.find(qn('a:srgbClr'))
        if clr is None:
            return
        for old in clr.findall(qn('a:alpha')):
            clr.remove(old)
        el = etree.SubElement(clr, qn('a:alpha'))
        el.set('val', str(int(max(0.0, min(1.0, alpha)) * 100000)))
    except Exception:
        pass


# ==============================================================================
# 카드 이미지 생성 (캐싱)
# ==============================================================================
@st.cache_data(show_spinner=False, max_entries=64)
def generate_single_card_png_bytes(card_json: str, idx: int, scripture_str: str, church_name: str) -> bytes:
    card_item = json.loads(card_json)
    bg_url = CARD_BACKGROUNDS[idx % len(CARD_BACKGROUNDS)]
    img_b = fetch_image_bytes(bg_url)

    if img_b:
        try:
            base_img = PIL.Image.open(io.BytesIO(img_b)).convert("RGBA").resize((1080, 1080))
        except Exception:
            base_img = PIL.Image.new("RGBA", (1080, 1080), (15, 23, 42, 255))
    else:
        base_img = PIL.Image.new("RGBA", (1080, 1080), (15, 23, 42, 255))

    overlay = PIL.Image.new("RGBA", (1080, 1080), (10, 15, 30, 200))
    combined = PIL.Image.alpha_composite(base_img, overlay)
    draw = PIL.ImageDraw.Draw(combined)

    M = 100                       # 좌우 여백
    MAXW = 1080 - M * 2
    head_txt = strip_emoji(str(card_item.get("headline", "")))
    body_txt = strip_emoji(str(card_item.get("body_text", "")))

    # 글자 수에 따라 폰트 크기를 자동 조절해 넘침/여백을 줄인다
    h_size = 52 if len(head_txt) <= 18 else (46 if len(head_txt) <= 28 else 38)
    b_size = 34 if len(body_txt) <= 90 else (30 if len(body_txt) <= 160 else 26)
    font_b, font_t, font_s = get_pil_font(h_size), get_pil_font(b_size), get_pil_font(26)

    head = wrap_korean_text(head_txt, font_b, MAXW, draw)
    body = wrap_korean_text(body_txt, font_t, MAXW, draw)

    hb = draw.multiline_textbbox((0, 0), head, font=font_b, spacing=14) if head else (0, 0, 0, 0)
    bb = draw.multiline_textbbox((0, 0), body, font=font_t, spacing=18) if body else (0, 0, 0, 0)
    h_h, b_h = hb[3] - hb[1], bb[3] - bb[1]
    gap = 46 if head and body else 0
    block = h_h + gap + b_h

    # 상단 배지/하단 출처 영역을 뺀 공간의 세로 중앙에 배치
    top_limit, bottom_limit = 200, 860
    y = max(top_limit, top_limit + ((bottom_limit - top_limit) - block) // 2)

    draw.text((M, 86), f"CARD {card_item.get('card_number', idx + 1)}",
              fill=(129, 140, 248, 255), font=font_s)
    draw.line([(M, 132), (M + 70, 132)], fill=(253, 224, 71, 255), width=4)

    if head:
        draw.multiline_text((M, y), head, fill=(253, 224, 71, 255), font=font_b, spacing=14)
    if body:
        draw.multiline_text((M, y + h_h + gap), body, fill=(241, 245, 249, 255), font=font_t, spacing=18)

    if scripture_str:
        draw.text((M, 906), f"「 {scripture_str} 」", fill=(253, 224, 71, 255), font=get_pil_font(30))
    if church_name:
        draw.text((M, 960), church_name, fill=(147, 197, 253, 255), font=font_s)

    out = io.BytesIO()
    combined.convert("RGB").save(out, format="PNG")
    return out.getvalue()


@st.cache_data(show_spinner=False, max_entries=16)
def generate_cardnews_zip_bytes(cards_json: str, scripture_str: str, church_name: str) -> bytes:
    cards = json.loads(cards_json)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for i, card in enumerate(cards):
            png = generate_single_card_png_bytes(json.dumps(card, ensure_ascii=False), i, scripture_str, church_name)
            zf.writestr(f"cardnews_{i+1:02d}.png", png)
    return buf.getvalue()


# ==============================================================================
# 문서 내보내기 (전부 캐싱 — 매 rerun 재생성 방지)
# ==============================================================================
@st.cache_data(show_spinner=False, max_entries=48)
def create_docx_bytes(title: str, content: str) -> bytes:
    try:
        doc = Document()
        tp = doc.add_paragraph()
        run = tp.add_run(title)
        run.font.size, run.font.bold = DocxPt(18), True
        run.font.color.rgb = DocxRGB(30, 58, 138)
        doc.add_paragraph(f"작성일: {datetime.now().strftime('%Y-%m-%d')} | MY 설교 AI 스튜디오\n")
        for line in (content or "").split("\n"):
            doc.add_paragraph(line.strip())
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception:
        return (content or "").encode("utf-8")


@st.cache_data(show_spinner=False, max_entries=48)
def create_pdf_bytes(title: str, content: str) -> bytes:
    try:
        font_to_use = init_korean_font()
        bio = io.BytesIO()
        doc = SimpleDocTemplate(bio, pagesize=letter, rightMargin=36, leftMargin=36,
                                topMargin=36, bottomMargin=36)
        t_style = ParagraphStyle("K_Title", fontName=font_to_use, fontSize=15, leading=20,
                                 textColor="#1e3a8a", spaceAfter=8)
        m_style = ParagraphStyle("K_Meta", fontName=font_to_use, fontSize=8, leading=12,
                                 textColor="#64748b", spaceAfter=12)
        b_style = ParagraphStyle("K_Body", fontName=font_to_use, fontSize=9.5, leading=15,
                                 textColor="#1e293b", spaceAfter=5)
        story = [Paragraph(f"<b>{title}</b>", t_style),
                 Paragraph(f"생성일: {datetime.now().strftime('%Y-%m-%d')} | MY 설교 AI 스튜디오", m_style),
                 Spacer(1, 8)]
        for line in (content or "").split("\n"):
            clean = line.strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            if clean:
                clean = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean)
                story.append(Paragraph(clean, b_style))
            else:
                story.append(Spacer(1, 4))
        doc.build(story)
        return bio.getvalue()
    except Exception:
        return (content or "").encode("utf-8")


def create_txt_bytes(title: str, content: str) -> bytes:
    return f"[{title}]\n작성일: {datetime.now().strftime('%Y-%m-%d')}\n\n{content}".encode("utf-8")


@st.cache_data(show_spinner=False, max_entries=32)
def create_document_pptx_bytes(title: str, content: str) -> bytes:
    """일반 문서형 PPT (QT/나눔지/가이드 등)"""
    try:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank = prs.slide_layouts[6]

        s1 = prs.slides.add_slide(blank)
        s1.background.fill.solid()
        s1.background.fill.fore_color.rgb = RGBColor(15, 23, 42)
        tb = s1.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(10.33), Inches(2.0))
        p = tb.text_frame.paragraphs[0]
        p.text = title
        p.font.size, p.font.bold = Pt(36), True
        p.font.color.rgb, p.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER

        chunks, cur = [], ""
        for para in [x.strip() for x in (content or "").split("\n") if x.strip()]:
            if len(cur) + len(para) > 300 and cur:
                chunks.append(cur)
                cur = para + "\n"
            else:
                cur += para + "\n"
        if cur:
            chunks.append(cur)

        for ch in chunks:
            slide = prs.slides.add_slide(blank)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(248, 250, 252)
            htx = slide.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.33), Inches(0.8))
            hp = htx.text_frame.paragraphs[0]
            hp.text = title
            hp.font.size, hp.font.bold = Pt(20), True
            hp.font.color.rgb = RGBColor(30, 58, 138)
            btx = slide.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.33), Inches(5.4))
            btx.text_frame.word_wrap = True
            bp = btx.text_frame.paragraphs[0]
            bp.text = ch.strip()
            bp.font.size = Pt(18)
            bp.font.color.rgb = RGBColor(30, 41, 59)

        bio = io.BytesIO()
        prs.save(bio)
        return bio.getvalue()
    except Exception:
        return (content or "").encode("utf-8")


# ------------------------------------------------------------------------------
# 설교 구조 PPT — 이제 요약 텍스트를 '진짜로' 파싱해서 슬라이드에 넣는다
# ------------------------------------------------------------------------------
def parse_sermon_content(title, scripture, summary_content, full_sermon=""):
    """요약문에서 명제/대지/적용/기도를 구조적으로 뽑아낸다. 실패 시 원고 분석으로 대체."""
    text = summary_content if summary_content and len(summary_content) > 40 else full_sermon
    text = text or ""

    # 명제
    prop = ""
    m = re.search(r'🎯[^\n]*\n(.+?)(?=\n\s*(?:🔑|📖|📌|💡|🙏|$))', text, re.DOTALL)
    if m:
        prop = m.group(1).strip()
    if not prop:
        m = re.search(r'(?:핵심\s*명제|중심\s*사상)[:：]?\s*(.+?)(?=\n\s*\n|\n[🔑📖📌💡🙏]|$)', text, re.DOTALL)
        prop = m.group(1).strip() if m else ""

    # 대지: "1. ..." ~ "3. ..." 블록 (📌 섹션 우선)
    points = []
    seg = text
    ms = re.search(r'📌(.+?)(?=\n\s*💡|\n\s*🙏|$)', text, re.DOTALL)
    if ms:
        seg = ms.group(1)
    for pm in re.finditer(r'(?m)^\s*([1-4])[\.\)]\s*(.+?)(?=\n\s*[1-4][\.\)]\s|\Z)', seg, re.DOTALL):
        block = pm.group(2).strip()
        if len(block) > 4:
            points.append(block)

    a = analyze_manuscript(full_sermon)
    if len(points) < 3:
        points = []
        for sec in a["sections"][:3]:
            points.append(f"{_trim_title(sec['headline'])}\n▸ 원고 근거: {sec['evidence'][:200]}")

    if not prop:
        prop = a["key_sentences"][0] if a["key_sentences"] else title

    # 적용
    app_text = ""
    m = re.search(r'💡[^\n]*\n(.+?)(?=\n\s*🙏|\Z)', text, re.DOTALL)
    if m:
        app_text = m.group(1).strip()
    if not app_text:
        if a["applications"]:
            app_text = "\n".join(f"{i}. {s}" for i, s in enumerate(a["applications"][:3], start=1))
        else:
            app_text = f"{scripture} 말씀을 이번 한 주 삶의 자리에서 한 가지 행동으로 옮기십시오."

    # 기도
    prayer_text = ""
    m = re.search(r'🙏[^\n]*\n(.+)$', text, re.DOTALL)
    if m:
        prayer_text = m.group(1).strip()
    if not prayer_text:
        prayer_text = a["prayer"] or (
            f"주님, 오늘 [{title}] 말씀({scripture})을 마음에 새기고 순종하게 하옵소서. "
            f"예수님의 이름으로 기도드립니다. 아멘.")

    while len(points) < 4:
        idx = len(points)
        extra = a["key_sentences"][idx] if len(a["key_sentences"]) > idx else prop
        points.append(extra)

    return {"prop": prop, "points": points, "app": app_text, "prayer": prayer_text}


@st.cache_data(show_spinner=False, max_entries=24)
def generate_sermon_structure_pptx_bytes(title: str, scripture: str, summary_content: str, full_sermon: str = "") -> bytes:
    try:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        blank = prs.slide_layouts[6]

        def image_dim_slide(slide, img_url=None, dim=0.62):
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(15, 23, 42)
            if img_url:
                b = fetch_image_bytes(img_url)
                if b:
                    slide.shapes.add_picture(io.BytesIO(b), 0, 0, width=Inches(13.333), height=Inches(7.5))
            ov = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
            ov.fill.solid()
            ov.fill.fore_color.rgb = RGBColor(10, 15, 30)
            set_shape_fill_alpha(ov, dim)          # ★ 이미지가 완전히 가려지던 버그 수정
            ov.line.fill.background()
            ov.shadow.inherit = False

        def light_slide(slide):
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(248, 250, 252)

        def body_slide(idx_label, heading, body, bg=None):
            s = prs.slides.add_slide(blank)
            light_slide(s)
            tb = s.shapes.add_textbox(Inches(0.9), Inches(0.7), Inches(11.5), Inches(6.0))
            tf = tb.text_frame
            tf.word_wrap = True
            h = tf.paragraphs[0]
            h.text = f"{idx_label} {heading}".strip()
            h.font.size, h.font.bold = Pt(30), True
            h.font.color.rgb = RGBColor(30, 58, 138)
            b = tf.add_paragraph()
            b.text = "\n" + (body or "")
            b.font.size = Pt(19)
            b.font.color.rgb = RGBColor(30, 41, 59)
            return s

        p = parse_sermon_content(title, scripture, summary_content, full_sermon)
        points = p["points"]

        # 1 표지
        s1 = prs.slides.add_slide(blank)
        image_dim_slide(s1, CARD_BACKGROUNDS[0], dim=0.66)
        tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(10.9), Inches(3.6))
        tb1.text_frame.word_wrap = True
        pp = tb1.text_frame.paragraphs[0]
        pp.text = f"주 일 설 교\n\n{title}\n\n본문 · {scripture}"
        pp.font.size, pp.font.bold = Pt(36), True
        pp.font.color.rgb, pp.alignment = RGBColor(253, 224, 71), PP_ALIGN.CENTER

        # 2 핵심 메시지
        body_slide("들어가며 ·", f"핵심 메시지 ({scripture})", p["prop"])

        # 3 흐름
        s3 = prs.slides.add_slide(blank)
        light_slide(s3)
        h3 = s3.shapes.add_textbox(Inches(0.9), Inches(0.7), Inches(11.5), Inches(0.9))
        hp3 = h3.text_frame.paragraphs[0]
        hp3.text = "설교의 흐름 (Sermon Outline)"
        hp3.font.size, hp3.font.bold = Pt(30), True
        hp3.font.color.rgb = RGBColor(30, 58, 138)
        for i in range(min(4, len(points))):
            shp = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9),
                                      Inches(2.0 + i * 1.15), Inches(11.5), Inches(0.95))
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(255, 255, 255)
            shp.line.color.rgb = RGBColor(203, 213, 225)
            shp.shadow.inherit = False
            cp = shp.text_frame.paragraphs[0]
            cp.text = f"  0{i+1}   {_trim_title(points[i].split(chr(10))[0], 52)}"
            cp.font.size, cp.font.bold = Pt(18), True
            cp.font.color.rgb = RGBColor(30, 41, 59)

        # 4~7 대지
        for i in range(min(4, len(points))):
            body_slide(f"0{i+1}.", "", points[i])

        # 8 적용
        body_slide("삶의 적용 ·", "이렇게 살아갑시다", p["app"])

        # 9 기도
        s9 = prs.slides.add_slide(blank)
        image_dim_slide(s9, CARD_BACKGROUNDS[3], dim=0.7)
        tb9 = s9.shapes.add_textbox(Inches(1.1), Inches(1.3), Inches(11.1), Inches(5.0))
        tb9.text_frame.word_wrap = True
        h9 = tb9.text_frame.paragraphs[0]
        h9.text = "결단과 마침 기도문"
        h9.font.size, h9.font.bold = Pt(28), True
        h9.font.color.rgb = RGBColor(253, 224, 71)
        b9 = tb9.text_frame.add_paragraph()
        b9.text = "\n" + p["prayer"]
        b9.font.size = Pt(18)
        b9.font.color.rgb = RGBColor(241, 245, 249)

        bio = io.BytesIO()
        prs.save(bio)
        return bio.getvalue()
    except Exception:
        return create_document_pptx_bytes(title, summary_content)


@st.cache_data(show_spinner=False, max_entries=16)
def generate_cardnews_pptx_bytes(cards_json: str, church_name: str = "") -> bytes:
    slides_data = json.loads(cards_json)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(10)
    blank = prs.slide_layouts[6]

    for idx, item in enumerate(slides_data):
        slide = prs.slides.add_slide(blank)
        b = fetch_image_bytes(CARD_BACKGROUNDS[idx % len(CARD_BACKGROUNDS)])
        if b:
            slide.shapes.add_picture(io.BytesIO(b), Inches(0), Inches(0), width=Inches(10), height=Inches(10))
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(15, 23, 42)

        ov = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(10))
        ov.fill.solid()
        ov.fill.fore_color.rgb = RGBColor(10, 15, 30)
        set_shape_fill_alpha(ov, 0.72)             # ★ 배경이 보이도록 반투명 처리
        ov.line.fill.background()
        ov.shadow.inherit = False

        bd = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(2.4), Inches(0.6))
        bp = bd.text_frame.paragraphs[0]
        bp.text = f"CARD {item.get('card_number', idx + 1)}"
        bp.font.size, bp.font.bold = Pt(14), True
        bp.font.color.rgb = RGBColor(129, 140, 248)

        tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(8.4), Inches(2.4))
        tbox.text_frame.word_wrap = True
        tp = tbox.text_frame.paragraphs[0]
        tp.text = item.get("headline", "")
        tp.font.size, tp.font.bold = Pt(30), True
        tp.font.color.rgb = RGBColor(253, 224, 71)

        bbox = slide.shapes.add_textbox(Inches(0.8), Inches(4.2), Inches(8.4), Inches(4.6))
        bbox.text_frame.word_wrap = True
        bb = bbox.text_frame.paragraphs[0]
        bb.text = item.get("body_text", "")
        bb.font.size = Pt(20)
        bb.font.color.rgb = RGBColor(241, 245, 249)

        if church_name:
            cbox = slide.shapes.add_textbox(Inches(0.8), Inches(9.0), Inches(8.4), Inches(0.6))
            cp = cbox.text_frame.paragraphs[0]
            cp.text = church_name
            cp.font.size = Pt(14)
            cp.font.color.rgb, cp.alignment = RGBColor(147, 197, 253), PP_ALIGN.CENTER

    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()


# ==============================================================================
# 말씀카드
# ==============================================================================
def generate_verse_card_png(text_str, scripture_str, bg_option="사진", custom_bg_file=None,
                            font_size=42, line_spacing=18, font_color="#FDE047",
                            stroke_color="#000000", overlay_opacity=0.6, church_name="",
                            bg_index=0):
    W = H = 1080
    if bg_option == "직접 업로드" and custom_bg_file is not None:
        try:
            base = PIL.Image.open(custom_bg_file).convert("RGBA").resize((W, H))
        except Exception:
            base = PIL.Image.new("RGBA", (W, H), (15, 23, 42, 255))
    elif bg_option == "기본":
        base = PIL.Image.new("RGBA", (W, H), (15, 23, 42, 255))
    else:
        b = fetch_image_bytes(CARD_BACKGROUNDS[bg_index % len(CARD_BACKGROUNDS)])
        try:
            base = PIL.Image.open(io.BytesIO(b)).convert("RGBA").resize((W, H)) if b \
                else PIL.Image.new("RGBA", (W, H), (15, 23, 42, 255))
        except Exception:
            base = PIL.Image.new("RGBA", (W, H), (15, 23, 42, 255))

    overlay = PIL.Image.new("RGBA", (W, H), (10, 15, 30, int(255 * overlay_opacity)))
    combined = PIL.Image.alpha_composite(base, overlay)
    draw = PIL.ImageDraw.Draw(combined)

    font_main = get_pil_font(int(font_size))
    font_sub = get_pil_font(30)

    def parse_hex(c):
        try:
            h = c.lstrip('#')
            return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (255,)
        except Exception:
            return (255, 255, 255, 255)

    f_color = parse_hex(font_color)
    s_color = parse_hex(stroke_color) if stroke_color else None

    wrapped = wrap_korean_text(strip_emoji(text_str), font_main, 880, draw)
    bbox = draw.multiline_textbbox((0, 0), wrapped, font=font_main, spacing=line_spacing)
    t_w, t_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (W - t_w) // 2
    y = max(120, (H - t_h) // 2 - 50)

    if s_color:
        for dx in (-2, -1, 0, 1, 2):
            for dy in (-2, -1, 0, 1, 2):
                if dx or dy:
                    draw.multiline_text((x + dx, y + dy), wrapped, font=font_main,
                                        fill=s_color, align="center", spacing=line_spacing)
    draw.multiline_text((x, y), wrapped, font=font_main, fill=f_color, align="center", spacing=line_spacing)

    if scripture_str:
        s_txt = f"「 {scripture_str} 」"
        sb = draw.textbbox((0, 0), s_txt, font=font_sub)
        draw.text(((W - (sb[2] - sb[0])) // 2, min(y + t_h + 50, H - 170)),
                  s_txt, fill=(253, 224, 71, 255), font=font_sub)
    if church_name:
        cb = draw.textbbox((0, 0), church_name, font=font_sub)
        draw.text(((W - (cb[2] - cb[0])) // 2, H - 100), church_name,
                  fill=(147, 197, 253, 255), font=font_sub)

    out = io.BytesIO()
    combined.convert("RGB").save(out, format="PNG")
    out.seek(0)
    return out


# ==============================================================================
# TTS
# ==============================================================================
async def _tts(text: str, voice: str, out_path: str):
    communicate = edge_tts.Communicate(text, voice)
    await communicate.save(out_path)


def generate_voiceover_audio(text: str, voice: str = "ko-KR-InJoonNeural") -> str:
    if not HAS_TTS:
        raise RuntimeError("edge-tts 패키지가 설치되어 있지 않습니다. (pip install edge-tts)")
    os.makedirs("./outputs", exist_ok=True)
    out_path = f"./outputs/voiceover_{int(datetime.now().timestamp())}.mp3"
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(_tts(text, voice, out_path))
        loop.close()
    except RuntimeError:
        asyncio.run(_tts(text, voice, out_path))
    return out_path


# ==============================================================================
# 유튜브 → 9:16 숏츠
# ==============================================================================
def extract_youtube_to_shorts(yt_url, start_sec, duration_sec, title, subtitle_text, church_name=""):
    if not (HAS_YTDLP and HAS_MOVIEPY and HAS_VIDEO_ENGINE):
        raise Exception("영상 처리 모듈(yt-dlp / moviepy / video_engine)이 준비되지 않았습니다.")

    out_dir = "./outputs"
    os.makedirs(out_dir, exist_ok=True)
    src_video = os.path.join(out_dir, "yt_raw_source.mp4")

    vid = re.search(r'(?:v=|/live/|/shorts/|youtu\.be/)([a-zA-Z0-9_-]{11})', yt_url.strip())
    clean_url = f"https://www.youtube.com/watch?v={vid.group(1)}" if vid else yt_url.strip()

    ok = False
    for clients in (['ios'], ['android_creator'], ['mweb'], ['tv_embedded'], ['web']):
        if os.path.exists(src_video):
            try:
                os.remove(src_video)
            except Exception:
                pass
        opts = {
            'format': 'best[ext=mp4][height<=720]/bestvideo[ext=mp4][height<=720]+bestaudio[ext=m4a]/best',
            'outtmpl': src_video, 'overwrites': True, 'quiet': True, 'no_warnings': True,
            'nocheckcertificate': True, 'geo_bypass': True,
            'http_headers': {'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_4_1 like Mac OS X)'},
            'extractor_args': {'youtube': {'player_client': clients}},
        }
        try:
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.download([clean_url])
            if os.path.exists(src_video) and os.path.getsize(src_video) > 100000:
                ok = True
                break
        except Exception:
            continue

    if not ok:
        raise Exception("유튜브 서버가 클라우드 IP를 차단했습니다. "
                        "[🎨 AI 나레이션 & 템플릿 숏츠 제작] 탭에서 영상을 직접 업로드하시면 즉시 9:16 렌더링이 가능합니다.")

    raw = VideoFileClip(src_video)
    start_sec = max(0, min(start_sec, int(raw.duration) - 5))
    end_sec = min(start_sec + duration_sec, int(raw.duration))
    sub = raw.subclip(start_sec, end_sec)

    w, h = sub.size
    target_w = int(h * 9 / 16)
    cropped = sub.crop(x1=w / 2 - target_w / 2, y1=0, x2=w / 2 + target_w / 2, y2=h) if w > target_w else sub
    final = cropped.resize((1080, 1920))
    dur = final.duration

    layers = [final,
              ColorClip(size=(1080, 240), color=(0, 0, 0), duration=dur).set_opacity(0.45).set_position(('center', 100)),
              create_pil_text_clip(title, fontsize=48, color="#FDE047", stroke_color="black",
                                   stroke_width=2, size=(920, None), duration=dur).set_position(("center", 140))]
    if subtitle_text:
        layers.append(create_pil_text_clip(subtitle_text, fontsize=40, color="white", stroke_color="black",
                                           stroke_width=2, size=(900, None), duration=dur).set_position(("center", 1400)))
    if church_name:
        layers.append(create_pil_text_clip(church_name, fontsize=28, color="#93C5FD", stroke_color="black",
                                           stroke_width=1, size=(800, None), duration=dur).set_position(("center", 1780)))

    comp = CompositeVideoClip(layers)
    out_file = os.path.join(out_dir, f"yt_shorts_{int(datetime.now().timestamp())}.mp4")
    comp.write_videofile(out_file, fps=24, codec="libx264", audio_codec="aac", threads=4, preset="ultrafast")
    raw.close(); sub.close(); comp.close()
    return out_file


# ==============================================================================
# 공통 툴바
# ==============================================================================
def render_section_top_toolbar(title: str, content: str, state_key: str, ppt_mode: str = "doc"):
    content = content or "내용 없음"
    c_title, c_btn = st.columns([1.2, 2.8])
    with c_title:
        st.markdown(f"<h3 style='margin:0;font-size:20px;font-weight:800;line-height:1.3;'>{title}</h3>",
                    unsafe_allow_html=True)
    with c_btn:
        b1, b2, b3, b4, b5, b6 = st.columns([1, 1, 1.1, 1.1, 1.1, 1])
        with b1:
            if st.button("✏️ 수정", key=f"edit_btn_{state_key}"):
                st.session_state[f"edit_mode_{state_key}"] = not st.session_state.get(f"edit_mode_{state_key}", False)
        with b2:
            if st.button("📋 복사", key=f"copy_btn_{state_key}"):
                st.session_state[f"show_copy_{state_key}"] = not st.session_state.get(f"show_copy_{state_key}", False)
        with b3:
            st.download_button("📥 워드", data=create_docx_bytes(title, content), file_name=f"{title}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key=f"dl_docx_{state_key}")
        with b4:
            st.download_button("📥 PDF", data=create_pdf_bytes(title, content), file_name=f"{title}.pdf",
                               mime="application/pdf", key=f"dl_pdf_{state_key}")
        with b5:
            if ppt_mode == "sermon":
                ppt_bytes = generate_sermon_structure_pptx_bytes(
                    st.session_state.get("sermon_title", title),
                    st.session_state.get("sermon_scripture", ""),
                    content,
                    st.session_state.get("full_sermon", "")
                )
            else:
                ppt_bytes = create_document_pptx_bytes(title, content)
            st.download_button("📥 PPT", data=ppt_bytes, file_name=f"{title}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               key=f"dl_ppt_{state_key}")
        with b6:
            st.download_button("📥 txt", data=create_txt_bytes(title, content), file_name=f"{title}.txt",
                               mime="text/plain", key=f"dl_txt_{state_key}")

    if st.session_state.get(f"show_copy_{state_key}", False):
        st.info("💡 아래 상자의 텍스트를 복사하여 사용하세요:")
        st.code(content, language="text")


def st_image_full(data, caption=None):
    """streamlit 버전에 따라 이미지 폭 옵션이 달라 호환 처리"""
    try:
        st.image(data, caption=caption, width='stretch')
    except Exception:
        st.image(data, caption=caption, use_container_width=True)


def show_ai_status():
    """AI 폴백이 쓰였으면 숨기지 않고 알린다."""
    if st.session_state.get("ai_fallback_used"):
        st.warning(
            "⚠️ **AI 서버 응답을 받지 못해, 설교 원고에서 직접 추출한 결과를 표시했습니다.**\n\n"
            f"사유: `{st.session_state.get('ai_last_error', '알 수 없음')}`\n\n"
            "→ 사이드바 **[⚙️ AI 연결 설정]** 에서 Gemini API 키를 확인해 주세요."
        )
    elif st.session_state.get("ai_model_used"):
        st.caption(f"✅ 생성 모델: `{st.session_state.ai_model_used}`")


def render_body(text: str):
    html = (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    html = re.sub(r'(\[인도자\s*팁[^\]]*\])', r"<span class='leader-tip'>\1</span>", html)
    html = re.sub(r'(▸\s*원고\s*근거[^\n]*)', r"<span class='ground-quote'>\1</span>", html)
    st.markdown(f"<div class='content-box'>{html}</div>", unsafe_allow_html=True)


def editable_section(state_key: str, session_field: str, label: str, height: int = 350,
                     persist_summary: bool = False):
    """수정 모드 공통 처리"""
    val = st.session_state.get(session_field, "")
    if st.session_state.get(f"edit_mode_{state_key}", False):
        edited = st.text_area(label, value=val, height=height, key=f"ta_{state_key}")
        if st.button("💾 저장", key=f"save_{state_key}"):
            st.session_state[session_field] = edited
            st.session_state[f"edit_mode_{state_key}"] = False
            if persist_summary:
                update_sermon_in_db(st.session_state.get("current_sermon_id", 1), updated_summary=edited)
            st.success("저장되었습니다.")
            st.rerun()
        return False
    return True


# ==============================================================================
# 세션 초기화
# ==============================================================================
if "sermon_library" not in st.session_state:
    st.session_state.sermon_library = get_db_sermons()
if "current_sermon_idx" not in st.session_state:
    st.session_state.current_sermon_idx = 0

_cur = st.session_state.sermon_library[st.session_state.current_sermon_idx] \
    if st.session_state.sermon_library else {}

st.session_state.setdefault("current_sermon_id", _cur.get("id", 1))
st.session_state.setdefault("sermon_title", _cur.get("title", "예배와 선교"))
st.session_state.setdefault("sermon_scripture", _cur.get("scripture", "이사야 59:21"))
st.session_state.setdefault("full_sermon", _cur.get("text", ""))
st.session_state.setdefault("sermon_summary_text", _cur.get("summary", ""))
st.session_state.setdefault("preacher_name", "김세훈목사")
st.session_state.setdefault("dash_active_view", "설교 요약")
st.session_state.setdefault("cn_church_name", "화광교회")
st.session_state.setdefault("ai_model_used", "")


# ==============================================================================
# 사이드바
# ==============================================================================
with st.sidebar.expander("⚙️ AI 연결 설정", expanded=not bool(get_resolved_api_key())):
    st.text_input("🔑 Gemini API Key", value=get_resolved_api_key(), type="password",
                  key="sidebar_api_key_input",
                  help="Google AI Studio(aistudio.google.com)에서 발급받은 키를 입력하세요.")
    if get_resolved_api_key():
        st.markdown("<span class='badge-ok'>키 등록됨</span>", unsafe_allow_html=True)
        if st.button("🔌 연결 테스트", key="btn_test_api"):
            with st.spinner("모델 조회 중..."):
                try:
                    genai.configure(api_key=get_resolved_api_key())
                    fp = hashlib.sha256(get_resolved_api_key().encode()).hexdigest()[:16]
                    discover_available_models.clear()
                    ms = discover_available_models(fp)
                    st.success("사용 가능 모델: " + ", ".join(ms[:5]))
                except Exception as e:
                    st.error(f"연결 실패: {e}")
    else:
        st.markdown("<span class='badge-bad'>키 없음 · AI 기능 제한</span>", unsafe_allow_html=True)

if not KOREAN_FONT_OK:
    st.sidebar.error(
        "⚠️ 서버에 한글 폰트가 없어 PDF·카드 이미지의 한글이 깨질 수 있습니다.\n\n"
        "Streamlit Cloud라면 저장소 루트에 `packages.txt` 파일을 만들고 한 줄만 넣어 재배포하세요:\n\n"
        "```\nfonts-nanum\n```"
    )

st.sidebar.markdown(
    f"**📖 현재 작업 설교**\n\n{st.session_state.sermon_title}\n\n`{st.session_state.sermon_scripture}` "
    f"· 원고 {len(st.session_state.full_sermon):,}자"
)

app_mode = st.sidebar.radio(
    "🕊️ 플랫폼 대메뉴",
    ["📊 설교 대시보드 (메인 작업실)",
     "📤 새 설교 등록/원고작성",
     "🎙️ AI 보이스오버 스튜디오",
     "🎬 쇼츠 만들기 (스튜디오)",
     "📷 말씀카드 이미지",
     "📚 설교 서재 (Sermon Library)"]
)


# ==============================================================================
# 1) 설교 대시보드
# ==============================================================================
if app_mode == "📊 설교 대시보드 (메인 작업실)":
    st.markdown(
        f"""<div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;">
        <h1 style="font-size:28px;font-weight:800;margin:0;color:#f8fafc;">{st.session_state.sermon_title}</h1>
        <span style="background:#2563eb;color:#fff;padding:4px 10px;border-radius:6px;font-size:13px;font-weight:bold;">
        본문 {st.session_state.sermon_scripture}</span></div>""",
        unsafe_allow_html=True
    )

    _an = analyze_manuscript(st.session_state.full_sermon)
    if _an["char_count"] < 80:
        st.error("설교 원고가 비어 있습니다. [📤 새 설교 등록/원고작성]에서 원고를 먼저 등록해 주세요.")
    else:
        st.caption(
            f"🔑 원고 자동 추출 키워드: {', '.join(_an['keywords'][:10]) or '(없음)'}　|　"
            f"📖 원고 인용 성구: {', '.join(_an['refs'][:6]) or '(없음)'}"
        )

    with st.expander("💡 설교를 더 풍성하게 — 참고 구절 & 예화", expanded=False):
        if st.button("✨ 참고 성구 및 신학적 예화 생성하기", key="btn_gen_rich"):
            with st.spinner("본문과 연관된 성구·예화 분석 중..."):
                task = """[출력 형식]
1. 📖 본문 연관 참고 성구 3가지
   - 각 항목: 성경 구절(장절) + 구절 내용 + "이 설교 원고의 어느 대목과 어떻게 연결되는지" (원고 문장 인용 필수)
2. 💡 이 설교의 논지를 뒷받침할 현대적 예화 2가지
   - 각 항목: 예화 제목 + 3~5문장 스토리 + 원고의 어떤 주장과 맞물리는지
3. 🏛️ 교회사·기독교 고전 명언 2가지
   - 인물 + 명언 + 이 원고의 어떤 대목을 강화하는지"""
                res = get_ai_response(build_grounded_prompt(task), is_json=False, kind="rich")
                st.session_state.rich_materials = res
                st.rerun()

        if st.session_state.get("rich_materials"):
            show_ai_status()
            render_section_top_toolbar(f"{st.session_state.sermon_title}_참고성구및예화",
                                       st.session_state.rich_materials, "rich_mat")
            render_body(st.session_state.rich_materials)

    with st.expander("🎵 추천 찬양 — 새찬송가 · 복음성가 · CCM", expanded=False):
        if st.button("🎶 맞춤 찬양 15곡 추천받기", key="btn_gen_praise"):
            with st.spinner("설교 메시지와 어울리는 찬양 선곡 중..."):
                task = """[출력 형식 - JSON만]
{"hymns": ["새찬송가 000장 - 제목", ...5곡],
 "gospel_songs": ["복음성가 제목", ...5곡],
 "ccm": ["CCM 제목", ...5곡]}
이 설교 원고의 정서·주제와 실제로 맞는 곡만 고르십시오."""
                res = get_ai_response(build_grounded_prompt(task, ctx_chars=4000), is_json=True, kind="praise")
                st.session_state.praise_list = res
                st.rerun()

        if st.session_state.get("praise_list"):
            show_ai_status()
            p = st.session_state.praise_list
            c1, c2, c3 = st.columns(3)
            for col, key, label in ((c1, "hymns", "📖 새찬송가"), (c2, "gospel_songs", "🕊️ 복음성가"), (c3, "ccm", "🎸 CCM")):
                with col:
                    st.markdown(f"#### {label}")
                    for song in p.get(key, []):
                        q = urllib.parse.quote(song)
                        st.markdown(f"- {song} [🔍](https://www.google.com/search?q={q}) "
                                    f"[▶️](https://www.youtube.com/results?search_query={q})")

    st.write("---")
    left, right = st.columns([1, 2.5])

    MENU = ["설교 요약", "소그룹 나눔", "QT 5일치", "카드뉴스", "쇼츠 대본",
            "🏡 세대별 가정예배지", "🔍 설교 점검 및 제안", "📖 소그룹 리더가이드"]

    with left:
        st.markdown("<p style='font-size:12px;font-weight:bold;color:#94a3b8;'>사역 메뉴 선택</p>",
                    unsafe_allow_html=True)
        sel = st.radio("사역 메뉴", MENU,
                       index=MENU.index(st.session_state.dash_active_view)
                       if st.session_state.dash_active_view in MENU else 0,
                       key="dash_menu_selector", label_visibility="collapsed")
        st.session_state.dash_active_view = sel

    with right:
        view = st.session_state.dash_active_view

        # ---------------- 설교 요약 ----------------
        if view == "설교 요약":
            if not st.session_state.get("sermon_summary_text") or len(st.session_state.sermon_summary_text) < 50:
                st.session_state.sermon_summary_text = build_local_summary(
                    st.session_state.sermon_title, st.session_state.sermon_scripture,
                    st.session_state.full_sermon)

            summary_val = st.session_state.sermon_summary_text
            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교요약",
                                       summary_val, "sermon_sum", ppt_mode="sermon")

            if st.button("✨ AI 강해 요약 생성 (원고 근거 인용 포함)", type="primary", key="btn_regen_ai_summary"):
                with st.spinner("원고를 읽고 대지·근거·적용을 추출하는 중..."):
                    task = """[출력 형식 - 아래 틀 그대로]
🎯 설교 핵심 명제
(이 원고 전체를 관통하는 중심 사상 1~2문장. 원고의 표현을 살려 쓸 것)

🔑 이 설교의 핵심 키워드
(원고에 실제로 반복 등장한 단어 6~8개, 쉼표 구분)

📌 강해적 3대지
1. (대지 제목 — 원고에 나온 표현으로. 상투어 금지) [근거 성구: 원고에 인용된 구절]
   ▸ 원고 근거: "원고에서 그대로 가져온 문장 1~2개"
   ▸ 주해 설명: (그 대목이 말하는 바를 2~3문장)
2. (동일 형식)
3. (동일 형식)

💡 성도를 위한 실천 적용 3가지
1. (원고의 권면을 구체적 행동으로 — 원고에 없는 일반론 금지)
2.
3.

🙏 결단 및 축복 기도문
(원고의 결론·기도 대목을 반영한 3~4문장)"""
                    res = get_ai_response(build_grounded_prompt(task, ctx_chars=11000), is_json=False,
                                          temperature=0.25, kind="summary")
                    st.session_state.sermon_summary_text = res
                    update_sermon_in_db(st.session_state.get("current_sermon_id", 1), updated_summary=res)
                    st.rerun()

            show_ai_status()
            if editable_section("sermon_sum", "sermon_summary_text", "설교 요약 편집",
                                height=400, persist_summary=True):
                render_body(st.session_state.sermon_summary_text)

            with st.expander("📜 설교문 원고 전문 보기", expanded=False):
                st.text(st.session_state.full_sermon)

        # ---------------- 소그룹 나눔 ----------------
        elif view == "소그룹 나눔":
            txt = st.session_state.get("small_group_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_소그룹나눔지", txt, "sm_grp")

            if st.button("✨ 소그룹 나눔지 생성", type="primary", key="btn_gen_sm_grp"):
                with st.spinner("소그룹 나눔지 작성 중..."):
                    task = """[출력 형식]
1. 마음 열기
- [인도자 팁 / 가이드]: (이 설교 주제와 이어지는 도입 멘트)
- 질문 1개 (이 설교의 키워드와 연결될 것)

2. 말씀 속으로
- [인도자 팁 / 가이드]: ...
- 1. (원고의 특정 대목을 직접 인용하며 묻는 질문)
- 2. (원고의 두 번째 대목을 인용하며 묻는 질문)

3. 삶 속으로
- [인도자 팁 / 가이드]: ...
- 1. (원고의 권면을 구체적 상황으로 바꾼 질문)
- 2. (이번 주 결단 질문)

4. 마침 합심 기도문 (원고 내용 반영, 3문장)

※ 모든 질문에는 이 설교 원고에만 나오는 표현이 최소 1개 포함되어야 합니다."""
                    st.session_state.small_group_text = get_ai_response(
                        build_grounded_prompt(task), is_json=False, kind="smallgroup")
                    st.rerun()

            if txt:
                show_ai_status()
                if editable_section("sm_grp", "small_group_text", "소그룹 나눔지 편집"):
                    render_body(st.session_state.small_group_text)
            else:
                st.caption("위 버튼을 눌러 소그룹 나눔지를 생성하세요.")

        # ---------------- QT 5일치 ----------------
        elif view == "QT 5일치":
            txt = st.session_state.get("qt5_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_주간QT5일치", txt, "qt5")

            if st.button("✨ 5일치 QT 묵상지 생성", type="primary", key="btn_gen_qt5"):
                with st.spinner("주간 5일치 QT 작성 중..."):
                    task = """[출력 형식 — 월~금 5일, 각 날짜마다]
📅 (요일): (그날의 소제목 — 원고의 서로 다른 대목에서 뽑을 것)
- 📖 본문 구절: (원고에 인용된 구절 중 하나, 또는 대표 성구)
- 💡 말씀 묵상: (원고의 해당 대목을 풀어 4~5문장)
- 🎯 삶의 적용: (구체적 행동 1가지)
- 🙏 오늘의 기도: (2문장)

※ 5일이 서로 다른 내용이어야 합니다. 같은 말을 다섯 번 반복하지 마십시오."""
                    st.session_state.qt5_text = get_ai_response(
                        build_grounded_prompt(task), is_json=False, kind="qt")
                    st.rerun()

            if txt:
                show_ai_status()
                if editable_section("qt5", "qt5_text", "QT 5일치 편집"):
                    render_body(st.session_state.qt5_text)
            else:
                st.caption("위 버튼을 눌러 5일치 QT를 생성하세요.")

        # ---------------- 카드뉴스 ----------------
        elif view == "카드뉴스":
            h1, h2 = st.columns([1.2, 1.8])
            with h1:
                st.markdown("<h2 style='margin:0;font-size:24px;font-weight:bold;'>카드뉴스</h2>",
                            unsafe_allow_html=True)
            with h2:
                e1, e2, e3 = st.columns([1, 1.3, 1.4])
                with e1:
                    if st.button("✏️ 편집", key="cn_edit_toggle_btn"):
                        st.session_state.cn_edit_mode = not st.session_state.get("cn_edit_mode", False)
                if st.session_state.get("card_list"):
                    cj = json.dumps(st.session_state.card_list, ensure_ascii=False)
                    with e2:
                        st.download_button("📥 PPT 전체",
                                           data=generate_cardnews_pptx_bytes(cj, st.session_state.cn_church_name),
                                           file_name=f"{st.session_state.sermon_title}_카드뉴스.pptx",
                                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                           key="cn_dl_ppt")
                    with e3:
                        st.download_button("📦 전체 PNG",
                                           data=generate_cardnews_zip_bytes(cj, st.session_state.sermon_scripture,
                                                                            st.session_state.cn_church_name),
                                           file_name=f"{st.session_state.sermon_title}_카드뉴스.zip",
                                           mime="application/zip", key="cn_dl_zip")

            o1, o2 = st.columns([1, 1])
            with o1:
                card_count = st.slider("카드 총 장수", 5, 12, 7, key="cn_count_slider")
            with o2:
                st.session_state.cn_church_name = st.text_input(
                    "교회명", value=st.session_state.cn_church_name, key="cn_church_input")

            if st.button(f"🎨 카드뉴스 {card_count}장 생성 / 다시 생성", type="primary", key="btn_gen_cardnews"):
                with st.spinner(f"원고에서 {card_count}장 카드뉴스 구성 중..."):
                    task = f"""[출력 형식 - JSON만]
{{"cards": [{{"card_number": 1, "headline": "...", "body_text": "..."}}, ...]}}

정확히 {card_count}장을 만드십시오. 구성 규칙:
- 1장: 표제 카드. headline은 설교 제목, body_text는 이 설교의 한 줄 요지(원고 문장 기반).
- 2장~{card_count-2}장: 원고의 서로 다른 대목을 하나씩 담습니다.
  headline은 그 대목의 핵심을 12~20자로 (원고 단어 사용), body_text는 60~110자로 원고 내용을 풀어 씁니다.
- {card_count-1}장: 💡 삶의 적용 — 원고의 권면 3가지를 번호로.
- {card_count}장: 🙏 기도 — 원고의 결론/기도를 반영한 2~3문장.

절대 금지: "말씀의 깊은 은혜", "믿음의 결단", "주님의 신실하신 은혜가 충만하기를" 같이
어느 설교에나 붙는 문구로 카드를 채우는 것. 모든 카드에 이 원고 고유의 단어가 들어가야 합니다."""
                    res = get_ai_response(build_grounded_prompt(task, ctx_chars=9000), is_json=True,
                                          temperature=0.3, kind="cards", card_count=card_count)
                    cards = (res or {}).get("cards", [])
                    if cards:
                        for i, c in enumerate(cards, start=1):
                            c["card_number"] = i
                        st.session_state.card_list = cards[:card_count]
                        st.session_state.cn_card_idx = 0
                    st.rerun()

            show_ai_status()

            if st.session_state.get("cn_edit_mode") and st.session_state.get("card_list"):
                st.markdown("#### ✏️ 카드 텍스트 편집기")
                for i, c in enumerate(st.session_state.card_list):
                    with st.expander(f"CARD {c.get('card_number', i+1)}", expanded=(i == 0)):
                        c["headline"] = st.text_input("헤드라인", value=c.get("headline", ""), key=f"cn_h_{i}")
                        c["body_text"] = st.text_area("본문", value=c.get("body_text", ""), key=f"cn_b_{i}")

            if st.session_state.get("card_list"):
                cards = st.session_state.card_list
                total = len(cards)
                st.session_state.setdefault("cn_card_idx", 0)
                idx = st.session_state.cn_card_idx % total

                n1, n2, n3 = st.columns([1, 4, 1])
                with n1:
                    if st.button("❮ 이전", key="cn_prev"):
                        st.session_state.cn_card_idx = (idx - 1) % total
                        st.rerun()
                with n3:
                    if st.button("다음 ❯", key="cn_next"):
                        st.session_state.cn_card_idx = (idx + 1) % total
                        st.rerun()
                with n2:
                    png = generate_single_card_png_bytes(
                        json.dumps(cards[idx], ensure_ascii=False), idx,
                        st.session_state.sermon_scripture, st.session_state.cn_church_name)
                    st_image_full(png, caption=f"{idx+1} / {total} — 실제 다운로드 결과와 동일")
                    st.download_button(f"🖼️ CARD {idx+1} PNG 다운로드", data=png,
                                       file_name=f"{st.session_state.sermon_title}_card_{idx+1}.png",
                                       mime="image/png", key=f"dl_card_{idx}")

                st.write("---")
                st.markdown("#### 인스타그램 캡션")
                _a = analyze_manuscript(st.session_state.full_sermon)
                cap = (f"[{st.session_state.sermon_title}] ({st.session_state.sermon_scripture})\n\n"
                       f"{(_a['key_sentences'][0] if _a['key_sentences'] else '')}\n\n"
                       + " ".join(["#주일설교", "#말씀묵상"] + [f"#{k}" for k in _a['keywords'][:6]]))
                st.code(cap, language="text")
            else:
                st.caption("위 버튼을 눌러 카드뉴스를 생성하세요.")

        # ---------------- 쇼츠 대본 ----------------
        elif view == "쇼츠 대본":
            txt = st.session_state.get("shorts_script_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_쇼츠대본", txt, "sh_script")

            if st.button("🎬 쇼츠 대본 3종 생성", type="primary", key="btn_gen_shorts_script"):
                with st.spinner("쇼츠 대본 작성 중..."):
                    task = """[출력 형식 — 60초 세로 쇼츠 대본 3종]
각 대본마다:
🎬 (유형명)
- 후킹(0~5초): (원고의 가장 강력한 문장에서 따온 한 마디)
- 본론(5~45초): (원고 내용 기반 3~4문장, 화면 자막 단위로 줄바꿈)
- 결단(45~60초): (한 문장 권면)
- 화면 자막 키워드: (5개)

3종 유형: 1) 감동·위로형  2) 질문·호기심형  3) 결단 선포형
세 대본은 원고의 서로 다른 대목을 사용해야 합니다."""
                    st.session_state.shorts_script_text = get_ai_response(
                        build_grounded_prompt(task), is_json=False, kind="shorts_script")
                    st.rerun()

            if txt:
                show_ai_status()
                if editable_section("sh_script", "shorts_script_text", "쇼츠 대본 편집"):
                    render_body(st.session_state.shorts_script_text)
            else:
                st.caption("위 버튼을 눌러 쇼츠 대본을 생성하세요.")

        # ---------------- 가정예배지 ----------------
        elif view == "🏡 세대별 가정예배지":
            age = st.selectbox("예배 대상 선택", ["👶 영유아용", "🧒 어린이용", "🧑 청소년용", "👨‍👩‍👧 청장년용"],
                               key="sel_age_group")
            fkey = f"family_worship_{age}"
            txt = st.session_state.get(fkey, "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_가정예배지_{age}", txt, f"fam_{age}")

            if st.button(f"✨ {age} 맞춤 가정예배지 생성", type="primary", key="btn_gen_fam"):
                with st.spinner(f"{age} 가정예배지 작성 중..."):
                    task = f"""[출력 형식 — 대상: {age}]
1. 찬양 및 신앙고백
- [인도자 팁 / 가이드]: (이 설교 주제에 맞는 찬송 제안과 시작 멘트)
2. 함께 읽는 성경 말씀
- [인도자 팁 / 가이드]: ...
3. {age} 눈높이 3분 가족 메시지
- (원고의 핵심 내용을 이 연령대 언어로. 원고의 예화·표현을 반드시 활용)
4. 온 가족 나눔 질문 2가지
- [인도자 팁 / 가이드]: ...
5. 가정을 축복하는 마무리 기도문

※ {age}의 이해 수준에 맞춘 어휘를 쓰되, 내용은 반드시 이 설교 원고에서 나와야 합니다."""
                    st.session_state[fkey] = get_ai_response(
                        build_grounded_prompt(task), is_json=False, kind="family")
                    st.rerun()

            if txt:
                show_ai_status()
                if editable_section(f"fam_{age}", fkey, "가정예배지 편집", height=320):
                    render_body(st.session_state[fkey])
            else:
                st.caption(f"위 버튼을 눌러 {age} 맞춤 가정예배지를 생성하세요.")

        # ---------------- 설교 점검 ----------------
        elif view == "🔍 설교 점검 및 제안":
            txt = st.session_state.get("sermon_audit_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_설교점검및제안", txt, "sermon_audit")

            if st.button("🔍 설교 정밀 점검", type="primary", key="btn_gen_audit"):
                with st.spinner("설교 분석 리포트 작성 중..."):
                    task = """[출력 형식 — 냉정하고 구체적으로. 칭찬만 나열하지 말 것]
1. 📖 본문 주해의 정확성 (점수/100 + 근거)
   - 잘된 점: 원고 문장 인용 + 이유
   - 아쉬운 점: 원고 문장 인용 + 구체적 수정 제안
2. 🏗️ 대지 전개와 구조 (점수/100 + 근거)
   - 실제 대지 구조를 요약한 뒤, 논리 비약이나 중복이 있으면 지적
3. 💡 예화·적용의 적절성 (점수/100)
   - 원고의 예화를 열거하고, 본문과의 연결 강도를 평가
4. 🎙️ 전달력 개선 제안 3가지 (원고의 특정 문장을 지목해서)
5. 📊 종합 총평 + 다음 설교를 위한 권고 3가지

※ 원고에 실제로 없는 내용을 지적하거나 칭찬하지 마십시오."""
                    st.session_state.sermon_audit_text = get_ai_response(
                        build_grounded_prompt(task, ctx_chars=12000), is_json=False,
                        temperature=0.3, kind="audit")
                    st.rerun()

            if txt:
                show_ai_status()
                if editable_section("sermon_audit", "sermon_audit_text", "설교 점검 편집"):
                    render_body(st.session_state.sermon_audit_text)
            else:
                st.caption("위 버튼을 눌러 설교 점검 리포트를 생성하세요.")

        # ---------------- 리더가이드 ----------------
        elif view == "📖 소그룹 리더가이드":
            txt = st.session_state.get("leader_guide_text", "")
            render_section_top_toolbar(f"{st.session_state.sermon_title}_소그룹리더가이드", txt, "ldr_guide")

            if st.button("📖 리더가이드 생성", type="primary", key="btn_gen_ldr_guide"):
                with st.spinner("소그룹 인도자 가이드 작성 중..."):
                    task = """[출력 형식]
1. 🎯 이번 주 모임의 핵심 목표
- [인도자 팁 / 가이드]: ...
2. 📖 본문 배경 및 신학적 핵심 해설 (리더 심화용)
- [인도자 팁 / 가이드]: (성도가 던질 만한 신학적 질문 2개와 답변 방향)
3. 💬 나눔 질문별 예상 답변 & 피드백 요령
- [인도자 팁 / 가이드]: ...
4. ⚠️ 침묵·돌발 상황 대처
- [인도자 팁 / 가이드]: ...
5. 🙏 소그룹 맞춤 중보기도 제목 3가지

※ 2번 해설은 반드시 이 설교의 본문과 원고 내용에 근거해야 합니다."""
                    st.session_state.leader_guide_text = get_ai_response(
                        build_grounded_prompt(task), is_json=False, kind="leader")
                    st.rerun()

            if txt:
                show_ai_status()
                if editable_section("ldr_guide", "leader_guide_text", "리더가이드 편집"):
                    render_body(st.session_state.leader_guide_text)
            else:
                st.caption("위 버튼을 눌러 소그룹 리더가이드를 생성하세요.")


# ==============================================================================
# 2) 새 설교 등록 / 원고작성
# ==============================================================================
elif app_mode == "📤 새 설교 등록/원고작성":
    st.markdown("<h1 style='font-size:28px;font-weight:800;'>📤 새 설교 등록 및 원고 작성</h1>",
                unsafe_allow_html=True)
    st.caption("직접 타이핑 · 파일 업로드 · AI 강해설교문 자동 생성으로 새 설교를 등록합니다.")

    t1, t2, t3 = st.tabs(["✍️ 직접 타이핑 작성",
                          "📁 파일 업로드 (.docx/.pdf/.txt)",
                          "📖 AI 강해설교문 생성"])

    with t1:
        c1, c2, c3 = st.columns([2, 1.5, 1])
        with c1:
            t_title = st.text_input("설교 제목", placeholder="예: 광야에서 만나는 하나님의 은혜", key="type_title")
        with c2:
            t_scrip = st.text_input("성경 본문", placeholder="예: 출애굽기 16:1-12", key="type_scrip")
        with c3:
            st.text_input("설교자", value=st.session_state.preacher_name, key="type_preach")

        t_tags = st.text_input("태그 (쉼표 구분)", value="주일설교", key="type_tags")
        t_content = st.text_area("설교문 본문 전문", height=380, key="type_text")
        st.caption(f"글자 수: **{len(t_content):,}자**")

        if st.button("💾 새 설교로 등록", type="primary", key="save_type_sermon"):
            if not t_title.strip() or len(t_content.strip()) < 50:
                st.warning("설교 제목과 본문(50자 이상)을 입력해주세요.")
            else:
                testament, book = classify_scripture(t_scrip.strip())
                entry = {"title": t_title.strip(), "scripture": t_scrip.strip(),
                         "testament": testament, "book": book,
                         "topic": (t_tags.split(",")[0].strip() if t_tags else "일반설교"),
                         "theology": "직접작성", "date": datetime.now().strftime("%Y-%m-%d"),
                         "tags": [x.strip() for x in t_tags.split(",") if x.strip()],
                         "summary": "", "text": t_content.strip()}
                saved = add_sermon_to_db(entry)
                load_sermon_to_workspace(saved, idx=len(st.session_state.sermon_library) - 1)
                st.success(f"'{t_title}' 등록 완료! 사이드바에서 [📊 설교 대시보드]로 이동하세요.")
                st.rerun()

    with t2:
        u_file = st.file_uploader("설교 파일 선택", type=["docx", "pdf", "txt"], key="up_sermon_file")
        f_title = st.text_input("설교 제목", value="", placeholder="업로드 설교 제목", key="up_title")
        f_scrip = st.text_input("성경 본문", value="", placeholder="예: 로마서 8:28-39", key="up_scrip")

        if u_file and st.button("📂 파일 읽어와서 등록", type="primary", key="save_up_sermon"):
            text = ""
            fn = u_file.name.lower()
            try:
                if fn.endswith('.txt'):
                    text = u_file.read().decode('utf-8', errors='ignore')
                elif fn.endswith('.docx'):
                    d = Document(u_file)
                    text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
                elif fn.endswith('.pdf'):
                    pdf = PdfReader(u_file)
                    text = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
            except Exception as e:
                st.error(f"파일 읽기 실패: {e}")
                text = ""

            if len(text.strip()) < 50:
                st.warning("본문을 읽지 못했습니다. 파일을 확인해 주세요.")
            else:
                title = f_title.strip() or os.path.splitext(u_file.name)[0]
                scrip = f_scrip.strip() or (extract_scripture_refs(text)[:1] or ["본문 미지정"])[0]
                testament, book = classify_scripture(scrip)
                entry = {"title": title, "scripture": scrip, "testament": testament, "book": book,
                         "topic": "파일등록", "theology": "파일업로드",
                         "date": datetime.now().strftime("%Y-%m-%d"), "tags": ["파일등록"],
                         "summary": "", "text": text.strip()}
                saved = add_sermon_to_db(entry)
                load_sermon_to_workspace(saved, idx=len(st.session_state.sermon_library) - 1)
                st.success(f"등록 완료! (본문 자동 인식: {scrip}) [📊 설교 대시보드]로 이동하세요.")
                st.rerun()

    with t3:
        a1, a2, a3 = st.columns([1.2, 1.5, 1.3])
        with a1:
            sel_book = st.selectbox("성경 66권", BIBLE_BOOKS, index=44, key="sel_ai_book")
        with a2:
            sel_cv = st.text_input("장·절", value="8장 28절~39절", key="sel_ai_cv")
        with a3:
            theology = st.selectbox("신학적 관점",
                                    ["개혁주의 (Reformed - 칼빈주의/하나님 주권)",
                                     "장로교 정통 (Presbyterian - 웨스트민스터/구속사)",
                                     "복음주의 (Evangelical - 십자가/은혜/복음선포)"],
                                    key="sel_ai_theology")
        b1, b2 = st.columns([2, 1])
        with b1:
            topic = st.text_input("설교 주제 / 강조 포인트",
                                  value="고난 속에서도 흔들리지 않는 하나님의 사랑과 구원의 확신",
                                  key="sel_ai_topic")
        with b2:
            style = st.selectbox("설교 형태", ["3대지 본문중심 강해설교", "구속사적 복음설교", "원어 주해 중심 강해설교"],
                                 key="sel_ai_style")

        full_scrip = f"{sel_book} {sel_cv}"

        if st.button("🚀 강해설교문 전문 작성 (25~30분 분량)", type="primary", key="btn_gen_ai_sermon"):
            if not get_resolved_api_key():
                st.error("이 기능은 AI 생성 전용입니다. 사이드바 [⚙️ AI 연결 설정]에서 Gemini API 키를 먼저 등록해 주세요.")
            else:
                with st.spinner(f"[{theology.split(' ')[0]}] 관점으로 강해설교문 작성 중... (1~2분 소요)"):
                    prompt = f"""당신은 한국 장로교 강단에서 30년간 설교해 온 목회자입니다.
아래 조건으로 실제 강단에서 그대로 선포할 수 있는 설교 원고 전문을 작성하십시오.

[조건]
- 성경 본문: {full_scrip}
- 설교 주제: {topic}
- 신학적 관점: {theology}
- 설교 형태: {style}
- 분량: 한국어 6,000자 이상 (25~30분 선포 분량)

[반드시 지킬 구조]
1. 제목
2. 본문 봉독 안내 ({full_scrip})
3. 서론 — 청중의 삶에서 출발하는 구체적 도입 (실제 있을 법한 상황 묘사)
4. 본문의 역사적·문학적 배경 설명
5. 제1대지 — (제목) / 본문 주해 / 예화 / 적용
6. 제2대지 — (제목) / 본문 주해 / 예화 / 적용
7. 제3대지 — (제목) / 본문 주해 / 예화 / 적용
8. 결론 — 메시지 요약과 결단 촉구
9. 마침 기도

[작성 규칙]
- 각 대지는 반드시 {full_scrip} 본문의 특정 구절을 인용하고 주해할 것.
- 예화는 구체적이고 현실적으로. "어떤 성도가 있었습니다" 식의 막연한 예화 금지.
- '은혜가 충만하기를' 같은 상투적 문구로 분량을 채우지 말 것.
- 100% 한국어. 영어 메모나 머리말 없이 설교 원고 본문만 출력.
- 마크다운 기호(#, **) 없이 평문으로 작성."""
                    res = get_ai_response(prompt, is_json=False, temperature=0.7, kind="sermon_write")
                    st.session_state.temp_generated_sermon = res
                    st.session_state.temp_ai_title = f"{sel_book} 강해: {topic[:30]}"
                    st.session_state.temp_ai_scrip = full_scrip
                    st.rerun()

        if st.session_state.get("temp_generated_sermon"):
            st.write("---")
            show_ai_status()
            st.caption(f"생성 분량: {len(st.session_state.temp_generated_sermon):,}자")
            render_section_top_toolbar(st.session_state.temp_ai_title,
                                       st.session_state.temp_generated_sermon, "ai_gen_sermon")
            edited = st.text_area("작성된 강해설교문 검토 및 수정",
                                  value=st.session_state.temp_generated_sermon,
                                  height=450, key="edit_ai_sermon_area")

            if st.button("✅ 서재와 대시보드에 최종 등록", type="primary", key="btn_save_ai_sermon"):
                testament, book = classify_scripture(st.session_state.temp_ai_scrip)
                entry = {"title": st.session_state.temp_ai_title,
                         "scripture": st.session_state.temp_ai_scrip,
                         "testament": testament, "book": book, "topic": sel_book,
                         "theology": theology.split(' ')[0],
                         "date": datetime.now().strftime("%Y-%m-%d"),
                         "tags": [sel_book, theology.split(' ')[0], "강해설교"],
                         "summary": "", "text": edited}
                saved = add_sermon_to_db(entry)
                load_sermon_to_workspace(saved, idx=len(st.session_state.sermon_library) - 1)
                st.success("등록 완료! [📊 설교 대시보드]로 이동하세요.")
                st.rerun()


# ==============================================================================
# 3) AI 보이스오버 스튜디오
# ==============================================================================
elif app_mode == "🎙️ AI 보이스오버 스튜디오":
    st.markdown("<h1 style='font-size:28px;font-weight:800;'>🎙️ AI 보이스오버 스튜디오</h1>",
                unsafe_allow_html=True)

    if not HAS_TTS:
        st.error("edge-tts 패키지가 설치되어 있지 않습니다. `pip install edge-tts` 후 다시 실행해 주세요.")

    c1, c2 = st.columns([1.5, 1])
    with c1:
        src = st.radio("변환할 원문", ["설교 원고 전문", "설교 요약", "직접 입력"], horizontal=True, key="vo_src")
        if src == "설교 원고 전문":
            default_txt = st.session_state.full_sermon
        elif src == "설교 요약":
            default_txt = st.session_state.get("sermon_summary_text", "")
        else:
            default_txt = ""

        vo_text = st.text_area("음성 변환 텍스트", value=default_txt, height=330, key=f"vo_text_{src}")
        vo_voice = st.selectbox("성우 보이스", ["인준 (남성 - 차분하고 신뢰감 있는 톤)", "선희 (여성 - 맑고 또렷한 톤)"],
                                key="vo_voice_sel")
        voice_id = "ko-KR-InJoonNeural" if "인준" in vo_voice else "ko-KR-SunHiNeural"

        st.caption(f"글자 수 {len(vo_text):,}자 · 예상 길이 약 {max(1, len(vo_text)//330)}분")

        if st.button("🎙️ 고음질 음성(TTS) 생성", type="primary", key="btn_gen_tts", disabled=not HAS_TTS):
            if not vo_text.strip():
                st.warning("텍스트를 입력해주세요.")
            else:
                with st.spinner("AI 음성 생성 중..."):
                    try:
                        st.session_state.vo_audio_path = generate_voiceover_audio(vo_text, voice_id)
                        st.success("보이스오버 생성 완료!")
                    except Exception as e:
                        st.error(f"음성 생성 실패: {e}")

    with c2:
        st.markdown("### 🎧 플레이어 & 다운로드")
        p = st.session_state.get("vo_audio_path")
        if p and os.path.exists(p):
            st.audio(p)
            with open(p, "rb") as af:
                st.download_button("📥 MP3 다운로드", data=af.read(),
                                   file_name=f"{st.session_state.sermon_title}_voice.mp3",
                                   mime="audio/mp3", key="dl_vo_mp3")
        else:
            st.info("왼쪽에서 생성하면 이곳에 플레이어가 나타납니다.")


# ==============================================================================
# 4) 쇼츠 만들기
# ==============================================================================
elif app_mode == "🎬 쇼츠 만들기 (스튜디오)":
    st.markdown("<h1 style='font-size:28px;font-weight:800;'>▶️ 쇼츠 만들기 스튜디오</h1>",
                unsafe_allow_html=True)

    if not HAS_VIDEO_ENGINE:
        st.warning("`video_engine.py` 모듈이 없어 영상 렌더링 기능이 비활성화됩니다. 같은 폴더에 파일을 두고 다시 실행해 주세요.")

    tab_yt, tab_ai = st.tabs(["🔗 유튜브 링크에서 숏츠 추출", "🎨 AI 나레이션 & 템플릿 숏츠 제작"])

    with tab_yt:
        st.caption("클라우드 서버 IP는 유튜브가 자주 차단합니다. 실패하면 오른쪽 탭에서 영상을 직접 업로드하세요.")
        yt_url = st.text_input("유튜브 영상 링크", placeholder="https://www.youtube.com/watch?v=...", key="yt_url_input")
        y1, y2, y3 = st.columns(3)
        with y1:
            s_min = st.number_input("시작 분", 0, 300, 12, key="yt_s_min")
        with y2:
            s_sec = st.number_input("시작 초", 0, 59, 30, key="yt_s_sec")
        with y3:
            dur = st.slider("길이(초)", 15, 60, 45, key="yt_dur")

        yt_title = st.text_input("상단 헤드라인", value=st.session_state.sermon_title, key="yt_title")
        yt_sub = st.text_input("강조 자막", value="", key="yt_sub")
        yt_church = st.text_input("교회명 워터마크", value=st.session_state.cn_church_name, key="yt_church")

        if st.button("🚀 9:16 세로 숏츠 추출", type="primary", key="btn_yt_extract"):
            if not yt_url.strip():
                st.warning("유튜브 링크를 입력해주세요.")
            else:
                with st.spinner("영상 추출 중... (20~40초)"):
                    try:
                        st.session_state.yt_extracted_result = extract_youtube_to_shorts(
                            yt_url.strip(), s_min * 60 + s_sec, dur, yt_title, yt_sub, yt_church)
                        st.success("추출 완료!")
                    except Exception as e:
                        st.warning(str(e))

        r = st.session_state.get("yt_extracted_result")
        if r and os.path.exists(r):
            st.write("---")
            v1, v2 = st.columns(2)
            with v1:
                st.video(r)
            with v2:
                with open(r, "rb") as f:
                    st.download_button("📥 MP4 다운로드", data=f.read(),
                                       file_name=f"{yt_title}_shorts.mp4", mime="video/mp4", key="dl_yt")

    with tab_ai:
        s1, s2, s3 = st.columns(3)
        with s1:
            st.link_button("🎥 픽사베이 영상", "https://pixabay.com/ko/videos/")
        with s2:
            st.link_button("📸 펙셀스 비디오", "https://www.pexels.com/ko-kr/videos/")
        with s3:
            st.link_button("🎵 무료 BGM", "https://pixabay.com/ko/music/")

        st.write("---")
        if st.button("✨ 쇼츠 제목 5개 & 해시태그 추천", key="btn_gen_shorts_meta"):
            with st.spinner("제목 분석 중..."):
                task = """[출력 형식 - JSON만]
{"titles": ["1. ...", "2. ...", "3. ...", "4. ...", "5. ..."], "hashtags": ["#...", ...8개]}
제목은 이 설교 원고의 실제 문장·키워드에서 뽑되, 25자 이내로 클릭을 부르게 다듬으십시오."""
                st.session_state.shorts_rec = get_ai_response(
                    build_grounded_prompt(task, ctx_chars=4000), is_json=True, kind="shorts_meta")
                st.rerun()

        selected_title = st.session_state.sermon_title
        if st.session_state.get("shorts_rec"):
            show_ai_status()
            rec = st.session_state.shorts_rec
            tc = st.radio("추천 제목", rec.get("titles", []), key="rad_shorts_title")
            if tc:
                selected_title = re.sub(r"^\d+\.\s*", "", tc)
            st.markdown("**추천 태그:** " + " ".join(rec.get("hashtags", [])))

        st.write("---")
        col_a, col_b = st.columns([1.2, 1])
        with col_a:
            v_title = st.text_input("쇼츠 제목", value=selected_title, key="in_shorts_title")
            _a = analyze_manuscript(st.session_state.full_sermon)
            default_script = "\n".join(_a["key_sentences"][:3]) or "말씀을 붙들고 이번 한 주를 살아갑시다."
            v_script = st.text_area("자막 대본 (줄바꿈 구분)", value=default_script, height=130, key="in_shorts_script")

            r1, r2 = st.columns(2)
            with r1:
                v_ratio = st.radio("비율", ["9:16 (세로)", "16:9 (가로)"], key="rad_ratio")
            with r2:
                v_voice = st.selectbox("보이스", ["인준 (남성)", "선희 (여성)"], key="sel_shorts_voice")

            with st.expander("🎨 폰트 크기 · 자막 위치", expanded=False):
                f1, f2 = st.columns(2)
                with f1:
                    t_fsize = st.slider("제목 폰트(pt)", 32, 72, 48, 2, key="sh_t_fsize")
                    t_ypos = st.slider("제목 Y 위치", 80, 500, 180, 10, key="sh_t_ypos")
                with f2:
                    s_fsize = st.slider("자막 폰트(pt)", 24, 60, 42, 2, key="sh_s_fsize")
                    s_ypos = st.slider("자막 Y 위치", 800, 1700, 1400, 20, key="sh_s_ypos")

            bg_media = st.file_uploader("배경 동영상/사진", type=["mp4", "mov", "jpg", "png"], key="up_shorts_bg")
            bgm_media = st.file_uploader("배경음악", type=["mp3", "wav"], key="up_shorts_bgm")

            if st.button("🚀 비디오 렌더링 시작", type="primary", key="btn_render_video",
                         disabled=not HAS_VIDEO_ENGINE):
                with st.spinner("자막 애니메이션 · BGM 믹싱 중..."):
                    try:
                        os.makedirs("./uploads", exist_ok=True)
                        bg_p = bgm_p = None
                        if bg_media:
                            bg_p = f"./uploads/{bg_media.name}"
                            with open(bg_p, "wb") as f:
                                f.write(bg_media.getbuffer())
                        if bgm_media:
                            bgm_p = f"./uploads/{bgm_media.name}"
                            with open(bgm_p, "wb") as f:
                                f.write(bgm_media.getbuffer())

                        lines = [l.strip() for l in v_script.split("\n") if l.strip()]
                        st.session_state.rendered_shorts_out = create_animated_video(
                            title=v_title, script_paragraphs=lines, bg_media_path=bg_p, bgm_path=bgm_p,
                            aspect_ratio=("9:16" if "9:16" in v_ratio else "16:9"),
                            voice=("ko-KR-InJoonNeural" if "인준" in v_voice else "ko-KR-SunHiNeural"),
                            title_fontsize=t_fsize, sub_fontsize=s_fsize, title_y=t_ypos, sub_y=s_ypos)
                        st.success("렌더링 완료!")
                    except Exception as e:
                        st.error(f"렌더링 실패: {e}")

        with col_b:
            st.markdown("### 🎬 완성된 영상")
            out = st.session_state.get("rendered_shorts_out")
            if out and os.path.exists(out):
                st.video(out)
                with open(out, "rb") as vf:
                    st.download_button("📥 MP4 다운로드", data=vf.read(), file_name="sermon_shorts.mp4",
                                       mime="video/mp4", key="dl_shorts_mp4")
            else:
                st.info("렌더링하면 이곳에 영상이 나타납니다.")


# ==============================================================================
# 5) 말씀카드 이미지
# ==============================================================================
elif app_mode == "📷 말씀카드 이미지":
    st.markdown("<h1 style='font-size:28px;font-weight:800;'>📷 말씀카드 이미지 스튜디오</h1>",
                unsafe_allow_html=True)

    _a = analyze_manuscript(st.session_state.full_sermon)
    candidates = _a["key_sentences"][:8] or [st.session_state.sermon_title]

    c1, c2 = st.columns([1.3, 1.2])
    with c1:
        st.markdown("#### ✏️ 문구 선택 (원고에서 자동 추출)")
        pick = st.selectbox("원고 속 핵심 문장에서 고르기",
                            ["(직접 입력)"] + [s[:70] + ("…" if len(s) > 70 else "") for s in candidates],
                            key="vc_pick")
        if pick != "(직접 입력)":
            idx = ["(직접 입력)"] + candidates
            init_text = candidates[[s[:70] + ("…" if len(s) > 70 else "") for s in candidates].index(pick)]
        else:
            init_text = st.session_state.sermon_title

        v_text = st.text_area("말씀문구 / 메시지", value=init_text, height=120, key=f"vc_text_{pick}")
        v_scrip = st.text_input("성경 구절", value=st.session_state.sermon_scripture, key="vc_scrip_in")
        v_church = st.text_input("교회명 배지", value=st.session_state.cn_church_name, key="vc_church_in")

        st.markdown("#### 🎨 배경 & 스타일")
        b1, b2 = st.columns(2)
        with b1:
            bg_opt = st.radio("배경", ["사진", "기본", "직접 업로드"], key="vc_bg_radio")
            bg_i = st.slider("배경 사진 번호", 0, len(CARD_BACKGROUNDS) - 1, 0, key="vc_bg_idx")
        with b2:
            up_file = st.file_uploader("배경 이미지", type=["jpg", "png", "jpeg"], key="vc_up_file") \
                if bg_opt == "직접 업로드" else None

        s1, s2 = st.columns(2)
        with s1:
            fsize = st.slider("폰트 크기", 28, 68, 42, 2, key="vc_fsize")
            lspace = st.slider("줄간격", 10, 40, 18, 2, key="vc_lspace")
        with s2:
            fcolor = st.color_picker("글자 색", "#FDE047", key="vc_fcolor")
            scolor = st.color_picker("테두리 색", "#000000", key="vc_scolor")
            opacity = st.slider("배경 어둡기", 0.2, 0.9, 0.6, 0.05, key="vc_op")

    with c2:
        st.markdown("### 🖼️ 미리보기 & 다운로드")
        png = generate_verse_card_png(v_text, v_scrip, bg_opt, up_file, fsize, lspace,
                                      fcolor, scolor, opacity, v_church, bg_index=bg_i)
        st_image_full(png.getvalue(), caption="1:1 고화질 말씀카드")
        st.download_button("📥 PNG 다운로드", data=png.getvalue(),
                           file_name=f"{st.session_state.sermon_title}_말씀카드.png",
                           mime="image/png", key="dl_verse_card")


# ==============================================================================
# 6) 설교 서재
# ==============================================================================
elif app_mode == "📚 설교 서재 (Sermon Library)":
    st.markdown("<h1 style='font-size:28px;font-weight:800;'>📚 설교 서재 (영구 기록보관소)</h1>",
                unsafe_allow_html=True)

    sermons_db = get_db_sermons()
    st.session_state.sermon_library = sermons_db

    t1, t2 = st.columns([1.5, 1.5])
    with t1:
        st.markdown(f"총 **{len(sermons_db):,}편**의 설교문이 보관되어 있습니다.")
    with t2:
        k1, k2 = st.columns(2)
        with k1:
            st.download_button("💾 전체 백업(.json)",
                               data=json.dumps(sermons_db, ensure_ascii=False, indent=2).encode('utf-8'),
                               file_name=f"설교서재_백업_{datetime.now().strftime('%Y%m%d')}.json",
                               mime="application/json", key="dl_backup")
        with k2:
            with st.popover("📂 백업 복원"):
                rf = st.file_uploader("백업 JSON", type=["json"], key="up_restore")
                mode = st.radio("복원 방식", ["기존에 병합(추가)", "완전 교체"], key="restore_mode")
                if rf and st.button("✅ 복원 실행", key="btn_restore"):
                    try:
                        data = json.load(rf)
                        if isinstance(data, list):
                            if mode == "기존에 병합(추가)":
                                exist = {(s.get("title"), s.get("scripture")) for s in sermons_db}
                                merged = list(sermons_db)
                                nid = max([int(s.get("id", 0)) for s in sermons_db] or [0])
                                for s in data:
                                    if (s.get("title"), s.get("scripture")) not in exist:
                                        nid += 1
                                        s["id"] = nid
                                        merged.append(s)
                                data = merged
                            save_db_sermons(data)
                            st.session_state.sermon_library = data
                            st.success(f"복원 완료! 현재 {len(data)}편")
                            st.rerun()
                        else:
                            st.error("올바른 백업 파일이 아닙니다.")
                    except Exception as e:
                        st.error(f"복원 실패: {e}")

    st.write("---")
    f1, f2, f3, f4 = st.columns([1.5, 1.2, 1.5, 1.2])
    with f1:
        kw = st.text_input("검색어 (제목/본문/키워드)", key="lib_kw")
    with f2:
        tf = st.selectbox("구약/신약", ["전체", "구약", "신약"], key="lib_testament")
    with f3:
        bf = st.selectbox("성경 66권", ["전체"] + BIBLE_BOOKS, key="lib_book")
    with f4:
        so = st.selectbox("정렬", ["최신순", "오래된순"], key="lib_sort")

    st.write("---")
    filtered = []
    for s in sermons_db:
        if "testament" not in s or "book" not in s:
            s["testament"], s["book"] = classify_scripture(s.get("scripture", ""))
        if kw:
            hay = " ".join([s.get("title", ""), s.get("scripture", ""), s.get("text", "")]
                           + list(s.get("tags", []))).lower()
            if kw.lower() not in hay:
                continue
        if tf != "전체" and s.get("testament") != tf:
            continue
        if bf != "전체" and s.get("book") != bf:
            continue
        filtered.append(s)

    filtered.sort(key=lambda x: x.get("id", 0), reverse=(so == "최신순"))
    st.markdown(f"**검색 결과:** `{len(filtered):,}편`")

    if not filtered:
        st.info("조건에 맞는 설교문이 없습니다.")
    else:
        for i, s in enumerate(filtered):
            tags_html = ' '.join(
                f'<span style="background:#1e293b;color:#38bdf8;padding:3px 10px;border-radius:6px;'
                f'font-size:12px;margin-right:4px;">#{t}</span>' for t in s.get('tags', []))
            st.markdown(
                f"""<div style="background:#0f172a;border:1px solid #334155;border-radius:14px;padding:20px;margin-bottom:10px;">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
                <h3 style="margin:0;font-size:20px;font-weight:bold;color:#f8fafc;">{s.get('title')}</h3>
                <span style="background:#1e3a8a;color:#fde047;padding:4px 12px;border-radius:12px;font-size:12px;font-weight:bold;">
                {s.get('testament','기타')} · {s.get('book','성경')}</span></div>
                <p style="margin:0 0 10px 0;font-size:14px;color:#94a3b8;">
                📖 <strong>{s.get('scripture')}</strong> · [{s.get('theology','-')}] · 📅 {s.get('date','-')}
                · 📝 {len(s.get('text','')):,}자</p>
                <div style="margin-bottom:6px;">{tags_html}</div></div>""",
                unsafe_allow_html=True
            )
            c1, c2 = st.columns([3, 1])
            with c1:
                if st.button("📖 대시보드로 불러오기", key=f"lib_load_{s.get('id')}_{i}"):
                    try:
                        real_idx = sermons_db.index(s)
                    except ValueError:
                        real_idx = 0
                    load_sermon_to_workspace(s, idx=real_idx)
                    st.success(f"'{s.get('title')}' 로 전체 메뉴가 동기화되었습니다.")
                    st.rerun()
            with c2:
                if st.button("🗑️ 삭제", key=f"lib_del_{s.get('id')}_{i}"):
                    updated = [x for x in get_db_sermons() if x.get('id') != s.get('id')]
                    save_db_sermons(updated)
                    st.session_state.sermon_library = updated
                    st.success("삭제되었습니다.")
                    st.rerun()
