import streamlit as st
import google.generativeai as genai
import json
import os
import io
import re
from docx import Document
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from video_engine import create_animated_video

st.set_page_config(page_title="MY 설교 AI 스튜디오 Pro", page_icon="🕊️", layout="wide")

# 1. 개인 보안 접속 인증
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

# 2. Gemini AI 안전 호출 래퍼 (NotFound 에러 완벽 방지)
GEMINI_KEY = st.secrets.get("GEMINI_API_KEY", "")
if not GEMINI_KEY:
    GEMINI_KEY = st.sidebar.text_input("Gemini API Key를 입력하세요", type="password")

def get_ai_response(prompt: str, is_json: bool = True):
    if not GEMINI_KEY:
        st.error("Gemini API Key가 등록되지 않았습니다.")
        return None
    
    genai.configure(api_key=GEMINI_KEY)
    
    # 404 NotFound 방지를 위한 우선순위 모델 리스트
    candidate_models = ["gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-1.5-pro", "gemini-pro"]
    
    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            if is_json:
                res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                return json.loads(res.text)
            else:
                res = model.generate_content(prompt)
                return res.text
        except Exception as e:
            # JSON 파싱 실패 시 정규식으로 JSON 블록 추출 시도
            try:
                model = genai.GenerativeModel(model_name)
                res = model.generate_content(prompt)
                if is_json:
                    match = re.search(r"\{.*\}|\[.*\]", res.text, re.DOTALL)
                    if match:
                        return json.loads(match.group(0))
                else:
                    return res.text
            except Exception:
                continue
    st.error("AI 모델 응답을 불러오지 못했습니다. 잠시 후 다시 시도해주세요.")
    return None

# 3. 파일 텍스트 추출 함수 (.docx, .pdf, .txt)
def extract_text_from_file(uploaded_file):
    text = ""
    file_name = uploaded_file.name.lower()
    if file_name.endswith('.txt'):
        text = uploaded_file.read().decode('utf-8', errors='ignore')
    elif file_name.endswith('.docx'):
        doc = Document(uploaded_file)
        text = "\n".join([p.text for p in doc.paragraphs if p.text])
    elif file_name.endswith('.pdf'):
        pdf = PdfReader(uploaded_file)
        for page in pdf.pages:
            t = page.extract_text()
            if t: text += t + "\n"
    return text

# 4. 16:9 예배용 PPT 생성기
def generate_church_pptx(title, scripture, preacher, points, conclusion):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
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
    p1.font.size = Pt(44)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(253, 224, 71)
    p1.alignment = PP_ALIGN.CENTER
    p2 = tx.text_frame.add_paragraph()
    p2.text = f"\n본문: {scripture} | 설교: {preacher}"
    p2.font.size = Pt(22)
    p2.font.color.rgb = RGBColor(226, 232, 240)
    p2.alignment = PP_ALIGN.CENTER

    # 대지별 슬라이드
    for idx, pt in enumerate(points, 1):
        slide = prs.slides.add_slide(blank_layout)
        apply_bg(slide, RGBColor(24, 32, 54))
        htx = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(1.2))
        hp = htx.text_frame.paragraphs[0]
        hp.text = pt.get("main_point", f"제 {idx} 대지")
        hp.font.size = Pt(36)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(253, 224, 71)

        btx = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(11.33), Inches(4.5))
        tf = btx.text_frame
        tf.word_wrap = True
        bp1 = tf.paragraphs[0]
        bp1.text = f"📖 성경적 해설:\n{pt.get('explanation', '')}\n"
        bp1.font.size = Pt(22)
        bp1.font.color.rgb = RGBColor(241, 245, 249)
        bp2 = tf.add_paragraph()
        bp2.text = f"\n💡 삶의 적용:\n{pt.get('application', '')}"
        bp2.font.size = Pt(22)
        bp2.font.color.rgb = RGBColor(147, 197, 253)

    # 결론 슬라이드
    slide_end = prs.slides.add_slide(blank_layout)
    apply_bg(slide_end, RGBColor(15, 23, 42))
    etx = slide_end.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.33), Inches(5.0))
    ep1 = etx.text_frame.paragraphs[0]
    ep1.text = "결론 및 결단의 기도"
    ep1.font.size = Pt(36)
    ep1.font.bold = True
    ep1.font.color.rgb = RGBColor(253, 224, 71)
    ep2 = etx.text_frame.add_paragraph()
    ep2.text = f"\n{conclusion}"
    ep2.font.size = Pt(22)
    ep2.font.color.rgb = RGBColor(226, 232, 240)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# 5. 1:1 정사각형 카드뉴스 PPT 생성기
def generate_cardnews_pptx(slides_data):
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(10)
    blank_layout = prs.slide_layouts[6]

    for item in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(15, 23, 42)

        # 번호 배지
        badge = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(2), Inches(0.8))
        bp = badge.text_frame.paragraphs[0]
        bp.text = f"CARD {item.get('card_number', '')}"
        bp.font.size = Pt(14)
        bp.font.bold = True
        bp.font.color.rgb = RGBColor(99, 102, 241)

        # 타이틀
        tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(2.2))
        t_frame = tbox.text_frame
        t_frame.word_wrap = True
        tp = t_frame.paragraphs[0]
        tp.text = item.get("headline", "")
        tp.font.size = Pt(32)
        tp.font.bold = True
        tp.font.color.rgb = RGBColor(253, 224, 71)

        # 본문 문구
        bbox = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), Inches(8.4), Inches(4.5))
        b_frame = bbox.text_frame
        b_frame.word_wrap = True
        bp_body = b_frame.paragraphs[0]
        bp_body.text = item.get("body_text", "")
        bp_body.font.size = Pt(22)
        bp_body.font.color.rgb = RGBColor(241, 245, 249)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

# 6. 세션 상태 초기화
if "outline_data" not in st.session_state: st.session_state.outline_data = None
if "full_sermon" not in st.session_state: st.session_state.full_sermon = ""
if "preacher_name" not in st.session_state: st.session_state.preacher_name = "담임목사"
if "cardnews_data" not in st.session_state: st.session_state.cardnews_data = None

# 7. 네비게이션
st.sidebar.title("🕊️ 설교 AI 스튜디오")
menu = st.sidebar.radio(
    "메뉴 선택",
    [
        "📤 기존 설교문 업로드",
        "📖 새 설교 개요 & 주해",
        "✍️ 설교문 전문 작성",
        "📋 설교문 요약 & 주보",
        "👥 소그룹 셀 나눔지",
        "🏡 가정예배 순서지",
        "📱 카드뉴스 (7~10장)",
        "🖥️ PPT 슬라이드 생성",
        "🎬 영상 & 쇼츠 스튜디오"
    ]
)

# --- 1. 기존 설교문 업로드 ---
if menu == "📤 기존 설교문 업로드":
    st.title("📤 기존 설교문 파일 업로드 & 연동")
    st.caption("이미 작성해두신 설교문 파일(.docx, .pdf, .txt)을 올리거나 텍스트를 붙여넣으면 모든 사역자료가 즉시 생성됩니다.")
    
    upload_col, direct_col = st.columns([1.2, 1])
    with upload_col:
        up_file = st.file_uploader("설교문 파일 업로드", type=["docx", "pdf", "txt"])
        if up_file is not None:
            extracted = extract_text_from_file(up_file)
            if extracted:
                st.session_state.full_sermon = extracted
                st.success(f"'{up_file.name}'에서 설교문을 성공적으로 불러왔습니다! ({len(extracted)}자)")

    with direct_col:
        direct_text = st.text_area("또는 원고 직접 붙여넣기", value=st.session_state.full_sermon, height=220)
        if st.button("원고 저장 및 동기화", type="primary"):
            st.session_state.full_sermon = direct_text
            st.success("설교문이 플랫폼에 등록되었습니다. 이제 왼쪽 메뉴에서 요약, 소그룹, 카드뉴스, PPT를 바로 생성하세요!")

# --- 2. 설교 개요 & 주해 ---
elif menu == "📖 새 설교 개요 & 주해":
    st.title("📖 설교 개요 및 본문 원어 주해")
    c1, c2, c3 = st.columns([1.5, 2, 1])
    with c1: scripture = st.text_input("성경 본문", value="로마서 8:28-39")
    with c2: topic = st.text_input("설교 주제", value="고난 속에서도 흔들리지 않는 하나님의 사랑")
    with c3: st.session_state.preacher_name = st.text_input("설교자 성함", value=st.session_state.preacher_name)

    if st.button("✨ 개요 및 원어 주해 생성", type="primary"):
        with st.spinner("본문 주해 및 대지 구조화 중..."):
            prompt = f"""
            성경 본문: {scripture}, 주제: {topic}
            복음주의 신학적 관점에서 아래 JSON 포맷으로 작성하세요:
            {{
                "title": "설교 제목",
                "historical_background": "역사적/문화적 배경 해설",
                "original_words": [{{"word": "원어단어", "meaning": "원어적 의미 및 설교적 적용"}}],
                "points": [
                    {{"main_point": "제1대지", "explanation": "성경적 해설", "application": "삶의 적용"}},
                    {{"main_point": "제2대지", "explanation": "성경적 해설", "application": "삶의 적용"}},
                    {{"main_point": "제3대지", "explanation": "성경적 해설", "application": "삶의 적용"}}
                ],
                "conclusion": "결론 요약 및 결단 기도"
            }}
            """
            data = get_ai_response(prompt, is_json=True)
            if data:
                st.session_state.outline_data = data
                st.success("개요 생성이 완료되었습니다!")

    if st.session_state.outline_data:
        data = st.session_state.outline_data
        st.subheader(f"📌 {data.get('title')}")
        with st.expander("📜 역사적 배경 및 원어 분석", expanded=True):
            st.write(data.get("historical_background", ""))
            for ow in data.get("original_words", []):
                st.markdown(f"- **{ow.get('word')}**: {ow.get('meaning')}")
        for pt in data.get("points", []):
            st.info(f"**{pt.get('main_point')}**\n\n{pt.get('explanation')}\n\n*적용:* {pt.get('application')}")

# --- 3. 설교문 전문 작성 ---
elif menu == "✍️ 설교문 전문 작성":
    st.title("✍️ 강단 선포용 설교문 전문 에디터")
    tone = st.selectbox("설교 어조", ["은혜롭고 따뜻한 목양적 선포체", "지성적이고 논리적인 강해체", "열정적이고 결단을 촉구하는 선포체"])
    
    if st.button("🚀 개요 기반으로 전체 원고 생성"):
        if not st.session_state.outline_data:
            st.warning("먼저 [설교 개요]를 생성하거나 [기존 설교문 업로드] 탭을 이용해주세요.")
        else:
            with st.spinner("25분 완성형 설교문 작성 중..."):
                prompt = f"""
                아래 개요를 바탕으로 25-30분 분량의 완성된 구어체 설교문 전문을 작성하세요:
                {json.dumps(st.session_state.outline_data, ensure_ascii=False)}
                어조: {tone}
                """
                res = get_ai_response(prompt, is_json=False)
                if res:
                    st.session_state.full_sermon = res
                    st.success("설교문 전문이 작성되었습니다!")

    st.session_state.full_sermon = st.text_area("설교 원고", value=st.session_state.full_sermon, height=450)

# --- 4. 설교문 요약 & 주보 ---
elif menu == "📋 설교문 요약 & 주보":
    st.title("📋 설교문 요약 및 주보 게재용 요약문")
    if st.button("📑 핵심 요약문 추출하기", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문을 입력하거나 업로드해주세요.")
        else:
            with st.spinner("주보용 핵심 요약 추출 중..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:4000]}
                아래 JSON 형식으로 응답하세요:
                {{
                    "three_lines": ["3줄 요약 1", "3줄 요약 2", "3줄 요약 3"],
                    "bulletin_summary": "주보 게재용 300자 정갈한 요약문",
                    "one_sentence_meditation": "성도들을 위한 한 줄 묵상"
                }}
                """
                summary = get_ai_response(prompt, is_json=True)
                if summary:
                    st.subheader("💡 3줄 핵심 요약")
                    for l in summary.get("three_lines", []): st.success(f"✓ {l}")
                    st.subheader("📰 주보 게재용 요약문")
                    st.text_area("복사용 텍스트", value=summary.get("bulletin_summary", ""), height=150)
                    st.subheader("🙏 한 줄 묵상")
                    st.info(summary.get("one_sentence_meditation", ""))

# --- 5. 소그룹 셀 나눔지 ---
elif menu == "👥 소그룹 셀 나눔지":
    st.title("👥 소그룹 / 구역예배 나눔지")
    if st.button("📑 소그룹 나눔지 자동 생성", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문을 입력하거나 업로드해주세요.")
        else:
            with st.spinner("소그룹 질문지 생성 중..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:4000]}
                JSON 포맷으로 출력:
                {{
                    "ice_breaker": "마음 열기 질문",
                    "word_questions": ["말씀 나눔 1", "말씀 나눔 2"],
                    "life_application": ["삶 적용 1", "삶 적용 2"],
                    "closing_prayer": "마무리 기도문"
                }}
                """
                group = get_ai_response(prompt, is_json=True)
                if group:
                    st.subheader("1. 마음 열기 (Ice Break)")
                    st.info(group.get("ice_breaker", ""))
                    st.subheader("2. 말씀 속으로 (Word & Insight)")
                    for q in group.get("word_questions", []): st.write(f"• {q}")
                    st.subheader("3. 삶 속으로 (Life Application)")
                    for q in group.get("life_application", []): st.write(f"• {q}")
                    st.subheader("4. 마침 기도")
                    st.caption(group.get("closing_prayer", ""))

# --- 6. 가정예배 순서지 ---
elif menu == "🏡 가정예배 순서지":
    st.title("🏡 온 가족이 함께 드리는 가정예배지")
    if st.button("📑 가정예배 순서지 생성", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문을 입력하거나 업로드해주세요.")
        else:
            with st.spinner("가정예배 순서지 작성 중..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:4000]}
                JSON 포맷으로 작성:
                {{
                    "hymn": "추천 찬송가 장수 및 제목",
                    "scripture_reading": "가족 함께 읽는 본문 구절",
                    "family_message": "어린이부터 부모까지 쉬운 3분 가족 메시지",
                    "family_sharing": "가족 나눔 질문 2가지",
                    "family_prayer": "가정 축복 기도문"
                }}
                """
                fam = get_ai_response(prompt, is_json=True)
                if fam:
                    st.markdown(f"### 🎵 찬양: {fam.get('hymn', '')}")
                    st.markdown(f"**📖 말씀:** {fam.get('scripture_reading', '')}")
                    st.subheader("👨‍👩‍👧‍👦 3분 가족 나눔")
                    st.write(fam.get("family_message", ""))
                    st.subheader("💬 우리 집 나눔")
                    st.info(fam.get("family_sharing", ""))
                    st.subheader("🙏 축복 기도")
                    st.caption(fam.get("family_prayer", ""))

# --- 7. 카드뉴스 (7~10장) ---
elif menu == "📱 카드뉴스 (7~10장)":
    st.title("📱 SNS / 인스타그램 카드뉴스 (7~10장) 스튜디오")
    st.caption("설교문에서 성도들이 SNS로 쉽게 읽을 수 있는 7~10장의 정갈한 카드뉴스를 자동 추출합니다.")
    
    card_count = st.slider("카드뉴스 장수 선택", min_value=7, max_value=10, value=8)
    
    if st.button("🎨 카드뉴스 일괄 생성하기", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문을 입력하거나 업로드해주세요.")
        else:
            with st.spinner(f"{card_count}장의 카드뉴스를 구성 중입니다..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:4000]}
                정확히 {card_count}장의 인스타그램 카드뉴스 형태로 작성해주세요.
                구조:
                - 1번 카드: 표지 (강렬한 헤드라인 + 부제목)
                - 2번~{card_count-1}번 카드: 본론 (문제 제기 - 공감 - 말씀 핵심 - 교훈 - 실천 방안)
                - {card_count}번 카드: 마무리 (결단 기도 및 축복)

                JSON 포맷:
                {{
                    "cards": [
                        {{
                            "card_number": 1,
                            "headline": "표지 제목",
                            "body_text": "표지 서브 문구"
                        }}
                    ]
                }}
                """
                res = get_ai_response(prompt, is_json=True)
                if res and "cards" in res:
                    st.session_state.cardnews_data = res["cards"]
                    st.success(f"{len(res['cards'])}장의 카드뉴스가 완성되었습니다!")

    if st.session_state.cardnews_data:
        cards = st.session_state.cardnews_data
        
        # 다운로드 버튼
        card_pptx = generate_cardnews_pptx(cards)
        st.download_button(
            label="📥 1:1 정사각형 카드뉴스 PPT(.pptx) 다운로드",
            data=card_pptx,
            file_name="설교_카드뉴스_1대1.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        st.write("---")

        # 카드 미리보기 뷰어 (2열 그리드)
        cols = st.columns(2)
        for idx, card in enumerate(cards):
            with cols[idx % 2]:
                st.markdown(
                    f"""
                    <div style="background-color: #0f172a; border: 1px solid #334155; border-radius: 16px; padding: 24px; margin-bottom: 20px;">
                        <span style="background-color: #4f46e5; color: white; padding: 4px 10px; border-radius: 8px; font-size: 12px; font-weight: bold;">CARD {card.get('card_number', idx+1)}</span>
                        <h3 style="color: #fde047; font-size: 20px; margin-top: 14px; margin-bottom: 10px; font-weight: bold;">{card.get('headline', '')}</h3>
                        <p style="color: #e2e8f0; font-size: 15px; line-height: 1.6; white-space: pre-wrap;">{card.get('body_text', '')}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

# --- 8. PPT 슬라이드 생성 ---
elif menu == "🖥️ PPT 슬라이드 생성":
    st.title("🖥️ 예배 설교용 와이드 PPT (.pptx) 생성")
    
    # 개요 데이터가 없으면 설교문에서 즉석 대지 추출
    if not st.session_state.outline_data and st.session_state.full_sermon:
        if st.button("📝 현재 설교문에서 PPT 슬라이드 구조 자동 추출"):
            with st.spinner("설교문에서 3대지 구조 추출 중..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:4000]}
                PPT 제작을 위해 JSON 형식으로 출력:
                {{
                    "title": "설교 제목",
                    "scripture": "본문 구절",
                    "points": [
                        {{"main_point": "제1대지", "explanation": "해설", "application": "적용"}},
                        {{"main_point": "제2대지", "explanation": "해설", "application": "적용"}},
                        {{"main_point": "제3대지", "explanation": "해설", "application": "적용"}}
                    ],
                    "conclusion": "결론 기도"
                }}
                """
                st.session_state.outline_data = get_ai_response(prompt, is_json=True)
                st.rerun()

    if st.session_state.outline_data:
        out = st.session_state.outline_data
        st.write(f"**설교 제목:** {out.get('title', '주일 설교')}")
        if st.button("🎨 16:9 와이드 파워포인트 파일 다운로드", type="primary"):
            pptx_file = generate_church_pptx(
                title=out.get("title", "주일 설교"),
                scripture=out.get("scripture", "본문 말씀"),
                preacher=st.session_state.preacher_name,
                points=out.get("points", []),
                conclusion=out.get("conclusion", "결단의 기도")
            )
            st.download_button(
                label="📥 파워포인트(.pptx) 파일 내려받기",
                data=pptx_file,
                file_name=f"설교_{out.get('title', '말씀')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.info("[기존 설교문 업로드] 메뉴에 원고를 넣거나 [새 설교 개요]를 먼저 만들어주세요.")

# --- 9. 영상 & 쇼츠 스튜디오 ---
elif menu == "🎬 영상 & 쇼츠 스튜디오":
    st.title("🎬 자막 애니메이션 & BGM 영상 스튜디오")
    v_title = st.text_input("영상 메인 헤드라인", value="고난 속에서 기억할 한 가지")
    v_script = st.text_area("자막/대본 (줄바꿈으로 문장 구분)", value="고난은 결코 우연이 아닙니다.\n하나님은 모든 것을 선으로 바꾸십니다.\n오늘 그 은혜 안에서 평안을 누리세요.", height=120)
    
    colA, colB = st.columns(2)
    with colA: v_ratio = st.radio("비율", ["9:16 (세로 쇼츠)", "16:9 (가로 영상)"])
    with colB: v_voice = st.selectbox("보이스", ["인준 (남성)", "선희 (여성)"])
        
    bg_file = st.file_uploader("배경 미디어 (선택)", type=["jpg", "png", "mp4"])
    bgm_file = st.file_uploader("배경음악 MP3 (선택)", type=["mp3"])

    if st.button("🚀 고화질 영상 렌더링 시작", type="primary"):
        with st.spinner("자막 애니메이션 및 BGM 믹싱 렌더링 중..."):
            bg_p, bgm_p = None, None
            if bg_file:
                bg_p = f"./uploads_{bg_file.name}"
                with open(bg_p, "wb") as f: f.write(bg_file.getbuffer())
            if bgm_file:
                bgm_p = f"./uploads_{bgm_file.name}"
                with open(bgm_p, "wb") as f: f.write(bgm_file.getbuffer())

            lines = [l.strip() for l in v_script.split("\n") if l.strip()]
            voice_code = "ko-KR-InJoonNeural" if "인준" in v_voice else "ko-KR-SunHiNeural"
            ratio_code = "9:16" if "9:16" in v_ratio else "16:9"

            out_vid = create_animated_video(v_title, lines, bg_p, bgm_p, ratio_code, voice=voice_code)
            st.session_state.rendered_vid = out_vid
            st.success("영상 렌더링이 완료되었습니다!")

    if "rendered_vid" in st.session_state and os.path.exists(st.session_state.rendered_vid):
        st.video(st.session_state.rendered_vid)
        with open(st.session_state.rendered_vid, "rb") as f:
            st.download_button("📥 MP4 비디오 다운로드", data=f, file_name="sermon_video.mp4", mime="video/mp4")
