import streamlit as st
import google.generativeai as genai
import json
import os
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from video_engine import create_animated_video

st.set_page_config(page_title="MY 설교 AI 스튜디오", page_icon="🕊️", layout="wide")

# 1. 보안 인증
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

# 2. Gemini AI 연동
GEMINI_KEY = st.secrets.get("GEMINI_API_KEY", "")
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel("gemini-1.5-flash")

# 3. PPT 자동 생성 함수 (.pptx 16:9 와이드스크린)
def generate_church_pptx(title, scripture, preacher, points, conclusion):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # 슬라이드 공통 배경 생성 함수
    def apply_bg(slide, color=RGBColor(24, 32, 54)):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color

    # [슬라이드 1: 표지]
    slide1 = prs.slides.add_slide(blank_layout)
    apply_bg(slide1, RGBColor(15, 23, 42))
    
    txBox = slide1.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(253, 224, 71)
    p.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = f"\n본문: {scripture} | 설교: {preacher}"
    p2.font.size = Pt(22)
    p2.font.color.rgb = RGBColor(226, 232, 240)
    p2.alignment = PP_ALIGN.CENTER

    # [슬라이드 2~N: 대지별 슬라이드]
    for idx, pt in enumerate(points, 1):
        slide = prs.slides.add_slide(blank_layout)
        apply_bg(slide, RGBColor(24, 32, 54))
        
        # 상단 타이틀 바
        header_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(1.2))
        htf = header_box.text_frame
        hp = htf.paragraphs[0]
        hp.text = pt.get("main_point", f"제 {idx} 대지")
        hp.font.size = Pt(36)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(253, 224, 71)
        
        # 본문 내용 박스
        body_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(11.33), Inches(4.5))
        btf = body_box.text_frame
        btf.word_wrap = True
        
        bp1 = btf.paragraphs[0]
        bp1.text = f"📖 성경적 해설:\n{pt.get('explanation', '')}\n"
        bp1.font.size = Pt(22)
        bp1.font.color.rgb = RGBColor(241, 245, 249)
        
        bp2 = btf.add_paragraph()
        bp2.text = f"\n💡 삶의 적용:\n{pt.get('application', '')}"
        bp2.font.size = Pt(22)
        bp2.font.color.rgb = RGBColor(147, 197, 253)

    # [마지막 슬라이드: 결론 및 결단 기도]
    slide_end = prs.slides.add_slide(blank_layout)
    apply_bg(slide_end, RGBColor(15, 23, 42))
    
    eBox = slide_end.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.33), Inches(5.0))
    etf = eBox.text_frame
    etf.word_wrap = True
    
    ep1 = etf.paragraphs[0]
    ep1.text = "결론 및 결단의 기도"
    ep1.font.size = Pt(36)
    ep1.font.bold = True
    ep1.font.color.rgb = RGBColor(253, 224, 71)
    
    ep2 = etf.add_paragraph()
    ep2.text = f"\n{conclusion}"
    ep2.font.size = Pt(22)
    ep2.font.color.rgb = RGBColor(226, 232, 240)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# 4. 네비게이션 메뉴
st.sidebar.title("🕊️ 설교 AI 스튜디오")
menu = st.sidebar.radio(
    "사역 플랫폼 메뉴",
    [
        "📊 대시보드",
        "📖 설교 개요 & 주해",
        "✍️ 설교문 전문 작성",
        "📋 설교문 요약 & 주보",
        "👥 소그룹 셀 나눔지",
        "🏡 가정예배 순서지",
        "🖥️ PPT 슬라이드 생성",
        "🎬 영상 & 쇼츠 스튜디오"
    ]
)

# 세션 데이터 저장소
if "outline_data" not in st.session_state:
    st.session_state.outline_data = None
if "full_sermon" not in st.session_state:
    st.session_state.full_sermon = ""
if "preacher_name" not in st.session_state:
    st.session_state.preacher_name = "담임목사"

# --- 1. 대시보드 ---
if menu == "📊 대시보드":
    st.title("목회 사역 올인원 대시보드")
    st.caption("성경 본문 한 구절로 주간 목회에 필요한 모든 문서와 미디어를 완성합니다.")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("📖 **설교 개요 & 원고**\n- 본문 역사/원어 분석\n- 3대지 설교 개요\n- 25분 완성형 원고 생성")
    with col2:
        st.success("📑 **목회 자료 파생**\n- 주보용 3줄 요약문\n- 소그룹 셀 모임 나눔지\n- 세대통합 가정예배지")
    with col3:
        st.warning("🖥️ **미디어 제작**\n- 16:9 와이드 PPT 다운로드\n- 자막 애니메이션 쇼츠\n- BGM 믹싱 1~2분 영상")

# --- 2. 설교 개요 & 주해 ---
elif menu == "📖 설교 개요 & 주해":
    st.title("📖 설교 개요 및 본문 원어 주해")
    col1, col2, col3 = st.columns([1.5, 2, 1])
    with col1:
        scripture = st.text_input("성경 본문", value="로마서 8:28-39")
    with col2:
        topic = st.text_input("설교 주제", value="고난 속에서도 흔들리지 않는 하나님의 사랑")
    with col3:
        st.session_state.preacher_name = st.text_input("설교자 성함", value=st.session_state.preacher_name)

    if st.button("✨ 개요 및 원어 주해 생성", type="primary"):
        with st.spinner("AI가 본문 역사적 배경과 원어 주해를 분석 중입니다..."):
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
            res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
            st.session_state.outline_data = json.loads(res.text)
            st.success("개요 및 주해 분석이 완료되었습니다!")

    if st.session_state.outline_data:
        data = st.session_state.outline_data
        st.subheader(f"📌 {data['title']}")
        with st.expander("📜 역사적 배경 및 원어 분석", expanded=True):
            st.write(data["historical_background"])
            for ow in data.get("original_words", []):
                st.markdown(f"- **{ow['word']}**: {ow['meaning']}")
        
        st.subheader("대지 구성")
        for pt in data.get("points", []):
            st.info(f"**{pt['main_point']}**\n\n{pt['explanation']}\n\n*적용:* {pt['application']}")

# --- 3. 설교문 전문 작성 ---
elif menu == "✍️ 설교문 전문 작성":
    st.title("✍️ 강단 선포용 설교문 전문 에디터")
    tone = st.selectbox("설교 어조", ["은혜롭고 따뜻한 목양적 선포체", "지성적이고 논리적인 강해체", "열정적이고 결단을 촉구하는 선포체"])
    
    if st.button("🚀 개요 기반으로 25분 분량 원고 작성", type="primary"):
        if not st.session_state.outline_data:
            st.warning("먼저 [설교 개요 & 주해] 메뉴에서 개요를 만들어주세요.")
        else:
            with st.spinner("예화와 적용이 포함된 완성형 설교문을 작성 중입니다..."):
                prompt = f"""
                다음 개요를 바탕으로 강단에서 실제로 선포할 수 있는 25-30분 분량의 완성된 설교문 전문을 작성하세요.
                개요: {json.dumps(st.session_state.outline_data, ensure_ascii=False)}
                어조: {tone}
                구어체(~합니다)와 실천적 일상 예화를 풍부하게 포함하세요.
                """
                res = model.generate_content(prompt)
                st.session_state.full_sermon = res.text
                st.success("설교문 전문이 완성되었습니다!")
                
    st.session_state.full_sermon = st.text_area("설교 원고", value=st.session_state.full_sermon, height=450)

# --- 4. 설교문 요약 & 주보 ---
elif menu == "📋 설교문 요약 & 주보":
    st.title("📋 설교문 요약 및 주보 게재용 요약문")
    if st.button("📑 설교문 기반 핵심 요약 추출", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 [설교문 전문 작성] 메뉴에서 원고를 작성해주세요.")
        else:
            with st.spinner("주보용 및 성도용 요약문 추출 중..."):
                prompt = f"""
                설교 원고: {st.session_state.full_sermon[:3500]}
                JSON 포맷으로 출력:
                {{
                    "three_lines": ["핵심 요약 1", "핵심 요약 2", "핵심 요약 3"],
                    "bulletin_summary": "주보에 실을 250~300자 내외의 정갈한 설교 요약문",
                    "one_sentence_meditation": "성도들이 일주일간 품고 기도할 한 줄 묵상 말씀"
                }}
                """
                res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                summary = json.loads(res.text)
                
                st.subheader("💡 3줄 핵심 요약")
                for line in summary.get("three_lines", []):
                    st.success(f"✓ {line}")
                    
                st.subheader("📰 주보 게재용 요약문 (복사하여 사용)")
                st.text_area("주보 본문", value=summary.get("bulletin_summary", ""), height=150)
                
                st.subheader("🙏 한 줄 묵상 포인트")
                st.info(summary.get("one_sentence_meditation", ""))

# --- 5. 소그룹 셀 나눔지 ---
elif menu == "👥 소그룹 셀 나눔지":
    st.title("👥 소그룹 / 구역예배 나눔지")
    if st.button("📑 소그룹 나눔지 자동 생성", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문 전문을 작성해주세요.")
        else:
            with st.spinner("소그룹 맞춤 질문지 구성 중..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:3500]}
                소그룹/구역 모임 나눔지를 JSON 형식으로 작성:
                {{
                    "ice_breaker": "마음 열기 (일상의 가벼운 나눔 질문)",
                    "word_questions": ["말씀 속으로 1 (본문 이해 질문)", "말씀 속으로 2 (설교 핵심 질문)"],
                    "life_application": ["삶 속으로 1 (구체적 실천 질문)", "삶 속으로 2 (나눔 및 기도제목)"],
                    "closing_prayer": "마무리 합심 기도문"
                }}
                """
                res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                group = json.loads(res.text)
                
                st.subheader("1. 마음 열기 (Ice Break)")
                st.info(group.get("ice_breaker", ""))
                
                st.subheader("2. 말씀 속으로 (Word & Insight)")
                for q in group.get("word_questions", []):
                    st.write(f"• {q}")
                    
                st.subheader("3. 삶 속으로 (Life Application)")
                for q in group.get("life_application", []):
                    st.write(f"• {q}")
                    
                st.subheader("4. 마침 기도")
                st.caption(group.get("closing_prayer", ""))

# --- 6. 가정예배 순서지 ---
elif menu == "🏡 가정예배 순서지":
    st.title("🏡 온 가족이 함께 드리는 가정예배 순서지")
    if st.button("📑 가정예배 순서지 생성", type="primary"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문 전문을 작성해주세요.")
        else:
            with st.spinner("부모와 자녀가 함께 드리는 예배 순서지 작성 중..."):
                prompt = f"""
                설교문: {st.session_state.full_sermon[:3500]}
                가정예배 순서지를 JSON 형식으로 작성:
                {{
                    "hymn": "추천 찬송가 (장 및 제목) 또는 복음성가",
                    "confession": "사도신경 고백 안내",
                    "scripture_reading": "가족 함께 읽는 핵심 성경 구절",
                    "family_message": "어린이부터 부모까지 이해하기 쉬운 3분 가족 메시지",
                    "family_sharing": "가족 간 나눔 질문 2가지",
                    "family_prayer": "가정을 축복하는 마무리 기도문"
                }}
                """
                res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                fam = json.loads(res.text)
                
                st.markdown(f"### 🎵 찬양: {fam.get('hymn', '')}")
                st.markdown(f"**📖 본문 말씀:** {fam.get('scripture_reading', '')}")
                
                st.subheader("👨‍👩‍👧‍👦 3분 가족 말씀 나눔")
                st.write(fam.get("family_message", ""))
                
                st.subheader("💬 오늘 우리 집 나눔")
                st.info(fam.get("family_sharing", ""))
                
                st.subheader("🙏 축복 기도")
                st.caption(fam.get("family_prayer", ""))

# --- 7. PPT 슬라이드 생성 ---
elif menu == "🖥️ PPT 슬라이드 생성":
    st.title("🖥️ 예배 설교용 PPT 슬라이드 (.pptx) 자동 생성")
    st.caption("16:9 와이드스크린 규격으로 가독성 높은 고화질 파워포인트 파일을 생성하여 다운로드합니다.")
    
    if not st.session_state.outline_data:
        st.warning("먼저 [설교 개요 & 주해] 메뉴에서 설교 개요를 생성해주세요.")
    else:
        out = st.session_state.outline_data
        st.write(f"**설교 제목:** {out.get('title', '')}")
        st.write(f"**대지 수:** {len(out.get('points', []))}개 대지")
        
        if st.button("🎨 16:9 예배용 PPT 파일 생성하기", type="primary"):
            pptx_file = generate_church_pptx(
                title=out.get("title", "주일 설교"),
                scripture=out.get("scripture", "본문 말씀"),
                preacher=st.session_state.preacher_name,
                points=out.get("points", []),
                conclusion=out.get("conclusion", "결단의 기도")
            )
            
            st.success("PPT 생성이 완료되었습니다! 아래 버튼을 눌러 파워포인트 파일을 다운로드하세요.")
            st.download_button(
                label="📥 파워포인트(.pptx) 파일 다운로드",
                data=pptx_file,
                file_name=f"설교_{out.get('title', '말씀')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

# --- 8. 영상 & 쇼츠 스튜디오 ---
elif menu == "🎬 영상 & 쇼츠 스튜디오":
    st.title("🎬 애니메이션 자막 & BGM 영상 스튜디오")
    v_title = st.text_input("영상 제목", value="고난 속에서 기억할 한 가지")
    v_script = st.text_area(
        "자막/대본 (줄바꿈으로 문장 구분)",
        value="고난은 결코 우연이 아닙니다.\n하나님은 모든 것을 선으로 바꾸십니다.\n오늘 그 은혜 안에서 참된 평안을 누리세요.",
        height=120
    )
    
    colA, colB = st.columns(2)
    with colA:
        v_ratio = st.radio("화면 비율", ["9:16 (세로 쇼츠/릴스)", "16:9 (가로 1~2분 영상)"])
    with colB:
        v_voice = st.selectbox("성우 목소리", ["인준 (남성 - 차분한 목소리)", "선희 (여성 - 또렷한 목소리)"])
        
    bg_file = st.file_uploader("배경 이미지/동영상 (선택)", type=["jpg", "png", "mp4"])
    bgm_file = st.file_uploader("배경음악 MP3 (선택)", type=["mp3"])

    if st.button("🚀 고화질 영상 렌더링 시작", type="primary"):
        with st.spinner("음성 합성, 자막 애니메이션, BGM 믹싱 렌더링 중... (약 30초 소요)"):
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

            out_video = create_animated_video(
                title=v_title,
                script_paragraphs=lines,
                bg_media_path=bg_p,
                bgm_path=bgm_p,
                aspect_ratio=ratio_code,
                voice=voice_code
            )
            st.session_state.rendered_vid = out_video
            st.success("영상 렌더링이 완료되었습니다!")

    if "rendered_vid" in st.session_state and os.path.exists(st.session_state.rendered_vid):
        st.video(st.session_state.rendered_vid)
        with open(st.session_state.rendered_vid, "rb") as f:
            st.download_button("📥 MP4 비디오 다운로드", data=f, file_name="sermon_video.mp4", mime="video/mp4")
