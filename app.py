import streamlit as st
import google.generativeai as genai
import json
import os
from video_engine import create_animated_video

st.set_page_config(page_title="MY 설교 AI 스튜디오", page_icon="🕊️", layout="wide")

# 보안 접속 PIN 확인
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

# Gemini 연동
GEMINI_KEY = st.secrets.get("GEMINI_API_KEY", "")
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel("gemini-1.5-flash")

st.sidebar.title("🕊️ 설교 AI 스튜디오")
menu = st.sidebar.radio(
    "사역 메뉴",
    ["📊 대시보드", "📖 설교 개요 & 주해", "✍️ 설교문 전문 작성", "🗂️ 사역자료 & 예화", "🎬 1~2분 영상 & 쇼츠 제작"]
)

if "outline_data" not in st.session_state:
    st.session_state.outline_data = None
if "full_sermon" not in st.session_state:
    st.session_state.full_sermon = ""

if menu == "📊 대시보드":
    st.title("목회 사역 대시보드")
    st.info("한 편의 본문으로 설교문, 주간 QT, 소그룹 나눔지, 쇼츠/영상 제작까지 원스톱으로 처리합니다.")

elif menu == "📖 설교 개요 & 주해":
    st.title("📖 설교 개요 및 본문 원어 주해")
    col1, col2 = st.columns(2)
    with col1:
        scripture = st.text_input("성경 본문", value="로마서 8:28-39")
    with col2:
        topic = st.text_input("설교 주제", value="고난 속에서도 흔들리지 않는 하나님의 사랑")

    if st.button("✨ 개요 및 주해 생성하기", type="primary"):
        with st.spinner("AI가 본문 주해와 3대지 구조를 설계 중입니다..."):
            prompt = f"본문: {scripture}, 주제: {topic}. JSON 포맷으로 출력: {{\"title\":\"제목\", \"historical_background\":\"배경\", \"original_words\":[{{\"word\":\"원어\", \"meaning\":\"의미\"}}], \"points\":[{{\"main_point\":\"대지\", \"explanation\":\"해설\", \"application\":\"적용\"}}]}}"
            res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
            st.session_state.outline_data = json.loads(res.text)
            st.success("개요 생성이 완료되었습니다!")

    if st.session_state.outline_data:
        data = st.session_state.outline_data
        st.subheader(f"📌 {data['title']}")
        st.write(data["historical_background"])
        for pt in data.get("points", []):
            st.info(f"**{pt['main_point']}**\n\n{pt['explanation']}\n\n*적용:* {pt['application']}")

elif menu == "✍️ 설교문 전문 작성":
    st.title("✍️ 25분 분량 설교문 전문 작성")
    tone = st.selectbox("설교 톤", ["은혜롭고 따뜻한 선포체", "지성적이고 논리적인 설명체", "열정적인 결단 선포체"])
    if st.button("🚀 개요 기반으로 전체 원고 작성"):
        if not st.session_state.outline_data:
            st.warning("먼저 [설교 개요 & 주해] 탭에서 개요를 만들어주세요.")
        else:
            with st.spinner("강단 선포용 전체 설교 원고 작성 중..."):
                prompt = f"다음 개요를 바탕으로 25분 분량의 완성된 구어체 설교문 전문을 작성하세요: {json.dumps(st.session_state.outline_data, ensure_ascii=False)}, 톤: {tone}"
                res = model.generate_content(prompt)
                st.session_state.full_sermon = res.text
                st.success("설교문이 완성되었습니다!")
    st.session_state.full_sermon = st.text_area("설교 원고 편집", value=st.session_state.full_sermon, height=400)

elif menu == "🗂️ 사역자료 & 예화":
    st.title("🗂️ 5일치 QT 및 소그룹 질문지 추출")
    if st.button("📑 설교문 기반 자동 추출"):
        if not st.session_state.full_sermon:
            st.warning("먼저 설교문 전문을 작성해주세요.")
        else:
            with st.spinner("사역 자료 생성 중..."):
                prompt = f"설교문({st.session_state.full_sermon[:3000]})을 바탕으로 5일치 주간 QT와 소그룹 질문 3개를 JSON으로 작성: {{\"qt\":[{{\"day\":\"월\", \"title\":\"제목\", \"text\":\"묵상\"}}], \"questions\":[\"질문1\", \"질문2\"]}}"
                res = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                mats = json.loads(res.text)
                for q in mats.get("qt", []):
                    st.write(f"**[{q['day']}] {q['title']}**: {q['text']}")
                st.subheader("소그룹 나눔 질문")
                for item in mats.get("questions", []):
                    st.write(f"- {item}")

elif menu == "🎬 1~2분 영상 & 쇼츠 제작":
    st.title("🎬 자막 애니메이션 & BGM 영상 스튜디오")
    v_title = st.text_input("영상 제목", value="고난 속에서 기억할 한 가지")
    v_script = st.text_area("자막/대본 (줄바꿈으로 문장 구분)", value="고난은 결코 우연이 아닙니다.\n하나님은 모든 것을 선으로 바꾸십니다.\n오늘 그 은혜 안에서 평안을 누리세요.", height=120)
    
    colA, colB = st.columns(2)
    with colA:
        v_ratio = st.radio("화면 비율", ["9:16 (세로 쇼츠)", "16:9 (가로 1~2분 영상)"])
    with colB:
        v_voice = st.selectbox("성우 목소리", ["인준 (남성 - 차분함)", "선희 (여성 - 또렷함)"])
        
    bg_file = st.file_uploader("배경 이미지/비디오 (선택)", type=["jpg", "png", "mp4"])
    bgm_file = st.file_uploader("배경음악 MP3 (선택)", type=["mp3"])

    if st.button("🚀 고화질 영상 렌더링 시작", type="primary"):
        with st.spinner("음성 합성, 자막 애니메이션, BGM 믹싱 렌더링 중... (약 20~40초)"):
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
            st.download_button("📥 MP4 비디오 파일 다운로드", data=f, file_name="sermon_video.mp4", mime="video/mp4")
