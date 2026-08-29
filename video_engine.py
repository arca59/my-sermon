# -*- coding: utf-8 -*-
"""
video_engine.py  (v2.0)

MY 설교 AI 스튜디오 Pro 전용 영상 엔진.

[이 버전이 해결하는 것]
 - moviepy 1.x 는 최신 파이썬(3.12/3.13)에서 아예 빌드가 되지 않아
   Streamlit Cloud 배포 시 "Error installing requirements" 가 발생했습니다.
 - 그래서 moviepy 2.x 기준으로 다시 작성하고, 1.x 에서도 그대로 돌아가도록
   호환 별칭(set_position / subclip / resize / crop ...)을 붙였습니다.
 - 자막은 ImageMagick(TextClip) 없이 PIL 로 직접 그립니다.
   → 서버에 ImageMagick 이 없어도 한글 자막이 정상 출력됩니다.

app.py 에서 쓰는 공개 함수
 - create_pil_text_clip(...)   : 한글 텍스트 이미지를 영상 클립으로
 - create_animated_video(...)  : 나레이션 + 자막 + BGM 합성 영상 렌더링
"""

import os
import re
import asyncio
import tempfile
from datetime import datetime

import numpy as np
from PIL import Image, ImageDraw, ImageFont

# ------------------------------------------------------------------------------
# moviepy 1.x / 2.x 양쪽 지원
# ------------------------------------------------------------------------------
MOVIEPY_MAJOR = 2
try:  # moviepy 2.x
    from moviepy import (
        VideoFileClip, ImageClip, ColorClip, CompositeVideoClip,
        AudioFileClip, CompositeAudioClip,
        concatenate_videoclips, concatenate_audioclips,
    )
except Exception:  # moviepy 1.x
    MOVIEPY_MAJOR = 1
    from moviepy.editor import (  # type: ignore
        VideoFileClip, ImageClip, ColorClip, CompositeVideoClip,
        AudioFileClip, CompositeAudioClip,
        concatenate_videoclips, concatenate_audioclips,
    )


def _install_compat_aliases():
    """moviepy 2.x 에 1.x 스타일 메서드 이름을 붙여 기존 코드가 그대로 돌게 한다."""
    pairs = [
        ("set_position", "with_position"),
        ("set_duration", "with_duration"),
        ("set_start", "with_start"),
        ("set_end", "with_end"),
        ("set_opacity", "with_opacity"),
        ("set_audio", "with_audio"),
        ("set_fps", "with_fps"),
        ("subclip", "subclipped"),
        ("resize", "resized"),
        ("crop", "cropped"),
        ("volumex", "with_volume_scaled"),
    ]
    classes = [VideoFileClip, ImageClip, ColorClip, CompositeVideoClip,
               AudioFileClip, CompositeAudioClip]
    for cls in classes:
        for old, new in pairs:
            if not hasattr(cls, old) and hasattr(cls, new):
                try:
                    setattr(cls, old, getattr(cls, new))
                except Exception:
                    pass


_install_compat_aliases()

try:
    import edge_tts
    HAS_TTS = True
except Exception:
    HAS_TTS = False


# ==============================================================================
# 한글 폰트
# ==============================================================================
_FONT_PATHS = [
    "./fonts/NanumGothic-Bold.ttf",
    "./fonts/NanumGothic-Regular.ttf",
    "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/opentype/nanum/NanumGothicBold.ttf",
    "/usr/share/fonts/opentype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "C:/Windows/Fonts/malgunbd.ttf",
    "C:/Windows/Fonts/malgun.ttf",
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
]
_FONT_CACHE = {}


def _font(size: int):
    if size in _FONT_CACHE:
        return _FONT_CACHE[size]
    for p in _FONT_PATHS:
        if os.path.exists(p):
            try:
                f = ImageFont.truetype(p, size)
                _FONT_CACHE[size] = f
                return f
            except Exception:
                continue
    f = ImageFont.load_default()
    _FONT_CACHE[size] = f
    return f


_EMOJI_RE = re.compile("[\U0001F000-\U0001FAFF\U00002600-\U000027BF\U0001F1E6-\U0001F1FF️‍]+")


def _clean(text: str) -> str:
    """한글 폰트에 없는 이모지는 □ 로 찍히므로 제거"""
    return re.sub(r'\s{2,}', ' ', _EMOJI_RE.sub('', str(text or ''))).strip()


def _wrap(text: str, font, max_width: int, draw) -> str:
    lines = []
    for para in str(text).split('\n'):
        cur = ""
        for w in para.split(' '):
            test = f"{cur} {w}".strip()
            bb = draw.textbbox((0, 0), test, font=font)
            if (bb[2] - bb[0]) > max_width and cur:
                lines.append(cur)
                cur = w
            else:
                cur = test
        lines.append(cur)
    return "\n".join(lines)


def render_text_rgba(text, fontsize=40, color="#FFFFFF", stroke_color="black",
                     stroke_width=2, max_width=900, align="center", line_spacing=None):
    """텍스트를 투명 배경 RGBA 넘파이 배열로 렌더링"""
    text = _clean(text)
    if not text:
        text = " "
    font = _font(int(fontsize))
    spacing = line_spacing if line_spacing is not None else int(fontsize * 0.45)

    probe = Image.new("RGBA", (8, 8), (0, 0, 0, 0))
    pd = ImageDraw.Draw(probe)
    wrapped = _wrap(text, font, max_width, pd)
    bb = pd.multiline_textbbox((0, 0), wrapped, font=font, spacing=spacing,
                               stroke_width=stroke_width, align=align)

    pad = int(stroke_width) * 2 + 6
    w = max(1, int(round(bb[2] - bb[0])) + pad * 2)
    h = max(1, int(round(bb[3] - bb[1])) + pad * 2)

    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.multiline_text((pad - bb[0], pad - bb[1]), wrapped, font=font, fill=color,
                     align=align, spacing=spacing,
                     stroke_width=stroke_width, stroke_fill=stroke_color)
    return np.array(img)


def create_pil_text_clip(text, fontsize=40, color="#FFFFFF", stroke_color="black",
                         stroke_width=2, size=(900, None), duration=5.0,
                         align="center", line_spacing=None):
    """
    한글 텍스트를 ImageMagick 없이 클립으로 만든다.
    size=(최대가로폭, None) 형식 — 세로는 내용에 맞춰 자동.
    """
    max_width = 900
    if isinstance(size, (tuple, list)) and size and size[0]:
        max_width = int(size[0])
    arr = render_text_rgba(text, fontsize, color, stroke_color, stroke_width,
                           max_width, align, line_spacing)
    clip = ImageClip(arr, transparent=True)
    return clip.set_duration(max(0.1, float(duration)))


# ==============================================================================
# 나레이션 (edge-tts)
# ==============================================================================
def _tts_to_file(text: str, voice: str, out_path: str) -> bool:
    if not HAS_TTS or not text.strip():
        return False

    async def _run():
        com = edge_tts.Communicate(text, voice)
        await com.save(out_path)

    try:
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            loop.run_until_complete(_run())
            loop.close()
        except RuntimeError:
            asyncio.run(_run())
        return os.path.exists(out_path) and os.path.getsize(out_path) > 500
    except Exception:
        return False


# ==============================================================================
# 배경 처리
# ==============================================================================
def _cover(clip, W, H):
    """비율 유지하며 캔버스를 꽉 채우도록 확대 후 중앙 크롭"""
    cw, ch = clip.size
    scale = max(W / cw, H / ch)
    resized = clip.resize((int(cw * scale) + 2, int(ch * scale) + 2))
    rw, rh = resized.size
    return resized.crop(x_center=rw / 2, y_center=rh / 2, width=W, height=H)


def _make_background(bg_media_path, W, H, total):
    if bg_media_path and os.path.exists(bg_media_path):
        ext = os.path.splitext(bg_media_path)[1].lower()
        if ext in (".mp4", ".mov", ".m4v", ".webm", ".avi"):
            try:
                src = VideoFileClip(bg_media_path).without_audio() \
                    if hasattr(VideoFileClip, "without_audio") else VideoFileClip(bg_media_path)
                if src.duration < total:
                    n = int(total // src.duration) + 1
                    src = concatenate_videoclips([src] * n)
                src = src.subclip(0, total)
                return _cover(src, W, H)
            except Exception:
                pass
        else:
            try:
                img = ImageClip(bg_media_path).set_duration(total)
                return _cover(img, W, H)
            except Exception:
                pass

    # 배경이 없으면 짙은 남색 단색
    return ColorClip(size=(W, H), color=(11, 19, 41), duration=total)


# ==============================================================================
# 메인 렌더러
# ==============================================================================
def create_animated_video(title,
                          script_paragraphs,
                          bg_media_path=None,
                          bgm_path=None,
                          aspect_ratio="9:16",
                          voice="ko-KR-InJoonNeural",
                          title_fontsize=48,
                          sub_fontsize=42,
                          title_y=180,
                          sub_y=1400,
                          church_name="",
                          out_dir="./outputs",
                          fps=24):
    """
    나레이션 + 자막 + BGM 을 합성해 mp4 를 만들고 파일 경로를 돌려준다.
    edge-tts 가 없거나 실패하면 자막만으로(무음) 렌더링한다.
    """
    os.makedirs(out_dir, exist_ok=True)
    W, H = (1080, 1920) if str(aspect_ratio).startswith("9") else (1920, 1080)

    lines = [str(p).strip() for p in (script_paragraphs or []) if str(p).strip()]
    if not lines:
        lines = [str(title or "말씀 묵상")]

    # ---- 1) 문장별 나레이션 생성 & 길이 측정 -------------------------------
    tmpdir = tempfile.mkdtemp(prefix="sermon_tts_")
    narrations, durations = [], []
    for i, line in enumerate(lines):
        mp3 = os.path.join(tmpdir, f"n{i}.mp3")
        if _tts_to_file(_clean(line), voice, mp3):
            try:
                ac = AudioFileClip(mp3)
                narrations.append(ac)
                durations.append(max(1.2, ac.duration + 0.35))
                continue
            except Exception:
                pass
        narrations.append(None)
        durations.append(max(2.2, min(7.0, len(line) / 9.0)))

    total = float(sum(durations)) + 0.8

    # ---- 2) 배경 + 어둡게 깔기 --------------------------------------------
    layers = [_make_background(bg_media_path, W, H, total)]
    layers.append(ColorClip(size=(W, H), color=(0, 0, 0), duration=total).set_opacity(0.42))

    # ---- 3) 제목 (상단 반투명 바 + 텍스트) --------------------------------
    if title:
        bar_h = int(title_fontsize * 3.2)
        layers.append(
            ColorClip(size=(W, bar_h), color=(0, 0, 0), duration=total)
            .set_opacity(0.38)
            .set_position(("center", max(0, int(title_y) - int(bar_h * 0.35))))
        )
        layers.append(
            create_pil_text_clip(title, fontsize=title_fontsize, color="#FDE047",
                                 stroke_color="black", stroke_width=2,
                                 size=(int(W * 0.86), None), duration=total)
            .set_position(("center", int(title_y)))
        )

    # ---- 4) 자막 (나레이션 길이에 맞춰 순차 등장) --------------------------
    t = 0.4
    for line, dur in zip(lines, durations):
        layers.append(
            create_pil_text_clip(line, fontsize=sub_fontsize, color="#FFFFFF",
                                 stroke_color="black", stroke_width=3,
                                 size=(int(W * 0.84), None), duration=dur)
            .set_position(("center", int(sub_y)))
            .set_start(t)
        )
        t += dur

    # ---- 5) 교회명 워터마크 ------------------------------------------------
    if church_name:
        layers.append(
            create_pil_text_clip(church_name, fontsize=max(22, int(sub_fontsize * 0.6)),
                                 color="#93C5FD", stroke_color="black", stroke_width=1,
                                 size=(int(W * 0.7), None), duration=total)
            .set_position(("center", int(H * 0.93)))
        )

    video = CompositeVideoClip(layers, size=(W, H)).set_duration(total)

    # ---- 6) 오디오 합성 ----------------------------------------------------
    audio_tracks = []
    valid_narrations = [n for n in narrations if n is not None]
    if valid_narrations and len(valid_narrations) == len(narrations):
        try:
            pieces, cur = [], 0.4
            for ac, dur in zip(narrations, durations):
                pieces.append(ac.set_start(cur))
                cur += dur
            audio_tracks.append(CompositeAudioClip(pieces))
        except Exception:
            pass

    if bgm_path and os.path.exists(bgm_path):
        try:
            bgm = AudioFileClip(bgm_path)
            if bgm.duration < total:
                n = int(total // bgm.duration) + 1
                bgm = concatenate_audioclips([bgm] * n)
            bgm = bgm.subclip(0, total)
            bgm = bgm.volumex(0.12) if hasattr(bgm, "volumex") else bgm
            audio_tracks.append(bgm)
        except Exception:
            pass

    if audio_tracks:
        try:
            video = video.set_audio(CompositeAudioClip(audio_tracks))
        except Exception:
            pass

    # ---- 7) 출력 -----------------------------------------------------------
    safe = re.sub(r'[^0-9A-Za-z가-힣_-]+', '_', str(title))[:40] or "shorts"
    out_file = os.path.join(out_dir, f"{safe}_{int(datetime.now().timestamp())}.mp4")

    kwargs = dict(fps=fps, codec="libx264", audio_codec="aac",
                  threads=4, preset="ultrafast")
    try:
        video.write_videofile(out_file, logger=None, **kwargs)
    except TypeError:
        video.write_videofile(out_file, **kwargs)

    try:
        video.close()
        for n in narrations:
            if n is not None:
                n.close()
    except Exception:
        pass

    return out_file
