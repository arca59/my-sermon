import asyncio
import os
import numpy as np
import PIL.Image
import PIL.ImageDraw
import PIL.ImageFont

# Pillow 호환성 패치
if not hasattr(PIL.Image, 'ANTIALIAS'):
    PIL.Image.ANTIALIAS = getattr(PIL.Image, 'Resampling', PIL.Image).LANCZOS

import edge_tts
from moviepy.editor import (
    VideoFileClip, ImageClip, ColorClip, AudioFileClip,
    CompositeAudioClip, CompositeVideoClip, afx, vfx
)

OUTPUT_DIR = "./outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- PIL 기반 안전 텍스트 클립 생성기 (ImageMagick OSError 완벽 방지) ---
def create_pil_text_clip(text, fontsize=40, color="white", stroke_color="black", stroke_width=2, size=(900, None), duration=None):
    font = None
    font_candidates = [
        "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/opentype/nanum/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf",
        "./NanumGothic.ttf"
    ]
    for f in font_candidates:
        if os.path.exists(f):
            try:
                font = PIL.ImageFont.truetype(f, fontsize)
                break
            except Exception:
                pass
    if font is None:
        font = PIL.ImageFont.load_default()

    max_w = size[0] if size and size[0] else 900
    dummy_img = PIL.Image.new('RGBA', (1, 1))
    dummy_draw = PIL.ImageDraw.Draw(dummy_img)

    lines = []
    for para in str(text).split('\n'):
        words = para.split(' ')
        curr_line = ""
        for w in words:
            test_line = f"{curr_line} {w}".strip()
            bbox = dummy_draw.textbbox((0, 0), test_line, font=font)
            if (bbox[2] - bbox[0]) > max_w - 40 and curr_line:
                lines.append(curr_line)
                curr_line = w
            else:
                curr_line = test_line
        if curr_line:
            lines.append(curr_line)

    wrapped_text = "\n".join(lines) if lines else str(text)

    bbox = dummy_draw.multiline_textbbox((0, 0), wrapped_text, font=font, spacing=12)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    canvas_w = max_w
    canvas_h = text_h + 40 + stroke_width * 2

    img = PIL.Image.new('RGBA', (canvas_w, canvas_h), (0, 0, 0, 0))
    draw = PIL.ImageDraw.Draw(img)

    def parse_color(c):
        if isinstance(c, tuple): return c
        if c == "white": return (255, 255, 255, 255)
        if c == "black": return (0, 0, 0, 255)
        if c == "#FDE047": return (253, 224, 71, 255)
        if c == "#93C5FD": return (147, 197, 253, 255)
        if c.startswith('#'):
            hex_c = c.lstrip('#')
            return tuple(int(hex_c[i:i+2], 16) for i in (0, 2, 4)) + (255,)
        return (255, 255, 255, 255)

    fill_color = parse_color(color)
    s_color = parse_color(stroke_color) if stroke_color else None

    x_pos = (canvas_w - text_w) // 2
    y_pos = 20

    if stroke_width > 0 and s_color:
        for dx in range(-stroke_width, stroke_width + 1):
            for dy in range(-stroke_width, stroke_width + 1):
                if dx != 0 or dy != 0:
                    draw.multiline_text((x_pos + dx, y_pos + dy), wrapped_text, font=font, fill=s_color, align="center", spacing=12)

    draw.multiline_text((x_pos, y_pos), wrapped_text, font=font, fill=fill_color, align="center", spacing=12)

    np_arr = np.array(img)
    clip = ImageClip(np_arr)
    if duration:
        clip = clip.set_duration(duration)
    return clip

async def generate_tts(text: str, voice: str = "ko-KR-InJoonNeural") -> str:
    audio_path = os.path.join(OUTPUT_DIR, "temp_voice.mp3")
    communicate = edge_tts.Communicate(text, voice)
    await communicate.save(audio_path)
    return audio_path

def create_animated_video(
    title: str,
    script_paragraphs: list,
    bg_media_path: str = None,
    bgm_path: str = None,
    aspect_ratio: str = "9:16",
    font_name: str = "NanumGothic-Bold",
    voice: str = "ko-KR-InJoonNeural"
):
    full_script = " ".join(script_paragraphs)
    voice_path = asyncio.run(generate_tts(full_script, voice))
    voice_audio = AudioFileClip(voice_path)
    total_duration = voice_audio.duration

    if aspect_ratio == "9:16":
        width, height = 1080, 1920
    else:
        width, height = 1920, 1080

    if bg_media_path and os.path.exists(bg_media_path):
        is_video = bg_media_path.lower().endswith(('.mp4', '.mov', '.avi', '.mkv'))
        if is_video:
            clip_raw = VideoFileClip(bg_media_path)
            if clip_raw.duration < total_duration:
                clip_raw = clip_raw.fx(vfx.loop, duration=total_duration)
            else:
                clip_raw = clip_raw.subclip(0, total_duration)
            
            w_raw, h_raw = clip_raw.size
            target_ratio = width / height
            current_ratio = w_raw / h_raw
            
            if current_ratio > target_ratio:
                new_w = int(h_raw * target_ratio)
                crop_x1 = int((w_raw - new_w) / 2)
                cropped = clip_raw.crop(x1=crop_x1, y1=0, x2=crop_x1 + new_w, y2=h_raw)
            else:
                new_h = int(w_raw / target_ratio)
                crop_y1 = int((h_raw - new_h) / 2)
                cropped = clip_raw.crop(x1=0, y1=crop_y1, x2=w_raw, y2=crop_y1 + new_h)
                
            bg = cropped.resize((width, height))
        else:
            bg = ImageClip(bg_media_path).set_duration(total_duration).resize((width, height))
            
        dim_layer = ColorClip(size=(width, height), color=(0, 0, 0), duration=total_duration).set_opacity(0.4)
        base_clip = CompositeVideoClip([bg, dim_layer])
    else:
        base_clip = ColorClip(size=(width, height), color=(15, 23, 42), duration=total_duration)

    title_clip = (
        create_pil_text_clip(
            title,
            fontsize=48 if aspect_ratio == "9:16" else 40,
            color="#FDE047",
            stroke_color="black",
            stroke_width=2,
            size=(width - 160, None),
            duration=total_duration
        )
        .set_position(("center", 180 if aspect_ratio == "9:16" else 90))
        .fadein(0.5)
    )

    segment_duration = total_duration / max(len(script_paragraphs), 1)
    subtitle_clips = []

    for idx, sentence in enumerate(script_paragraphs):
        start_time = idx * segment_duration
        sub_text = (
            create_pil_text_clip(
                sentence.strip(),
                fontsize=42 if aspect_ratio == "9:16" else 36,
                color="white",
                stroke_color="black",
                stroke_width=2,
                size=(width - 180, None),
                duration=segment_duration
            )
            .set_start(start_time)
            .set_position(("center", height // 2 if aspect_ratio == "9:16" else height * 0.65))
            .fadein(0.2)
            .fadeout(0.2)
        )
        subtitle_clips.append(sub_text)

    audio_tracks = [voice_audio]
    if bgm_path and os.path.exists(bgm_path):
        bgm = (
            AudioFileClip(bgm_path)
            .loop(duration=total_duration)
            .fx(afx.volumex, 0.13)
            .audio_fadein(1.5)
            .audio_fadeout(2.0)
        )
        audio_tracks.append(bgm)

    final_audio = CompositeAudioClip(audio_tracks)
    video = CompositeVideoClip([base_clip, title_clip] + subtitle_clips)
    video = video.set_audio(final_audio)

    output_file = os.path.join(OUTPUT_DIR, f"sermon_{aspect_ratio.replace(':', '_')}.mp4")
    video.write_videofile(output_file, fps=24, codec="libx264", audio_codec="aac", threads=4, preset="ultrafast")
    
    voice_audio.close()
    video.close()
    return output_file
