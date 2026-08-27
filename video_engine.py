import asyncio
import os
import edge_tts
from moviepy.editor import (
    VideoFileClip, ImageClip, ColorClip, TextClip,
    AudioFileClip, CompositeAudioClip, CompositeVideoClip, afx
)

OUTPUT_DIR = "./outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

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
        if bg_media_path.lower().endswith(('.mp4', '.mov', '.avi')):
            bg = VideoFileClip(bg_media_path).resize((width, height)).loop(duration=total_duration)
        else:
            bg = ImageClip(bg_media_path).set_duration(total_duration).resize((width, height))
        dim_layer = ColorClip(size=(width, height), color=(0, 0, 0), duration=total_duration).set_opacity(0.4)
        base_clip = CompositeVideoClip([bg, dim_layer])
    else:
        base_clip = ColorClip(size=(width, height), color=(15, 23, 42), duration=total_duration)

    title_clip = (
        TextClip(
            title,
            fontsize=48 if aspect_ratio == "9:16" else 40,
            color="#FDE047",
            font=font_name,
            method="caption",
            size=(width - 160, None)
        )
        .set_position(("center", 180 if aspect_ratio == "9:16" else 90))
        .set_duration(total_duration)
        .fadein(0.5)
    )

    segment_duration = total_duration / max(len(script_paragraphs), 1)
    subtitle_clips = []

    for idx, sentence in enumerate(script_paragraphs):
        start_time = idx * segment_duration
        sub_text = (
            TextClip(
                sentence.strip(),
                fontsize=42 if aspect_ratio == "9:16" else 36,
                color="white",
                font=font_name,
                method="caption",
                size=(width - 180, None),
                stroke_color="black",
                stroke_width=1.5
            )
            .set_start(start_time)
            .set_duration(segment_duration)
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
