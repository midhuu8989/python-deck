# streamlit_app.py
# -------------------------------------------------------------
# PPT → Review Narration → Preview Voice → Download PPT / MP4
# (NO AZURE | OpenAI LLM + OpenAI TTS)
# -------------------------------------------------------------

import os
import tempfile
import time
import subprocess
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pydub import AudioSegment

from dotenv import load_dotenv
from openai import OpenAI

# ===================== CONFIG =====================
MAX_FILE_MB = 20
NARRATION_PREFIX = "In this slide we will look at "

# ===================== ENV ========================
load_dotenv()
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    st.error("Missing OPENAI_API_KEY")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

# ================= UI =============================
st.set_page_config("PPT Narration Studio", layout="wide")
st.title("🎬 PPT Narration Studio")
st.caption("Preview narration per slide • Download PPT with voice-over or MP4")

# ================= SESSION STATE ==================
if "slides" not in st.session_state:
    st.session_state.slides = []
if "ppt_loaded" not in st.session_state:
    st.session_state.ppt_loaded = False
if "ppt_path" not in st.session_state:
    st.session_state.ppt_path = None
if "ppt_name" not in st.session_state:
    st.session_state.ppt_name = None

# ================= HELPERS ========================
def generate_narration(slide_text: str) -> str:
    prompt = f"""
Create a narration suitable for self-directed learning.
Rules:
- Start exactly with: \"{NARRATION_PREFIX}\"
- No headings
- No bullet references
- Conversational teaching tone

Slide content:
{slide_text}
"""
    res = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return res.choices[0].message.content.strip()


def openai_tts(text: str, out_mp3: Path, speed: int):
    # OpenAI TTS does not have SSML; speed is approximated via instruction
    paced_text = f"Speak at {speed}% speed. {text}"

    with client.audio.speech.with_streaming_response.create(
        model="gpt-4o-mini-tts",
        voice="alloy",
        input=paced_text,
    ) as response:
        response.stream_to_file(out_mp3)


def add_audio_to_slide(slide, audio_path: Path):
    # Official python-pptx supported method (audio treated as movie)
    slide.shapes.add_movie(
        movie_file=str(audio_path),
        left=Inches(0.3),
        top=Inches(0.3),
        width=Inches(1),
        height=Inches(1),
        mime_type="audio/mpeg",
    )

# ================= FILE UPLOAD ====================
ppt_file = st.file_uploader("Upload PPTX", type=["pptx"])

if ppt_file and not st.session_state.ppt_loaded:
    if ppt_file.size > MAX_FILE_MB * 1024 * 1024:
        st.error("File too large")
        st.stop()

    workdir = Path(tempfile.mkdtemp())
    ppt_path = workdir / ppt_file.name
    ppt_path.write_bytes(ppt_file.read())

    prs = Presentation(ppt_path)
    st.session_state.slides.clear()

    for idx, slide in enumerate(prs.slides):
        slide_text = " ".join(
            shape.text for shape in slide.shapes if hasattr(shape, "text")
        )

        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()

        if not notes:
            notes = generate_narration(slide_text)

        st.session_state.slides.append({
            "index": idx,
            "text": slide_text,
            "notes": notes,
            "audio": None,
            "duration": 0,
        })

    st.session_state.ppt_loaded = True
    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_name = ppt_file.name
    st.success("PPT loaded successfully")

# ================= REVIEW + PREVIEW ===============
if st.session_state.ppt_loaded:
    st.subheader("📝 Review & Preview Narration")

    speed = st.slider("Narration Speed (%)", 80, 120, 100)

    for slide in st.session_state.slides:
        with st.expander(f"Slide {slide['index'] + 1}", expanded=False):
            st.write(slide["text"] or "_No visible text_")

            slide["notes"] = st.text_area(
                "Narration Text",
                slide["notes"],
                key=f"n_{slide['index']}",
                height=120,
            )

            if st.button("🎧 Preview Voice", key=f"p_{slide['index']}"):
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                    openai_tts(slide["notes"], Path(f.name), speed)
                    audio = AudioSegment.from_mp3(f.name)
                    slide["audio"] = f.name
                    slide["duration"] = audio.duration_seconds
                    st.audio(f.name)

# ================= FINAL GENERATION ===============
st.divider()

if st.session_state.ppt_loaded:
    col1, col2 = st.columns(2)

    # ---------- PPT WITH VOICE ----------
    with col1:
        if st.button("📥 Download PPT with Voice-over"):
            prs = Presentation(st.session_state.ppt_path)
            outdir = Path(tempfile.mkdtemp())

            for slide_data in st.session_state.slides:
                slide = prs.slides[slide_data["index"]]

                mp3 = outdir / f"slide_{slide_data['index']}.mp3"
                openai_tts(slide_data["notes"], mp3, speed)

                add_audio_to_slide(slide, mp3)
                slide.notes_slide.notes_text_frame.text = slide_data["notes"]

            final_ppt = outdir / st.session_state.ppt_name
            prs.save(final_ppt)

            st.download_button(
                "⬇ Download PPT",
                final_ppt.read_bytes(),
                file_name=st.session_state.ppt_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

    # ---------- MP4 VIDEO ----------
    with col2:
        if st.button("🎞 Download MP4 Video"):
            st.info("MP4 generation requires FFmpeg + LibreOffice locally")
            st.warning("MP4 is not supported on Streamlit Cloud")
            st.stop()
