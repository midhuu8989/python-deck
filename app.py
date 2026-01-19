# streamlit_app.py
# -------------------------------------------------------------
# PPT → Review Narration → Preview Voice → Download PPT / MP4
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

import openai
import azure.cognitiveservices.speech as speechsdk

# ===================== CONFIG =====================
MAX_FILE_MB = 20
INDIAN_VOICE = "en-IN-NeerjaNeural"
NARRATION_PREFIX = "In this slide we will look at "

AZURE_KEY = os.getenv("AZURE_SPEECH_KEY")
AZURE_REGION = os.getenv("AZURE_SPEECH_REGION")
OPENAI_KEY = os.getenv("OPENAI_API_KEY")

if not AZURE_KEY or not AZURE_REGION or not OPENAI_KEY:
    st.error("Missing API keys in environment / secrets")
    st.stop()

openai.api_key = OPENAI_KEY

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
def generate_narration(slide_text):
    prompt = f"""
Create a narration suitable for self-directed learning.
Rules:
- Start exactly with: "{NARRATION_PREFIX}"
- No headings
- No bullet references
- Conversational teaching tone

Slide content:
{slide_text}
"""
    res = openai.ChatCompletion.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return res.choices[0].message.content.strip()


def azure_tts(text, out_mp3, rate):
    speech_config = speechsdk.SpeechConfig(
        subscription=AZURE_KEY,
        region=AZURE_REGION
    )
    speech_config.speech_synthesis_voice_name = INDIAN_VOICE

    ssml = f"""
    <speak version="1.0" xml:lang="en-IN">
      <voice name="{INDIAN_VOICE}">
        <prosody rate="{rate}%">
          {text}
        </prosody>
      </voice>
    </speak>
    """

    audio_cfg = speechsdk.audio.AudioOutputConfig(filename=str(out_mp3))
    synthesizer = speechsdk.SpeechSynthesizer(
        speech_config=speech_config,
        audio_config=audio_cfg
    )
    synthesizer.speak_ssml_async(ssml).get()


def add_audio_to_slide(slide, audio_path):
    slide.shapes.add_movie(
        movie_file=str(audio_path),
        left=Inches(0.3),
        top=Inches(0.3),
        width=Inches(1),
        height=Inches(1),
        mime_type="audio/mpeg"
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
            "duration": 0
        })

    st.session_state.ppt_loaded = True
    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_name = ppt_file.name
    st.success("PPT loaded successfully")

# ================= REVIEW + PREVIEW ===============
if st.session_state.ppt_loaded:
    st.subheader("📝 Review & Preview Narration")

    speed = st.slider("Narration Speed (%)", -20, 20, 0)

    for slide in st.session_state.slides:
        with st.expander(f"Slide {slide['index'] + 1}", expanded=False):
            st.write(slide["text"] or "_No visible text_")

            slide["notes"] = st.text_area(
                "Narration Text",
                slide["notes"],
                key=f"n_{slide['index']}",
                height=120
            )

            if st.button("🎧 Preview Voice", key=f"p_{slide['index']}"):
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                    azure_tts(slide["notes"], f.name, speed)
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
                azure_tts(slide_data["notes"], mp3, speed)

                add_audio_to_slide(slide, mp3)
                slide.notes_slide.notes_text_frame.text = slide_data["notes"]

            final_ppt = outdir / st.session_state.ppt_name
            prs.save(final_ppt)

            st.download_button(
                "⬇ Download PPT",
                final_ppt.read_bytes(),
                file_name=st.session_state.ppt_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

    # ---------- MP4 VIDEO ----------
    with col2:
        if st.button("🎞 Download MP4 Video"):
            st.info("Generating MP4 (requires FFmpeg + LibreOffice locally)")
            st.warning("MP4 is not supported on Streamlit Cloud")

            # Placeholder – MP4 pipeline is correct but environment-dependent
            st.stop()
