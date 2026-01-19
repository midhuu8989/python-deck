# -------------------------------------------------------------
# Streamlit App: PPT Upload -> Edit Narration -> PPT with Audio
# (Cloud-safe | OpenAI TTS | NO video | NO private APIs)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.util import Inches

from dotenv import load_dotenv
from openai import OpenAI

# ================= ENV SETUP =====================
load_dotenv()

OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    st.error("❌ OPENAI_API_KEY not configured")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

# ================= UI ============================
st.set_page_config(page_title="PPT Voice Narration", layout="wide")
st.title("🎤 PPT Voice Narration Generator")
st.caption("Generates PPT with embedded voice-over (audio per slide)")

# ================= SESSION STATE =================
if "slides" not in st.session_state:
    st.session_state.slides = []

if "ppt_path" not in st.session_state:
    st.session_state.ppt_path = None

if "ppt_name" not in st.session_state:
    st.session_state.ppt_name = None

# ================= HELPERS =======================
def generate_narration(text: str) -> str:
    prompt = f"Create a clear narration script for this slide:\n{text}"
    res = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return res.choices[0].message.content.strip()


def text_to_speech(text: str, out_mp3: Path):
    with client.audio.speech.with_streaming_response.create(
        model="gpt-4o-mini-tts",
        voice="alloy",
        input=text,
    ) as response:
        response.stream_to_file(out_mp3)


def add_audio_to_slide(slide, audio_path: Path):
    """
    SAFE method: PowerPoint treats audio as a movie object.
    This is the only supported way in python-pptx.
    """
    slide.shapes.add_movie(
        movie_file=str(audio_path),
        left=Inches(0.5),
        top=Inches(0.5),
        width=Inches(1),
        height=Inches(1),
        poster_frame_image=None,
        mime_type="audio/mpeg",
    )

# ================= UPLOAD PPT ====================
uploaded = st.file_uploader("Upload PPTX", type=["pptx"])

if uploaded:
    temp_dir = Path(tempfile.mkdtemp())
    ppt_path = temp_dir / uploaded.name
    ppt_path.write_bytes(uploaded.read())

    prs = Presentation(ppt_path)
    st.session_state.slides.clear()

    for i, slide in enumerate(prs.slides):
        text = " ".join(
            s.text for s in slide.shapes if hasattr(s, "text")
        ).strip()

        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()

        if not notes:
            notes = generate_narration(text)

        st.session_state.slides.append({
            "index": i,
            "text": text,
            "notes": notes,
        })

    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_name = uploaded.name
    st.success("✅ PPT loaded")

# ================= EDIT UI =======================
if st.session_state.slides:
    st.subheader("✏️ Edit Narration")

    for s in st.session_state.slides:
        with st.expander(f"Slide {s['index'] + 1}"):
            st.write(s["text"] or "_No visible text_")
            s["notes"] = st.text_area(
                "Narration",
                s["notes"],
                key=f"n_{s['index']}",
                height=120,
            )

    st.divider()

    if st.button("🎧 Generate PPT with Voice-over"):
        st.info("Generating narration audio...")

        prs = Presentation(st.session_state.ppt_path)
        work = Path(tempfile.mkdtemp())

        for s in st.session_state.slides:
            slide = prs.slides[s["index"]]

            mp3 = work / f"slide_{s['index']}.mp3"
            text_to_speech(s["notes"], mp3)

            add_audio_to_slide(slide, mp3)
            slide.notes_slide.notes_text_frame.text = s["notes"]

        final_ppt = work / st.session_state.ppt_name
        prs.save(final_ppt)

        st.success("✅ PPT with voice-over ready")

        st.download_button(
            "⬇ Download PPT (with narration)",
            final_ppt.read_bytes(),
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
