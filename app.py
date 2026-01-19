# -------------------------------------------------------------
# Streamlit App: PPT → Voice-over Preview → Download PPT
# (OpenAI LLM + OpenAI TTS | Indian narration style)
# -------------------------------------------------------------

# ===================== IMPORTS =====================
import os
import tempfile
import time
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pydub import AudioSegment

from dotenv import load_dotenv
from openai import OpenAI

# ===================== ENV ========================
load_dotenv()
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    st.error("❌ OPENAI_API_KEY not configured")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

# ================= UI SETUP ======================
st.set_page_config(page_title="PPT Voice Over Studio", layout="wide")
st.title("🎤 PPT Voice Over Studio")
st.caption("Preview voice per slide • Generate PPT with Indian-style narration")

st.divider()

# ================= SESSION STATE =================
if "slides" not in st.session_state:
    st.session_state.slides = []
if "ppt_loaded" not in st.session_state:
    st.session_state.ppt_loaded = False
if "ppt_path" not in st.session_state:
    st.session_state.ppt_path = None
if "ppt_name" not in st.session_state:
    st.session_state.ppt_name = None

# ================= HELPERS =======================
def generate_narration(slide_text: str, slide_index: int) -> str:
    """
    Slide narration rules:
    - Slide 1: Today we are going to start with <topic>
    - Other slides: In this slide we are going to look into
    """
    if slide_index == 0:
        prefix = "Today we are going to start with "
    else:
        prefix = "In this slide we are going to look into "

    prompt = f"""
Generate narration for self-directed learning.
Rules:
- Start exactly with: "{prefix}"
- Use simple Indian teaching tone
- No headings
- No bullet points

Slide content:
{slide_text}
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content.strip()


def openai_tts(text: str, out_mp3: Path):
    """
    OpenAI TTS – neutral voice, Indian-style narration via text
    """
    with client.audio.speech.with_streaming_response.create(
        model="gpt-4o-mini-tts",
        voice="alloy",
        input=text,
    ) as response:
        response.stream_to_file(out_mp3)


def add_audio_to_slide(slide, audio_path: Path):
    """
    Official python-pptx supported audio embedding
    """
    slide.shapes.add_movie(
        movie_file=str(audio_path),
        left=Inches(0.3),
        top=Inches(0.3),
        width=Inches(1),
        height=Inches(1),
        mime_type="audio/mpeg",
    )

# ================= FILE UPLOAD ====================
ppt_file = st.file_uploader("📤 Upload PPTX", type=["pptx"])

if ppt_file and not st.session_state.ppt_loaded:
    st.info("📄 PPT uploaded. Reading slides…")

    workdir = Path(tempfile.mkdtemp())
    ppt_path = workdir / ppt_file.name
    ppt_path.write_bytes(ppt_file.read())

    prs = Presentation(ppt_path)
    st.session_state.slides.clear()

    for idx, slide in enumerate(prs.slides):
        slide_text = " ".join(
            shape.text for shape in slide.shapes if hasattr(shape, "text")
        ).strip()

        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()

        if not notes:
            notes = generate_narration(slide_text, idx)

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
    st.success("✅ PPT loaded successfully")

# ================= PREVIEW ========================
if st.session_state.ppt_loaded:
    st.subheader("🎧 Preview Voice per Slide")

    for slide in st.session_state.slides:
        with st.expander(f"Slide {slide['index'] + 1}", expanded=False):
            st.write(slide["text"] or "_No visible text_")

            slide["notes"] = st.text_area(
                "Narration Text",
                slide["notes"],
                key=f"notes_{slide['index']}",
                height=120,
            )

            if st.button("▶ Preview Voice", key=f"preview_{slide['index']}"):
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                    with st.spinner("Generating voice…"):
                        openai_tts(slide["notes"], Path(f.name))
                        audio = AudioSegment.from_mp3(f.name)
                        slide["audio"] = f.name
                        slide["duration"] = audio.duration_seconds
                    st.audio(f.name)

# ================= FINAL GENERATION =================
st.divider()

if st.session_state.ppt_loaded:
    if st.button("📥 Generate & Download PPT with Voice-over"):
        prs = Presentation(st.session_state.ppt_path)
        outdir = Path(tempfile.mkdtemp())

        total = len(st.session_state.slides)
        progress = st.progress(0.0)
        status = st.empty()

        for idx, slide_data in enumerate(st.session_state.slides, start=1):
            status.info(f"🔄 Generating voice for slide {idx} of {total}")

            slide = prs.slides[slide_data["index"]]
            mp3_path = outdir / f"slide_{slide_data['index']}.mp3"

            openai_tts(slide_data["notes"], mp3_path)
            add_audio_to_slide(slide, mp3_path)
            slide.notes_slide.notes_text_frame.text = slide_data["notes"]

            progress.progress(idx / total)
            time.sleep(0.1)

        status.success("✅ Voice-over added to all slides")

        final_ppt = outdir / st.session_state.ppt_name
        prs.save(final_ppt)

        st.download_button(
            "⬇ Download PPT with Voice-over",
            final_ppt.read_bytes(),
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
