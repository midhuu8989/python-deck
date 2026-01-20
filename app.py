# -------------------------------------------------------------
# Streamlit App: PPT → Fully Narrated Voice PPT (ELEVENLABS SAFE)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path
import requests

import streamlit as st
from pptx import Presentation
from pptx.util import Inches

from dotenv import load_dotenv
from groq import Groq
from openai import OpenAI

# ===================== ENV ========================
load_dotenv()

OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY") or os.getenv("GROQ_API_KEY")
ELEVENLABS_API_KEY = st.secrets.get("ELEVENLABS_API_KEY") or os.getenv("ELEVENLABS_API_KEY")

if not ELEVENLABS_API_KEY:
    st.error("❌ ElevenLabs API key missing")
    st.stop()

openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
groq_client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

# ================= UI =============================
st.set_page_config(page_title="PPT Voice Over Studio", layout="wide")
st.title("🎤 PPT Voice Over Studio")
st.caption("ElevenLabs • Error-safe • PPT always downloadable")

st.divider()

# ================= SIDEBAR ========================
st.sidebar.header("🎙 Voice Settings")

voice_type = st.sidebar.radio("Voice Type", ["Female", "Male"])

VOICE_ID_MAP = {
    "Female": "21m00Tcm4TlvDq8ikWAM",  # Rachel
    "Male": "29vD33N1CtxCmqQRPOHJ",    # Adam
}

selected_voice_id = VOICE_ID_MAP[voice_type]

# ================= SESSION ========================
if "slides" not in st.session_state:
    st.session_state.slides = []
if "ppt_loaded" not in st.session_state:
    st.session_state.ppt_loaded = False

# ================= HELPERS ========================
def get_slide_title(slide):
    try:
        if slide.shapes.title:
            return slide.shapes.title.text.strip()
    except Exception:
        pass
    return ""

def create_silent_mp3(path: Path):
    silent_mp3_bytes = (
        b"\xFF\xFB\x90\x64\x00\x0F\xFF\xFA\x92\x40\x00\x0F"
    )
    with open(path, "wb") as f:
        f.write(silent_mp3_bytes)

# ================= LLM ============================
def call_llm(prompt: str) -> str:
    if openai_client:
        try:
            r = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
            )
            return r.choices[0].message.content.strip()
        except Exception:
            pass

    if groq_client:
        r = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
        )
        return r.choices[0].message.content.strip()

    return "In this slide we are going to look into the given topic."

def generate_intro(title):
    return call_llm(
        f"""Generate narration.
Start exactly with: Today we are going to explore on {title}
Simple Indian teaching tone.
3 to 4 sentences."""
    )

def generate_narration(text):
    return call_llm(
        f"""Generate narration.
Start exactly with: In this slide we are going to look into
Simple Indian teaching tone.
{text}"""
    )

# ================= ELEVENLABS TTS =================
def generate_audio(text: str, out_mp3: Path):
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{selected_voice_id}"
    headers = {
        "xi-api-key": ELEVENLABS_API_KEY,
        "Content-Type": "application/json",
        "Accept": "audio/mpeg",
    }

    payload = {
        "text": text,
        "model_id": "eleven_monolingual_v1",
        "voice_settings": {"stability": 0.6, "similarity_boost": 0.75},
    }

    try:
        response = requests.post(url, json=payload, headers=headers, timeout=60)
        if response.status_code == 200 and response.content:
            with open(out_mp3, "wb") as f:
                f.write(response.content)
        else:
            create_silent_mp3(out_mp3)
    except Exception:
        create_silent_mp3(out_mp3)

def add_audio_to_slide(slide, audio_path):
    slide.shapes.add_movie(
        movie_file=str(audio_path),
        left=Inches(0.2),
        top=Inches(0.2),
        width=Inches(1),
        height=Inches(1),
        mime_type="audio/mpeg",
    )

# ================= UPLOAD =========================
ppt_file = st.file_uploader("📤 Upload PPTX", type=["pptx"])

if ppt_file and not st.session_state.ppt_loaded:
    temp_dir = Path(tempfile.mkdtemp())
    ppt_path = temp_dir / ppt_file.name
    ppt_path.write_bytes(ppt_file.read())

    prs = Presentation(ppt_path)
    st.session_state.slides.clear()

    for i, slide in enumerate(prs.slides):
        slide_text = " ".join(
            shape.text for shape in slide.shapes if hasattr(shape, "text")
        ).strip()

        title = get_slide_title(slide)

        notes = generate_intro(title) if i == 0 and title else generate_narration(slide_text or title)

        st.session_state.slides.append({"index": i, "notes": notes})

    st.session_state.ppt_loaded = True
    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_name = ppt_file.name

    st.success("✅ Narration generated")

# ================= FINAL ==========================
if st.session_state.ppt_loaded:
    if st.button("📥 Download PPT with Voice-over"):
        prs = Presentation(st.session_state.ppt_path)
        out_dir = Path(tempfile.mkdtemp())

        for i, slide_data in enumerate(st.session_state.slides):
            slide = prs.slides[slide_data["index"]]
            mp3 = out_dir / f"slide_{i}.mp3"

            generate_audio(slide_data["notes"], mp3)
            add_audio_to_slide(slide, mp3)

            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_data["notes"]

        final_ppt = out_dir / st.session_state.ppt_name
        prs.save(final_ppt)

        st.download_button(
            "⬇ Download PPT with Voice-over",
            final_ppt.read_bytes(),
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
