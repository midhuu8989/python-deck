# -------------------------------------------------------------
# Streamlit App: PPT → Fully Narrated Voice PPT (STABLE FINAL)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.util import Inches

from dotenv import load_dotenv
from groq import Groq
from openai import OpenAI
from gtts import gTTS

# ===================== ENV ========================
load_dotenv()

OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY") or os.getenv("GROQ_API_KEY")

openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
groq_client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

if not openai_client and not groq_client:
    st.error("❌ No LLM key configured")
    st.stop()

# ================= UI =============================
st.set_page_config(page_title="PPT Voice Over Studio", layout="wide")
st.title("🎤 PPT Voice Over Studio")
st.caption("Notes-safe • Cloud-safe • All slides narrated")

st.divider()

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
            pass  # silent fallback

    if groq_client:
        r = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
        )
        return r.choices[0].message.content.strip()

    raise RuntimeError("No LLM available")

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

# ================= TTS ============================
def generate_audio(text: str, out_mp3: Path):
    tts = gTTS(text=text, lang="en", slow=False)
    tts.save(str(out_mp3))

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

        if i == 0 and title:
            notes = generate_intro(title)
        else:
            notes = generate_narration(slide_text or title)

        st.session_state.slides.append({
            "index": i,
            "notes": notes,
        })

    st.session_state.ppt_loaded = True
    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_name = ppt_file.name

    st.success("✅ Narration generated for all slides")

# ================= PREVIEW ========================
if st.session_state.ppt_loaded:
    st.subheader("🎧 Preview Narration")

    for slide in st.session_state.slides:
        with st.expander(f"Slide {slide['index'] + 1}"):
            slide["notes"] = st.text_area(
                "Narration",
                slide["notes"],
                key=f"n_{slide['index']}",
            )

            if st.button("▶ Preview", key=f"p_{slide['index']}"):
                tmp = Path(tempfile.mktemp(suffix=".mp3"))
                generate_audio(slide["notes"], tmp)
                st.audio(str(tmp))

# ================= FINAL ==========================
st.divider()

if st.session_state.ppt_loaded:
    if st.button("📥 Download PPT with Voice-over"):
        prs = Presentation(st.session_state.ppt_path)
        out_dir = Path(tempfile.mkdtemp())

        progress = st.progress(0.0)
        total = len(st.session_state.slides)

        for i, slide_data in enumerate(st.session_state.slides):
            slide = prs.slides[slide_data["index"]]
            mp3 = out_dir / f"slide_{i}.mp3"

            generate_audio(slide_data["notes"], mp3)
            add_audio_to_slide(slide, mp3)

            # ✅ EXPLICIT NOTES CREATION (NO ERROR)
            if slide.notes_slide is None:
                notes_slide = slide.notes_slide
            else:
                notes_slide = slide.notes_slide

            notes_slide.notes_text_frame.text = slide_data["notes"]

            progress.progress((i + 1) / total)

        final_ppt = out_dir / st.session_state.ppt_name
        prs.save(final_ppt)

        st.download_button(
            "⬇ Download Narrated PPT",
            final_ppt.read_bytes(),
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
