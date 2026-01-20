# -------------------------------------------------------------
# Streamlit App: PPT → Voice Preview → Download PPT with Voice
# -------------------------------------------------------------

# ===================== IMPORTS =====================
import os
import tempfile
import time
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.util import Inches

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
st.caption("Title-based narration • Corruption-safe • Cloud-safe")

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
def is_text_clear(text: str) -> bool:
    return bool(text and len(text.strip()) >= 20)


def get_slide_title(slide) -> str:
    try:
        if slide.shapes.title and slide.shapes.title.text.strip():
            return slide.shapes.title.text.strip()
    except Exception:
        pass
    return ""


def generate_slide1_narration(title: str) -> str:
    prompt = f"""
Generate narration for self-directed learning.

Rules:
- Start exactly with: "Today we are going to explore on {title}"
- Simple Indian teaching tone
- Explain briefly what this topic is
- Explain where it is used in real life
- 3 to 4 sentences only
- No headings
- No bullet points
"""
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content.strip()


def generate_narration(slide_text: str, slide_index: int, slide_title: str = "") -> str:
    if slide_index == 0:
        prompt = f"""
Generate narration for self-directed learning.

Rules:
- Start exactly with: "Today we are going to explore on {slide_title}"
- Simple Indian teaching tone
- No headings
- No bullet points
"""
    else:
        prompt = f"""
Generate narration for self-directed learning.

Rules:
- Start exactly with: "In this slide we are going to look into "
- Simple Indian teaching tone
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


# ================= SAFE TTS ======================
def chunk_text(text, max_chars=900):
    chunks, current = [], ""
    for sentence in text.split(". "):
        if len(current) + len(sentence) < max_chars:
            current += sentence + ". "
        else:
            chunks.append(current.strip())
            current = sentence + ". "
    if current.strip():
        chunks.append(current.strip())
    return chunks


def openai_tts(text: str, out_mp3: Path, retries=3):
    chunks = chunk_text(text)

    with open(out_mp3, "wb") as f:
        for chunk in chunks:
            attempt = 0
            while attempt < retries:
                try:
                    with client.audio.speech.with_streaming_response.create(
                        model="gpt-4o-mini-tts",
                        voice="alloy",
                        input=chunk,
                    ) as response:
                        for audio_bytes in response.iter_bytes():
                            f.write(audio_bytes)
                    break
                except Exception:
                    attempt += 1
                    time.sleep(2 * attempt)
                    if attempt == retries:
                        raise


def add_audio_to_slide(slide, audio_path: Path):
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
    workdir = Path(tempfile.mkdtemp())
    ppt_path = workdir / ppt_file.name
    ppt_path.write_bytes(ppt_file.read())

    prs = Presentation(ppt_path)
    st.session_state.slides.clear()

    for idx, slide in enumerate(prs.slides):
        slide_text = " ".join(
            shape.text for shape in slide.shapes if hasattr(shape, "text")
        ).strip()

        slide_title = get_slide_title(slide)

        # 🔥 ONLY CHANGE: SLIDE 1 EXTENDED TITLE-BASED NARRATION
        if idx == 0:
            if slide_title and len(slide_title.strip()) >= 5:
                notes = generate_slide1_narration(slide_title)
                skip = False
            else:
                notes = ""
                skip = True

        elif not is_text_clear(slide_text):
            skip = True
            notes = ""

        else:
            notes = generate_narration(slide_text, idx)
            skip = False

        st.session_state.slides.append({
            "index": idx,
            "text": slide_text or slide_title,
            "notes": notes,
            "skip": skip,
        })

    st.session_state.ppt_loaded = True
    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_name = ppt_file.name
    st.success("✅ PPT loaded successfully")

# ================= PREVIEW ========================
if st.session_state.ppt_loaded:
    st.subheader("🎧 Preview Voice")

    for slide in st.session_state.slides:
        if slide["skip"]:
            continue

        with st.expander(f"Slide {slide['index'] + 1}"):
            slide["notes"] = st.text_area(
                "Narration Text",
                slide["notes"],
                key=f"notes_{slide['index']}",
                height=120,
            )

            if st.button("▶ Preview Voice", key=f"preview_{slide['index']}"):
                with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as f:
                    try:
                        openai_tts(slide["notes"], Path(f.name))
                        st.audio(f.name)
                    except Exception:
                        st.error("⚠️ Voice preview failed. Please retry.")

# ================= FINAL GENERATION =================
st.divider()

if st.session_state.ppt_loaded:
    if st.button("📥 Download PPT with Voice-over"):
        prs = Presentation(st.session_state.ppt_path)
        outdir = Path(tempfile.mkdtemp())

        total = len(st.session_state.slides)
        progress = st.progress(0.0)

        for i, slide_data in enumerate(st.session_state.slides, start=1):
            progress.progress(i / total)

            if slide_data["skip"]:
                continue

            slide = prs.slides[slide_data["index"]]
            mp3_path = outdir / f"slide_{slide_data['index']}.mp3"

            try:
                openai_tts(slide_data["notes"], mp3_path)
                add_audio_to_slide(slide, mp3_path)
            except Exception:
                st.warning(f"⚠️ Audio skipped for slide {slide_data['index'] + 1}")

            try:
                slide.notes_slide.notes_text_frame.text = slide_data["notes"]
            except Exception:
                pass

        final_ppt = outdir / st.session_state.ppt_name
        prs.save(final_ppt)

        st.download_button(
            "⬇ Download PPT with Voice-over",
            final_ppt.read_bytes(),
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
