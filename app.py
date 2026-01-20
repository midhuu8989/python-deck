# -------------------------------------------------------------
# PPT → Notes → Narration → Voice → Download (SELF LEARNING)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path
import requests

import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from dotenv import load_dotenv
from openai import OpenAI
from groq import Groq

# ================= ENV =================
load_dotenv()

ELEVENLABS_API_KEY = st.secrets.get("ELEVENLABS_API_KEY") or os.getenv("ELEVENLABS_API_KEY")
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
GROQ_API_KEY = st.secrets.get("GROQ_API_KEY") or os.getenv("GROQ_API_KEY")

if not ELEVENLABS_API_KEY:
    st.error("❌ ElevenLabs API key missing")
    st.stop()

openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
groq_client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

# ================= UI =================
st.set_page_config(page_title="PPT Self Learning Narrator", layout="wide")
st.title("🎓 PPT Self Learning Narrator")
st.caption("Uses Notes → Generates narration → Embeds voice → Download PPT")

# ================= HELPERS =================
def get_slide_text(slide):
    texts = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            texts.append(shape.text.strip())
    return " ".join(texts)

def get_notes_text_frame(slide):
    return slide.notes_slide.notes_text_frame  # auto-creates

def read_notes_safely(notes_tf):
    try:
        if hasattr(notes_tf, "paragraphs"):
            return " ".join(p.text for p in notes_tf.paragraphs).strip()
    except Exception:
        pass
    return ""

def write_notes_safely(notes_tf, text):
    # ✅ SAFE WRITE (FIXES YOUR ERROR)
    notes_tf.clear()
    p = notes_tf.add_paragraph()
    p.text = text

def call_llm(prompt):
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

    return "This slide explains the given topic in simple terms."

def generate_slide1_narration(title):
    return call_llm(
        f"""
Today we are going to explore on {title}.
Explain what this topic is.
Explain where it is used in real life.
Simple Indian teaching tone.
3 to 4 sentences.
"""
    )

def generate_narration(text):
    return call_llm(
        f"""
Create narration for self learning.
Simple Indian teaching tone.
3 to 4 sentences.
Content:
{text}
"""
    )

def create_silent_mp3(path: Path):
    silent = b"\xFF\xFB\x90\x64\x00\x0F\xFF\xFA\x92\x40\x00\x0F"
    with open(path, "wb") as f:
        f.write(silent)

def elevenlabs_tts(text, out_mp3):
    url = "https://api.elevenlabs.io/v1/text-to-speech/21m00Tcm4TlvDq8ikWAM"
    headers = {
        "xi-api-key": ELEVENLABS_API_KEY,
        "Content-Type": "application/json",
        "Accept": "audio/mpeg",
    }
    payload = {"text": text, "model_id": "eleven_monolingual_v1"}

    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        if r.status_code == 200 and r.content:
            with open(out_mp3, "wb") as f:
                f.write(r.content)
        else:
            create_silent_mp3(out_mp3)
    except Exception:
        create_silent_mp3(out_mp3)

def add_audio(slide, audio_path):
    slide.shapes.add_movie(
        movie_file=str(audio_path),
        left=Inches(0.2),
        top=Inches(0.2),
        width=Inches(1),
        height=Inches(1),
        mime_type="audio/mpeg",
    )

# ================= MAIN =================
ppt_file = st.file_uploader("📤 Upload PPTX", type=["pptx"])

if ppt_file:
    workdir = Path(tempfile.mkdtemp())
    ppt_path = workdir / ppt_file.name
    ppt_path.write_bytes(ppt_file.read())

    prs = Presentation(ppt_path)

    slide_data = []

    for idx, slide in enumerate(prs.slides):
        notes_tf = get_notes_text_frame(slide)
        existing_notes = read_notes_safely(notes_tf)

        if not existing_notes:
            if idx == 0:
                title = slide.shapes.title.text if slide.shapes.title else "this topic"
                narration = generate_slide1_narration(title)
            else:
                narration = generate_narration(get_slide_text(slide))
        else:
            narration = existing_notes

        slide_data.append((idx, narration))

    for idx, narration in slide_data:
        with st.expander(f"Slide {idx + 1}"):
            updated_text = st.text_area(
                "Narration / Notes",
                narration,
                height=120,
                key=f"note_{idx}",
            )

            if st.button("▶ Preview Voice", key=f"preview_{idx}"):
                tmp = workdir / f"preview_{idx}.mp3"
                elevenlabs_tts(updated_text, tmp)
                st.audio(str(tmp))

            slide_data[idx] = (idx, updated_text)

    if st.button("📥 Generate & Download PPT with Audio"):
        for idx, narration in slide_data:
            slide = prs.slides[idx]
            notes_tf = get_notes_text_frame(slide)

            # ✅ FIXED LINE
            write_notes_safely(notes_tf, narration)

            audio_path = workdir / f"slide_{idx}.mp3"
            elevenlabs_tts(narration, audio_path)
            add_audio(slide, audio_path)

        final_ppt = workdir / ppt_file.name
        prs.save(final_ppt)

        st.success("✅ PPT generated successfully")
        st.download_button(
            "⬇ Download PPT with Audio",
            final_ppt.read_bytes(),
            file_name=ppt_file.name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
