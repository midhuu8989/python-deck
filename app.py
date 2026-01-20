# -------------------------------------------------------------
# PPT → Narration → Audio → Download PPT (SELF LEARNING)
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
    st.error("ElevenLabs API key missing")
    st.stop()

openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None
groq_client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None

# ================= UI ==================
st.set_page_config(page_title="PPT Self Learning Narrator")
st.title("🎓 PPT Self Learning Narrator")
st.caption("Upload PPT → Auto narration → Audio embedded → Download")

# ================= HELPERS ==================
def get_slide_text(slide):
    texts = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            texts.append(shape.text.strip())
    return " ".join(texts)

def get_notes_text_frame(slide):
    # Always creates notes slide if not present
    return slide.notes_slide.notes_text_frame

def read_notes_text(notes_tf):
    return " ".join(p.text for p in notes_tf.paragraphs).strip()

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

def generate_narration(slide_text):
    return call_llm(
        f"""
Create narration for self learning.
Simple Indian teaching tone.
3 to 4 sentences.
Content:
{slide_text}
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

# ================= MAIN ==================
ppt_file = st.file_uploader("Upload PPTX", type=["pptx"])

if ppt_file:
    workdir = Path(tempfile.mkdtemp())
    ppt_path = workdir / ppt_file.name
    ppt_path.write_bytes(ppt_file.read())

    prs = Presentation(ppt_path)

    for idx, slide in enumerate(prs.slides):
        notes_tf = get_notes_text_frame(slide)

        narration = read_notes_text(notes_tf)

        if not narration:
            slide_text = get_slide_text(slide)
            narration = generate_narration(slide_text)
            notes_tf.text = narration  # safe write

        audio_path = workdir / f"slide_{idx}.mp3"
        elevenlabs_tts(narration, audio_path)
        add_audio(slide, audio_path)

    final_ppt = workdir / ppt_file.name
    prs.save(final_ppt)

    st.success("✅ PPT narration completed")
    st.download_button(
        "⬇ Download PPT with Audio",
        final_ppt.read_bytes(),
        file_name=ppt_file.name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
