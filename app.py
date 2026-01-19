# -------------------------------------------------------------
# Streamlit App: PPT Upload -> Edit Narration -> PPT with Voice
# (Cloud-safe | OpenAI only | No FFmpeg / No pydub)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

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
st.caption("Generates PPT with embedded voice-over (Cloud-safe)")

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


def embed_audio(slide, audio_path: Path):
    audio_part = slide.part.package._add_media_part(
        audio_path.read_bytes(),
        content_type="audio/mpeg",
    )

    rId = slide.part.relate_to(
        audio_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
    )

    audio_xml = f"""
    <p:pic {nsdecls('a','p','r')}>
      <p:nvPicPr>
        <p:cNvPr id="1" name="Narration"/>
        <p:cNvPicPr/>
        <p:nvPr>
          <a:audioFile r:link="{rId}"/>
        </p:nvPr>
      </p:nvPicPr>
      <p:blipFill/>
      <p:spPr/>
    </p:pic>
    """

    slide.shapes._spTree.insert_element_before(
        parse_xml(audio_xml), "p:extLst"
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

            embed_audio(slide, mp3)
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

        st.warning(
            "ℹ️ MP4 generation is disabled on Streamlit Cloud. "
            "Run locally or via Docker for video export."
        )
