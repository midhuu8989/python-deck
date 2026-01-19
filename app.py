# streamlit_app.py
# -------------------------------------------------------------
# Streamlit App: PPT Upload -> Review/Edit -> Narration ->
# 1) PPT with embedded voice-over
# 2) MP4 video with narration
# (OpenAI LLM + OpenAI TTS | Local + Streamlit Cloud Safe)
# -------------------------------------------------------------

import os
import tempfile
import subprocess
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pydub import AudioSegment

from dotenv import load_dotenv
from openai import OpenAI

# ===================== ENV SETUP ===================
load_dotenv()
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    st.error("OPENAI_API_KEY not configured")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

# ================= UI SETUP ======================
st.set_page_config(page_title="PPT Narration Studio", layout="wide")
st.title("🎬 PPT Narration Studio")
st.caption("Generate PPT with voice-over and MP4 video")

# ================= SESSION STATE =================
if "slides_data" not in st.session_state:
    st.session_state.slides_data = []
if "ppt_loaded" not in st.session_state:
    st.session_state.ppt_loaded = False
if "ppt_path" not in st.session_state:
    st.session_state.ppt_path = None
if "ppt_filename" not in st.session_state:
    st.session_state.ppt_filename = None

# ================= FILE UPLOAD ===================
uploaded_file = st.file_uploader("Upload PPTX", type=["pptx"])

# ================= HELPERS =======================
def generate_notes(slide_text: str) -> str:
    prompt = f"Create a clear narration script for this slide:\n{slide_text}"
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.choices[0].message.content.strip()


def openai_tts(text: str, out_mp3: Path):
    with client.audio.speech.with_streaming_response.create(
        model="gpt-4o-mini-tts",
        voice="alloy",
        input=text,
    ) as response:
        response.stream_to_file(out_mp3)


def embed_audio_in_slide(slide, audio_path: Path):
    """Best-effort audio embedding into PPT slide"""
    with open(audio_path, "rb") as f:
        audio_bytes = f.read()

    audio_part = slide.part.package._add_media_part(
        audio_bytes, content_type="audio/mpeg"
    )
    rId = slide.part.relate_to(audio_part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio")

    audio_xml = f"""
    <p:pic {nsdecls('a','p','r')}>
      <p:nvPicPr>
        <p:cNvPr id="999" name="Narration"/>
        <p:cNvPicPr/>
        <p:nvPr>
          <a:audioFile r:link="{rId}"/>
        </p:nvPr>
      </p:nvPicPr>
      <p:blipFill/>
      <p:spPr/>
    </p:pic>
    """

    slide.shapes._spTree.insert_element_before(parse_xml(audio_xml), 'p:extLst')

# ================= LOAD PPT ======================
if uploaded_file and not st.session_state.ppt_loaded:
    temp_dir = Path(tempfile.mkdtemp())
    ppt_path = temp_dir / uploaded_file.name
    ppt_path.write_bytes(uploaded_file.read())

    prs = Presentation(ppt_path)
    st.session_state.slides_data.clear()

    for idx, slide in enumerate(prs.slides):
        slide_text = " ".join(
            shape.text for shape in slide.shapes if hasattr(shape, "text")
        ).strip()

        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        if not notes:
            notes = generate_notes(slide_text)

        st.session_state.slides_data.append({
            "index": idx,
            "slide_text": slide_text,
            "notes": notes,
            "audio": None,
            "duration": 0
        })

    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_filename = uploaded_file.name
    st.session_state.ppt_loaded = True

# ================= REVIEW & EDIT =================
if st.session_state.ppt_loaded:
    st.subheader("📝 Review & Edit Narration")

    for slide in st.session_state.slides_data:
        with st.expander(f"Slide {slide['index'] + 1}"):
            st.write(slide["slide_text"] or "(No text)")
            slide["notes"] = st.text_area(
                "Narration",
                slide["notes"],
                key=f"notes_{slide['index']}"
            )

    st.divider()

    if st.button("🎤 Generate Voice-over + Export PPT & MP4"):
        st.info("Generating narration audio…")

        workdir = Path(tempfile.mkdtemp())
        prs = Presentation(st.session_state.ppt_path)

        audio_files = []
        image_files = []

        for slide_data in st.session_state.slides_data:
            idx = slide_data["index"]
            slide = prs.slides[idx]

            mp3_path = workdir / f"slide_{idx}.mp3"
            openai_tts(slide_data["notes"], mp3_path)

            audio = AudioSegment.from_mp3(mp3_path)
            slide_data["duration"] = audio.duration_seconds
            slide_data["audio"] = mp3_path

            embed_audio_in_slide(slide, mp3_path)
            slide.notes_slide.notes_text_frame.text = slide_data["notes"]

            audio_files.append(mp3_path)

        # -------- SAVE PPT WITH VOICE-OVER --------
        ppt_out = workdir / st.session_state.ppt_filename
        prs.save(ppt_out)

        # -------- GENERATE VIDEO (MP4) --------
        st.info("Rendering MP4 video…")

        slides_dir = workdir / "slides"
        slides_dir.mkdir(exist_ok=True)

        # Convert PPT to images using LibreOffice (must exist on system)
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "png",
            str(ppt_out), "--outdir", str(slides_dir)
        ], check=False)

        slide_images = sorted(slides_dir.glob("*.png"))

        concat_file = workdir / "concat.txt"
        with open(concat_file, "w") as f:
            for i, img in enumerate(slide_images):
                f.write(f"file '{img.as_posix()}'\n")
                f.write(f"duration {st.session_state.slides_data[i]['duration'] + 0.5}\n")

        mp4_out = workdir / st.session_state.ppt_filename.replace('.pptx', '.mp4')

        subprocess.run([
            "ffmpeg", "-y", "-f", "concat", "-safe", "0",
            "-i", str(concat_file), "-vsync", "vfr",
            str(mp4_out)
        ], check=False)

        st.success("✅ PPT with voice-over and MP4 generated")

        st.download_button(
            "⬇ Download PPT (with narration)",
            ppt_out.read_bytes(),
            file_name=st.session_state.ppt_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        st.download_button(
            "⬇ Download MP4",
            mp4_out.read_bytes(),
            file_name=mp4_out.name,
            mime="video/mp4"
        )
