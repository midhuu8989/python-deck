# streamlit_app.py
# -------------------------------------------------------------
# Streamlit App: PPT Upload -> Review/Edit -> Export PPT
# (OpenAI LLM + OpenAI TTS | Local + Cloud Safe)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path

import streamlit as st
from pptx import Presentation

from dotenv import load_dotenv
from openai import OpenAI

# ===================== ENV SETUP ===================
# Load .env for LOCAL use (ignored on Streamlit Cloud)
load_dotenv()

# Read key from Streamlit Cloud OR .env
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    st.error("❌ OPENAI_API_KEY not configured.\n\nAdd it to `.env` (local) or Streamlit Secrets (cloud).")
    st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

# ================= UI SETUP ======================
st.set_page_config(page_title="PPT Narration Studio", layout="wide")
st.title("🎬 PPT Narration Studio")
st.caption("Upload • Review • Edit narration • Download PPT")

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
uploaded_file = st.file_uploader("📤 Upload PPTX file", type=["pptx"])

# ================= HELPERS =======================
def generate_notes(slide_text: str) -> str:
    """Generate narration using OpenAI if notes are missing"""
    prompt = f"Create a clear, concise narration script for this slide:\n{slide_text}"

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
    )

    return response.choices[0].message.content.strip()

# ================= LOAD PPT ======================
if uploaded_file and not st.session_state.ppt_loaded:
    temp_dir = tempfile.mkdtemp()
    ppt_path = Path(temp_dir) / uploaded_file.name
    ppt_path.write_bytes(uploaded_file.read())

    prs = Presentation(ppt_path)

    st.session_state.slides_data.clear()

    for idx, slide in enumerate(prs.slides):
        slide_text = " ".join(
            shape.text for shape in slide.shapes if hasattr(shape, "text")
        ).strip()

        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()

        if not notes:
            notes = generate_notes(slide_text)

        st.session_state.slides_data.append({
            "index": idx,
            "slide_text": slide_text,
            "notes": notes
        })

    st.session_state.ppt_path = ppt_path
    st.session_state.ppt_filename = uploaded_file.name
    st.session_state.ppt_loaded = True

    st.success("✅ PPT loaded successfully")

# ================= REVIEW & EDIT =================
if st.session_state.ppt_loaded:
    st.subheader("📝 Review & Edit Narration (Slide-wise)")

    for slide in st.session_state.slides_data:
        with st.expander(f"Slide {slide['index'] + 1}", expanded=False):
            st.markdown("**Slide Content (reference):**")
            st.write(slide["slide_text"] or "_No visible text on slide_")

            slide["notes"] = st.text_area(
                "Edit Narration",
                slide["notes"],
                key=f"notes_{slide['index']}",
                height=120
            )

    # ================= EXPORT PPT =================
    st.divider()

    if st.button("📥 Generate & Download Narrated PPT"):
        st.info("Generating narrated PPT…")

        # Reload original PPT
        prs = Presentation(st.session_state.ppt_path)

        for slide_data in st.session_state.slides_data:
            slide = prs.slides[slide_data["index"]]
            slide.notes_slide.notes_text_frame.text = slide_data["notes"]

        # Save final PPT
        output_dir = tempfile.mkdtemp()
        final_ppt_path = Path(output_dir) / st.session_state.ppt_filename
        prs.save(final_ppt_path)

        ppt_bytes = final_ppt_path.read_bytes()

        st.success("✅ Narrated PPT ready")

        st.download_button(
            label="⬇ Download PPT",
            data=ppt_bytes,
            file_name=st.session_state.ppt_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
