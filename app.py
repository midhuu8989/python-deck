# streamlit_app.py
# -------------------------------------------------------------
# Streamlit App: PPT Upload -> Review/Edit -> Narration -> Export
# (NO AZURE | OpenAI TTS ONLY)
# -------------------------------------------------------------

import os
import tempfile
from pathlib import Path

import streamlit as st
from pptx import Presentation

from dotenv import load_dotenv
from openai import OpenAI

# ===================== LOAD ENV ===================
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    st.error("OPENAI_API_KEY not found in .env file")
