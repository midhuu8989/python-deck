st.divider()

if st.session_state.ppt_loaded:

    if st.button("📥 Generate PPT with Voice-over (Sequential)"):
        prs = Presentation(st.session_state.ppt_path)
        outdir = Path(tempfile.mkdtemp())

        total_slides = len(st.session_state.slides)

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, slide_data in enumerate(st.session_state.slides, start=1):
            status_text.info(f"🔄 Processing slide {idx} of {total_slides}")

            slide = prs.slides[slide_data["index"]]

            # Generate audio only for THIS slide
            mp3_path = outdir / f"slide_{slide_data['index']}.mp3"
            openai_tts(slide_data["notes"], mp3_path, speed)

            # Optional: calculate duration (safe fallback if pydub not available)
            try:
                audio = AudioSegment.from_mp3(mp3_path)
                slide_data["duration"] = audio.duration_seconds
            except Exception:
                slide_data["duration"] = 3.0  # fallback

            # Embed audio into slide
            add_audio_to_slide(slide, mp3_path)

            # Update notes
            slide.notes_slide.notes_text_frame.text = slide_data["notes"]

            # Update progress bar
            progress_bar.progress(idx / total_slides)

            # Small delay for UI smoothness (optional)
            time.sleep(0.1)

        status_text.success("✅ All slides processed successfully")

        # Save final PPT
        final_ppt = outdir / st.session_state.ppt_name
        prs.save(final_ppt)

        st.download_button(
            "⬇ Download PPT with Voice-over",
            final_ppt.read_bytes(),
            file_name=st.session_state.ppt_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
