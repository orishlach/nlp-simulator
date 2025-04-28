
import streamlit as st
import json
from protocol_parser import process_document
import tempfile
import os
from PIL import Image

st.set_page_config(page_title="🎤 Knesset Protocols Analyzer", page_icon="🎤", layout="wide")

# הצגת לוגו
logo = Image.open("logo.png")
st.image(logo, width=100)

st.title("🎤 NLP Final Project - Knesset Protocols Analyzer")
st.write("Upload Knesset protocol .docx files, and get structured NLP output!")

uploaded_files = st.file_uploader("Choose one or more .docx files", type="docx", accept_multiple_files=True)

if uploaded_files:
    all_sentences = []

    for uploaded_file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name

        protocol_name = uploaded_file.name
        knesset_number, protocol_type = 1, 'plenary'  # ניתן להרחיב
        protocol = process_document(tmp_file_path, protocol_name, knesset_number, protocol_type)

        if protocol:
            for sentence in protocol.sentences:
                all_sentences.append({
                    'protocol_name': protocol.protocol_name,
                    'knesset_number': protocol.knesset_number,
                    'protocol_type': protocol.protocol_type,
                    'protocol_number': protocol.protocol_number,
                    'speaker_name': sentence.speaker_name,
                    'sentence_text': sentence.sentence_text,
                })

        os.unlink(tmp_file_path)

    if all_sentences:
        st.success(f"✅ Processed {len(all_sentences)} sentences!")

        unique_speakers = sorted(set(s['speaker_name'] for s in all_sentences))
        speaker_filter = st.selectbox("Filter by speaker (optional):", ["All"] + unique_speakers)

        filtered_sentences = all_sentences
        if speaker_filter != "All":
            filtered_sentences = [s for s in all_sentences if s['speaker_name'] == speaker_filter]

        st.subheader("Example sentences:")
        for i, sentence in enumerate(filtered_sentences[:30]):
            st.write(f"**{sentence['speaker_name']}**: {sentence['sentence_text']}")

        jsonl_data = "\n".join(json.dumps(line, ensure_ascii=False) for line in filtered_sentences)
        st.download_button("📥 Download JSONL", jsonl_data, file_name="protocol_sentences.jsonl", mime="application/jsonl")
