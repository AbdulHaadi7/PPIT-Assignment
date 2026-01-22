import streamlit as st
import easyocr
from PIL import Image
import numpy as np

st.set_page_config(page_title="Handwritten Notes to Digital Text")

st.title("✍️ Handwritten Notes → Editable Digital Text")

uploaded_file = st.file_uploader(
    "Upload handwritten notes image",
    type=["png", "jpg", "jpeg"]
)

if uploaded_file is not None:
    image = Image.open(uploaded_file)
    st.image(image, caption="Uploaded Image", use_column_width=True)

    st.write("🔍 Extracting text...")

    reader = easyocr.Reader(['en'], gpu=False)
    text = reader.readtext(np.array(image), detail=0)

    extracted_text = "\n".join(text)

    st.subheader("📝 Editable Digital Notes")
    st.text_area(
        "You can edit the extracted text below:",
        extracted_text,
        height=300
    )
