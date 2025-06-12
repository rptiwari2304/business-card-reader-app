import streamlit as st
import zipfile
import os
import requests
from PIL import Image
import pandas as pd
import io

# OCR.space API details (free API key)
OCR_API_KEY = "helloworld"  # Use 'helloworld' key for testing (limit: 10 req/day)

# OCR function
def ocr_space_image(image_bytes):
    url_api = "https://api.ocr.space/parse/image"
    response = requests.post(
        url_api,
        files={"filename": image_bytes},
        data={"apikey": OCR_API_KEY, "language": "eng"},
    )
    result = response.json()
    if result["IsErroredOnProcessing"]:
        return ""
    return result["ParsedResults"][0]["ParsedText"]

# Extraction logic
def extract_info(text):
    import re

    email = re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", text)
    phone = re.findall(r"\+?\d[\d\s\-]{8,}", text)

    lines = text.split("\n")
    lines = [line.strip() for line in lines if line.strip()]
    name = lines[0] if lines else ""
    designation = lines[1] if len(lines) > 1 else ""
    address = ", ".join(lines[2:5]) if len(lines) > 4 else ""

    return {
        "Name": name,
        "Email": email[0] if email else "",
        "Mobile": phone[0] if phone else "",
        "Designation": designation,
        "Address": address
    }

# Streamlit App
st.set_page_config(page_title="📇 Card Reader", layout="centered")
st.title("📇 Business Card Reader (Web App)")
st.markdown("Upload a ZIP file of business card images (JPG/PNG). This app extracts name, email, phone, etc. and exports to Excel.")

uploaded_zip = st.file_uploader("📂 Upload ZIP of cards", type=["zip"])

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip, "r") as zf:
        zf.extractall("cards")

    data = []
    for filename in os.listdir("cards"):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            with open(os.path.join("cards", filename), "rb") as f:
                image_bytes = f.read()
                result_text = ocr_space_image(image_bytes)
                info = extract_info(result_text)
                data.append(info)

    df = pd.DataFrame(data)
    st.dataframe(df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Cards')
        writer.save()

    st.download_button(
        label="📥 Download Excel",
        data=buffer.getvalue(),
        file_name="Card_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
