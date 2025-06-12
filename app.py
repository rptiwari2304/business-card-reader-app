import streamlit as st
import zipfile
import os
import requests
from PIL import Image
import pandas as pd
import io
import re

OCR_API_KEY = "helloworld"  # Free API key from OCR.space

# OCR function
def ocr_space_image(image_bytes):
    url_api = "https://api.ocr.space/parse/image"
    response = requests.post(
        url_api,
        files={"filename": image_bytes},
        data={"apikey": OCR_API_KEY, "language": "eng"},
    )
    result = response.json()
    if result.get("IsErroredOnProcessing", True):
        return ""
    return result["ParsedResults"][0]["ParsedText"]

# Extract info from text
def extract_info(text):
    email = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    phone = re.findall(r"\+?\d[\d\s\-]{8,}", text)

    lines = text.split("\n")
    lines = [line.strip() for line in lines if line.strip()]

    name = lines[0] if len(lines) > 0 else ""
    designation = lines[1] if len(lines) > 1 else ""
    address = ", ".join(lines[2:5]) if len(lines) >= 5 else ""

    # Identify airline
    airline_keywords = [
        "IndiGo", "Air India", "SpiceJet", "GoAir", "Vistara", "Qatar Airways",
        "Emirates", "Etihad", "Turkish Airlines", "Lufthansa", "Singapore Airlines",
        "British Airways", "Air France", "KLM", "Saudi Airlines", "Flynas"
    ]
    airline = ""
    for keyword in airline_keywords:
        if keyword.lower() in text.lower():
            airline = keyword
            break

    return {
        "Name": name,
        "Designation": designation,
        "Email": email[0] if email else "",
        "Mobile": phone[0] if phone else "",
        "Address": address,
        "Airline": airline
    }

# Streamlit App
st.set_page_config(page_title="📇 Business Card Reader", layout="centered")
st.title("📇 Business Card Reader Web App")
st.markdown("Upload a ZIP file of business card images. The app reads the cards and extracts details to Excel.")

uploaded_zip = st.file_uploader("📂 Upload ZIP of Business Cards", type=["zip"])

if uploaded_zip:
    extract_path = "cards"
    if not os.path.exists(extract_path):
        os.makedirs(extract_path)

    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        zip_ref.extractall(extract_path)

    card_data = []
    st.info("🔍 Processing images...")
    for file in os.listdir(extract_path):
        if file.lower().endswith((".jpg", ".jpeg", ".png")):
            filepath = os.path.join(extract_path, file)
            with open(filepath, "rb") as f:
                image_bytes = f.read()
                text = ocr_space_image(image_bytes)

                if not text:
                    st.warning(f"❗ No text found in {file}")
                    continue

                st.markdown(f"**📎 File:** `{file}`")
                st.text_area("📤 OCR Output", text, height=150)

                details = extract_info(text)
                details["File Name"] = file
                card_data.append(details)

    if card_data:
        df = pd.DataFrame(card_data)
        st.subheader("✅ Extracted Card Details")
        st.dataframe(df)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='BusinessCards')

        st.download_button(
            label="📥 Download Extracted Excel",
            data=buffer.getvalue(),
            file_name="Extracted_Card_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid data found in cards.")
