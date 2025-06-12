import streamlit as st
import zipfile
import os
import requests
from PIL import Image
import pandas as pd
import io
import re

OCR_API_KEY = "helloworld"  # Free OCR.space key (limit: 10 requests/day)

def ocr_space_image(image_bytes):
    url_api = "https://api.ocr.space/parse/image"
    try:
        response = requests.post(
            url_api,
            files={"filename": image_bytes},
            data={"apikey": OCR_API_KEY, "language": "eng"},
        )
        result = response.json()
        if result.get("IsErroredOnProcessing", True):
            return ""
        return result["ParsedResults"][0]["ParsedText"]
    except Exception as e:
        return ""

def extract_info(text):
    email = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    phone = re.findall(r"\+?\d[\d\s\-]{8,}", text)

    lines = text.split("\n")
    lines = [line.strip() for line in lines if line.strip()]

    name = lines[0] if len(lines) > 0 else ""
    designation = lines[1] if len(lines) > 1 else ""
    address = ", ".join(lines[2:5]) if len(lines) >= 5 else ""

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

st.set_page_config(page_title="📇 Business Card Reader", layout="centered")
st.title("📇 Business Card Reader Web App")
st.markdown("Upload a ZIP file of business card images (JPG/PNG). This app extracts details and exports them to Excel.")

uploaded_zip = st.file_uploader("📂 Upload ZIP of Business Cards", type=["zip"])

if uploaded_zip:
    extract_path = "cards"
    if not os.path.exists(extract_path):
        os.makedirs(extract_path)

    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        zip_ref.extractall(extract_path)

    card_data = []
    no_text_files = []

    for file in os.listdir(extract_path):
        if file.lower().endswith((".jpg", ".jpeg", ".png")):
            filepath = os.path.join(extract_path, file)
            with open(filepath, "rb") as f:
                image_bytes = f.read()
                text = ocr_space_image(image_bytes)

                st.subheader(f"📎 {file}")
                if not text.strip():
                    st.error("❌ No text extracted from this image.")
                    no_text_files.append(file)
                    continue

                st.text_area("📝 Extracted Text", text, height=150)

                details = extract_info(text)
                details["File Name"] = file
                card_data.append(details)

    if card_data:
        df = pd.DataFrame(card_data)
        st.success(f"✅ Extracted data from {len(card_data)} cards.")
        st.dataframe(df)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='BusinessCards')

        st.download_button(
            label="📥 Download Excel",
            data=buffer.getvalue(),
            file_name="Card_Details.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid text extracted from uploaded cards. Try different images or clearer scans.")
