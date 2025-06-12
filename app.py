import streamlit as st
import zipfile
import os
import requests
from PIL import Image
import pandas as pd
import io
import re

# -------------------------------
# OCR.space API Key (Free usage)
# -------------------------------
OCR_API_KEY = "helloworld"  # Free test key (limit: 10 requests/day)

# -------------------------------
# OCR Function to read image
# -------------------------------
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

# -------------------------------
# Extract useful information
# -------------------------------
def extract_info(text):
    # List of known airline names
    airline_keywords = [
        "IndiGo", "Air India", "SpiceJet", "GoAir", "Vistara", "Qatar Airways",
        "Emirates", "Etihad", "Turkish Airlines", "Lufthansa", "Singapore Airlines",
        "British Airways", "Air France", "KLM", "Saudi Airlines", "Flynas"
    ]

    # Extract email and phone
    email = re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", text)
    phone = re.findall(r"\+?\d[\d\s\-]{8,}", text)

    # Clean and split text into lines
    lines = text.split("\n")
    lines = [line.strip() for line in lines if line.strip()]

    # Basic structure assumption
    name = lines[0] if len(lines) > 0 else ""
    designation = lines[1] if len(lines) > 1 else ""
    address = ", ".join(lines[2:5]) if len(lines) > 4 else ""

    # Match airline
    airline = ""
    for word in airline_keywords:
        if word.lower() in text.lower():
            airline = word
            break

    return {
        "Name": name,
        "Email": email[0] if email else "",
        "Mobile": phone[0] if phone else "",
        "Designation": designation,
        "Address": address,
        "Airline": airline
    }

# -------------------------------
# Streamlit App Configuration
# -------------------------------
st.set_page_config(page_title="📇 Business Card Reader", layout="centered")
st.title("📇 Business Card Reader Web App")
st.markdown(
    "Upload a **ZIP file** containing business card images (JPG/PNG). "
    "The app extracts Name, Email, Phone, Designation, Address, and Airline Name. "
    "You can download the extracted data as an Excel file."
)

# -------------------------------
# Upload ZIP file
# -------------------------------
uploaded_zip = st.file_uploader("📂 Upload ZIP of Business Cards", type=["zip"])

if uploaded_zip:
    # Create directory to extract cards
    extract_path = "cards"
    if not os.path.exists(extract_path):
        os.makedirs(extract_path)

    # Unzip uploaded file
    with zipfile.ZipFile(uploaded_zip, "r") as zf:
        zf.extractall(extract_path)

    # Process each image file
    extracted_data = []
    for filename in os.listdir(extract_path):
        if filename.lower().endswith((".jpg", ".jpeg", ".png")):
            with open(os.path.join(extract_path, filename), "rb") as f:
                image_bytes = f.read()
                text = ocr_space_image(image_bytes)
                info = extract_info(text)
                extracted_data.append(info)

    # Show data in table
    df = pd.DataFrame(extracted_data)
    st.subheader("📋 Extracted Card Details")
    st.dataframe(df)

    # Download Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="BusinessCards")
        writer.save()

    st.download_button(
        label="📥 Download Excel File",
        data=buffer.getvalue(),
        file_name="Business_Card_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
