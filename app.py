import streamlit as st
import pdfplumber
import re
import os
import datetime
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO

st.set_page_config(page_title="JCR Generator", layout="centered")

st.title("📄 Job Completion Report Generator")

TEMPLATE_PATH = "JCR.docx"

def extract_data_from_pdf(file):
    with pdfplumber.open(file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    return {
        "WO_NUMBER": re.search(r"WO Number: (\d+)", text).group(1) if re.search(r"WO Number: (\d+)", text) else "",
        "FACILITY_CODE": re.search(r"(FM\d{4})", text).group(1) if re.search(r"(FM\d{4})", text) else "",
        "FACILITY_LOCATION": re.search(r"FM\d{4} (.+)", text).group(1).strip() if re.search(r"FM\d{4} (.+)", text) else "",
        "REPORTED_ON": re.findall(r"(\d{2}\.\d{2}\.\d{4})", text)[-1] if re.findall(r"(\d{2}\.\d{2}\.\d{4})", text) else "",
        "ISSUED_ON": re.findall(r"(\d{2}\.\d{2}\.\d{4})", text)[0] if re.findall(r"(\d{2}\.\d{2}\.\d{4})", text) else "",
        "EST_COMPLETION_DATE": re.findall(r"(\d{2}\.\d{2}\.\d{4})", text)[1] if len(re.findall(r"(\d{2}\.\d{2}\.\d{4})", text)) > 1 else "",
    }

def fill_word_template(data, doc_id):
    doc = DocxTemplate(TEMPLATE_PATH)
    data["DOC_ID"] = doc_id
    data["GENERATED_DATE"] = datetime.datetime.now().strftime("%d.%m.%Y")
    doc.render(data)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Optional DOC_ID input
start_doc_id = st.text_input("🔢 Enter Starting DOC_ID (optional)", "")

# Upload PDFs
uploaded_files = st.file_uploader("📎 Upload PDF Files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    try:
        start_num = int(start_doc_id) if start_doc_id else None
    except ValueError:
        st.error("DOC_ID must be a number if provided.")
        st.stop()

    word_files = []
    summary_data = []
    current = start_num if start_num else None

    for idx, uploaded_file in enumerate(uploaded_files):
        fields = extract_data_from_pdf(uploaded_file)
        doc_id = str(current + idx) if current else ""
        summary_data.append({**fields, "DOC_ID": doc_id, "GENERATED_DATE": datetime.datetime.now().strftime("%d.%m.%Y")})
        word_output = fill_word_template(fields, doc_id)
        word_files.append((f"JCR_{doc_id if doc_id else 'no_id'}_{idx+1}.docx", word_output))

    df = pd.DataFrame(summary_data)
    st.success("✅ All PDFs processed and Word files generated.")

    st.download_button("⬇️ Download CSV Summary", df.to_csv(index=False).encode(), "summary.csv", "text/csv")

    for filename, filedata in word_files:
        st.download_button(f"⬇️ Download {filename}", filedata, file_name=filename)