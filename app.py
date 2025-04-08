
import streamlit as st
import pdfplumber
import re
import os
import datetime
import json
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO

st.set_page_config(page_title="JCR Generator", layout="centered")

st.title("📄 Job Completion Report Generator")

# Initialize counter file
COUNTER_FILE = "counter.json"
MONTH = datetime.datetime.now().month

def load_counter():
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE, "r") as f:
            data = json.load(f)
        if str(MONTH) not in data:
            data[str(MONTH)] = 1
    else:
        data = {str(MONTH): 1}
    return data

def save_counter(data):
    with open(COUNTER_FILE, "w") as f:
        json.dump(data, f)

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

def fill_word_template(data, doc_id, template_bytes):
    doc = DocxTemplate(template_bytes)
    data["DOC_ID"] = doc_id
    data["GENERATED_DATE"] = datetime.datetime.now().strftime("%d.%m.%Y")
    doc.render(data)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Upload template
st.sidebar.header("📄 Upload Template")
template_file = st.sidebar.file_uploader("Upload JCR Template (.docx)", type=["docx"])

# Reset counter
if st.sidebar.button("🔄 Reset Counter for New Month"):
    save_counter({str(MONTH): 1})
    st.sidebar.success("Counter reset to 1 for this month.")

# Upload PDFs
uploaded_files = st.file_uploader("Upload PDF Files", type=["pdf"], accept_multiple_files=True)

if uploaded_files and template_file:
    counter = load_counter()
    current = counter[str(MONTH)]
    summary_data = []
    word_files = []

    for uploaded_file in uploaded_files:
        fields = extract_data_from_pdf(uploaded_file)
        doc_id = f"{MONTH}{str(current).zfill(3)}"
        summary_data.append({**fields, "DOC_ID": doc_id})
        word_output = fill_word_template(fields, doc_id, template_file)
        word_files.append((f"JCR_{doc_id}.docx", word_output))
        current += 1

    counter[str(MONTH)] = current
    save_counter(counter)

    df = pd.DataFrame(summary_data)
    st.success("✅ All PDFs processed and Word files generated.")

    st.download_button("⬇️ Download CSV Summary", df.to_csv(index=False).encode(), "summary.csv", "text/csv")

    for filename, filedata in word_files:
        st.download_button(f"⬇️ Download {filename}", filedata, file_name=filename)
else:
    st.info("Please upload both PDF files and a Word template.")
