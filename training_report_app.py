
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from PyPDF2 import PdfReader
import tempfile
import io
import math

st.set_page_config(page_title="Training Evaluation Report Generator", layout="centered")

st.title("📊 Training Evaluation Report Generator")
st.markdown("Upload your **Excel evaluation file** and an optional **PDF template**. The app will generate a PDF report with charts and paginated open-text answers.")

excel_file = st.file_uploader("Upload Excel file", type=["xlsx"])
template_file = st.file_uploader("Upload PDF template (optional)", type=["pdf"])

MAX_LINES_PER_PAGE = 16

def generate_chart(df):
    ratings_df = df.iloc[1:].copy()
    question_cols = [col for col in ratings_df.columns if col.startswith("Q17") or col.startswith("Q18") or col.startswith("Q19")
                     or col.startswith("Q20") or col.startswith("Q21") or col.startswith("Q22")]
    ratings = ratings_df[question_cols].apply(pd.to_numeric, errors='coerce')
    avg_scores = ratings.mean().round(2)

    fig, ax = plt.subplots(figsize=(8, 4))
    avg_scores.plot(kind='bar', ax=ax, ylim=(0, 5), color='skyblue')
    plt.title("Average Ratings per Question")
    plt.ylabel("Average Score")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    
    chart_buf = io.BytesIO()
    plt.savefig(chart_buf, format='png')
    chart_buf.seek(0)
    return chart_buf

def split_text_to_chunks(text_list, max_lines):
    chunks = []
    for i in range(0, len(text_list), max_lines):
        chunk = text_list[i:i + max_lines]
        chunks.append(chunk)
    return chunks

def add_template_background(pdf, template_reader, page_index):
    if template_reader and page_index < len(template_reader.pages):
        page = template_reader.pages[page_index]
        txt = page.extract_text()
        pdf.set_font("Arial", size=12)
        pdf.set_text_color(200, 200, 200)
        pdf.multi_cell(0, 10, txt)

def create_pdf(chart_img, q24_list, q25_list, template_reader=None):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Title Page
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Training evaluation - Client", ln=True)
    pdf.ln(10)
    pdf.image(chart_img, w=180)

    # Q24
    q24_chunks = split_text_to_chunks(q24_list, MAX_LINES_PER_PAGE)
    for i, chunk in enumerate(q24_chunks):
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"Q24 Responses (Page {i+1})", ln=True)
        pdf.set_font("Arial", '', 11)
        for line in chunk:
            pdf.multi_cell(0, 8, f"• {line}")
        pdf.ln(5)

    # Q25
    q25_chunks = split_text_to_chunks(q25_list, MAX_LINES_PER_PAGE)
    for i, chunk in enumerate(q25_chunks):
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, f"Q25 Responses (Page {i+1})", ln=True)
        pdf.set_font("Arial", '', 11)
        for line in chunk:
            pdf.multi_cell(0, 8, f"• {line}")
        pdf.ln(5)

    return pdf

if excel_file:
    with st.spinner("Generating report..."):
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        df = excel_data['Summary']

        q24_col = "Q24: O que poderia mudar no seu contexto de trabalho para melhorar ainda mais o seu índice de performance?"
        q25_col = "Q25: Se já tivesse um maior indíce de produtividade e performance onde e como notaria?"

        q24_responses = df[q24_col].dropna().astype(str).str.strip().tolist()
        q25_responses = df[q25_col].dropna().astype(str).str.strip().tolist()

        chart_buf = generate_chart(df)

        template_reader = PdfReader(template_file) if template_file else None
        pdf = create_pdf(chart_buf, q24_responses, q25_responses, template_reader)

        output_path = tempfile.mktemp(suffix=".pdf")
        pdf.output(output_path)

        with open(output_path, "rb") as f:
            st.success("✅ Report generated!")
            st.download_button("📥 Download Report PDF", f, file_name="training_evaluation_report.pdf")
