import pandas as pd
from docxtpl import DocxTemplate
from datetime import date
import streamlit as st
import io
import zipfile
import subprocess
import tempfile
import os

st.set_page_config(page_title="Settlement Agreement Generator", layout="centered", page_icon="📄")

st.markdown("""
<style>
    .main {
        background-color: #f4f6fa;
        font-family: 'Segoe UI', sans-serif;
    }
    .title {
        font-size: 2.2em;
        font-weight: bold;
        color: #2c3e50;
    }
    .subtitle {
        font-size: 1.05em;
        color: #7f8c8d;
        margin-bottom: 10px;
    }
    .info-box {
        background-color: #eaf0fb;
        border-left: 4px solid #3a86ff;
        padding: 12px 16px;
        border-radius: 6px;
        margin-bottom: 16px;
        font-size: 0.95em;
        color: #2c3e50;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="title">📄 Settlement Agreement Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Coded by Char &nbsp;·&nbsp; ARETIS LIMITED</div>', unsafe_allow_html=True)
st.markdown("---")

st.markdown("""
<div class="info-box">
    <b>How it works:</b> Upload your CSV file and Word template below.
    The app will generate one Settlement Agreement PDF per row and bundle them into a ZIP file for download.
</div>
""", unsafe_allow_html=True)

with st.expander("📋 Required CSV Columns"):
    st.markdown("""
    Your CSV file must contain the following columns:

    | Column | Description |
    |---|---|
    | `Name` | Influencer / Party B full name |
    | `Platform username` | Social media username |
    | `Links` | Media account link |
    | `Total videos` | Number of videos in the cooperation |
    | `Rate` | USD rate per video |
    | `Total rate` | Total fee (rate × videos) |
    | `PayPal email` | Party B's PayPal email for payment |
    """)

st.markdown("### 📤 Upload Files")
col1, col2 = st.columns(2)
with col1:
    uploaded_csv = st.file_uploader("📑 Upload CSV File", type=["csv"])
with col2:
    uploaded_template = st.file_uploader("📄 Upload Word Template (.docx)", type=["docx"])

if uploaded_csv:
    try:
        preview_df = pd.read_csv(uploaded_csv)
        preview_df.columns = preview_df.columns.str.strip()
        uploaded_csv.seek(0)
        with st.expander(f"👀 Preview CSV ({len(preview_df)} rows)"):
            st.dataframe(preview_df, use_container_width=True)
    except Exception as e:
        st.warning(f"Could not preview CSV: {e}")


def docx_to_pdf(docx_path: str, output_dir: str) -> str:
    """Convert a .docx file to PDF using LibreOffice and return the PDF path."""
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        capture_output=True,
        text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found after conversion: {pdf_path}")
    return pdf_path


if uploaded_csv and uploaded_template:
    st.success("✅ Both files uploaded successfully!")

    if st.button("🚀 Generate Agreements", type="primary"):
        try:
            df = pd.read_csv(uploaded_csv)
            df.columns = df.columns.str.strip()

            today = date.today().isoformat()
            zip_buffer = io.BytesIO()
            errors = []
            success_count = 0

            progress = st.progress(0, text="Generating agreements...")

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df.iterrows():
                    try:
                        template = DocxTemplate(uploaded_template)
                        context = {
                            'Influencer_name':   row['Name'],
                            'platform_username': row['Platform username'],
                            'Influencer_links':  row['Links'],
                            'total_video':       row['Total videos'],
                            'rate':              row['Rate'],
                            'total_rate':        row['Total rate'],
                            'paypal_email':      row['PayPal email'],
                        }
                        template.render(context)

                        influencer_name = str(row['Name'])
                        username        = str(row['Platform username'])

                        # Safe temp names (no special chars for filesystem)
                        safe_name     = influencer_name.replace(" ", "_").replace("/", "-")
                        safe_username = username.replace(" ", "_").replace("/", "-")

                        # Final PDF filename as requested
                        pdf_filename = f"BlockBlast X {influencer_name} ({username}) Settlement Agreement.pdf"

                        with tempfile.TemporaryDirectory() as tmpdir:
                            docx_path = os.path.join(tmpdir, f"{safe_name}_{safe_username}.docx")
                            template.save(docx_path)

                            pdf_path = docx_to_pdf(docx_path, tmpdir)

                            with open(pdf_path, "rb") as f:
                                zip_file.writestr(pdf_filename, f.read())

                        success_count += 1

                    except Exception as e:
                        errors.append(f"Row {index} ({row.get('Name', 'Unknown')}): {e}")

                    progress.progress((index + 1) / len(df), text=f"Processing {index + 1} of {len(df)}...")

            progress.empty()

            if errors:
                st.warning(f"⚠️ {len(errors)} error(s) occurred:")
                for err in errors:
                    st.error(f"❌ {err}")

            if success_count > 0:
                zip_buffer.seek(0)
                st.markdown(f"### ✅ {success_count} agreement(s) generated!")
                st.download_button(
                    label="📥 Download ZIP of All Agreements",
                    data=zip_buffer,
                    file_name=f"BlockBlast_Settlement_Agreements_{today}.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"❌ Failed to process files: {e}")

else:
    st.info("⬆️ Upload both files above to get started.")
