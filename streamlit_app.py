"""
Astoria White-Label Tool — Streamlit Web App
=============================================
Batch-replace the Astoria logo in monthly commentary DOCX
with client logos and export as PDF.

Deploy: streamlit run streamlit_app.py
"""

import io
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime

import streamlit as st
from PIL import Image

from white_label import (
    replace_logo_in_docx,
    convert_docx_to_pdf,
    get_client_name_from_logo,
    find_logo_image,
    SUPPORTED_LOGO_FORMATS,
)

# ───────────────────────────────────────────
# Page Config
# ───────────────────────────────────────────
st.set_page_config(
    page_title="Astoria White-Label Tool",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ───────────────────────────────────────────
# Custom Styling
# ───────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');

    /* Global */
    .stApp {
        font-family: 'DM Sans', sans-serif;
    }

    /* Header */
    .app-header {
        text-align: center;
        padding: 1.5rem 0 2rem;
    }
    .app-badge {
        display: inline-block;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        color: #3b82f6;
        background: rgba(59, 130, 246, 0.12);
        padding: 5px 14px;
        border-radius: 100px;
        margin-bottom: 12px;
    }
    .app-title {
        font-size: 28px;
        font-weight: 700;
        margin: 0 0 8px;
        letter-spacing: -0.3px;
    }
    .app-subtitle {
        font-size: 14px;
        opacity: 0.6;
        max-width: 450px;
        margin: 0 auto;
        line-height: 1.6;
    }

    /* Section headers */
    .section-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 4px;
    }
    .section-num {
        width: 28px;
        height: 28px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 50%;
        font-size: 13px;
        font-weight: 600;
        background: #3b82f6;
        color: white;
        flex-shrink: 0;
    }
    .section-num.done {
        background: #22c55e;
    }
    .section-title {
        font-size: 16px;
        font-weight: 600;
        margin: 0;
    }
    .section-sub {
        font-size: 13px;
        opacity: 0.5;
        margin: 0 0 16px 40px;
    }

    /* Logo grid */
    .logo-grid-container {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(130px, 1fr));
        gap: 10px;
        margin: 12px 0;
    }
    .logo-card {
        background: rgba(128, 128, 128, 0.08);
        border-radius: 10px;
        padding: 12px 8px 10px;
        text-align: center;
        border: 1px solid rgba(128, 128, 128, 0.12);
    }
    .logo-card img {
        max-height: 45px;
        max-width: 100%;
        object-fit: contain;
        margin-bottom: 6px;
    }
    .logo-card .logo-label {
        font-size: 11px;
        font-weight: 500;
        opacity: 0.7;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    /* Result rows */
    .result-row {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 14px;
        background: rgba(128, 128, 128, 0.06);
        border-radius: 8px;
        margin-bottom: 6px;
        font-size: 13px;
    }
    .result-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        flex-shrink: 0;
    }
    .result-dot.success { background: #22c55e; }
    .result-dot.error { background: #ef4444; }
    .result-name { font-weight: 500; flex: 1; }
    .result-file { font-size: 11px; opacity: 0.5; font-family: monospace; }
    .result-error { font-size: 11px; color: #ef4444; }

    /* Hide default Streamlit elements */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header { visibility: hidden; }

    /* Divider */
    .custom-divider {
        height: 1px;
        background: rgba(128, 128, 128, 0.15);
        margin: 28px 0;
    }
</style>
""", unsafe_allow_html=True)


# ───────────────────────────────────────────
# Header
# ───────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="app-badge">Internal Tool</div>
    <h1 class="app-title">White-Label Generator</h1>
    <p class="app-subtitle">Upload the monthly commentary and client logos.<br>Get branded PDFs for every client in one click.</p>
</div>
""", unsafe_allow_html=True)


# ───────────────────────────────────────────
# Session State
# ───────────────────────────────────────────
if "commentary_file" not in st.session_state:
    st.session_state.commentary_file = None
if "commentary_info" not in st.session_state:
    st.session_state.commentary_info = None
if "logo_files" not in st.session_state:
    st.session_state.logo_files = {}
if "results" not in st.session_state:
    st.session_state.results = None
if "zip_data" not in st.session_state:
    st.session_state.zip_data = None
if "zip_name" not in st.session_state:
    st.session_state.zip_name = None


# ───────────────────────────────────────────
# Step 1: Upload Commentary
# ───────────────────────────────────────────
commentary_done = st.session_state.commentary_file is not None
num_class_1 = "done" if commentary_done else ""

st.markdown(f"""
<div class="section-header">
    <div class="section-num {num_class_1}">{"✓" if commentary_done else "1"}</div>
    <h3 class="section-title">Monthly Commentary</h3>
</div>
<p class="section-sub">Upload the Astoria commentary DOCX for this month</p>
""", unsafe_allow_html=True)

uploaded_docx = st.file_uploader(
    "Upload Commentary DOCX",
    type=["docx"],
    key="docx_uploader",
    label_visibility="collapsed",
)

if uploaded_docx is not None and (
    st.session_state.commentary_file is None
    or st.session_state.commentary_file.name != uploaded_docx.name
):
    # Validate the DOCX
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp.write(uploaded_docx.read())
    tmp.close()
    uploaded_docx.seek(0)

    try:
        from docx import Document
        doc = Document(tmp.name)
        rel_id, target = find_logo_image(doc)

        month_match = re.search(
            r'(January|February|March|April|May|June|July|August|September|October|November|December)',
            uploaded_docx.name, re.IGNORECASE,
        )
        year_match = re.search(r'(20\d{2})', uploaded_docx.name)

        st.session_state.commentary_file = uploaded_docx
        st.session_state.commentary_info = {
            "filename": uploaded_docx.name,
            "month": month_match.group(1) if month_match else datetime.now().strftime("%B"),
            "year": year_match.group(1) if year_match else str(datetime.now().year),
            "logo_target": target,
            "tmp_path": tmp.name,
        }
        # Reset results when new file uploaded
        st.session_state.results = None
        st.session_state.zip_data = None
    except Exception as e:
        st.error(f"Could not process DOCX: {e}")
    finally:
        pass

if st.session_state.commentary_info:
    info = st.session_state.commentary_info
    st.success(
        f"**{info['filename']}** — {info['month']} {info['year']} · Logo detected"
    )


st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)


# ───────────────────────────────────────────
# Step 2: Upload Logos
# ───────────────────────────────────────────
logos_done = len(st.session_state.logo_files) > 0
num_class_2 = "done" if logos_done else ""

st.markdown(f"""
<div class="section-header">
    <div class="section-num {num_class_2}">{"✓" if logos_done else "2"}</div>
    <h3 class="section-title">Client Logos</h3>
</div>
<p class="section-sub">Import client logos — one branded PDF per logo</p>
""", unsafe_allow_html=True)

uploaded_logos = st.file_uploader(
    "Upload Client Logos",
    type=["png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp"],
    accept_multiple_files=True,
    key="logos_uploader",
    label_visibility="collapsed",
)

# Process newly uploaded logos
if uploaded_logos:
    for logo_file in uploaded_logos:
        if logo_file.name not in st.session_state.logo_files:
            st.session_state.logo_files[logo_file.name] = {
                "file": logo_file,
                "client_name": get_client_name_from_logo(logo_file.name),
            }
    # Reset results when logos change
    if st.session_state.results is not None:
        st.session_state.results = None
        st.session_state.zip_data = None

# Display logo grid
if st.session_state.logo_files:
    logo_items = list(st.session_state.logo_files.items())

    # Build grid HTML with base64 thumbnails
    import base64

    grid_html = '<div class="logo-grid-container">'
    for filename, data in logo_items:
        try:
            data["file"].seek(0)
            img_bytes = data["file"].read()
            b64 = base64.b64encode(img_bytes).decode()
            ext = Path(filename).suffix.lower().replace(".", "")
            mime = f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
            img_tag = f'<img src="data:{mime};base64,{b64}" />'
        except Exception:
            img_tag = '<div style="height:45px;display:flex;align-items:center;justify-content:center;opacity:0.4;font-size:11px;">No preview</div>'

        grid_html += f"""
        <div class="logo-card">
            {img_tag}
            <div class="logo-label" title="{data['client_name']}">{data['client_name']}</div>
        </div>
        """
    grid_html += "</div>"
    st.markdown(grid_html, unsafe_allow_html=True)

    st.caption(f"**{len(logo_items)}** client logos loaded")

    if st.button("Clear All Logos", type="secondary"):
        st.session_state.logo_files = {}
        st.session_state.results = None
        st.session_state.zip_data = None
        st.rerun()


st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)


# ───────────────────────────────────────────
# Step 3: Generate
# ───────────────────────────────────────────
can_generate = (
    st.session_state.commentary_info is not None
    and len(st.session_state.logo_files) > 0
)

num_class_3 = "done" if st.session_state.results else ""
st.markdown(f"""
<div class="section-header">
    <div class="section-num {num_class_3}">{"✓" if st.session_state.results else "3"}</div>
    <h3 class="section-title">Generate Branded PDFs</h3>
</div>
<p class="section-sub">One PDF per client with their logo replacing the Astoria logo</p>
""", unsafe_allow_html=True)

if not can_generate:
    st.info("Upload a commentary DOCX and at least one client logo to get started.")

if can_generate and st.session_state.results is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        generate_clicked = st.button(
            "⚡ Generate All PDFs",
            type="primary",
            use_container_width=True,
        )

    if generate_clicked:
        info = st.session_state.commentary_info
        logo_items = list(st.session_state.logo_files.items())

        results = []
        output_dir = tempfile.mkdtemp(prefix="astoria_output_")
        tmp_dir = tempfile.mkdtemp(prefix="astoria_tmp_")

        # Save commentary to temp file (re-save from buffer)
        commentary_tmp = os.path.join(tmp_dir, "commentary.docx")
        st.session_state.commentary_file.seek(0)
        with open(commentary_tmp, "wb") as f:
            f.write(st.session_state.commentary_file.read())

        progress_bar = st.progress(0, text="Starting...")
        total = len(logo_items)

        for i, (filename, data) in enumerate(logo_items):
            client_name = data["client_name"]
            base_name = f"{info['month']}_{info['year']}_Monthly_Commentary_Report_{client_name.replace(' ', '_')}"

            progress_bar.progress(
                (i) / total,
                text=f"Processing {client_name} ({i+1}/{total})...",
            )

            try:
                # Save logo to temp
                logo_tmp = os.path.join(tmp_dir, filename)
                data["file"].seek(0)
                with open(logo_tmp, "wb") as f:
                    f.write(data["file"].read())

                # Replace logo
                docx_output = os.path.join(tmp_dir, f"{base_name}.docx")
                replace_logo_in_docx(commentary_tmp, logo_tmp, docx_output)

                # Convert to PDF
                pdf_path = convert_docx_to_pdf(docx_output, output_dir)

                results.append({
                    "client": client_name,
                    "filename": os.path.basename(pdf_path),
                    "pdf_path": pdf_path,
                    "status": "success",
                })
            except Exception as e:
                results.append({
                    "client": client_name,
                    "status": "error",
                    "error": str(e),
                })

        progress_bar.progress(1.0, text="Complete!")

        # Create ZIP
        zip_buffer = io.BytesIO()
        zip_name = f"{info['month']}_{info['year']}_Monthly_Commentary_All_Clients.zip"
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for r in results:
                if r["status"] == "success":
                    zf.write(r["pdf_path"], r["filename"])
        zip_buffer.seek(0)

        st.session_state.results = results
        st.session_state.zip_data = zip_buffer.getvalue()
        st.session_state.zip_name = zip_name

        # Cleanup
        shutil.rmtree(tmp_dir, ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)

        st.rerun()


# ───────────────────────────────────────────
# Results
# ───────────────────────────────────────────
if st.session_state.results:
    results = st.session_state.results
    success_count = sum(1 for r in results if r["status"] == "success")
    error_count = sum(1 for r in results if r["status"] == "error")

    # Summary
    if error_count == 0:
        st.success(f"All **{success_count}** branded PDFs generated successfully!")
    else:
        st.warning(f"**{success_count}** succeeded · **{error_count}** failed")

    # Result list
    results_html = ""
    for r in results:
        if r["status"] == "success":
            results_html += f"""
            <div class="result-row">
                <div class="result-dot success"></div>
                <div class="result-name">{r['client']}</div>
                <div class="result-file">{r['filename']}</div>
            </div>
            """
        else:
            results_html += f"""
            <div class="result-row">
                <div class="result-dot error"></div>
                <div class="result-name">{r['client']}</div>
                <div class="result-error">{r.get('error', 'Unknown error')}</div>
            </div>
            """
    st.markdown(results_html, unsafe_allow_html=True)

    st.markdown("")

    # Download button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.download_button(
            label="📥 Download All PDFs (.zip)",
            data=st.session_state.zip_data,
            file_name=st.session_state.zip_name,
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

    # Option to re-generate
    st.markdown("")
    if st.button("↻ Start Over", type="secondary"):
        st.session_state.results = None
        st.session_state.zip_data = None
        st.session_state.zip_name = None
        st.rerun()
