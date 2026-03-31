"""
Astoria White-Label Tool — Streamlit Web App
=============================================
Upload DOCX + logos → download branded DOCX files.
Open in Word → Save As PDF for perfect formatting.
"""

import io
import os
import re
import base64
import tempfile
import shutil
import zipfile
from pathlib import Path
from datetime import datetime

import streamlit as st

from white_label import (
    replace_logo_in_docx,
    find_logo_image,
    get_client_name_from_logo,
    SUPPORTED_LOGO_FORMATS,
)

# ───────────────────────────────────────────
st.set_page_config(
    page_title="Astoria White-Label Tool",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ───────────────────────────────────────────
# Styling
# ───────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
    .stApp { font-family: 'DM Sans', sans-serif; }
    .app-header { text-align: center; padding: 1.5rem 0 2rem; }
    .app-badge {
        display: inline-block; font-size: 11px; font-weight: 600;
        letter-spacing: 1.5px; text-transform: uppercase;
        color: #3b82f6; background: rgba(59, 130, 246, 0.12);
        padding: 5px 14px; border-radius: 100px; margin-bottom: 12px;
    }
    .app-title { font-size: 28px; font-weight: 700; margin: 0 0 8px; letter-spacing: -0.3px; }
    .app-subtitle { font-size: 14px; opacity: 0.6; max-width: 480px; margin: 0 auto; line-height: 1.6; }
    .section-header { display: flex; align-items: center; gap: 12px; margin-bottom: 4px; }
    .section-num {
        width: 28px; height: 28px; display: flex; align-items: center;
        justify-content: center; border-radius: 50%; font-size: 13px;
        font-weight: 600; background: #3b82f6; color: white; flex-shrink: 0;
    }
    .section-num.done { background: #22c55e; }
    .section-title { font-size: 16px; font-weight: 600; margin: 0; }
    .section-sub { font-size: 13px; opacity: 0.5; margin: 0 0 16px 40px; }
    .logo-grid-container {
        display: grid; grid-template-columns: repeat(auto-fill, minmax(130px, 1fr));
        gap: 10px; margin: 12px 0;
    }
    .logo-card {
        background: rgba(128, 128, 128, 0.08); border-radius: 10px;
        padding: 12px 8px 10px; text-align: center;
        border: 1px solid rgba(128, 128, 128, 0.12);
    }
    .logo-card img { max-height: 45px; max-width: 100%; object-fit: contain; margin-bottom: 6px; }
    .logo-card .logo-label {
        font-size: 11px; font-weight: 500; opacity: 0.7;
        white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .result-row {
        display: flex; align-items: center; gap: 10px;
        padding: 10px 14px; background: rgba(128, 128, 128, 0.06);
        border-radius: 8px; margin-bottom: 6px; font-size: 13px;
    }
    .result-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
    .result-dot.success { background: #22c55e; }
    .result-dot.error { background: #ef4444; }
    .result-name { font-weight: 500; flex: 1; }
    .result-file { font-size: 11px; opacity: 0.5; font-family: monospace; }
    .result-error { font-size: 11px; color: #ef4444; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header { visibility: hidden; }
    .custom-divider { height: 1px; background: rgba(128, 128, 128, 0.15); margin: 28px 0; }
</style>
""", unsafe_allow_html=True)


# ───────────────────────────────────────────
# Header
# ───────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="app-badge">Internal Tool</div>
    <h1 class="app-title">White-Label Generator</h1>
    <p class="app-subtitle">Upload the monthly commentary DOCX and client logos.<br>Download branded DOCX files, then Save As PDF in Word.</p>
</div>
""", unsafe_allow_html=True)


# ───────────────────────────────────────────
# Session State — bytes only
# ───────────────────────────────────────────
for key, default in [
    ("docx_bytes", None), ("docx_info", None), ("logo_data", {}),
    ("results", None), ("zip_data", None), ("zip_name", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ───────────────────────────────────────────
# Step 1: Upload DOCX
# ───────────────────────────────────────────
docx_done = st.session_state.docx_bytes is not None

st.markdown(f"""
<div class="section-header">
    <div class="section-num {"done" if docx_done else ""}">{"✓" if docx_done else "1"}</div>
    <h3 class="section-title">Monthly Commentary</h3>
</div>
<p class="section-sub">Upload the Astoria commentary DOCX</p>
""", unsafe_allow_html=True)

uploaded_docx = st.file_uploader("Upload DOCX", type=["docx"], key="docx_up", label_visibility="collapsed")

if uploaded_docx is not None:
    new_bytes = uploaded_docx.read()
    uploaded_docx.seek(0)

    if (
        st.session_state.docx_bytes is None
        or st.session_state.docx_info is None
        or st.session_state.docx_info.get("filename") != uploaded_docx.name
    ):
        tmp_path = None
        try:
            tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
            tmp_path = tmp.name
            tmp.write(new_bytes)
            tmp.close()

            image_zip_path = find_logo_image(tmp_path)

            month_match = re.search(
                r'(January|February|March|April|May|June|July|August|September|October|November|December)',
                uploaded_docx.name, re.IGNORECASE,
            )
            year_match = re.search(r'(20\d{2})', uploaded_docx.name)

            st.session_state.docx_bytes = new_bytes
            st.session_state.docx_info = {
                "filename": uploaded_docx.name,
                "month": month_match.group(1) if month_match else datetime.now().strftime("%B"),
                "year": year_match.group(1) if year_match else str(datetime.now().year),
                "image_zip_path": image_zip_path,
            }
            st.session_state.results = None
            st.session_state.zip_data = None
        except Exception as e:
            st.error(f"Could not process DOCX: {e}")
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)

if st.session_state.docx_info:
    info = st.session_state.docx_info
    st.success(f"**{info['filename']}** — {info['month']} {info['year']} · Logo detected at `{info['image_zip_path']}`")

st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)


# ───────────────────────────────────────────
# Step 2: Upload Logos
# ───────────────────────────────────────────
logos_done = len(st.session_state.logo_data) > 0

st.markdown(f"""
<div class="section-header">
    <div class="section-num {"done" if logos_done else ""}">{"✓" if logos_done else "2"}</div>
    <h3 class="section-title">Client Logos</h3>
</div>
<p class="section-sub">Import client logos — one branded DOCX per logo</p>
""", unsafe_allow_html=True)

uploaded_logos = st.file_uploader(
    "Upload Logos",
    type=["png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp"],
    accept_multiple_files=True,
    key="logos_up",
    label_visibility="collapsed",
)

if uploaded_logos:
    for logo_file in uploaded_logos:
        if logo_file.name not in st.session_state.logo_data:
            logo_bytes = logo_file.read()
            logo_file.seek(0)
            st.session_state.logo_data[logo_file.name] = {
                "bytes": logo_bytes,
                "client_name": get_client_name_from_logo(logo_file.name),
            }

if st.session_state.logo_data:
    logo_items = list(st.session_state.logo_data.items())

    grid_html = '<div class="logo-grid-container">'
    for filename, data in logo_items:
        try:
            b64 = base64.b64encode(data["bytes"]).decode()
            ext = Path(filename).suffix.lower().replace(".", "")
            mime = f"image/{'jpeg' if ext in ('jpg', 'jpeg') else ext}"
            img_tag = f'<img src="data:{mime};base64,{b64}" />'
        except Exception:
            img_tag = '<div style="height:45px;opacity:0.4;font-size:11px;">?</div>'
        grid_html += f'<div class="logo-card">{img_tag}<div class="logo-label" title="{data["client_name"]}">{data["client_name"]}</div></div>'
    grid_html += "</div>"
    st.markdown(grid_html, unsafe_allow_html=True)
    st.caption(f"**{len(logo_items)}** client logos loaded")

    if st.button("Clear All Logos", type="secondary"):
        st.session_state.logo_data = {}
        st.session_state.results = None
        st.session_state.zip_data = None
        st.rerun()

st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)


# ───────────────────────────────────────────
# Step 3: Generate
# ───────────────────────────────────────────
can_generate = st.session_state.docx_bytes is not None and len(st.session_state.logo_data) > 0

st.markdown(f"""
<div class="section-header">
    <div class="section-num {"done" if st.session_state.results else ""}">{"✓" if st.session_state.results else "3"}</div>
    <h3 class="section-title">Generate Branded Files</h3>
</div>
<p class="section-sub">One DOCX per client — logo swapped, everything else untouched</p>
""", unsafe_allow_html=True)

if not can_generate:
    st.info("Upload a commentary DOCX and at least one client logo to get started.")

if can_generate and st.session_state.results is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⚡ Generate All Files", type="primary", use_container_width=True):

            info = st.session_state.docx_info
            logo_items = list(st.session_state.logo_data.items())
            total = len(logo_items)

            results = []
            tmp_dir = tempfile.mkdtemp(prefix="astoria_wl_")

            try:
                # Write source DOCX to disk
                src_path = os.path.join(tmp_dir, "source.docx")
                with open(src_path, "wb") as f:
                    f.write(st.session_state.docx_bytes)

                progress_bar = st.progress(0, text="Starting...")
                docx_outputs = {}

                for i, (filename, data) in enumerate(logo_items):
                    client_name = data["client_name"]
                    safe_client = re.sub(r'[^\w\s-]', '', client_name).replace(' ', '_')
                    out_filename = f"{info['month']}_{info['year']}_Monthly_Commentary_Report_{safe_client}.docx"
                    out_path = os.path.join(tmp_dir, out_filename)

                    progress_bar.progress(i / total, text=f"Processing {client_name} ({i + 1}/{total})...")

                    try:
                        # Write logo to disk
                        logo_tmp = os.path.join(tmp_dir, f"logo_{i}{Path(filename).suffix}")
                        with open(logo_tmp, "wb") as f:
                            f.write(data["bytes"])

                        # Swap logo
                        replace_logo_in_docx(src_path, logo_tmp, out_path, info["image_zip_path"])

                        # Read output bytes
                        with open(out_path, "rb") as f:
                            docx_outputs[out_filename] = f.read()

                        results.append({"client": client_name, "filename": out_filename, "status": "success"})
                    except Exception as e:
                        results.append({"client": client_name, "status": "error", "error": str(e)})

                progress_bar.progress(1.0, text="Complete!")

                # Create ZIP
                zip_buffer = io.BytesIO()
                zip_name = f"{info['month']}_{info['year']}_Monthly_Commentary_All_Clients.zip"
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, docx_bytes in docx_outputs.items():
                        zf.writestr(fname, docx_bytes)
                zip_buffer.seek(0)

                st.session_state.results = results
                st.session_state.zip_data = zip_buffer.getvalue()
                st.session_state.zip_name = zip_name

            except Exception as e:
                import traceback
                st.error(f"Generation failed: {e}")
                st.code(traceback.format_exc())
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)


# ───────────────────────────────────────────
# Results
# ───────────────────────────────────────────
if st.session_state.results:
    results = st.session_state.results
    success_count = sum(1 for r in results if r["status"] == "success")
    error_count = sum(1 for r in results if r["status"] == "error")

    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

    if error_count == 0:
        st.success(f"All **{success_count}** branded DOCX files generated!")
    else:
        st.warning(f"**{success_count}** succeeded · **{error_count}** failed")

    results_html = ""
    for r in results:
        if r["status"] == "success":
            results_html += f'<div class="result-row"><div class="result-dot success"></div><div class="result-name">{r["client"]}</div><div class="result-file">{r["filename"]}</div></div>'
        else:
            results_html += f'<div class="result-row"><div class="result-dot error"></div><div class="result-name">{r["client"]}</div><div class="result-error">{r.get("error", "Unknown")}</div></div>'
    st.markdown(results_html, unsafe_allow_html=True)

    st.markdown("")

    if st.session_state.zip_data and len(st.session_state.zip_data) > 22:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.download_button(
                label="📥 Download All DOCX Files (.zip)",
                data=st.session_state.zip_data,
                file_name=st.session_state.zip_name,
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

        st.caption("💡 **Tip:** To convert to PDF, open each DOCX in Word → File → Save As → PDF")

    st.markdown("")
    if st.button("↻ Start Over", type="secondary"):
        st.session_state.results = None
        st.session_state.zip_data = None
        st.session_state.zip_name = None
        st.rerun()
