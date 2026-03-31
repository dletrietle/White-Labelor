#!/usr/bin/env python3
"""
Astoria White-Label Tool — Web Interface
=========================================
Flask web app for drag-and-drop batch white-labeling.

Usage:
    python app.py
    # Then open http://localhost:5000 in your browser
"""

import os
import io
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from white_label import (
    replace_logo_in_docx,
    convert_docx_to_pdf,
    get_client_name_from_logo,
    find_logo_image,
    SUPPORTED_LOGO_FORMATS,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB max upload

# Temp storage for the current session's files
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'astoria_whitelabel')


def ensure_clean_dir(path: str) -> str:
    """Create or clean a directory."""
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path, exist_ok=True)
    return path


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload-commentary', methods=['POST'])
def upload_commentary():
    """Upload the monthly commentary DOCX file."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({'error': 'File must be a .docx file'}), 400

    # Save to temp directory
    commentary_dir = ensure_clean_dir(os.path.join(UPLOAD_DIR, 'commentary'))
    filename = secure_filename(file.filename)
    filepath = os.path.join(commentary_dir, filename)
    file.save(filepath)

    # Validate — check if we can find a logo in it
    try:
        from docx import Document
        doc = Document(filepath)
        rel_id, target = find_logo_image(doc)
        
        # Extract month/year from filename
        month_match = re.search(
            r'(January|February|March|April|May|June|July|August|September|October|November|December)',
            filename, re.IGNORECASE
        )
        year_match = re.search(r'(20\d{2})', filename)
        month = month_match.group(1) if month_match else None
        year = year_match.group(1) if year_match else None

        return jsonify({
            'success': True,
            'filename': filename,
            'month': month,
            'year': year,
            'logo_found': True,
            'logo_target': target,
        })
    except Exception as e:
        return jsonify({'error': f'Could not process DOCX: {str(e)}'}), 400


@app.route('/api/upload-logos', methods=['POST'])
def upload_logos():
    """Upload one or more client logo files."""
    if 'files' not in request.files:
        return jsonify({'error': 'No files provided'}), 400

    files = request.files.getlist('files')
    logos_dir = os.path.join(UPLOAD_DIR, 'logos')
    os.makedirs(logos_dir, exist_ok=True)

    uploaded = []
    for file in files:
        ext = Path(file.filename).suffix.lower()
        if ext not in SUPPORTED_LOGO_FORMATS:
            continue

        filename = secure_filename(file.filename)
        filepath = os.path.join(logos_dir, filename)
        file.save(filepath)

        client_name = get_client_name_from_logo(filepath)
        uploaded.append({
            'filename': filename,
            'client_name': client_name,
            'size': os.path.getsize(filepath),
        })

    return jsonify({
        'success': True,
        'logos': uploaded,
        'total': len(uploaded),
    })


@app.route('/api/remove-logo', methods=['POST'])
def remove_logo():
    """Remove a single logo from the uploaded set."""
    data = request.get_json()
    filename = data.get('filename')
    if not filename:
        return jsonify({'error': 'No filename provided'}), 400

    filepath = os.path.join(UPLOAD_DIR, 'logos', secure_filename(filename))
    if os.path.exists(filepath):
        os.remove(filepath)

    return jsonify({'success': True})


@app.route('/api/clear-logos', methods=['POST'])
def clear_logos():
    """Clear all uploaded logos."""
    logos_dir = os.path.join(UPLOAD_DIR, 'logos')
    if os.path.exists(logos_dir):
        shutil.rmtree(logos_dir)
    os.makedirs(logos_dir, exist_ok=True)
    return jsonify({'success': True})


@app.route('/api/generate', methods=['POST'])
def generate():
    """Generate all white-labeled PDFs and return as a ZIP."""
    commentary_dir = os.path.join(UPLOAD_DIR, 'commentary')
    logos_dir = os.path.join(UPLOAD_DIR, 'logos')

    # Find the commentary file
    docx_files = [f for f in os.listdir(commentary_dir) if f.endswith('.docx')]
    if not docx_files:
        return jsonify({'error': 'No commentary file uploaded'}), 400
    commentary_path = os.path.join(commentary_dir, docx_files[0])

    # Find all logos
    logo_files = sorted([
        f for f in os.listdir(logos_dir)
        if Path(f).suffix.lower() in SUPPORTED_LOGO_FORMATS
    ])
    if not logo_files:
        return jsonify({'error': 'No client logos uploaded'}), 400

    # Extract month/year from filename
    input_stem = Path(commentary_path).stem
    month_match = re.search(
        r'(January|February|March|April|May|June|July|August|September|October|November|December)',
        input_stem, re.IGNORECASE
    )
    year_match = re.search(r'(20\d{2})', input_stem)
    month = month_match.group(1) if month_match else datetime.now().strftime('%B')
    year = year_match.group(1) if year_match else str(datetime.now().year)

    # Process each logo
    output_dir = ensure_clean_dir(os.path.join(UPLOAD_DIR, 'output'))
    tmp_dir = ensure_clean_dir(os.path.join(UPLOAD_DIR, 'tmp'))

    results = []
    for logo_filename in logo_files:
        logo_path = os.path.join(logos_dir, logo_filename)
        client_name = get_client_name_from_logo(logo_path)
        base_name = f"{month}_{year}_Monthly_Commentary_Report_{client_name.replace(' ', '_')}"

        try:
            # Replace logo
            docx_output = os.path.join(tmp_dir, f"{base_name}.docx")
            replace_logo_in_docx(commentary_path, logo_path, docx_output)

            # Convert to PDF
            pdf_path = convert_docx_to_pdf(docx_output, output_dir)

            results.append({
                'client': client_name,
                'filename': os.path.basename(pdf_path),
                'status': 'success',
            })
        except Exception as e:
            results.append({
                'client': client_name,
                'status': 'error',
                'error': str(e),
            })

    # Create ZIP of all PDFs
    zip_buffer = io.BytesIO()
    zip_name = f"{month}_{year}_Monthly_Commentary_All_Clients.zip"
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in os.listdir(output_dir):
            if f.endswith('.pdf'):
                zf.write(os.path.join(output_dir, f), f)

    zip_buffer.seek(0)
    zip_path = os.path.join(UPLOAD_DIR, zip_name)
    with open(zip_path, 'wb') as f:
        f.write(zip_buffer.getvalue())

    success_count = sum(1 for r in results if r['status'] == 'success')
    error_count = sum(1 for r in results if r['status'] == 'error')

    return jsonify({
        'success': True,
        'results': results,
        'summary': {
            'total': len(results),
            'success': success_count,
            'errors': error_count,
        },
        'zip_filename': zip_name,
    })


@app.route('/api/download')
def download():
    """Download the generated ZIP file."""
    filename = request.args.get('file')
    if not filename:
        return jsonify({'error': 'No filename specified'}), 400

    filepath = os.path.join(UPLOAD_DIR, secure_filename(filename))
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404

    return send_file(filepath, as_attachment=True, download_name=filename)


if __name__ == '__main__':
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    print("\n  Astoria White-Label Tool")
    print("  Open http://localhost:5000 in your browser\n")
    app.run(debug=True, port=5000)
