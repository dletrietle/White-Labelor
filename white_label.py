#!/usr/bin/env python3
"""
Astoria White-Label Tool
========================
Batch-replace the Astoria logo in monthly commentary DOCX files
with client logos. Outputs branded DOCX files that you open in
Word and Save As PDF — preserving perfect formatting.

Usage:
    python white_label.py --input commentary.docx --logos ./client_logos/ --output ./output/

Requirements:
    pip install python-docx Pillow
"""

import argparse
import io
import os
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from PIL import Image

# Supported image formats for client logos
SUPPORTED_LOGO_FORMATS = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp'}


def find_logo_image(doc_path: str) -> str:
    """
    Find the logo image path inside the DOCX ZIP.

    Returns:
        The ZIP-internal path like 'word/media/image1.jpeg'
    """
    doc = Document(doc_path)
    nsmap = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }

    for para_idx in range(min(5, len(doc.paragraphs))):
        para = doc.paragraphs[para_idx]
        drawings = para._element.findall('.//w:drawing', nsmap)
        for drawing in drawings:
            blips = drawing.findall('.//a:blip', nsmap)
            for blip in blips:
                embed_id = blip.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
                )
                if embed_id and embed_id in doc.part.rels:
                    rel = doc.part.rels[embed_id]
                    target = rel.target_ref
                    if any(ext in target.lower() for ext in ['.jpeg', '.jpg', '.png', '.bmp', '.gif']):
                        return f"word/{target}"

    raise ValueError(
        "Could not find the Astoria logo in the document. "
        "Make sure the DOCX has an image in the first few paragraphs."
    )


def replace_logo_in_docx(input_path: str, logo_path: str, output_path: str, image_zip_path: str = None) -> None:
    """
    Replace the Astoria logo in a DOCX with a client logo.

    Works at the ZIP level — swaps just the image binary while
    keeping everything else byte-identical. Zero formatting changes.
    """
    # Find the logo if not provided
    if image_zip_path is None:
        image_zip_path = find_logo_image(input_path)

    original_ext = Path(image_zip_path).suffix.lower()

    # Extract original logo dimensions from the DOCX
    with zipfile.ZipFile(input_path, 'r') as z:
        original_data = z.read(image_zip_path)
    original_img = Image.open(io.BytesIO(original_data))
    target_w, target_h = original_img.size

    # Open the client logo
    logo_img = Image.open(logo_path)

    # Convert palette mode to RGBA for proper handling
    if logo_img.mode == 'P':
        logo_img = logo_img.convert('RGBA')

    # Resize client logo to FIT within original dimensions (preserve aspect ratio)
    logo_w, logo_h = logo_img.size
    scale = min(target_w / logo_w, target_h / logo_h)
    new_w = int(logo_w * scale)
    new_h = int(logo_h * scale)
    logo_img = logo_img.resize((new_w, new_h), Image.LANCZOS)

    # Center on a canvas matching the original logo's exact dimensions
    if original_ext in ('.jpg', '.jpeg'):
        # JPEG: white background, RGB mode
        canvas = Image.new('RGB', (target_w, target_h), (255, 255, 255))
        # Flatten transparency if present
        if logo_img.mode == 'RGBA':
            temp_bg = Image.new('RGB', logo_img.size, (255, 255, 255))
            temp_bg.paste(logo_img, mask=logo_img.split()[3])
            logo_img = temp_bg
        elif logo_img.mode != 'RGB':
            logo_img = logo_img.convert('RGB')
    else:
        # PNG: transparent background, RGBA mode
        canvas = Image.new('RGBA', (target_w, target_h), (255, 255, 255, 0))
        if logo_img.mode != 'RGBA':
            logo_img = logo_img.convert('RGBA')

    # Paste centered
    x_offset = (target_w - new_w) // 2
    y_offset = (target_h - new_h) // 2
    if logo_img.mode == 'RGBA' and canvas.mode == 'RGBA':
        canvas.paste(logo_img, (x_offset, y_offset), mask=logo_img.split()[3])
    else:
        canvas.paste(logo_img, (x_offset, y_offset))

    # Save in the same format as the original
    img_buffer = io.BytesIO()
    if original_ext in ('.jpg', '.jpeg'):
        canvas.save(img_buffer, format='JPEG', quality=95)
    else:
        canvas.save(img_buffer, format='PNG')

    replacement_bytes = img_buffer.getvalue()

    # Rebuild the DOCX ZIP, replacing only the logo image
    with zipfile.ZipFile(input_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == image_zip_path:
                    zout.writestr(item, replacement_bytes)
                else:
                    zout.writestr(item, data)


def get_client_name_from_logo(logo_path: str) -> str:
    """Extract a clean client name from the logo filename."""
    stem = Path(logo_path).stem
    for remove in ['logo', 'Logo', 'LOGO', '_logo', '-logo', 'logo_', 'logo-']:
        stem = stem.replace(remove, '')
    stem = stem.replace('_', ' ').replace('-', ' ').strip()
    return stem.title() if stem else Path(logo_path).stem


def process_batch(
    input_docx: str,
    logos_dir: str,
    output_dir: str,
    naming_template: str = '{month}_{year}_Monthly_Commentary_Report_{client}'
) -> list[dict]:
    """Process a batch of client logos against a single commentary DOCX."""
    if not os.path.exists(input_docx):
        raise FileNotFoundError(f"Input DOCX not found: {input_docx}")
    if not os.path.isdir(logos_dir):
        raise NotADirectoryError(f"Logos directory not found: {logos_dir}")

    os.makedirs(output_dir, exist_ok=True)

    # Find the logo once (reuse for all clients)
    image_zip_path = find_logo_image(input_docx)
    print(f"  Logo found: {image_zip_path}")

    logo_files = sorted([
        os.path.join(logos_dir, f)
        for f in os.listdir(logos_dir)
        if Path(f).suffix.lower() in SUPPORTED_LOGO_FORMATS
    ])

    if not logo_files:
        raise ValueError(f"No logo files found in {logos_dir}.")

    # Extract month/year from filename
    input_stem = Path(input_docx).stem
    month_match = re.search(
        r'(January|February|March|April|May|June|July|August|September|October|November|December)',
        input_stem, re.IGNORECASE
    )
    year_match = re.search(r'(20\d{2})', input_stem)
    month = month_match.group(1) if month_match else 'Monthly'
    year = year_match.group(1) if year_match else ''

    results = []
    total = len(logo_files)

    for i, logo_path in enumerate(logo_files, 1):
        client_name = get_client_name_from_logo(logo_path)
        base_name = naming_template.format(
            month=month, year=year, client=client_name.replace(' ', '_')
        )
        docx_output = os.path.join(output_dir, f"{base_name}.docx")

        result = {'client': client_name, 'logo': logo_path, 'status': 'pending', 'outputs': []}

        try:
            print(f"  [{i}/{total}] Processing: {client_name}...", end=' ', flush=True)
            replace_logo_in_docx(input_docx, logo_path, docx_output, image_zip_path)
            result['status'] = 'success'
            result['outputs'].append(docx_output)
            print("✓")
        except Exception as e:
            result['status'] = 'error'
            result['error'] = str(e)
            print(f"✗ ({e})")

        results.append(result)

    return results


def main():
    parser = argparse.ArgumentParser(
        description='Astoria White-Label Tool — Batch brand commentary with client logos',
    )
    parser.add_argument('--input', '-i', required=True, help='Path to the Astoria commentary DOCX')
    parser.add_argument('--logos', '-l', required=True, help='Directory containing client logo images')
    parser.add_argument('--output', '-o', default='./output', help='Output directory (default: ./output)')
    parser.add_argument('--naming', '-n', default='{month}_{year}_Monthly_Commentary_Report_{client}')

    args = parser.parse_args()

    print("=" * 60)
    print("  Astoria White-Label Tool")
    print("=" * 60)
    print(f"  Input:   {args.input}")
    print(f"  Logos:   {args.logos}")
    print(f"  Output:  {args.output}")
    print("=" * 60)

    logo_count = len([f for f in os.listdir(args.logos) if Path(f).suffix.lower() in SUPPORTED_LOGO_FORMATS])
    print(f"\n  Found {logo_count} client logos.\n")

    results = process_batch(args.input, args.logos, args.output, args.naming)

    success = sum(1 for r in results if r['status'] == 'success')
    errors = sum(1 for r in results if r['status'] == 'error')

    print(f"\n{'=' * 60}")
    print(f"  COMPLETE: {success} succeeded, {errors} failed")
    print(f"  Output: {os.path.abspath(args.output)}")
    print(f"{'=' * 60}")

    return 0 if errors == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
