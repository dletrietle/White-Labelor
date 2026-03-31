# Astoria White-Label Tool

Batch-replace the Astoria logo in monthly commentary DOCX files with client logos and export as PDF.

## How It Works

1. Drop your monthly commentary DOCX into the tool
2. Point it at a folder of client logos (PNG, JPG, etc.)
3. Get back one branded PDF per client — logo swapped, everything else preserved

## Quick Start

### Prerequisites

```bash
# Python 3.10+
pip install python-docx Pillow

# LibreOffice (required for PDF conversion)
# macOS:
brew install --cask libreoffice
# Ubuntu/Debian:
sudo apt install libreoffice
# Windows: download from https://www.libreoffice.org
```

### Setup

```bash
git clone https://github.com/yourusername/astoria-white-label.git
cd astoria-white-label
pip install -r requirements.txt
```

### Usage

```bash
# Basic usage — outputs PDFs
python white_label.py --input commentary.docx --logos ./client_logos/ --output ./output/

# Output as DOCX instead
python white_label.py --input commentary.docx --logos ./client_logos/ --output ./output/ --format docx

# Output both DOCX and PDF
python white_label.py --input commentary.docx --logos ./client_logos/ --output ./output/ --format both
```

### Folder Structure

```
astoria-white-label/
├── white_label.py          # Main script
├── requirements.txt
├── client_logos/            # Drop client logos here
│   ├── Bluedoor.png
│   ├── Summit_Wealth.png
│   ├── Pinnacle_Advisory.jpg
│   └── ...
├── input/                   # Monthly commentary DOCX files
│   └── March_2026_Monthly_Commentary.docx
└── output/                  # Branded outputs appear here
    ├── March_2026_Monthly_Commentary_Report_Bluedoor.pdf
    ├── March_2026_Monthly_Commentary_Report_Summit_Wealth.pdf
    └── ...
```

### Logo Naming Convention

The client name is derived from the logo filename:
- `Bluedoor.png` → client name: **Bluedoor**
- `Summit_Wealth.png` → client name: **Summit Wealth**
- `Pinnacle_Advisory_Logo.jpg` → client name: **Pinnacle Advisory**

### Custom Output Naming

```bash
python white_label.py \
  --input commentary.docx \
  --logos ./client_logos/ \
  --output ./output/ \
  --naming "{month}_{year}_Commentary_{client}"
```

Available template variables: `{month}`, `{year}`, `{client}`

## Web Interface (Recommended)

For a visual drag-and-drop experience:

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000 in your browser
```

The web UI gives you three simple steps:
1. **Upload** — drag in your monthly commentary DOCX
2. **Import** — drag in all client logos at once
3. **Generate** — click one button, download the ZIP of all branded PDFs

## Monthly Workflow

Each month when the commentary is ready:

```bash
# 1. Save the new commentary DOCX to input/
# 2. Run the tool
python white_label.py -i input/April_2026_Monthly_Commentary.docx -l client_logos/ -o output/

# 3. All 20 branded PDFs are ready in output/
```

## Supported Formats

**Logo formats:** PNG, JPG, JPEG, BMP, GIF, TIFF, WEBP  
**Input:** DOCX only  
**Output:** PDF, DOCX, or both

## How Logo Detection Works

The tool automatically finds the first image in the first few paragraphs of the DOCX body — which is where the Astoria logo is placed in the monthly commentary template. The image binary is swapped with the client logo while preserving the original size and position in the document layout.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `LibreOffice not found` | Install LibreOffice and ensure it's on your PATH |
| `Could not find the Astoria logo` | Make sure the DOCX has an image in the first few paragraphs |
| Logo looks stretched | Use a logo with similar aspect ratio to the Astoria logo (~3.6:1 width:height) |
| PDF formatting differs | LibreOffice rendering may vary slightly from Word — check output |

## License

Internal tool — Astoria Portfolio Advisors LLC
