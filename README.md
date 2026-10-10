# Document Creator Studio

Automated tool and web application for generating website Word documents (`.docx`) from structured JSON packages (`content.json`).

---

## Features

- **Automated Document Assembly**: Clones and populates Word stencils for Home, Generic/Service, and Contact pages.
- **Dynamic Button Cycling**: Avoids repetitive button text by cycling through approved lead and in-text button pools (`LEAD_BUTTON_POOL`, `INTEXT_BUTTON_POOL`).
- **Typography & Formatting Preservation**: WordprocessingML schema-compliant styling, preserving bold spans, font hierarchies, and multi-line paragraphs.
- **Interactive Web Interface (FastAPI)**:
  - Drag-and-drop or paste JSON content.
  - Live pre-flight preview of detected pages, meta titles, H1 headings, and USPs.
  - Built-in default master template with optional custom `.docx` upload.
  - 1-click document generation and download.
- **Streamlit Version**: Ready for deployment to Streamlit Community Cloud.
- **Docker & Cloud Ready**: Includes `Dockerfile` and `requirements.txt` for deployment on Render, Railway, Fly.io, or AWS.

---

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Web Application
```bash
python -m uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```
Open **[http://localhost:8000](http://localhost:8000)** in your browser.

*(On Windows, you can simply double-click `run_server.bat`)*

---

## CLI Usage

You can also run the generator directly from the command line:

```bash
python fill_template.py template.docx content.json output.docx
```

Or generate directly without a template using:

```bash
python generate_doc.py
```

---

## Repository Structure

```
├── app.py               # FastAPI backend server
├── static/
│   └── index.html       # Responsive web frontend (Tailwind CSS)
├── fill_template.py     # Core stencil cloning & DOCX generator engine
├── generate_doc.py      # Standalone document generation script
├── streamlit_app.py     # Streamlit web app alternative
├── template.docx        # Master Word template
├── content.json         # Sample JSON content package
├── run_server.bat       # 1-click Windows launcher
├── Dockerfile           # Container build configuration
├── requirements.txt     # Python package requirements
└── README.md
```
