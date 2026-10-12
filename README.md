# Document Creator Studio

Automated tool and web application designed to generate professionally formatted, client-ready website copy Word documents (`.docx`) from structured JSON content packages (`content.json`).

**Live Web Application:** [https://content-document-maker.streamlit.app/](https://content-document-maker.streamlit.app/)

---

## Overview

In content operations, AI-generated or structured website copy is frequently delivered as structured JSON payloads containing multi-page content (Home, Service, Area, and Contact pages). Formatting this copy manually into client-ready documents is time-consuming and error-prone.

**Document Creator Studio** automates this entire pipeline:
1. Slices reusable page stencils from a master Word template document (`template.docx`).
2. Accurately maps JSON fields to designated slots (Meta titles, H1 headings, SQOT subheadlines, USPs, content sections, FAQs, process steps, value propositions, and CTA tables).
3. Preserves typography, hierarchy, and WordprocessingML schema integrity without style degradation.
4. Provides dual web interfaces (FastAPI and Streamlit) alongside a command-line utility.

---

## Key Features

- **Automated Stencil Cloning**: Clones and populates Word stencils for Home, Service/Location (generic), and Contact pages.
- **Dynamic Button Variation Engine**: Eliminates repetitive button text by cycling through approved call-to-action phrasing (`LEAD_BUTTON_POOL` and `INTEXT_BUTTON_POOL`) with per-page collision avoidance.
- **Schema-Compliant XML Typography**: Strictly follows the WordprocessingML `<w:rPr>` schema hierarchy, preserving fonts, sizes, inline `**bold**` markups, bulleted lists, and multiline spacing.
- **Pre-Flight Content Inspector**: Live UI previews detected pages, badges (`HOME`, `SERVICE`, `CONTACT`), meta descriptions, headlines, and key selling points before generation.
- **Dual Web Interfaces**:
  - **Streamlit App**: Zero-configuration, public-facing web app deployed on Streamlit Cloud.
  - **FastAPI + Modern Web App**: High-performance backend with drag-and-drop, paste editor, and RESTful API endpoints.
- **Flexible Template Modes**: Use the embedded master Word stencil by default or upload a custom organizational `.docx` template.
- **CLI & Script Automation**: Fully scriptable for batch document generation in CI/CD or local automation pipelines.

---

## Architecture & How It Works

```
                        JSON Content Package (content.json)
                                       │
                                       ▼
┌────────────────────────────────────────────────────────────────────────┐
│                        Document Creator Studio                         │
├──────────────────────────┬─────────────────────────────┬───────────────┤
│    FastAPI / Streamlit   │      fill_template.py       │ template.docx │
│   • Drag & Drop File     │   • Stencil Slicing Engine  │  • Front-page │
│   • Paste JSON Editor    │   • Button Rotation Pool    │  • Home Block │
│   • Live Pre-flight View │   • Run Property Formatting │  • Generic    │
│   • REST API Endpoints   │   • Multiline Expansion     │  • Contact    │
└──────────────────────────┴─────────────────────────────┴───────────────┘
                                       │
                                       ▼
                     Formatted Output Document (.docx)
```

### Stencil Mapping Table

| Page Type | JSON Source | Generated Document Sections |
| :--- | :--- | :--- |
| **Front Matter** | Top-level metadata | Client details, opening hours, QA notes, and auto-populated Site Map |
| **Home** (`home`) | `AIContent` | Header table, SQOT line, hero CTA button, 3-column USP table, Section 1, Service summaries (1–6), "Why choose us" bullets, Process steps (1–3), Propositions (1–3), and CTA table |
| **Generic** (`generic`) | `AIContent` | Header table, SQOT line, hero CTA button, Sections 1–4, in-text buttons, Propositions (1–3), FAQs (1–3), and CTA table |
| **Contact** (`contact`) | `AIContent` | Header table, 100-word website conclusion, hero CTA button, Section 1, 3-column USP table, and Contact Form blueprint |

---

## Quick Start

### 1. Try Online
Access the hosted web application directly:
**[https://content-document-maker.streamlit.app/](https://content-document-maker.streamlit.app/)**

### 2. Local Setup

Clone the repository and install dependencies:

```bash
git clone https://github.com/anubhavbaghel/content-document-maker.git
cd content-document-maker
pip install -r requirements.txt
```

### 3. Run the Web Application

#### Option A: FastAPI Web App (Recommended for local / team server)
```bash
python -m uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```
Open **[http://localhost:8000](http://localhost:8000)** in your browser.  
*(On Windows, you can double-click `run_server.bat`)*

#### Option B: Streamlit Web App
```bash
streamlit run streamlit_app.py
```
Open **[http://localhost:8501](http://localhost:8501)** in your browser.

---

## Command Line Usage

Generate documents directly via terminal without launching a web server:

```bash
python fill_template.py template.docx content.json output.docx
```

Arguments:
1. `template.docx`: Path to the master Word stencil document.
2. `content.json`: Path to the structured JSON content package.
3. `output.docx`: Path where the stitched Word document will be written.

Alternatively, build a document programmatically without relying on an existing Word template:

```bash
python generate_doc.py
```

---

## API Documentation

When running `app.py`, automated interactive API documentation is available at **[http://localhost:8000/docs](http://localhost:8000/docs)**.

### Primary Endpoints

- **`GET /api/health`**: Service status and master template verification.
- **`GET /api/default-template`**: Downloads the active master `.docx` template.
- **`POST /api/preview`**: Validates JSON syntax and returns parsed page metadata, counts, headlines, and badges.
- **`POST /api/generate`**: Accepts JSON content (file upload or raw text) and optional custom `.docx` template, returning the compiled Word document stream.

---

## JSON Schema Example

The generator expects a top-level `"pages"` array containing page definitions:

```json
{
  "pages": [
    {
      "PageName": "Home",
      "PageType": "home",
      "AIContent": {
        "headline": "Residential care home in Chertsey with 24 hour nursing care",
        "subheadline": "Good CQC rating. Low staff turnover. Home cooked meals",
        "USP1": "Family run since 1965",
        "USP2": "Historic home with en suites",
        "USP3": "Open door visitor policy",
        "service_title": "The care and support that make daily life feel steadier",
        "service_01": {
          "headline": "Residential care",
          "summary": "We support day to day living with person centred care."
        },
        "section1": {
          "heading": "A home with character, care and real continuity",
          "description": "Finding the right care home can feel heavy..."
        },
        "CTA": {
          "heading": "Let's talk about the right care",
          "action": "Call The Grange Retirement Home today."
        },
        "meta_data": {
          "heading": "Residential care home | Chertsey | The Grange Retirement Home",
          "description": "Residential care home in Chertsey with 24 hour nursing care."
        }
      }
    }
  ]
}
```

---

## Repository Structure

```
├── app.py               # FastAPI backend server & REST API
├── static/
│   └── index.html       # Web UI (Tailwind CSS, drag & drop, live preview)
├── fill_template.py     # Core stencil-cloning & Word generator engine
├── generate_doc.py      # Standalone document builder script
├── streamlit_app.py     # Streamlit web app (cloud & local)
├── template.docx        # Master CCT Word stencil template
├── content.json         # Reference content package
├── run_server.bat       # 1-click Windows startup script
├── Dockerfile           # Container build configuration
├── requirements.txt     # Python dependencies
└── README.md            # Project documentation
```

---

## Deployment

### Deploy to Streamlit Community Cloud
1. Fork or push this repository to GitHub.
2. Sign in to [Streamlit Community Cloud](https://share.streamlit.io/).
3. Select this repository and set `streamlit_app.py` as the entrypoint file.

### Deploy with Docker (Render, Railway, Fly.io, AWS)
```bash
docker build -t document-creator-studio .
docker run -p 8000:8000 document-creator-studio
```
The application will be live at `http://localhost:8000`.

---

## License

This project is licensed under the MIT License.
