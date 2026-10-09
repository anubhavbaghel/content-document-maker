"""
app.py — Web Application Server for CCT Document Generator
Provides an interactive web interface and REST API for generating Word documents from JSON.
"""

import io
import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Optional

import docx
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

from fill_template import build_document

app = FastAPI(title="CCT Content Creator Studio", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
DEFAULT_TEMPLATE_PATH = BASE_DIR / "template.docx"

STATIC_DIR.mkdir(parents=True, exist_ok=True)


@app.get("/api/health")
def health_check():
    return {
        "status": "healthy",
        "has_default_template": DEFAULT_TEMPLATE_PATH.exists(),
        "template_size_bytes": DEFAULT_TEMPLATE_PATH.stat().st_size if DEFAULT_TEMPLATE_PATH.exists() else 0,
    }


@app.get("/api/default-template")
def download_default_template():
    if not DEFAULT_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=404, detail="Default template.docx not found on server.")
    return FileResponse(
        path=DEFAULT_TEMPLATE_PATH,
        filename="master_template.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/api/preview")
async def preview_content(
    json_file: Optional[UploadFile] = File(None),
    json_text: Optional[str] = Form(None),
):
    try:
        raw_json = ""
        if json_file:
            content_bytes = await json_file.read()
            raw_json = content_bytes.decode("utf-8")
        elif json_text:
            raw_json = json_text
        else:
            raise HTTPException(status_code=400, detail="Please upload a JSON file or paste JSON content.")

        data = json.loads(raw_json)
        pages = data.get("pages", [])
        if not pages:
            raise HTTPException(status_code=400, detail="JSON must contain a non-empty 'pages' array.")

        preview_pages = []
        for idx, p in enumerate(pages, start=1):
            ai = p.get("AIContent", {})
            meta = ai.get("meta_data", {})
            preview_pages.append({
                "index": idx,
                "name": p.get("PageName", f"Page {idx}"),
                "type": p.get("PageType", "generic"),
                "headline": ai.get("headline", ""),
                "meta_title": meta.get("heading", ""),
                "meta_description": meta.get("description", ""),
                "usps": [ai.get("USP1"), ai.get("USP2"), ai.get("USP3")],
                "has_faqs": bool(ai.get("FAQs1")),
                "services_count": sum(1 for i in range(1, 10) if ai.get(f"service_0{i}") or ai.get(f"service_{i}")),
            })

        return {
            "valid": True,
            "total_pages": len(pages),
            "pages": preview_pages,
            "company_name": preview_pages[0]["meta_title"].split("|")[-1].strip() if "|" in preview_pages[0]["meta_title"] else "Website CCT",
        }
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON format: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Preview error: {str(e)}")


@app.post("/api/generate")
async def generate_document_endpoint(
    json_file: Optional[UploadFile] = File(None),
    json_text: Optional[str] = Form(None),
    template_file: Optional[UploadFile] = File(None),
    output_filename: Optional[str] = Form(None),
):
    temp_dir = tempfile.mkdtemp(prefix="cct_gen_")
    try:
        # 1. Resolve JSON content
        raw_json = ""
        if json_file:
            content_bytes = await json_file.read()
            raw_json = content_bytes.decode("utf-8")
        elif json_text:
            raw_json = json_text
        else:
            raise HTTPException(status_code=400, detail="No JSON content provided.")

        try:
            parsed_json = json.loads(raw_json)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid JSON syntax: {str(e)}")

        json_path = os.path.join(temp_dir, "input_content.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(parsed_json, f, ensure_ascii=False, indent=2)

        # 2. Resolve Template
        template_path = os.path.join(temp_dir, "template.docx")
        if template_file and template_file.filename:
            tpl_bytes = await template_file.read()
            with open(template_path, "wb") as f:
                f.write(tpl_bytes)
        elif DEFAULT_TEMPLATE_PATH.exists():
            shutil.copyfile(str(DEFAULT_TEMPLATE_PATH), template_path)
        else:
            raise HTTPException(
                status_code=400,
                detail="No template uploaded and default master template.docx is missing on the server.",
            )

        # 3. Build Document
        out_doc_path = os.path.join(temp_dir, "generated_output.docx")
        build_document(template_path, json_path, out_doc_path)

        # 4. Stream response
        with open(out_doc_path, "rb") as f:
            doc_bytes = f.read()

        # Filename calculation
        first_page = parsed_json.get("pages", [{}])[0]
        meta_title = first_page.get("AIContent", {}).get("meta_data", {}).get("heading", "")
        safe_name = "Website_CCT_Document"
        if "|" in meta_title:
            safe_name = meta_title.split("|")[-1].strip().replace(" ", "_")
        elif first_page.get("PageName"):
            safe_name = first_page.get("PageName").replace(" ", "_")

        download_name = output_filename.strip() if output_filename and output_filename.strip() else f"{safe_name}_CCT.docx"
        if not download_name.endswith(".docx"):
            download_name += ".docx"

        return Response(
            content=doc_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# Serve Frontend
@app.get("/", response_class=HTMLResponse)
def serve_index():
    index_path = STATIC_DIR / "index.html"
    if index_path.exists():
        return HTMLResponse(content=index_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h2>Frontend is loading...</h2>")


app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

if __name__ == "__main__":
    import uvicorn
    print("Starting CCT Online Generator on http://127.0.0.1:8000 ...")
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
