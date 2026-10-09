"""
streamlit_app.py — Streamlit Web Application for CCT Document Creator Studio
Can be run locally or deployed for free with 1 click to Streamlit Community Cloud.
"""

import io
import json
import os
import shutil
import tempfile
from pathlib import Path

import streamlit as st
from fill_template import build_document

st.set_page_config(
    page_title="CCT Document Creator Studio",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE_PATH = BASE_DIR / "template.docx"

# Header
st.title("📄 CCT Document Creator Studio")
st.markdown("Generate fully formatted, high-quality **Content Creation Template (CCT)** Word documents from JSON packages in seconds.")

with st.sidebar:
    st.header("⚙️ Configuration")
    
    # Template Selection
    template_option = st.radio(
        "Master Template Source",
        ["Use Default Master Template", "Upload Custom Template (.docx)"],
        index=0,
    )
    
    custom_template_file = None
    if template_option == "Upload Custom Template (.docx)":
        custom_template_file = st.file_uploader("Upload .docx Template", type=["docx"])
    else:
        if DEFAULT_TEMPLATE_PATH.exists():
            st.success("✅ Default master template loaded")
            with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
                st.download_button(
                    "📥 Download Master Template",
                    data=f.read(),
                    file_name="master_template.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        else:
            st.warning("⚠️ Default template.docx not found.")

    st.markdown("---")
    st.markdown("### 💡 Quick Actions")
    if st.button("✨ Load Example JSON"):
        st.session_state["json_input_text"] = json.dumps({
            "pages": [
                {
                    "AIContent": {
                        "headline": "Residential care home in Chertsey with 24 hour nursing care",
                        "subheadline": "Good CQC rating. Low staff turnover. Home cooked meals",
                        "calltoaction": "Call now to arrange a visit",
                        "USP1": "Family run since 1965",
                        "USP2": "Historic home with en suites",
                        "USP3": "Open door visitor policy",
                        "service_title": "The care and support that make daily life feel steadier",
                        "service_01": { "headline": "Residential care", "summary": "We support day to day living with person centred care." },
                        "service_02": { "headline": "24 hour nursing care", "summary": "Our qualified nurses are on duty around the clock." },
                        "service_03": { "headline": "Care home facilities", "summary": "En suite bedrooms, welcoming communal spaces, landscaped gardens." },
                        "section1": { "heading": "A home with character, care and real continuity", "description": "Finding the right care home can feel heavy..." },
                        "section2": { "heading": "Why families feel reassured here", "description": "- Good CQC rated home\n- Qualified nurses day and night\n- Up to 62 residents" },
                        "section3": { "heading": "Support that adapts as life changes", "description": "Your family may need short term support after illness or full time care." },
                        "CTA": { "heading": "Let’s talk about the right care", "action": "Call The Grange Retirement Home today." },
                        "meta_data": { "heading": "Residential care home | Chertsey | The Grange Retirement Home", "description": "Residential care home in Chertsey with 24 hour nursing care." }
                    },
                    "PageName": "Home",
                    "PageType": "home"
                },
                {
                    "AIContent": {
                        "headline": "Residential care services in Chertsey with a Good CQC rating",
                        "subheadline": "24 hour nursing care. En suite accommodation. Home cooked nutritious meals",
                        "calltoaction": "Call now to discuss care",
                        "USP1": "Low staff turnover",
                        "USP2": "Person centred care plans",
                        "USP3": "Warm homely surroundings",
                        "section1": { "heading": "Support that feels personal every day", "description": "Residential care services at The Grange Retirement Home in Chertsey..." },
                        "section2": { "heading": "A home with care close at hand", "description": "With roots as a long established care home since 1965..." },
                        "section3": { "heading": "What you can expect from daily life here", "description": "- Individual care plans tailored to your personal needs\n- En suite rooms" },
                        "section4": { "heading": "The little things that help somewhere feel like home", "description": "Your surroundings can shape how secure and relaxed you feel." },
                        "CTA": { "heading": "Talk through care with people who understand", "action": "Call The Grange Retirement Home today." },
                        "meta_data": { "heading": "Residential care services | Chertsey | The Grange Retirement Home", "description": "Residential care services in Chertsey from The Grange." }
                    },
                    "PageName": "Residential Care",
                    "PageType": "generic"
                },
                {
                    "AIContent": {
                        "headline": "Residential care for elderly in Chertsey with easy visiting access",
                        "website_conclusion": "We welcome enquiries about residential care, 24-hour nursing care, and respite support.",
                        "calltoaction": "Call us to discuss care",
                        "section1": { "heading": "Call us to discuss care options", "description": "Our home is in Chertsey at the foot of St. Anne's Hill." },
                        "USP1": "Easy to reach in Chertsey",
                        "USP2": "Family run home",
                        "USP3": "Visits warmly welcomed",
                        "meta_data": { "heading": "Residential care for elderly | Chertsey | The Grange Retirement Home", "description": "Residential care for elderly in Chertsey." }
                    },
                    "PageName": "Contact",
                    "PageType": "contact"
                }
            ]
        }, indent=2)

col1, col2 = st.columns([1, 1], gap="large")

parsed_data = None
raw_json_str = ""

with col1:
    st.subheader("1. Input JSON Package")
    input_mode = st.radio("Input Method", ["Upload JSON File", "Paste JSON Text"], horizontal=True)

    if input_mode == "Upload JSON File":
        uploaded_json_file = st.file_uploader("Upload content.json", type=["json"])
        if uploaded_json_file is not None:
            raw_json_str = uploaded_json_file.read().decode("utf-8")
    else:
        raw_json_str = st.text_area(
            "Paste JSON content",
            value=st.session_state.get("json_input_text", ""),
            height=320,
            placeholder='{"pages": [{"AIContent": {...}}]}',
        )

    output_filename = st.text_input("Custom Output Filename (Optional)", placeholder="e.g. MyClient_CCT.docx")

    # Validate JSON
    if raw_json_str.strip():
        try:
            parsed_data = json.loads(raw_json_str)
            if "pages" not in parsed_data or not parsed_data["pages"]:
                st.error("JSON is valid, but missing a non-empty 'pages' list.")
                parsed_data = None
            else:
                st.success(f"✅ JSON Valid: {len(parsed_data['pages'])} pages detected")
        except json.JSONDecodeError as e:
            st.error(f"❌ Invalid JSON Syntax: {e}")
            parsed_data = None

with col2:
    st.subheader("2. Document Structure & Preview")
    
    if parsed_data:
        pages = parsed_data.get("pages", [])
        for idx, p in enumerate(pages, 1):
            p_name = p.get("PageName", f"Page {idx}")
            p_type = p.get("PageType", "generic")
            ai = p.get("AIContent", {})
            meta = ai.get("meta_data", {})
            
            with st.expander(f"**Page {idx}: {p_name}** ({p_type.upper()})", expanded=(idx == 1)):
                st.markdown(f"**H1:** {ai.get('headline', '*(None)*')}")
                st.markdown(f"**Meta Title:** {meta.get('heading', '*(None)*')}")
                st.markdown(f"**Meta Description:** {meta.get('description', '*(None)*')}")
                
                usps = [ai.get("USP1"), ai.get("USP2"), ai.get("USP3")]
                usps = [u for u in usps if u]
                if usps:
                    st.markdown("**USPs:** " + " • ".join(f"`{u}`" for u in usps))
    else:
        st.info("👈 Upload or paste a JSON package on the left to preview document structure.")

# Generation Section
st.markdown("---")
gen_col1, gen_col2, gen_col3 = st.columns([1, 2, 1])

with gen_col2:
    if st.button("🚀 Generate CCT Word Document", type="primary", use_container_width=True, disabled=(parsed_data is None)):
        temp_dir = tempfile.mkdtemp(prefix="cct_streamlit_")
        try:
            with st.spinner("Generating document and stitching pages..."):
                # Save JSON
                json_path = os.path.join(temp_dir, "content.json")
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(parsed_data, f, ensure_ascii=False, indent=2)

                # Save Template
                template_path = os.path.join(temp_dir, "template.docx")
                if custom_template_file is not None:
                    with open(template_path, "wb") as f:
                        f.write(custom_template_file.read())
                elif DEFAULT_TEMPLATE_PATH.exists():
                    shutil.copyfile(str(DEFAULT_TEMPLATE_PATH), template_path)
                else:
                    st.error("No template available.")
                    st.stop()

                # Generate
                output_docx_path = os.path.join(temp_dir, "output.docx")
                build_document(template_path, json_path, output_docx_path)

                # Determine download name
                first_page = parsed_data.get("pages", [{}])[0]
                meta_title = first_page.get("AIContent", {}).get("meta_data", {}).get("heading", "")
                safe_name = "Generated_CCT"
                if "|" in meta_title:
                    safe_name = meta_title.split("|")[-1].strip().replace(" ", "_")
                elif first_page.get("PageName"):
                    safe_name = first_page.get("PageName").replace(" ", "_")

                final_name = output_filename.strip() if output_filename.strip() else f"{safe_name}_CCT.docx"
                if not final_name.endswith(".docx"):
                    final_name += ".docx"

                with open(output_docx_path, "rb") as f:
                    doc_bytes = f.read()

                st.success("🎉 Document generated successfully!")
                st.download_button(
                    label=f"⬇️ Download {final_name}",
                    data=doc_bytes,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"Generation error: {e}")
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
