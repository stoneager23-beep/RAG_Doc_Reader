import io
import os
import tempfile

import pandas as pd
import streamlit as st
from docx import Document as DocxDocument
from langchain_community.document_loaders import PyPDFLoader
from langchain_community.vectorstores import FAISS
from langchain_core.documents import Document
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnablePassthrough
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_text_splitters import RecursiveCharacterTextSplitter
from pptx import Presentation

# ----------------------------------------------------------
# Page Setup
# ----------------------------------------------------------
st.set_page_config(
    page_title="DocuVerse 3D Studio",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------
# File Format Metadata & Custom 3D Animation Themes
# ----------------------------------------------------------
FILE_THEMES = {
    "pdf": {
        "name": "PDF Document",
        "icon": "📕",
        "unit": "Page",
        "color": "#EF4444",
        "bg": "#FEF2F2",
        "border": "#FCA5A5",
        "shadow": "rgba(239, 68, 68, 0.22)",
        "anim_class": "anim-pdf",
        "badge_text": "Laser OCR & Page Indexing",
    },
    "docx": {
        "name": "Word Document",
        "icon": "📘",
        "unit": "Section",
        "color": "#0EA5E9",
        "bg": "#F0F9FF",
        "border": "#7DD3FC",
        "shadow": "rgba(14, 165, 233, 0.22)",
        "anim_class": "anim-docx",
        "badge_text": "Paragraph Stream Parser",
    },
    "pptx": {
        "name": "PowerPoint Deck",
        "icon": "📙",
        "unit": "Slide",
        "color": "#F97316",
        "bg": "#FFF7ED",
        "border": "#FDBA74",
        "shadow": "rgba(249, 115, 22, 0.22)",
        "anim_class": "anim-pptx",
        "badge_text": "Slide Deck Extractor",
    },
    "csv": {
        "name": "CSV Dataset",
        "icon": "📗",
        "unit": "Row Batch",
        "color": "#10B981",
        "bg": "#ECFDF5",
        "border": "#6EE7B7",
        "shadow": "rgba(16, 185, 129, 0.22)",
        "anim_class": "anim-grid",
        "badge_text": "Tabular Data Vectorizer",
    },
    "xlsx": {
        "name": "Excel Workbook",
        "icon": "📊",
        "unit": "Sheet",
        "color": "#059669",
        "bg": "#ECFDF5",
        "border": "#34D399",
        "shadow": "rgba(5, 150, 105, 0.22)",
        "anim_class": "anim-grid",
        "badge_text": "Multi-Sheet Matrix Engine",
    },
}

# ----------------------------------------------------------
# Figma 3D Tactile CSS & Format-Specific Keyframe Animations
# ----------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0F172A;
    }

    .stApp {
        background-color: #F4F6F9;
        background-image: radial-gradient(#CBD5E1 1.25px, transparent 1.25px);
        background-size: 24px 24px;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .block-container {
        padding-top: 2rem;
        padding-bottom: 6rem;
        max-width: 1120px;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
        border-right: 1px solid #E2E8F0;
        box-shadow: 10px 0 30px rgba(15, 23, 42, 0.03);
    }

    .sidebar-brand {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 18px;
        padding: 16px 18px;
        margin-bottom: 18px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.06),
                    0 4px 0px #E2E8F0;
        display: flex;
        align-items: center;
        gap: 14px;
        transition: transform 0.25s cubic-bezier(0.34, 1.56, 0.64, 1);
    }
    .sidebar-brand:hover {
        transform: translateY(-3px);
    }
    .brand-cube {
        width: 44px;
        height: 44px;
        border-radius: 13px;
        background: linear-gradient(135deg, #FF5A36 0%, #FF8A00 100%);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 22px;
        color: white;
        box-shadow: 0 6px 14px rgba(255, 90, 54, 0.35),
                    inset 0 2px 2px rgba(255, 255, 255, 0.4);
    }
    .brand-title {
        font-weight: 800;
        font-size: 1.05rem;
        letter-spacing: -0.02em;
        color: #0F172A;
        line-height: 1.2;
    }
    .brand-sub {
        font-size: 0.74rem;
        font-weight: 600;
        color: #64748B;
    }

    .status-pill {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 6px 12px;
        border-radius: 999px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-bottom: 14px;
        width: 100%;
        box-sizing: border-box;
    }
    .status-online {
        background: #ECFDF5;
        color: #059669;
        border: 1px solid #A7F3D0;
    }
    .status-offline {
        background: #FFFBEB;
        color: #D97706;
        border: 1px solid #FDE68A;
    }
    .pulse-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: currentColor;
        animation: pulse-ring 1.6s infinite;
    }
    @keyframes pulse-ring {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 currentColor; }
        70% { transform: scale(1); box-shadow: 0 0 0 6px transparent; }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 transparent; }
    }

    .format-grid {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 6px;
        margin: 10px 0 18px 0;
    }
    .fmt-chip {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-bottom: 3px solid #CBD5E1;
        border-radius: 10px;
        padding: 7px 4px;
        text-align: center;
        font-size: 0.68rem;
        font-weight: 700;
        color: #475569;
        transition: all 0.2s cubic-bezier(0.34, 1.56, 0.64, 1);
        cursor: default;
    }
    .fmt-chip:hover {
        transform: translateY(-3px) scale(1.05);
        border-color: #0EA5E9;
        border-bottom-color: #0284C7;
        color: #0F172A;
        box-shadow: 0 6px 12px rgba(14, 165, 233, 0.15);
    }
    .fmt-chip.active-fmt {
        background: #FFF7ED;
        border-color: #FF5A36;
        border-bottom-color: #EA580C;
        color: #C2410C;
        transform: translateY(-2px);
    }

    .hero-bar {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-bottom: 4px solid #E2E8F0;
        border-radius: 24px;
        padding: 24px 30px;
        margin-bottom: 24px;
        box-shadow: 0 18px 35px -10px rgba(15, 23, 42, 0.06);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 20px;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #FFF1EE;
        color: #FF5A36;
        border: 1px solid #FFD5CC;
        padding: 5px 12px;
        border-radius: 999px;
        font-size: 0.74rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 8px;
    }
    .hero-title {
        font-size: 1.75rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        color: #0F172A;
        margin: 0;
    }
    .hero-sub {
        font-size: 0.92rem;
        color: #64748B;
        margin-top: 4px;
        font-weight: 500;
    }

    .stage-3d-card {
        background: #FFFFFF;
        border: 2px solid var(--theme-border, #E2E8F0);
        border-bottom: 6px solid var(--theme-color, #CBD5E1);
        border-radius: 24px;
        padding: 24px 28px;
        margin-bottom: 26px;
        box-shadow: 0 20px 40px -12px var(--theme-shadow, rgba(15,23,42,0.08));
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 24px;
        position: relative;
        overflow: hidden;
        transition: transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
    }
    .stage-3d-card:hover {
        transform: translateY(-4px);
    }

    .icon-3d-stage {
        width: 84px;
        height: 84px;
        border-radius: 22px;
        background: var(--theme-bg, #F8FAFC);
        border: 2px solid var(--theme-border, #E2E8F0);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 40px;
        position: relative;
        overflow: hidden;
        flex-shrink: 0;
        box-shadow: 0 12px 24px -6px var(--theme-shadow),
                    inset 0 2px 4px rgba(255,255,255,0.8);
        animation: float-3d 3.2s ease-in-out infinite;
    }

    @keyframes float-3d {
        0%, 100% { transform: translateY(0px) rotate(0deg); }
        50% { transform: translateY(-7px) rotate(2deg); }
    }

    .anim-pdf::after {
        content: "";
        position: absolute;
        left: 0;
        right: 0;
        height: 4px;
        background: linear-gradient(90deg, transparent, #EF4444, #FCA5A5, #EF4444, transparent);
        box-shadow: 0 0 12px #EF4444;
        animation: pdf-laser 2s ease-in-out infinite;
    }
    @keyframes pdf-laser {
        0%, 100% { top: 8%; opacity: 0.9; }
        50% { top: 88%; opacity: 1; }
    }

    .anim-docx::after {
        content: "";
        position: absolute;
        inset: 0;
        border-radius: 20px;
        border: 3px solid #0EA5E9;
        animation: docx-wave 2.2s ease-out infinite;
    }
    @keyframes docx-wave {
        0% { transform: scale(0.75); opacity: 0.9; }
        100% { transform: scale(1.25); opacity: 0; }
    }

    .anim-pptx::after {
        content: "";
        position: absolute;
        top: 0;
        left: -120%;
        width: 80%;
        height: 100%;
        background: linear-gradient(
            110deg,
            transparent,
            rgba(249, 115, 22, 0.35),
            transparent
        );
        animation: pptx-slide 2.1s cubic-bezier(0.4, 0, 0.2, 1) infinite;
    }
    @keyframes pptx-slide {
        0% { left: -100%; }
        100% { left: 150%; }
    }

    .anim-grid::after {
        content: "";
        position: absolute;
        inset: 0;
        background-image: linear-gradient(to right, rgba(16, 185, 129, 0.22) 1px, transparent 1px),
                          linear-gradient(to bottom, rgba(16, 185, 129, 0.22) 1px, transparent 1px);
        background-size: 14px 14px;
        animation: grid-pulse 2s infinite alternate;
    }
    @keyframes grid-pulse {
        0% { opacity: 0.25; transform: scale(0.95); }
        100% { opacity: 0.85; transform: scale(1.05); }
    }

    .telemetry-row {
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
    }
    .stat-box {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-bottom: 3px solid #CBD5E1;
        border-radius: 14px;
        padding: 10px 16px;
        min-width: 95px;
        text-align: center;
        transition: transform 0.2s ease;
    }
    .stat-box:hover {
        transform: translateY(-2px);
    }
    .stat-val {
        font-size: 1.05rem;
        font-weight: 800;
        color: #0F172A;
    }
    .stat-lbl {
        font-size: 0.68rem;
        font-weight: 700;
        text-transform: uppercase;
        color: #64748B;
        letter-spacing: 0.04em;
    }

    .showcase-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
        gap: 16px;
        margin-top: 20px;
    }
    .showcase-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-bottom: 4px solid var(--card-accent, #CBD5E1);
        border-radius: 20px;
        padding: 20px 18px;
        text-align: left;
        box-shadow: 0 10px 25px -8px rgba(15, 23, 42, 0.06);
        transition: all 0.25s cubic-bezier(0.34, 1.56, 0.64, 1);
    }
    .showcase-card:hover {
        transform: translateY(-6px) rotate(-0.5deg);
        box-shadow: 0 20px 30px -10px rgba(15, 23, 42, 0.12);
    }
    .sc-icon {
        width: 46px;
        height: 46px;
        border-radius: 14px;
        background: var(--card-bg, #F8FAFC);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 24px;
        margin-bottom: 12px;
    }
    .sc-title {
        font-weight: 800;
        font-size: 0.95rem;
        color: #0F172A;
        margin-bottom: 4px;
    }
    .sc-desc {
        font-size: 0.76rem;
        color: #64748B;
        line-height: 1.4;
    }

    .citation-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #FF5A36;
        border-radius: 12px;
        padding: 12px 14px;
        margin-bottom: 10px;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03);
    }
    .citation-chip {
        display: inline-block;
        background: #FFF1EE;
        color: #EA580C;
        font-size: 0.7rem;
        font-weight: 800;
        padding: 3px 9px;
        border-radius: 6px;
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .citation-text {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.78rem;
        color: #334155;
        line-height: 1.5;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------
# Multi-Format Document Loader Engine
# ----------------------------------------------------------
def load_any_document(file_bytes: bytes, filename: str):
    ext = filename.rsplit(".", 1)[-1].lower()
    docs = []

    if ext == "pdf":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        try:
            raw_docs = PyPDFLoader(tmp_path).load()
            for d in raw_docs:
                page_num = d.metadata.get("page", 0) + 1
                d.metadata["unit_label"] = f"Page {page_num}"
                docs.append(d)
        finally:
            os.unlink(tmp_path)

    elif ext == "docx":
        doc = DocxDocument(io.BytesIO(file_bytes))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                if row_text:
                    paragraphs.append(row_text)

        group_size = 6
        for idx, i in enumerate(range(0, max(1, len(paragraphs)), group_size), start=1):
            block = "\n\n".join(paragraphs[i : i + group_size])
            if block:
                docs.append(
                    Document(
                        page_content=block,
                        metadata={"page": idx - 1, "unit_label": f"Section {idx}"},
                    )
                )

    elif ext == "pptx":
        prs = Presentation(io.BytesIO(file_bytes))
        for slide_idx, slide in enumerate(prs.slides, start=1):
            slide_texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    txt = shape.text_frame.text.strip()
                    if txt:
                        slide_texts.append(txt)
            if slide_texts:
                docs.append(
                    Document(
                        page_content="\n".join(slide_texts),
                        metadata={"page": slide_idx - 1, "unit_label": f"Slide {slide_idx}"},
                    )
                )

    elif ext == "csv":
        df = pd.read_csv(io.BytesIO(file_bytes))
        rows_per_chunk = 25
        for idx, start in enumerate(range(0, len(df), rows_per_chunk), start=1):
            sub_df = df.iloc[start : start + rows_per_chunk]
            end_row = min(start + rows_per_chunk, len(df))
            text_table = sub_df.to_string(index=False)
            docs.append(
                Document(
                    page_content=f"Columns: {', '.join(map(str, df.columns))}\n\n{text_table}",
                    metadata={
                        "page": idx - 1,
                        "unit_label": f"Rows {start + 1}-{end_row}",
                    },
                )
            )

    elif ext == "xlsx":
        excel = pd.ExcelFile(io.BytesIO(file_bytes))
        for sheet_idx, sheet_name in enumerate(excel.sheet_names, start=1):
            df = excel.parse(sheet_name)
            if not df.empty:
                rows_per_chunk = 30
                for start in range(0, len(df), rows_per_chunk):
                    sub_df = df.iloc[start : start + rows_per_chunk]
                    end_row = min(start + rows_per_chunk, len(df))
                    docs.append(
                        Document(
                            page_content=(
                                f"Sheet: {sheet_name} | Columns: {', '.join(map(str, df.columns))}\n\n"
                                f"{sub_df.to_string(index=False)}"
                            ),
                            metadata={
                                "page": sheet_idx - 1,
                                "unit_label": f"Sheet {sheet_name} (Rows {start + 1}-{end_row})",
                            },
                        )
                    )

    return docs, ext


def build_vectorstore(file_bytes: bytes, filename: str, size: int, overlap: int):
    docs, ext = load_any_document(file_bytes, filename)
    if not docs:
        raise ValueError("No readable text content was found inside this file.")

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=size,
        chunk_overlap=overlap,
        separators=["\n\n", "\n", ". ", " ", ""],
    )
    chunks = splitter.split_documents(docs)

    embeddings = OpenAIEmbeddings(model="text-embedding-3-small")
    vectorstore = FAISS.from_documents(chunks, embeddings)
    return vectorstore, len(docs), len(chunks), ext


# ----------------------------------------------------------
# Interactive 3D Sidebar
# ----------------------------------------------------------
with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="brand-cube">✨</div>
            <div>
                <div class="brand-title">DocuVerse Studio</div>
                <div class="brand-sub">Multi-Format 3D RAG Engine</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    api_key = st.text_input(
        "🔑 OpenAI API Key",
        type="password",
        value=os.getenv("OPENAI_API_KEY", ""),
        placeholder="Paste your sk-... key",
    )

    if api_key:
        st.markdown(
            '<div class="status-pill status-online"><span class="pulse-dot"></span>Neural Engine Online</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="status-pill status-offline"><span class="pulse-dot"></span>Awaiting API Key</div>',
            unsafe_allow_html=True,
        )

    st.markdown("##### 📂 Upload Knowledge Source")

    uploaded_file = st.file_uploader(
        "Drop PDF, Word, PPT, CSV, or Excel",
        type=["pdf", "docx", "pptx", "csv", "xlsx"],
        label_visibility="collapsed",
    )

    active_ext = (
        uploaded_file.name.rsplit(".", 1)[-1].lower()
        if uploaded_file is not None
        else None
    )

    pdf_cls = "active-fmt" if active_ext == "pdf" else ""
    docx_cls = "active-fmt" if active_ext == "docx" else ""
    pptx_cls = "active-fmt" if active_ext == "pptx" else ""
    csv_cls = "active-fmt" if active_ext == "csv" else ""
    xlsx_cls = "active-fmt" if active_ext == "xlsx" else ""

    st.markdown(
        f"""
        <div class="format-grid">
            <div class="fmt-chip {pdf_cls}">📕<br>PDF</div>
            <div class="fmt-chip {docx_cls}">📘<br>DOCX</div>
            <div class="fmt-chip {pptx_cls}">📙<br>PPTX</div>
            <div class="fmt-chip {csv_cls}">📗<br>CSV</div>
            <div class="fmt-chip {xlsx_cls}">📊<br>XLSX</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("##### 🧠 Intelligence & Retrieval")
    model_name = st.selectbox(
        "Chat Model",
        ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"],
    )

    with st.expander("⚙️ Vector Chunking Controls", expanded=False):
        chunk_size = st.slider("Chunk Size (chars)", 500, 2000, 1000, step=100)
        chunk_overlap = st.slider("Chunk Overlap", 0, 500, 200, step=50)
        top_k = st.slider("Retrieved Passages (k)", 2, 10, 4)

    if "messages" in st.session_state and len(st.session_state.messages) > 0:
        st.divider()
        if st.button("🧹 Reset Conversation", use_container_width=True):
            st.session_state.messages = []
            st.rerun()


# ----------------------------------------------------------
# Top Hero Bar
# ----------------------------------------------------------
st.markdown(
    """
    <div class="hero-bar">
        <div>
            <div class="hero-badge">⚡ Grounded Multi-Format RAG</div>
            <h1 class="hero-title">Interactive Document Studio</h1>
            <div class="hero-sub">Upload PDFs, Word docs, PowerPoint decks, CSVs, or Excel sheets — answers cite exact pages, slides, or rows.</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if not api_key:
    st.warning("👈 Enter your **OpenAI API Key** in the left studio dock to unlock multi-format indexing.")
    st.stop()

os.environ["OPENAI_API_KEY"] = api_key


# ----------------------------------------------------------
# File Indexing + Custom Format Animation Trigger
# ----------------------------------------------------------
if uploaded_file is not None:
    ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
    theme = FILE_THEMES.get(ext, FILE_THEMES["pdf"])
    file_sig = (uploaded_file.name, uploaded_file.size, chunk_size, chunk_overlap)

    if st.session_state.get("file_sig") != file_sig:
        status_placeholder = st.empty()
        status_placeholder.markdown(
            f"""
            <div class="stage-3d-card" style="--theme-color:{theme['color']}; --theme-bg:{theme['bg']}; --theme-border:{theme['border']}; --theme-shadow:{theme['shadow']};">
                <div style="display:flex; align-items:center; gap:20px;">
                    <div class="icon-3d-stage {theme['anim_class']}">{theme['icon']}</div>
                    <div>
                        <div style="font-size:0.75rem; font-weight:800; color:{theme['color']}; text-transform:uppercase; letter-spacing:0.05em;">
                            ⚡ {theme['badge_text']} Active...
                        </div>
                        <div style="font-size:1.25rem; font-weight:800; color:#0F172A; margin-top:2px;">
                            Scanning & Vectorizing {uploaded_file.name}
                        </div>
                        <div style="font-size:0.85rem; color:#64748B; margin-top:4px;">
                            Splitting into {chunk_size}-character semantic chunks & generating OpenAI embeddings...
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        try:
            vs, num_units, num_chunks, parsed_ext = build_vectorstore(
                uploaded_file.getvalue(), uploaded_file.name, chunk_size, chunk_overlap
            )
            st.session_state.vectorstore = vs
            st.session_state.file_sig = file_sig
            st.session_state.file_meta = {
                "name": uploaded_file.name,
                "size_kb": round(uploaded_file.size / 1024, 1),
                "units": num_units,
                "chunks": num_chunks,
                "ext": parsed_ext,
            }
            st.session_state.messages = []
            status_placeholder.empty()
            st.toast(f"{theme['icon']} Indexed {uploaded_file.name} ({num_chunks} vectors ready!)")
        except Exception as e:
            status_placeholder.empty()
            st.error(f"Could not parse `{uploaded_file.name}`: {e}")
            st.stop()

    meta = st.session_state.file_meta
    theme = FILE_THEMES.get(meta["ext"], FILE_THEMES["pdf"])
    st.markdown(
        f"""
        <div class="stage-3d-card" style="--theme-color:{theme['color']}; --theme-bg:{theme['bg']}; --theme-border:{theme['border']}; --theme-shadow:{theme['shadow']};">
            <div style="display:flex; align-items:center; gap:20px;">
                <div class="icon-3d-stage {theme['anim_class']}">{theme['icon']}</div>
                <div>
                    <div style="display:inline-block; background:{theme['bg']}; color:{theme['color']}; border:1px solid {theme['border']}; font-size:0.7rem; font-weight:800; padding:3px 10px; border-radius:999px; text-transform:uppercase; margin-bottom:6px;">
                        {theme['name']} &bull; {theme['badge_text']}
                    </div>
                    <div style="font-size:1.25rem; font-weight:800; color:#0F172A;">
                        {meta['name']}
                    </div>
                    <div style="font-size:0.82rem; color:#64748B; margin-top:2px;">
                        Vector index locked & ready — ask anything below.
                    </div>
                </div>
            </div>
            <div class="telemetry-row">
                <div class="stat-box">
                    <div class="stat-val">{meta['units']}</div>
                    <div class="stat-lbl">{theme['unit']}s</div>
                </div>
                <div class="stat-box">
                    <div class="stat-val">{meta['chunks']}</div>
                    <div class="stat-lbl">Vectors</div>
                </div>
                <div class="stat-box">
                    <div class="stat-val">{meta['size_kb']} KB</div>
                    <div class="stat-lbl">File Size</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------
# RAG Chain Setup
# ----------------------------------------------------------
RAG_PROMPT = ChatPromptTemplate.from_messages([
    ("system",
     "You are a helpful assistant that answers questions strictly using the "
     "provided context from an uploaded document.\n"
     "Rules:\n"
     "1. Answer ONLY from the context below.\n"
     "2. If the answer is not in the context, say: "
     "\"I couldn't find that in the document.\"\n"
     "3. Always cite the exact source unit tag you used in parentheses, e.g. (Page 3), (Slide 4), (Section 2), or (Rows 1-25).\n\n"
     "Context:\n{context}"),
    ("human", "{question}"),
])


def format_docs(docs) -> str:
    """Tag each retrieved chunk with its unit label (Page, Slide, Section, or Row batch)."""
    formatted = []
    for d in docs:
        page_num = d.metadata.get("page", 0) + 1
        label = d.metadata.get("unit_label", f"Page {page_num}")
        formatted.append(f"[{label}]\n{d.page_content}")
    return "\n\n".join(formatted)


def get_chain(vectorstore: FAISS, model: str, k: int):
    retriever = vectorstore.as_retriever(search_kwargs={"k": k})
    llm = ChatOpenAI(model=model, temperature=0)
    return (
        {"context": retriever | format_docs, "question": RunnablePassthrough()}
        | RAG_PROMPT
        | llm
        | StrOutputParser()
    ), retriever


def render_citations(sources: list):
    if not sources:
        return
    with st.expander(f"🔎 Inspect {len(sources)} Verified Source Passages"):
        for src in sources:
            st.markdown(
                f"""
                <div class="citation-card">
                    <div class="citation-chip">📍 {src['label']}</div>
                    <div class="citation-text">{src['excerpt']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# ----------------------------------------------------------
# Chat Interface or 3D Multi-Format Onboarding Showcase
# ----------------------------------------------------------
if "vectorstore" in st.session_state and uploaded_file is not None:
    if "messages" not in st.session_state:
        st.session_state.messages = []

    if len(st.session_state.messages) == 0:
        st.markdown("###### 💡 Quick-Start Queries")
        c1, c2, c3 = st.columns(3)
        quick_q = None
        if c1.button("📌 Summarize key takeaways", use_container_width=True):
            quick_q = "Summarize the main points and key takeaways of this document."
        if c2.button("📊 Extract core numbers & metrics", use_container_width=True):
            quick_q = "List the most important numbers, dates, or metrics mentioned in this document."
        if c3.button("🎯 List action items or conclusions", use_container_width=True):
            quick_q = "What are the primary conclusions, findings, or action items in this document?"
    else:
        quick_q = None

    for msg in st.session_state.messages:
        avatar = "🧑‍💻" if msg["role"] == "user" else "✨"
        with st.chat_message(msg["role"], avatar=avatar):
            st.markdown(msg["content"])
            if msg["role"] == "assistant" and msg.get("sources"):
                render_citations(msg["sources"])

    chat_q = st.chat_input(f"Ask a question about {uploaded_file.name}...")
    question = quick_q or chat_q

    if question:
        st.session_state.messages.append({"role": "user", "content": question})
        with st.chat_message("user", avatar="🧑‍💻"):
            st.markdown(question)

        chain, retriever = get_chain(
            st.session_state.vectorstore, model_name, top_k
        )

        with st.chat_message("assistant", avatar="✨"):
            with st.spinner("Scanning vector index & synthesizing answer..."):
                answer = chain.invoke(question)
                retrieved_docs = retriever.invoke(question)

            st.markdown(answer)

            formatted_sources = []
            for doc in retrieved_docs:
                page_num = doc.metadata.get("page", 0) + 1
                label = doc.metadata.get("unit_label", f"Page {page_num}")
                formatted_sources.append({
                    "label": label,
                    "excerpt": doc.page_content[:450].strip() + "...",
                })

            render_citations(formatted_sources)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer,
            "sources": formatted_sources,
        })
        if quick_q:
            st.rerun()

else:
    st.markdown(
        """
        <div style="margin-top: 8px;">
            <div style="font-weight: 800; font-size: 1.05rem; color: #0F172A;">
                🚀 Choose a Format & Drop Your File in the Left Dock
            </div>
            <div style="font-size: 0.85rem; color: #64748B;">
                Each file format triggers its own custom parser, citation tracker, and 3D stage animation:
            </div>
            <div class="showcase-grid">
                <div class="showcase-card" style="--card-accent: #EF4444; --card-bg: #FEF2F2;">
                    <div class="sc-icon">📕</div>
                    <div class="sc-title">PDF Documents</div>
                    <div class="sc-desc">Laser-scanned page extraction with exact <b>(Page X)</b> citations.</div>
                </div>
                <div class="showcase-card" style="--card-accent: #0EA5E9; --card-bg: #F0F9FF;">
                    <div class="sc-icon">📘</div>
                    <div class="sc-title">Word (.docx)</div>
                    <div class="sc-desc">Streams paragraphs & tables with structured <b>(Section X)</b> tags.</div>
                </div>
                <div class="showcase-card" style="--card-accent: #F97316; --card-bg: #FFF7ED;">
                    <div class="sc-icon">📙</div>
                    <div class="sc-title">PowerPoint (.pptx)</div>
                    <div class="sc-desc">Extracts slide decks & speaker frames with <b>(Slide X)</b> citations.</div>
                </div>
                <div class="showcase-card" style="--card-accent: #10B981; --card-bg: #ECFDF5;">
                    <div class="sc-icon">📗</div>
                    <div class="sc-title">CSV Datasets</div>
                    <div class="sc-desc">Vectorizes tabular rows with <b>(Rows X-Y)</b> batch references.</div>
                </div>
                <div class="showcase-card" style="--card-accent: #059669; --card-bg: #ECFDF5;">
                    <div class="sc-icon">📊</div>
                    <div class="sc-title">Excel (.xlsx)</div>
                    <div class="sc-desc">Parses multi-sheet workbooks with exact <b>(Sheet + Row)</b> citations.</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )