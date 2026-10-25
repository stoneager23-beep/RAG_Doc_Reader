<div align="center">

# ✨ DocuVerse 3D Studio
### **Grounded Multi-Format RAG Engine with Tactile 3D UI & Exact Unit Citations**

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![LangChain](https://img.shields.io/badge/LangChain-LCEL-1C3C3C?style=for-the-badge&logo=chainlink&logoColor=white)](https://www.langchain.com/)
[![OpenAI](https://img.shields.io/badge/OpenAI-Embeddings_%26_GPT--4o-412991?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com/)
[![FAISS](https://img.shields.io/badge/Vector_DB-FAISS-00599C?style=for-the-badge)](https://github.com/facebookresearch/faiss)

<p align="center">
  <b>Upload PDFs, Word Documents, PowerPoint Decks, CSVs, or Excel Workbooks.</b><br>
  Get zero-hallucination answers strictly grounded in your files — with verifiable Page, Slide, Section, and Row citations.
</p>

</div>

---

## 🌟 Overview

**DocuVerse 3D Studio** is an interactive Retrieval-Augmented Generation (RAG) workspace built with **Streamlit**, **LangChain (LCEL)**, **FAISS**, and **OpenAI**. 

Unlike standard single-PDF chatbots, DocuVerse features a **Multi-Format Ingestion Engine** paired with a bespoke **Figma-inspired 3D tactile interface**. Every supported file format triggers its own specialized parser, citation metadata tracker, and custom CSS keyframe stage animation.

---

## 🚀 Key Features

* **📂 Multi-Format Document Engine:** Seamlessly parses and indexes `.pdf`, `.docx`, `.pptx`, `.csv`, and `.xlsx` files into semantic vector chunks.
* **🎨 Tactile 3D Studio UI:** Designed with floating claymorphic cards, spring-physics hover states, custom telemetry badges, and format-specific upload animations:
  * 📕 **PDF:** Laser OCR scan beam animation + `(Page X)` citations.
  * 📘 **DOCX:** Ripple wave stream animation + `(Section X)` paragraph/table citations.
  * 📙 **PPTX:** Holographic slide shimmer + `(Slide X)` deck citations.
  * 📗 **CSV / 📊 XLSX:** Matrix grid pulse animation + `(Rows X–Y)` and `(Sheet + Row)` citations.
* **🛡️ Zero-Hallucination Guardrails:** Strict system prompting ensures the assistant answers *only* from retrieved document context and explicitly states when information is absent.
* **🔎 Persistent Source Inspector:** Every assistant response includes an interactive citation drawer displaying the exact raw text chunks and location tags used to generate the answer.
* **⚙️ Live Chunking & Retrieval Telemetry:** Tune `chunk_size`, `chunk_overlap`, and top-$k$ retrieved passages in real time from the studio dock.

---

## 🏗️ System Architecture

```text
+-------------------------------------------------------------------------+
|                        1. MULTI-FORMAT INGESTION                        |
|   [📕 PDF]       [📘 DOCX]       [📙 PPTX]       [📗 CSV]     [📊 XLSX]  |
+------+---------------+---------------+---------------+------------+-----+
       |               |               |               |            |
       v               v               v               v            v
+-------------------------------------------------------------------------+
|               2. UNIT-AWARE PARSING & METADATA TAGGING                  |
|        Tags chunks by Page #, Section #, Slide #, or Sheet/Row #        |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                  3. SEMANTIC CHUNKING & VECTORIZATION                   |
|        RecursiveCharacterTextSplitter  -->  text-embedding-3-small      |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                     4. IN-MEMORY FAISS VECTOR INDEX                     |
|             Top-k Similarity Search (Configurable k = 2..10)            |
+------------------------------------+------------------------------------+
                                     |
                                     v
+-------------------------------------------------------------------------+
|                5. GROUNDED LCEL RAG CHAIN & CITATION UI                 |
|      Context + Strict Prompt --> GPT-4o / 4o-mini --> Cited Answer      |
+-------------------------------------------------------------------------+
```

---

## 📄 Supported Formats & Citation Mapping

| Format | Extension | Parser Used | Citation Unit Tag | Stage Animation |
| :--- | :---: | :--- | :---: | :--- |
| **PDF Document** | `.pdf` | `PyPDFLoader` | `[Page X]` | 🔴 Laser Scanner Beam |
| **Word Document** | `.docx` | `python-docx` | `[Section X]` | 🔵 Sonic Ripple Wave |
| **PowerPoint Deck** | `.pptx` | `python-pptx` | `[Slide X]` | 🟠 Slide Shimmer Sweep |
| **CSV Dataset** | `.csv` | `pandas` | `[Rows X-Y]` | 🟢 Matrix Grid Pulse |
| **Excel Workbook** | `.xlsx` | `openpyxl` + `pandas` | `[Sheet Name (Rows X-Y)]` | 🟢 Matrix Grid Pulse |

---

## 🛠️ Local Installation & Setup

### 1. Clone the Repository
```bash
git clone [https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git)
cd YOUR_REPO_NAME
```

### 2. Create & Activate a Virtual Environment
```powershell
# Windows (PowerShell)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the Application
```bash
python -m streamlit run rag.py
```
The studio will open automatically in your browser at `http://localhost:8501`.

---

## 📁 Project Structure

```text
├── .streamlit/
│   └── config.toml        # Custom bright coral & slate studio theme
├── rag.py                 # Main 3D Studio UI & Multi-Format LCEL RAG application
├── requirements.txt       # Python dependencies for cloud & local deployment
├── .gitignore             # Excludes .env, virtual environments, and temp files
└── README.md              # Project documentation
```

---

## ☁️ Deploying on Streamlit Community Cloud

1. Push this repository to your GitHub account.
2. Log in to [share.streamlit.io](https://share.streamlit.io/) and click **Create app**.
3. Select your repository, branch (`main`), and main file path (`rag.py`).
4. *(Optional)* In **Advanced settings -> Secrets**, add your OpenAI key if you want it pre-configured:
   ```toml
   OPENAI_API_KEY = "sk-proj-your-api-key"
   ```
5. Click **Deploy!**

---

## 🤝 License & Credits

Built with ❤️ using **Streamlit**, **LangChain**, **FAISS**, and **OpenAI**. Feel free to fork, star ⭐, and customize for your own document intelligence workflows!