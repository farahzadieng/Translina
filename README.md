<div align="center">

# Translina

**Local models. Local documents. Private translation.**

`Farsi` · `Arabic` · `English`

</div>

---

# About

> **Private, local document translation powered by local LLMs.**

**Translina** is a local-first translation application for translating **Microsoft Word documents (`.docx`)**, **digital PDFs**, and **plain text** between:

- 🇮🇷 **Farsi / Persian**
- 🇸🇦 **Arabic**
- 🇬🇧 **English**

Translina runs translation models **locally on your machine**, keeping documents and text private without requiring a cloud translation service.

The translation pipeline uses multiple stages of translation, editing, and verification to improve **accuracy, fluency, terminology consistency, and preservation of meaning**.

---

## ✨ Features

- 🔒 **Local & private** — documents are processed on your machine
- 🌐 **Farsi, Arabic, and English**
- 📄 **DOCX translation**
- 📕 **Digital PDF translation**
- 📝 **Plain-text translation**
- 🧠 Powered by local **GGUF LLMs**
- ✅ Multi-stage translation verification
- 🎯 Configurable translation tone
- 🔄 Translation progress tracking
- ⛔ Translation cancellation
- 💾 Checkpoint-based processing
- ↔️ Automatic **RTL / LTR** handling
- 🖥️ macOS Apple Silicon and Windows support
- 🚫 No cloud translation API required

---

## 🧠 Translation Pipeline

Translina does more than send a piece of text to an LLM and return the result.

Documents pass through a structured translation pipeline:

```text
Document
   │
   ▼
Structural Analysis
   │
   ▼
Document & Context Analysis
   │
   ▼
Glossary / Translation Memory
   │
   ▼
Initial Translation
   │
   ▼
Target-Language Editing
   │
   ▼
Fidelity Verification
   │
   ▼
Consistency Verification
   │
   ▼
Document Reconstruction
   │
   ▼
Translated Document
```

### Three quality layers

After the initial translation, Translina uses multiple verification stages.

**1. Language Editing**

Improves sentence structure, readability, fluency, and target-language style without intentionally changing the original meaning.

**2. Fidelity Verification**

Checks important translation details such as:

- meaning
- numbers
- names
- dates
- negation
- terminology
- missing information
- unintended additions

**3. Consistency Verification**

Checks the complete document for consistent:

- terminology
- tone
- names
- punctuation
- translation decisions
- RTL / LTR direction

Suspicious sections can be corrected independently instead of translating the entire document again.

---

## 🤖 Models

Translina is designed around the following local GGUF models:

```text
Qwen2.5-14B-Instruct
Gemma 3 12B Instruct
Gemma 3 4B Instruct
```

For example:

```text
Qwen2.5-14B-Instruct-Q4_K_M.gguf
google_gemma-3-12b-it-Q4_K_M.gguf
google_gemma-3-4b-it-Q4_K_M.gguf
```

### Recommended model profiles

#### 16 GB RAM

```text
Qwen 2.5 14B
    ↓
Document analysis
Glossary extraction
Initial translation

Gemma 3 12B
    ↓
Target-language editing
Quality verification
```

The models are loaded sequentially so memory can be released before the next model is loaded.

#### 8 GB RAM

```text
Gemma 3 4B
    ↓
Analysis
Translation
Editing
Verification
```

The 4B model performs the different pipeline roles in separate passes and is the recommended low-memory configuration.

---

# 🚀 Setup

## Requirements

Recommended:

- **Python 3.12**
- `llama-cpp-python`
- Flask
- lxml
- python-docx
- PyMuPDF
- Local GGUF models

You will need local copies of the required models before starting Translina.

Add the model locations to:

```text
translator.config.json
```

---

# 🍎 macOS — Apple Silicon

Translina can use the **Metal** backend through `llama.cpp` on Apple Silicon.

## 1. Create a virtual environment

```bash
python3.12 -m venv .venv
```

## 2. Activate it

```bash
source .venv/bin/activate
```

## 3. Upgrade pip

```bash
python -m pip install --upgrade pip
```

## 4. Install dependencies

```bash
python -m pip install Flask lxml python-docx PyMuPDF
```

## 5. Install `llama-cpp-python` with Metal support

```bash
python -m pip install llama-cpp-python \
  --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/metal
```

---

## 🔍 Check the installation

Before starting the application, run the built-in diagnostic and model probe:

```bash
python app.py --config translator.config.json --doctor --probe-model
```

This can be used to verify the application environment and model configuration before starting the server.

---

## ▶️ Run Translina

```bash
python app.py --config translator.config.json
```

Then open:

```text
http://127.0.0.1:5000
```

---

# 🪟 Windows

## 1. Create a virtual environment

```powershell
py -3.12 -m venv .venv
```

## 2. Activate it

```powershell
.venv\Scripts\Activate.ps1
```

## 3. Upgrade pip

```powershell
python -m pip install --upgrade pip
```

## 4. Install dependencies

```powershell
python -m pip install Flask lxml python-docx PyMuPDF llama-cpp-python
```

## 5. Check the installation

```powershell
python app.py --config translator.config.json --doctor --probe-model
```

## 6. Run Translina

```powershell
python app.py --config translator.config.json
```

Then open:

```text
http://127.0.0.1:5000
```

---

# ⚙️ Configuration

Translina reads its configuration from:

```text
translator.config.json
```

The configuration is used to define settings such as:

- model locations
- model roles
- inference backend
- context settings
- GPU layer settings
- application data paths
- fallback fonts
- optional LibreOffice location

Model paths can be configured according to the machine running Translina.

> Keep model files outside the Git repository. GGUF models are large and should not normally be committed to source control.

---

# 📄 Supported Inputs

| Input                  | Supported | Notes                                              |
| ---------------------- | :-------: | -------------------------------------------------- |
| Plain text             |    ✅     | Preserves paragraphs and blank lines               |
| Microsoft Word `.docx` |    ✅     | Attempts to preserve document structure and styles |
| Digital PDF            |    ✅     | PDF must contain a text layer                      |
| Scanned PDF            |    ❌     | OCR is not currently included                      |
| PowerPoint             |    ❌     | Not currently supported                            |
| Excel                  |    ❌     | Not currently supported                            |

---

# 📝 DOCX Translation

Translina translates complete semantic units rather than independently translating every Word formatting run.

The document engine is designed to preserve as much of the original structure as possible, including:

- paragraphs
- headings
- styles
- tables
- headers and footers
- hyperlinks
- bold and italic text
- page and section breaks
- document direction

For Persian and Arabic output, document direction can be changed to **RTL**.

English output uses **LTR**.

Because translated text may be shorter or longer than the original, exact pixel-perfect pagination is not guaranteed.

---

# 📕 PDF Translation

Translina supports **digital PDFs containing a text layer**.

The PDF pipeline analyzes elements such as:

- text blocks
- fonts
- colors
- images
- margins
- headings
- repeated document elements

The translated PDF may use **reflow**, meaning the number of pages can change depending on the translated text.

The goal is to preserve the overall visual hierarchy and readability rather than reproduce the source PDF pixel-for-pixel.

> **Scanned PDFs and OCR are not currently supported.**

---

# 🎭 Translation Tone

Translations can be adapted for different styles of writing.

Supported tone profiles include:

```text
Automatic
Formal
Legal
Business
Academic
Literary
Conversational
Screenplay
```

An optional custom description can also be used to provide additional context about the intended audience or writing style.

---

# 🌍 Language Direction

Translina handles both right-to-left and left-to-right languages.

```text
Farsi   → RTL
Arabic  → RTL
English → LTR
```

Direction-sensitive document elements are adjusted when producing translated output.

---

# 🔐 Privacy

Translina is designed to operate locally.

By default:

```text
Your document
     │
     ▼
Your machine
     │
     ▼
Local LLM
     │
     ▼
Translated document
```

Your document does **not need to be uploaded to a cloud translation provider**.

The application listens on the local loopback interface by default:

```text
127.0.0.1
```

This makes Translina suitable for documents where privacy or confidentiality is important.

---

# 🏗️ Architecture

At a high level, Translina consists of:

```text
┌──────────────────────────┐
│       Flask UI/API       │
└────────────┬─────────────┘
             │
┌────────────▼─────────────┐
│       Job Manager        │
└────────────┬─────────────┘
             │
    ┌────────▼────────┐
    │ Document Engine │
    └────────┬────────┘
             │
┌────────────▼─────────────┐
│   Translation Pipeline   │
└────────────┬─────────────┘
             │
     ┌───────▼────────┐
     │ Context Manager│
     └───────┬────────┘
             │
      ┌──────▼───────┐
      │ Model Manager │
      │ llama.cpp     │
      └──────┬───────┘
             │
   ┌─────────▼──────────┐
   │ Quality Controller │
   └─────────┬──────────┘
             │
   ┌─────────▼──────────┐
   │ Output Validation  │
   └────────────────────┘
```

The application uses `llama.cpp` through `llama-cpp-python` for local GGUF model inference.

---

# 💾 Long Documents

Long documents are split using natural boundaries such as:

- sentences
- paragraphs
- sections
- scenes
- dialogue boundaries

Translina maintains contextual information including:

```text
Document summary
Section summary
Glossary
Names
Translation decisions
Nearby context
Current text
```

This helps maintain terminology and translation consistency across larger documents while remaining within the model's context window.

---

# 🩺 Diagnostics

If Translina is not starting correctly, run:

```bash
python app.py --config translator.config.json --doctor --probe-model
```

This should be the first troubleshooting step when checking model or runtime configuration.

---

# ⚠️ Limitations

Translina is designed to improve translation quality through multiple LLM passes, but automatic translation should not be considered infallible.

Human review is still recommended for sensitive or high-stakes material such as:

- legal documents
- medical documents
- contracts
- regulatory content
- safety-critical material

Current limitations also include:

- no OCR for scanned PDFs
- no PowerPoint translation
- no Excel translation
- PDF layout may reflow
- translated DOCX pagination may differ from the original

---

# 🛠️ Tech Stack

```text
Python
Flask
llama.cpp
llama-cpp-python
python-docx
lxml
PyMuPDF
GGUF
```

Optional components may include **LibreOffice** for document rendering where available.

---

# 📁 Basic Usage

```text
1. Install the required Python packages
             ↓
2. Download / host the GGUF models locally
             ↓
3. Add model paths to translator.config.json
             ↓
4. Run the diagnostic
             ↓
5. Start app.py
             ↓
6. Open http://127.0.0.1:5000
             ↓
7. Select text or a document
             ↓
8. Choose source language
             ↓
9. Choose target language
             ↓
10. Select a tone
             ↓
11. Start translation
             ↓
12. Download the translated result
```

---

# 🤝 Contributing

Contributions, bug reports, testing, and suggestions are welcome.

If you find an issue, please include useful information such as:

- operating system
- Python version
- model being used
- available RAM
- inference backend
- input document type
- relevant error output

Please do **not** include confidential source documents in public bug reports.

---
