::: {align="center"}

# Translina

**Local models. Local documents. Private translation.**

`Persian` · `Arabic` · `English`
:::

## About

**Translina** is a local-first document translation web app.

It translates:

- plain text
- Microsoft Word (`.docx`)
- digital PDF files with a text layer

All translation runs locally. No cloud translation API is required.

The current pipeline uses three specialized local models instead of
asking one LLM to do everything:

```text
Document
   ↓
Gemma 3 12B
Context / terminology analysis
   ↓
MADLAD-400 7B MT-BT
Initial machine translation
   ↓
Gemma 3 12B
Semantic post-editing
   ↓
Deterministic QA
   ↓
Dorna Llama 3 8B
Persian-language polishing
   ↓
Final QA
   ↓
TXT / DOCX / PDF
```

Models are loaded sequentially so large models do not need to stay in
memory together.

## Features

- Fully local translation
- Persian, Arabic, and English
- TXT, DOCX, and digital PDF input
- DOCX structure/style preservation
- RTL/LTR handling
- Translation glossary
- Translation memory
- Number, URL, email, and terminology checks
- Job progress tracking
- Pause, resume, and cancel
- Checkpoint-based recovery
- Local job history
- Configurable translation tone
- Apple Silicon Metal acceleration for GGUF LLMs
- NVIDIA CUDA support for GGUF LLMs when `llama-cpp-python` is built
  with CUDA
- No Ollama dependency

## Models

Translina currently expects:

Role Model Backend

---

Initial translation MADLAD-400 7B MT-BT CTranslate2
Analysis + semantic review Gemma 3 12B Instruct Q4_K_M llama.cpp
Persian polishing Dorna Llama 3 8B Instruct Q5_K_M llama.cpp

Recommended model layout:

```text
Ai-Models/
├── madlad/
│   ├── model.bin
│   ├── config.json
│   ├── shared_vocabulary.json
│   ├── spiece.model
│   └── ...
├── gemma/
│   └── google_gemma-3-12b-it-Q4_K_M.gguf
└── dorna/
    └── dorna-llama3-8b-instruct.Q5_K_M.gguf
```

> Models are not included in this repository. Do not commit large
> GGUF/model files to Git.

## Requirements

Recommended:

- Python 3.12
- 16 GB RAM or more
- about 25 GB free disk space for the recommended models
- modern Apple Silicon Mac or NVIDIA GPU recommended

Python packages are listed in:

```text
requirements.txt
```

## 1. Get the models

Create a model directory outside the Git repository.

Example:

```text
~/Ai-Models
```

### MADLAD

Translina expects a **CTranslate2-converted** MADLAD-400 7B MT-BT model.

The configured path must point to the **model directory**, not to
`model.bin`.

Example:

```text
/Users/you/Ai-Models/madlad
```

### Gemma

Expected GGUF:

```text
google_gemma-3-12b-it-Q4_K_M.gguf
```

### Dorna

Expected GGUF:

```text
dorna-llama3-8b-instruct.Q5_K_M.gguf
```

Dorna may require accepting its Hugging Face access conditions before
downloading.

## 2. Configure model paths

Open:

```text
translator.config.json
```

Set:

```json
{
  "paths": {
    "model_root": "/absolute/path/to/Ai-Models",
    "data_dir": "./translator_data"
  }
}
```

Model entries should point to:

```text
${MODEL_ROOT}/madlad
${MODEL_ROOT}/gemma/google_gemma-3-12b-it-Q4_K_M.gguf
${MODEL_ROOT}/dorna/dorna-llama3-8b-instruct.Q5_K_M.gguf
```

Use forward slashes in JSON paths on Windows if possible:

```text
D:/AI/Ai-Models
```

## macOS --- Apple Silicon

Gemma and Dorna can use **Metal** through `llama.cpp`.

MADLAD uses CTranslate2 and currently runs on CPU on macOS.

### 1. Create the environment

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
```

### 2. Install normal dependencies

```bash
python -m pip install -r requirements.txt
```

### 3. Install `llama-cpp-python` with Metal

If the project includes `install_mac.sh`, the easiest option is:

```bash
chmod +x install_mac.sh
./install_mac.sh
```

Otherwise build/install `llama-cpp-python` with Metal enabled:

```bash
CMAKE_ARGS="-DGGML_METAL=on" \
python -m pip install --upgrade --force-reinstall llama-cpp-python --no-cache-dir
```

### 4. Verify Metal

```bash
python -c "from llama_cpp import llama_cpp; print(llama_cpp.llama_supports_gpu_offload())"
```

Expected:

```text
True
```

### 5. Run diagnostics

```bash
python app.py --config translator.config.json --doctor --probe-model
```

A healthy setup should report:

```text
status: ok
backend: ctranslate2+llama.cpp(metal)
llm_acceleration: metal
probe.ok: true
errors: []
```

### 6. Start Translina

```bash
python app.py --config translator.config.json
```

Open:

```text
http://127.0.0.1:5000
```

## Windows

On Windows, Gemma and Dorna can run through either CPU or an NVIDIA GPU.

For an NVIDIA GPU, install a CUDA-enabled build of `llama-cpp-python`.
MADLAD can also use the CTranslate2 CUDA backend when the required
NVIDIA/CUDA runtime is available.

### 1. Create the environment

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

### 2. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

### 3. Install `llama-cpp-python`

For CPU-only testing:

```powershell
python -m pip install llama-cpp-python
```

For NVIDIA/CUDA, install or build `llama-cpp-python` with CUDA support
for your local CUDA environment.

Example source build:

```powershell
$env:CMAKE_ARGS="-DGGML_CUDA=on"
python -m pip install --upgrade --force-reinstall llama-cpp-python --no-cache-dir
```

### 4. Configure MADLAD

For CPU:

```json
"device": "cpu",
"compute_type": "int8"
```

For a supported NVIDIA/CUDA setup:

```json
"device": "cuda",
"compute_type": "int8_float16"
```

The included `auto` profile can choose the appropriate configured
profile, but always confirm it with `--doctor`.

### 5. Run diagnostics

```powershell
python app.py --config translator.config.json --doctor --probe-model
```

Do not continue until:

```text
probe.ok: true
errors: []
```

### 6. Start Translina

```powershell
python app.py --config translator.config.json
```

Open:

```text
http://127.0.0.1:5000
```

## Using `uv`

`uv` is optional.

If the project environment is already active:

```bash
uv run --active app.py --config translator.config.json --doctor --probe-model
```

Run the server:

```bash
uv run --active app.py --config translator.config.json
```

If `uv` warns that `VIRTUAL_ENV` points to a different project, either
activate the correct `.venv` or omit `--active` and let `uv` use the
project environment.

## Diagnostics

Always run this first when moving the project to another computer:

```bash
python app.py --config translator.config.json --doctor --probe-model
```

It checks:

- Python dependencies
- model paths
- CTranslate2
- llama.cpp
- Metal/CUDA availability
- MADLAD translation
- Gemma generation
- Dorna generation

A successful probe is the best indication that the local model stack is
ready.

## Usage

```text
1. Start Translina
2. Open http://127.0.0.1:5000
3. Paste text or upload DOCX/PDF
4. Choose source language
5. Choose target language
6. Select a tone
7. Add glossary terms if needed
8. Start translation
9. Watch progress
10. Download the result
```

## Supported files

Input Support Notes

---

Plain text ✅ Preserves paragraphs
DOCX ✅ Attempts to preserve structure and formatting
Digital PDF ✅ Requires a text layer
Scanned PDF ❌ OCR is not included
PPTX ❌ Not currently supported
XLSX ❌ Not currently supported

## DOCX

Translina attempts to preserve:

- paragraphs
- headings
- styles
- tables
- headers and footers
- hyperlinks
- bold/italic formatting
- section/page breaks
- RTL/LTR direction

Exact pagination may change after translation.

## PDF

Only digital PDFs with extractable text are supported.

The output may reflow. The goal is readable translated layout, not
pixel-perfect reproduction of the original PDF.

## Translation modes

The application supports quality profiles such as:

```text
fast
balanced
strict
```

`balanced` is the normal default.

Higher-quality modes perform more review/polishing work and therefore
take longer.

## Tone

Available tone choices include:

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

A custom style/context description can also be supplied.

## Privacy

By default, Translina listens only on:

```text
127.0.0.1
```

Documents, extracted text, model prompts, and translation output can
remain on the local machine.

No cloud translation API is required.

## Project structure

A typical installation looks like:

```text
Translina/
├── app.py
├── translator.config.json
├── requirements.txt
├── install_mac.sh
├── font/
└── translator_data/
```

Runtime data is stored under `translator_data/`, including job state,
uploads, outputs, checkpoints, and translation memory.

Do not commit runtime data or local model files to Git.

Recommended `.gitignore` entries:

```gitignore
.venv/
__pycache__/
*.pyc
.DS_Store

translator_data/

*.gguf
*.bin
*.safetensors

.env
```

## Architecture

```text
Browser
   ↓
Flask UI / API
   ↓
Job Manager
   ↓
Document Processor
   ↓
Translation Pipeline
   ├── Gemma 3 12B / llama.cpp
   ├── MADLAD / CTranslate2
   ├── Deterministic QA
   └── Dorna 8B / llama.cpp
   ↓
Document Reconstruction
   ↓
Output
```

Only the model needed for the current stage is kept loaded where
possible. This reduces memory pressure on machines with limited
RAM/unified memory.

## Troubleshooting

### `Selected profile is missing usable models`

Check model paths:

```bash
python app.py --config translator.config.json --doctor
```

The MADLAD path must point to its directory:

```text
.../Ai-Models/madlad
```

not:

```text
.../Ai-Models/madlad/model.bin
```

### Metal is not being used

Run:

```bash
python -c "from llama_cpp import llama_cpp; print(llama_cpp.llama_supports_gpu_offload())"
```

If it returns `False`, reinstall `llama-cpp-python` with Metal support.

### `PyTorch was not found`

This is normally harmless for the current pipeline. Translina does not
use PyTorch for model inference.

### Out of memory

Try:

- closing other large applications
- using the low-memory profile
- reducing model context size
- using smaller GGUF quantizations
- processing fewer units per review request

### Translation stopped midway

Translina uses checkpoints. Open the job again and use resume when
available.

### PDF output differs from the source

This is expected for some PDFs. Translation changes text length, so the
reconstructed PDF may reflow.

## Moving to another computer

You need to copy:

```text
Translina/
Ai-Models/
```

Then on the new computer:

```text
1. Install Python 3.12
2. Create a fresh virtual environment
3. Install dependencies
4. Update model_root in translator.config.json
5. Install Metal or CUDA support if applicable
6. Run --doctor --probe-model
7. Start the app
```

Do **not** copy `.venv` between macOS and Windows. Recreate it on the
destination machine.

## Limitations

- Automatic translation can still make mistakes.
- Human review is recommended for medical, legal, regulatory,
  contractual, and safety-critical documents.
- Scanned PDF OCR is not included.
- PDF layout may reflow.
- DOCX pagination can change.
- Translation speed depends heavily on hardware and selected quality
  mode.

## Contributing

Bug reports and pull requests are welcome.

For useful bug reports, include:

- operating system
- CPU/GPU
- RAM
- Python version
- model/quantization
- selected profile
- input type
- output of `--doctor`
- relevant error message

Do not attach confidential source documents to public issues.

## License

Add the project's license here before publishing the repository
publicly.

If you have not selected a license yet, do not assume that public source
code is automatically open source.

---

**Recommended first test after every installation:**

```bash
python app.py --config translator.config.json --doctor --probe-model
```

If the probe succeeds, start the server and run a short real translation
before testing large documents.
