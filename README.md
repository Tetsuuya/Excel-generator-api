# 📊 AI-Powered Secure Python Excel Generator

A high-performance Python microservice that generates executive-grade Excel (`.xlsx`) workbooks from natural language prompts using **Groq LPUs** (`llama-3.3-70b-versatile`), protected by **AST static code analysis** and **isolated subprocess sandboxing**.

---

## 🔒 Security Architecture (Defense in Depth)

1. **Layer 1: AST Code Sanitizer (`core/sanitizer.py`)**
   - Parses the generated Python code into an Abstract Syntax Tree (AST) before execution.
   - **Allowed Root Modules**: `openpyxl`, `xlsxwriter`, `pandas`, `numpy`, `datetime`, `math`, `random`, `decimal`, `itertools`, `collections`, `re`, `json`, `string`.
   - **Banned**: `os`, `sys`, `subprocess`, `socket`, `urllib`, `requests`, `eval`, `exec`, `open`, `__import__`, `globals()`, and dunder attribute traversal (`__subclasses__`, `__class__`, etc.).
   
2. **Layer 2: Isolated Subprocess Execution (`core/executor.py`)**
   - Runs code in an ephemeral temporary directory.
   - Strips sensitive host environment variables (API keys, credentials).
   - Hard execution timeout (15s) to eliminate infinite loops.

3. **Layer 3: Self-Healing LLM Loop (`services/excel_service.py`)**
   - If an error occurs (e.g. typo in cell formula or `openpyxl` coordinate), the error is fed back to the LLM to fix itself automatically within ~1 second.

---

## 🚀 Getting Started

### 1. Activate the Virtual Environment (`venv`)

#### **Windows (PowerShell - Recommended)**
```powershell
.\venv\Scripts\Activate.ps1
```
> *Tip: If you get a script execution policy error in PowerShell, run this first:*
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> .\venv\Scripts\Activate.ps1
> ```

#### **Windows (Command Prompt / CMD)**
```cmd
venv\Scripts\activate.bat
```

#### **Git Bash / Linux / macOS**
```bash
source venv/Scripts/activate
# or on Linux/macOS:
source venv/bin/activate
```

*(Once activated, you will see `(venv)` in your terminal prompt).*

---

### 2. Configure `.env`
Add your API key (OpenRouter, DeepSeek, or Groq) in `.env`:
```env
# OpenRouter Free Tier
OPENAI_BASE_URL=https://openrouter.ai/api/v1
OPENAI_API_KEY=sk-or-v1-your_openrouter_key_here

PORT=8001
HOST=0.0.0.0
```

---

### 3. Run Verification Tests
```powershell
python test_cli.py
```
*(Or without activating venv: `.\venv\Scripts\python test_cli.py`)*

---

### 4. Start the FastAPI Server
```powershell
python app.py
# or using uvicorn:
uvicorn app:app --port 8001 --reload
```
* **Web UI Dashboard**: Open [http://localhost:8001](http://localhost:8001) in your browser.
* **Interactive Swagger API Docs**: [http://localhost:8001/docs](http://localhost:8001/docs)

---

## 📡 API Endpoints

### 1. `POST /api/generate-excel` (Download .xlsx directly)
**Request Body:**
```json
{
  "prompt": "Create a 2026 SaaS Financial Model with revenue, COGS, OpEx, net income, dynamic formulas, and a bar chart comparing quarters.",
  "filename": "SaaS_Financial_Model_2026.xlsx"
}
```
**Response:** Binary `.xlsx` file stream with content type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.

### 2. `POST /api/generate-excel/preview` (Preview Code & Metadata)
**Request Body:**
```json
{
  "prompt": "Create an invoice template with item description, quantity, unit price, and total."
}
```
**Response:**
```json
{
  "success": true,
  "model": "llama-3.3-70b-versatile",
  "attempts": 1,
  "duration_seconds": 1.12,
  "code": "def generate_excel(output_path: str): ...",
  "file_size_bytes": 6240
}
```

---

## 🌐 Calling from Node.js (e.g. `Docs-Service`)

```javascript
const axios = require('axios');
const fs = require('fs');

async function createSpreadsheet(prompt, outputPath) {
  const response = await axios.post('http://localhost:8001/api/generate-excel', {
    prompt: prompt,
    filename: 'output.xlsx'
  }, {
    responseType: 'arraybuffer'
  });

  fs.writeFileSync(outputPath, response.data);
  console.log(`Excel file saved to ${outputPath}`);
}
```
