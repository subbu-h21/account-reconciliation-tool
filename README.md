# Account Reconciliation Tool

A Flask web application that reconciles bank account statements against internal accounting books — with optional AI-assisted transaction matching powered by **Google Gemini** or **OpenAI (via OpenRouter)**.

---

## Features

- **Upload & Compare** — Upload two Excel files (bank statement + internal books) and get an instant discrepancy report
- **Three-sheet Output** — Results exported as a formatted Excel file with separate sheets for Receivements, Payments, and Summary
- **AI Matching** — Intelligently matches transactions with different naming conventions (e.g. `"NEFT-KAPILA"` → `"Kapila Pharma"`)
- **Real-time Logs** — Live terminal in the UI streams AI matching progress via Server-Sent Events (SSE)
- **Dual AI Provider** — Switch between Google Gemini (yellow highlights) and OpenAI GPT via OpenRouter (orange highlights)
- **Dark UI** — Clean, responsive dark-theme interface with drag-and-drop file upload

---

## Screenshots

> Upload your files, optionally enable AI matching, and download the reconciliation report.

| Upload Screen | AI Matching Terminal |
|---|---|
| *(drag & drop Excel files, toggle AI)* | *(live logs streamed during processing)* |

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python, Flask |
| Data Processing | pandas, numpy, openpyxl, xlrd |
| AI Providers | Google Gemini 2.5 Flash, OpenAI GPT-4o-mini (OpenRouter) |
| Frontend | Vanilla JS, IBM Plex fonts, SSE |
| Output | Excel (.xlsx) with conditional formatting |

---

## Getting Started

### Prerequisites

- Python 3.9+
- A [Google Gemini API key](https://aistudio.google.com/apikey) and/or an [OpenRouter API key](https://openrouter.ai/keys)

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/subbu-h21/account-reconciliation-tool.git
cd account-reconciliation-tool/flask-backend

# 2. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate        # Windows
# source venv/bin/activate   # macOS/Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set up environment variables
cp .env.example .env
# Edit .env and add your API keys
```

### Configure API Keys

Edit `flask-backend/.env`:

```env
GEMINI_API_KEY=your_gemini_api_key_here
OPEN_ROUTER_API_KEY=your_open_router_api_key_here
```

### Run the App

```bash
python app.py
```

Open your browser at **http://localhost:5000**

---

## Usage

1. **Upload** your bank account statement (`.xlsx` / `.xls`)
2. **Upload** your internal books file (`.xlsx` / `.xls`)
3. *(Optional)* Toggle **AI Matching** and select a provider (Gemini or OpenAI)
4. Click **Process**
5. Watch the live terminal for AI matching progress
6. Click **Download** to get your reconciliation report

### Output Excel Structure

| Sheet | Contents |
|---|---|
| `Receivements` | Inbound transactions — entries in bank not in books, and vice versa |
| `Payments` | Outbound transactions — same discrepancy breakdown |
| `Summary` | Side-by-side metrics: counts, totals, closing balances, and differences |

AI-matched entries are highlighted:
- **Yellow** — matched by Gemini
- **Orange** — unmatched flagged by OpenAI

---

## Project Structure

```
flask-backend/
├── app.py                  # Flask app, routes, session management
├── tasks.py                # Background task processor
├── requirements.txt
├── .env.example            # API key template
├── templates/
│   └── index.html          # Frontend UI
├── services/
│   ├── reconciliation.py   # Core reconciliation logic
│   └── matcher.py          # AI matching (Gemini + OpenRouter)
└── uploads/                # Temporary output files (git-ignored)
```

---

## How It Works

```
User uploads 2 Excel files
        ↓
Flask creates a session ID and spawns a background thread
        ↓
pandas loads and filters both files
        ↓
reconciliation.py identifies unmatched entries (by name + amount + date)
        ↓
[If AI enabled] matcher.py batches entries (10 dates/batch) → sends to AI
        ↓
AI returns match flags → openpyxl applies color highlights
        ↓
Output .xlsx written → streamed to user for download
```

---

## Configuration

| Variable | Description |
|---|---|
| `GEMINI_API_KEY` | Google AI Studio API key for Gemini 2.5 Flash |
| `OPEN_ROUTER_API_KEY` | OpenRouter key for GPT-4o-mini access |

Max upload size: **50 MB** per file. Processing timeout: **600 seconds**.

---

## License

MIT License — free to use and modify.
