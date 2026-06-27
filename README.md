# Account Reconciliation Tool

A Flask web app that reconciles a bank account statement against internal accounting books. Upload two Excel files and get an instant discrepancy report — with optional AI-assisted transaction matching powered by **OpenRouter (GPT-4o-mini)**.

---

## Features

- **Upload & Compare** — Upload two Excel files (bank statement + internal books) and get a formatted discrepancy report
- **Three-sheet Output** — Results exported as an Excel file with separate sheets for Receivements, Payments, and Summary
- **AI Matching** — Optionally uses AI to match transactions with different naming conventions (e.g. `NEFT-KAPILA-PHARMA` → `Kapila Pharma`) across three tiers:
  - 🟡 **Yellow** — exact name + amount match
  - 🔵 **Blue** — AI-matched (name similarity within ₹1 tolerance)
  - ⬜ **White** — fuzzy fallback (Jaccard word similarity)
- **Real-time Logs** — Live terminal in the UI streams AI matching progress via Server-Sent Events (SSE)
- **Dark UI** — Clean, responsive dark-theme interface

---

## Prerequisites

- **Python 3.9+** — [Download](https://www.python.org/downloads/) (check *Add to PATH* during install)
- **OpenRouter API key** — [Get one free](https://openrouter.ai/keys) *(only needed if you want AI matching)*

---

## Setup (Windows)

### Option A — Batch scripts (easiest)

```
1. Double-click setup.bat    ← installs everything
2. Edit flask-backend\.env   ← paste your OpenRouter API key
3. Double-click start.bat    ← launches the app and opens browser
```

### Option B — Manual

```bash
# 1. Clone the repo
git clone https://github.com/subbu-h21/account-reconciliation-tool.git
cd account-reconciliation-tool

# 2. Create and activate a virtual environment
cd flask-backend
python -m venv venv
venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set up environment variables
copy .env.example .env
# Open .env and paste your OpenRouter API key

# 5. Run
python app.py
```

Open your browser at **http://localhost:5000**

---

## Environment Variables

Edit `flask-backend/.env`:

```env
OPEN_ROUTER_API_KEY=your_openrouter_key_here
```

The app runs fine without this key — AI matching will simply be unavailable.

---

## Usage

1. **Upload** your bank account statement (`.xlsx` / `.xls`)
2. **Upload** your internal books file (`.xlsx` / `.xls`)
3. *(Optional)* Toggle **AI Matching**
4. Click **Run Reconciliation**
5. Watch the live terminal for AI matching progress (AI mode only)
6. Click **Download** to get the reconciliation report

### Output Excel Structure

| Sheet | Contents |
|---|---|
| `Receivements` | Side-by-side inbound transactions (Books Debit vs Bank Received) |
| `Payments` | Side-by-side outbound transactions (Books Credit vs Bank Given) |
| `Summary` | Per-day metrics: entry counts, totals, closing balances, and differences |

---

## Project Structure

```
account-reconciliation-tool/
├── setup.bat                   # First-time setup script
├── start.bat                   # Launch the app
└── flask-backend/
    ├── app.py                  # Flask routes + SSE streaming
    ├── tasks.py                # Background task orchestrator
    ├── requirements.txt
    ├── .env.example            # API key template
    ├── templates/
    │   └── index.html          # Frontend UI (vanilla JS)
    ├── services/
    │   ├── reconciliation.py   # Core reconciliation logic
    │   └── matcher.py          # 3-tier AI matching engine
    └── uploads/                # Temporary output files (git-ignored)
```

---

## License

MIT License — free to use and modify.
