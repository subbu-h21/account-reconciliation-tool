"""
Quick test to send a matching prompt directly to Ollama models.

Usage:
    python test_matcher.py

Edit the SAMPLE_DATA dict below with real names/amounts from your Excel.
"""
import json
import time
import requests
from services.matcher import OLLAMA_MODELS

OLLAMA_URL = "http://localhost:11434"

SAMPLE_DATA = {
    # Easy: near-identical names, exact amounts
    "2025-08-01": {
        "bank": [
            {"name": "NEFT-KAPILA-PHARMA",      "amount": 15000.0},
            {"name": "MClick/To-SHARMA-TRADERS", "amount":  8500.0},
            {"name": "NEFT-SUNRISE-AGENCIES",    "amount": 42000.0},
        ],
        "books": [
            {"name": "Kapila Pharma",    "amount": 15000.0},
            {"name": "Sharma Traders",   "amount":  8500.0},
            {"name": "Sunrise Agencies", "amount": 42000.0},
        ]
    },
    # Tricky: abbreviations, extra words, same amount
    "2025-08-02": {
        "bank": [
            {"name": "NEFT-ABC-SUPPLIES",       "amount": 22000.0},
            {"name": "NEFT-MEHTA-MEDICO-STORE", "amount": 11250.0},
            {"name": "MClick/To-RAGHAV-ENTP",   "amount":  5000.0},
        ],
        "books": [
            {"name": "ABC Supplies Pvt Ltd", "amount": 22000.0},
            {"name": "Mehta Medicos",        "amount": 11250.0},
            {"name": "Raghav Enterprises",   "amount":  5000.0},
        ]
    },
    # Tricky: amount mismatch on one pair — should NOT match that one
    "2025-08-03": {
        "bank": [
            {"name": "NEFT-KRISHNA-DIST",   "amount": 18000.0},
            {"name": "NEFT-PATEL-BROTHERS", "amount":  9300.0},
        ],
        "books": [
            {"name": "Krishna Distributors", "amount": 18000.0},
            {"name": "Patel Bros",           "amount":  7800.0},
        ]
    },
    # Hard: heavily abbreviated bank names
    "2025-08-04": {
        "bank": [
            {"name": "NEFT-OM-SAI-ENT",       "amount": 33500.0},
            {"name": "MClick/To-JSK-TRD-PVT", "amount": 14000.0},
            {"name": "NEFT-SHREE-RAM-PHARMA",  "amount":  6750.0},
        ],
        "books": [
            {"name": "Om Sai Enterprises",          "amount": 33500.0},
            {"name": "JSK Trading Pvt Ltd",         "amount": 14000.0},
            {"name": "Shree Ram Pharmaceutical Co", "amount":  6750.0},
        ]
    },
    # Swapped: names match but amounts are crossed — should NOT match
    "2025-08-05": {
        "bank": [
            {"name": "NEFT-VERMA-CHEMICALS", "amount": 25000.0},
            {"name": "NEFT-GUPTA-FOODS",     "amount": 12000.0},
        ],
        "books": [
            {"name": "Gupta Foods",     "amount": 25000.0},
            {"name": "Verma Chemicals", "amount": 12000.0},
        ]
    },
    # Hard: 10 entries, mixed difficulty
    # - 6 clean matches
    # - 2 tricky (heavy abbreviation + pvt/ltd suffix noise)
    # - 1 amount mismatch (similar name, wrong amount — should NOT match)
    # - 1 completely unrelated (different supplier entirely — should NOT match)
    "2025-08-06": {
        "bank": [
            {"name": "NEFT-KAPILA-PHARMA",        "amount": 15000.0},  # 0 — easy match
            {"name": "NEFT-KRISHNA-DIST",          "amount": 18000.0},  # 1 — easy match
            {"name": "MClick/To-JSK-TRD-PVT",     "amount": 14000.0},  # 2 — tricky abbreviation
            {"name": "NEFT-SHREE-RAM-PHARMA",      "amount":  6750.0},  # 3 — tricky abbreviation
            {"name": "NEFT-SUNRISE-AGENCIES",      "amount": 42000.0},  # 4 — easy match
            {"name": "NEFT-MEHTA-MEDICO-STORE",    "amount": 11250.0},  # 5 — easy match
            {"name": "MClick/To-RAGHAV-ENTP",      "amount":  5000.0},  # 6 — easy match
            {"name": "NEFT-OM-SAI-ENT",            "amount": 33500.0},  # 7 — easy match
            {"name": "NEFT-PATEL-BROTHERS",        "amount":  9300.0},  # 8 — amount mismatch, should NOT match
            {"name": "NEFT-SHARMA-ELECTRICALS",    "amount": 21000.0},  # 9 — unrelated to books entry, should NOT match
        ],
        "books": [
            {"name": "Kapila Pharma",              "amount": 15000.0},  # 0
            {"name": "Krishna Distributors",       "amount": 18000.0},  # 1
            {"name": "JSK Trading Pvt Ltd",        "amount": 14000.0},  # 2
            {"name": "Shree Ram Pharmaceutical Co","amount":  6750.0},  # 3
            {"name": "Sunrise Agencies",           "amount": 42000.0},  # 4
            {"name": "Mehta Medicos",              "amount": 11250.0},  # 5
            {"name": "Raghav Enterprises",         "amount":  5000.0},  # 6
            {"name": "Om Sai Enterprises",         "amount": 33500.0},  # 7
            {"name": "Patel Bros",                 "amount":  7800.0},  # 8 — amount mismatch
            {"name": "National Traders",           "amount": 21000.0},  # 9 — unrelated name
        ]
    },
}

PROMPT_TEMPLATE = (
    "You are matching transactions between a bank statement and internal accounting books. "
    "The same entity is often written differently on each side "
    "(e.g. 'NEFT-KAPILA-PHARMA' on the bank side matches 'Kapila Pharma' on the books side).\n\n"
    "Rules:\n"
    "  1. Names must look like the same supplier or entity.\n"
    "  2. Amounts must be equal or very close.\n"
    "  Only match if BOTH agree.\n\n"
    "Input JSON:\n{data}\n\n"
    "Go through each object in the bank array and each object in the books array. "
    "For each object, decide if it has a match on the other side. "
    "Replace each object with 1 if it matched, 0 if it did not. "
    "Return the same JSON structure with objects replaced by 0s and 1s. "
    "No markdown, no explanation.\n"
    'Example: {{"date": "2025-08-01", "bank": [1, 0, 1], "books": [1, 0, 1]}}'
)

RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {
        "date":  {"type": "string"},
        "bank":  {"type": "array", "items": {"type": "integer", "enum": [0, 1]}},
        "books": {"type": "array", "items": {"type": "integer", "enum": [0, 1]}},
    },
    "required": ["date", "bank", "books"],
}

EXPECTED = {
    "2025-08-01": {"bank": [1, 1, 1], "books": [1, 1, 1]},
    "2025-08-02": {"bank": [1, 1, 1], "books": [1, 1, 1]},
    "2025-08-03": {"bank": [1, 0],    "books": [1, 0]},
    "2025-08-04": {"bank": [1, 1, 1], "books": [1, 1, 1]},
    "2025-08-05": {"bank": [0, 0],                            "books": [0, 0]},
    "2025-08-06": {"bank": [1, 1, 1, 1, 1, 1, 1, 1, 0, 0],  "books": [1, 1, 1, 1, 1, 1, 1, 1, 0, 0]},
}


def _build_llm_input(date_str, bank_entries, books_entries):
    return {
        "date":  date_str,
        "bank":  [{"id": i, "name": e["name"], "amount": e["amount"]} for i, e in enumerate(bank_entries)],
        "books": [{"id": i, "name": e["name"], "amount": e["amount"]} for i, e in enumerate(books_entries)],
    }


def _validate_and_extract(llm_input, llm_output, bank_entries, books_entries):
    bank_flags  = llm_output["bank"]
    books_flags = llm_output["books"]

    if len(bank_flags) != len(llm_input["bank"]):
        raise ValueError(f"bank length mismatch: expected {len(llm_input['bank'])}, got {len(bank_flags)}")
    if len(books_flags) != len(llm_input["books"]):
        raise ValueError(f"books length mismatch: expected {len(llm_input['books'])}, got {len(books_flags)}")
    if not all(v in (0, 1) for v in bank_flags + books_flags):
        raise ValueError("response contains values other than 0 or 1")

    matched_bank  = [bank_entries[i]["name"]  for i, v in enumerate(bank_flags)  if v == 1]
    matched_books = [books_entries[i]["name"] for i, v in enumerate(books_flags) if v == 1]
    return matched_bank, matched_books


def _check_expected(date_str, llm_output):
    exp = EXPECTED.get(date_str)
    if not exp:
        return
    bank_ok  = llm_output["bank"]  == exp["bank"]
    books_ok = llm_output["books"] == exp["books"]
    status   = "PASS" if bank_ok and books_ok else "FAIL"
    print(f"  Expected : bank={exp['bank']}  books={exp['books']}")
    print(f"  Got      : bank={llm_output['bank']}  books={llm_output['books']}")
    print(f"  Result   : {status}")


def call_model(model_key, model_name):
    model_total_matches = 0
    model_pass          = 0

    print(f"\n{'='*55}")
    print(f"  Model : {model_key} ({model_name})")
    print(f"{'='*55}")

    for date_str, sides in SAMPLE_DATA.items():
        llm_input = _build_llm_input(date_str, sides["bank"], sides["books"])
        prompt    = PROMPT_TEMPLATE.replace("{data}", json.dumps(llm_input))

        payload = {
            "model":    model_name,
            "messages": [
                {
                    "role":    "system",
                    "content": "You are a JSON-only assistant. Output only a raw JSON object — no markdown, no explanation.",
                },
                {"role": "user", "content": prompt},
            ],
            "stream": False,
            "format": RESPONSE_SCHEMA,
        }

        print(f"\n  ── Date: {date_str} ──")

        try:
            t0      = time.perf_counter()
            resp    = requests.post(f"{OLLAMA_URL}/api/chat", json=payload, timeout=600)
            elapsed = time.perf_counter() - t0

            if resp.status_code != 200:
                print(f"  ERROR: HTTP {resp.status_code} — {resp.text}")
                continue

            raw = resp.json()["message"]["content"]
            print(f"  Time : {elapsed:.2f}s")
            print(f"  Raw  : {raw}")

            try:
                llm_output = json.loads(raw.strip())
                matched_bank, matched_books = _validate_and_extract(
                    llm_input, llm_output, sides["bank"], sides["books"]
                )
                print(f"  Matched bank  : {matched_bank}")
                print(f"  Matched books : {matched_books}")
                _check_expected(date_str, llm_output)
                model_total_matches += len(matched_bank)
                if llm_output["bank"] == EXPECTED[date_str]["bank"]:
                    model_pass += 1

            except (json.JSONDecodeError, ValueError, KeyError) as e:
                print(f"  PARSE/VALIDATION FAILED: {e}")

        except requests.exceptions.ConnectionError:
            print("  CONNECTION ERROR — is Ollama running?")
        except requests.exceptions.Timeout:
            print("  TIMEOUT — model took too long (>600s)")
        except Exception as e:
            print(f"  UNEXPECTED ERROR: {e}")

    print(f"\n  ── Summary for {model_key} ──")
    print(f"  Dates passed : {model_pass}/{len(SAMPLE_DATA)}")
    print(f"  Total matched entries : {model_total_matches}")


def check_ollama():
    try:
        r = requests.get(OLLAMA_URL, timeout=5)
        if "Ollama is running" in r.text:
            print("Ollama: running")
            return True
    except Exception:
        pass
    print("Ollama: NOT RUNNING — start it with `ollama serve` or open the app")
    return False


def list_local_models():
    try:
        r      = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        models = [m["name"] for m in r.json().get("models", [])]
        print(f"Pulled models: {models if models else 'none'}")
        return models
    except Exception:
        return []


if __name__ == "__main__":
    print("\n── Ollama status ──────────────────────────────────────")
    if not check_ollama():
        raise SystemExit(1)

    pulled = list_local_models()

    for key, name in OLLAMA_MODELS.items():
        if not any(name in p for p in pulled):
            print(f"\n  SKIP {key} — '{name}' not pulled yet. Run: ollama pull {name}")
            continue
        call_model(key, name)

    print(f"\n{'='*55}")
    print("  Done")
    print(f"{'='*55}\n")
