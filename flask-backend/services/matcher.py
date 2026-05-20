import json
import logging
from datetime import datetime, date as date_type
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

logger = logging.getLogger(__name__)

YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
ORANGE = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

SHEET_COLS = {
    "Receivements": {"books_col": 2, "books_amt_col": 3, "bank_col": 5, "bank_amt_col": 6},
    "Payments":     {"books_col": 2, "books_amt_col": 3, "bank_col": 5, "bank_amt_col": 6},
}

BATCH_SIZE = 10

GEMINI_PROMPT_TEMPLATE = (
    "You are matching transactions between a bank statement and internal accounting books. "
    "The same entity is often written differently on each side "
    "(e.g. 'NEFT-KAPILA-PHARMA (5000.0)' on the bank side matches 'Kapila Pharma (5000.0)' on the books side).\n\n"
    "Rules:\n"
    "  1. Names must refer to the same supplier or entity.\n"
    "  2. Amounts (shown in brackets) must be exactly equal. Different amounts = no match.\n"
    "  Both name AND amount must agree.\n\n"
    "Input: a JSON object where each key is a date and the value has 'bank' and 'books' arrays of strings.\n"
    "Each string is 'EntityName (amount)'.\n\n"
    "For each date, go through each entry in bank[] and books[]. "
    "Output 1 if the entry has a matching counterpart on the other side, 0 if it does not. "
    "The two arrays can have different lengths — match independently.\n\n"
    "Return a JSON object with the same date keys. For each date, return 'bank' and 'books' arrays "
    "of integers (0 or 1), one per input entry, in the same order.\n\n"
    "Input:\n{data}\n\n"
    "Example input:  {\"2025-08-01\": {\"bank\": [\"NEFT-KAPILA (5000.0)\", \"UBI-SHARMA (200.0)\"], \"books\": [\"Kapila Pharma (5000.0)\"]}}\n"
    "Example output: {\"2025-08-01\": {\"bank\": [1, 0], \"books\": [1]}}"
)

OPENAI_SYSTEM_PROMPT = (
    "You are matching transactions between a bank statement and internal accounting books. "
    "The same entity is often written differently on each side "
    "(e.g. 'NEFT-KAPILA-PHARMA (5000.0)' on the bank side matches 'Kapila Pharma (5000.0)' on the books side).\n\n"
    "Rules:\n"
    "  1. Names must refer to the same supplier or entity.\n"
    "  2. Amounts (shown in brackets) must be exactly equal. Different amounts = no match.\n"
    "  Both name AND amount must agree for a match.\n\n"
    "Input: a JSON object where each key is a date and the value has 'bank' and 'books' arrays of strings. "
    "Each string is 'EntityName (amount)'.\n\n"
    "For each date, identify entries that have NO match on the other side. "
    "Return a JSON object with the same date keys. For each date, return:\n"
    "  - 'bank': list of strings from the input 'bank' array that have NO match in 'books'\n"
    "  - 'books': list of strings from the input 'books' array that have NO match in 'bank'\n\n"
    "Return only strings that appeared in the original input, word for word. "
    "If all entries on a side matched, return an empty array [] for that side.\n\n"
    "Example input:  {\"2025-08-01\": {\"bank\": [\"NEFT-KAPILA (5000.0)\", \"UBI-SHARMA (200.0)\"], \"books\": [\"Kapila Pharma (5000.0)\"]}}\n"
    "Example output: {\"2025-08-01\": {\"bank\": [\"UBI-SHARMA (200.0)\"], \"books\": []}}"
)


def _date_str(val):
    if isinstance(val, datetime):
        return val.date().isoformat()
    if isinstance(val, date_type):
        return val.isoformat()
    return str(val)


def _fmt(e):
    amt = e["amount"]
    return f"{e['name']} ({amt})" if amt is not None else e["name"]


def _build_groups(ws, books_col, books_amt_col, bank_col, bank_amt_col):
    groups = {}
    for row in ws.iter_rows(min_row=2):
        date_val = row[0].value
        if date_val is None:
            continue
        ds = _date_str(date_val)
        if ds not in groups:
            groups[ds] = {"bank": [], "books": []}
        row_num    = row[0].row
        books_name = row[books_col     - 1].value
        books_amt  = row[books_amt_col - 1].value
        bank_name  = row[bank_col      - 1].value
        bank_amt   = row[bank_amt_col  - 1].value
        if books_name is not None:
            groups[ds]["books"].append({
                "name":   str(books_name),
                "amount": float(books_amt) if books_amt is not None else None,
                "row":    row_num,
            })
        if bank_name is not None:
            groups[ds]["bank"].append({
                "name":   str(bank_name),
                "amount": float(bank_amt) if bank_amt is not None else None,
                "row":    row_num,
            })
    return groups


def _build_llm_batch(batch):
    return {
        ds: {
            "bank":  [_fmt(e) for e in sides["bank"]],
            "books": [_fmt(e) for e in sides["books"]],
        }
        for ds, sides in batch
    }


def _build_schema(batch):
    date_schema = {
        "type": "object",
        "properties": {
            "bank":  {"type": "array", "items": {"type": "integer"}},
            "books": {"type": "array", "items": {"type": "integer"}},
        },
        "required": ["bank", "books"],
    }
    return {
        "type": "object",
        "properties": {ds: date_schema for ds, _ in batch},
        "required":   [ds for ds, _ in batch],
    }


def _call_gemini(batch_input, batch, client):
    from google.genai import types

    prompt   = GEMINI_PROMPT_TEMPLATE.replace("{data}", json.dumps(batch_input))
    schema   = _build_schema(batch)
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=schema,
        ),
    )
    logger.debug("Gemini raw response: %s", response.text)
    return json.loads(response.text)


def _call_openai(batch_input, client):
    response = client.chat.completions.create(
        model="openai/gpt-5-nano",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": OPENAI_SYSTEM_PROMPT},
            {"role": "user",   "content": json.dumps(batch_input)},
        ],
    )
    logger.debug("OpenAI raw response: %s", response.choices[0].message.content)
    return json.loads(response.choices[0].message.content)


def _validate_one(date_str, llm_out, sides):
    """Gemini: returns (matched_bank, matched_books) from 0/1 flag arrays."""
    bank_flags  = llm_out["bank"]
    books_flags = llm_out["books"]
    if len(bank_flags) != len(sides["bank"]):
        raise ValueError(f"bank length mismatch: expected {len(sides['bank'])}, got {len(bank_flags)}")
    if len(books_flags) != len(sides["books"]):
        raise ValueError(f"books length mismatch: expected {len(sides['books'])}, got {len(books_flags)}")
    if not all(v in (0, 1) for v in bank_flags + books_flags):
        raise ValueError("values other than 0/1 in response")
    matched_bank  = [sides["bank"][i]  for i, v in enumerate(bank_flags)  if v == 1]
    matched_books = [sides["books"][i] for i, v in enumerate(books_flags) if v == 1]
    return matched_bank, matched_books


def _normalize(s):
    """Collapse whitespace so Excel trailing/double spaces don't break lookups."""
    return " ".join(s.strip().split())


def _validate_unmatched(date_str, llm_out, sides):
    """OpenAI: returns (unmatched_bank, unmatched_books) from name lists."""
    bank_names  = llm_out.get("bank", [])
    books_names = llm_out.get("books", [])

    bank_lookup  = {_normalize(_fmt(e)): e for e in sides["bank"]}
    books_lookup = {_normalize(_fmt(e)): e for e in sides["books"]}

    unmatched_bank, unmatched_books = [], []

    for name in bank_names:
        entry = bank_lookup.get(_normalize(name))
        if entry is None:
            raise ValueError(f"bank name not in input: {name!r}")
        unmatched_bank.append(entry)

    for name in books_names:
        entry = books_lookup.get(_normalize(name))
        if entry is None:
            raise ValueError(f"books name not in input: {name!r}")
        unmatched_books.append(entry)

    return unmatched_bank, unmatched_books


def apply_highlights(output_path, api_key, provider="gemini", log=None):
    def emit(msg):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        logger.debug(line)
        if log:
            log(line)

    emit(f"AI matching started — provider: {provider}")

    if provider == "gemini":
        from google import genai
        client = genai.Client(api_key=api_key)
        fill_color = YELLOW
    else:
        from openai import OpenAI
        client = OpenAI(api_key=api_key, base_url="https://openrouter.ai/api/v1")
        fill_color = ORANGE

    wb = load_workbook(output_path)
    total_highlighted = 0

    for sheet_name, cols in SHEET_COLS.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws       = wb[sheet_name]
        groups   = _build_groups(ws, cols["books_col"], cols["books_amt_col"],
                                 cols["bank_col"], cols["bank_amt_col"])
        eligible = [(ds, sides) for ds, sides in groups.items() if sides["bank"] and sides["books"]]
        emit(f"Sheet: {sheet_name} — {len(eligible)} dates to process")

        for batch_start in range(0, len(eligible), BATCH_SIZE):
            batch      = eligible[batch_start:batch_start + BATCH_SIZE]
            batch_num  = batch_start // BATCH_SIZE + 1
            date_range = f"{batch[0][0]} … {batch[-1][0]}"
            emit(f"Batch {batch_num}: {len(batch)} dates ({date_range}) → sending to {provider}...")

            batch_input = _build_llm_batch(batch)
            try:
                if provider == "gemini":
                    batch_output = _call_gemini(batch_input, batch, client)
                else:
                    batch_output = _call_openai(batch_input, client)
            except Exception as e:
                logger.warning("Batch %d %s call failed: %s", batch_num, provider, e)
                emit(f"Batch {batch_num} ERROR: {e} — skipping {len(batch)} dates")
                continue

            for date_str, sides in batch:
                if date_str not in batch_output:
                    emit(f"  {date_str} → missing from response — skipping")
                    continue
                try:
                    if provider == "gemini":
                        to_highlight_bank, to_highlight_books = _validate_one(date_str, batch_output[date_str], sides)
                    else:
                        to_highlight_bank, to_highlight_books = _validate_unmatched(date_str, batch_output[date_str], sides)
                except Exception as e:
                    emit(f"  {date_str} → validation ERROR: {e} — skipping")
                    continue

                bank_names  = [e["name"] for e in to_highlight_bank]
                books_names = [e["name"] for e in to_highlight_books]
                emit(f"  {date_str} → bank: {bank_names}  books: {books_names}")

                for entry in to_highlight_bank:
                    ws.cell(row=entry["row"], column=cols["bank_col"]).fill = fill_color
                    total_highlighted += 1
                for entry in to_highlight_books:
                    ws.cell(row=entry["row"], column=cols["books_col"]).fill = fill_color
                    total_highlighted += 1

    emit(f"Done. {total_highlighted} entries highlighted.")
    wb.save(output_path)
