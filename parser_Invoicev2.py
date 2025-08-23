# parsers/invoice.py
"""
Parser brick for legacy “Invoice …” PDFs
========================================
Now also extracts *case_descr* – the text that sits in the first row of the
“Description / Amount” table (e.g. «repellents» in the sample invoice).

---------------------------------------------------------------------------
parse_invoice(path: str | pathlib.Path) -> dict
---------------------------------------------------------------------------

Returns
-------
dict with **four** keys:

    {
      "invoice_number": "<str>",
      "date":           "YYYY-MM-DD",
      "amount":         <float>,
      "case_descr":     "<str>"      # new, mandatory
    }

If any of the four fields are missing → `ValueError`.
"""

from __future__ import annotations

import re
from datetime import datetime as _dt
from pathlib import Path
from typing import Final

import pdfplumber  # external dep

# ────────── regexes for old-layout invoice ──────────
_RE_DATE: Final = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")
_RE_INVNO: Final = re.compile(r"Invoice\s*No\.?\s*0*(\d{3,})", re.I)
_RE_INVNO_FBK: Final = re.compile(r"\b000\d{5}\b")
_RE_TOTALUSD: Final = re.compile(
    r"Total\s+amount:?\s*USD\s+([\d\s,]+(?:\.\d{2})?)", re.I
)
_RE_ANYUSD: Final = re.compile(r"USD\s+([\d\s,]+(?:\.\d{2})?)", re.I)
_RE_NON_DIGIT_DOT: Final = re.compile(r"[^\d\.]")

# NEW: capture the first non-empty line between the header row
# “Description  Amount” and the first “USD …” money cell.
# Works for the sample where we have:
#   Description Amount
#
#   repellents
#   USD 4000
_RE_CASEDESCR: Final = re.compile(
    r"Description\s+Amount\s+(.+?)\s+USD\s*\d", re.I | re.S
)


def _find_case_descr(text: str) -> str | None:
    """Return description line, stripped, or None if not found."""
    if (m := _RE_CASEDESCR.search(text)):
        # Take the inner capture, split by newlines, pick first non-blank line.
        raw = m.group(1)
        for line in (l.strip() for l in raw.splitlines()):
            if line:
                return line
    return None


# ────────── public API ──────────
def parse_invoice(path: str | Path) -> dict:
    """
    Extract *invoice_number*, *date*, *amount* and *case_descr* from legacy
    invoice PDFs.

    Algorithm
    ---------
    1. Check file exists.
    2. Join text of all pages via pdfplumber.
    3. Parse date (first MM/DD/YYYY).
    4. Parse invoice number (strict / fallback).
    5. Parse amount (prefer “Total amount: USD …”, else first “USD …”).
    6. Parse case description (first line inside Description/Amount table).
    7. Validate all four fields.
    8. Return dict.
    """
    pdf_path = Path(path)
    if not pdf_path.is_file():
        raise FileNotFoundError(pdf_path)

    with pdfplumber.open(pdf_path) as pdf:
        full_text: str = "\n".join(p.extract_text() or "" for p in pdf.pages)

    # 3️⃣ date
    date_iso = None
    if (m_date := _RE_DATE.search(full_text)):
        mm, dd, yyyy = map(int, m_date.groups())
        try:
            date_iso = _dt(yyyy, mm, dd).date().isoformat()
        except ValueError:
            pass

    # 4️⃣ invoice number
    invoice_number = None
    if (m_inv := _RE_INVNO.search(full_text)):
        invoice_number = m_inv.group(1)
    elif (m_fb := _RE_INVNO_FBK.search(full_text)):
        invoice_number = m_fb.group(0).lstrip("0")

    # 5️⃣ amount
    amount = None
    m_amt = _RE_TOTALUSD.search(full_text) or _RE_ANYUSD.search(full_text)
    if m_amt:
        raw = _RE_NON_DIGIT_DOT.sub("", m_amt.group(1))
        try:
            amount = float(raw)
        except ValueError:
            pass

    # 6️⃣ description
    case_descr = _find_case_descr(full_text)

    # 7️⃣ validation
    missing = [
        k
        for k, v in {
            "invoice_number": invoice_number,
            "date": date_iso,
            "amount": amount,
            "case_descr": case_descr,
        }.items()
        if v in (None, "", [])
    ]
    if missing:
        raise ValueError(f"parse_invoice: missing {', '.join(missing)}")

    # 8️⃣ return
    return {
        "invoice_number": invoice_number,  # type: ignore[arg-type]
        "date": date_iso,                  # type: ignore[arg-type]
        "amount": amount,                  # type: ignore[arg-type]
        "case_descr": case_descr,          # type: ignore[arg-type]
    }
