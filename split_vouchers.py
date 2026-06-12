#!/usr/bin/env python3
"""
Split scanned voucher PDFs (Q<q><yy>_5_Vouchers_No_<start>-<end>.pdf) into
per-voucher sub-files (Q<q><yy>_6_Voucher_<num>_<type>.pdf), then compress
each output to ebook quality. The Q<q><yy> quarter prefix (e.g. Q126,
Q227, Q325) is read from each input filename and reused on its outputs.

Suffix convention:
    _1  main voucher sheet          (Standard / Collective / Travel /
                                     Reimbursement-private-km voucher form)
    _2  receipts, invoices, slips   (third-party documents)
    _3  supporting documents        (other internal forms attached to the voucher)

Cross-platform: works on macOS, Linux, and Windows. Requires:
    * Python 3.9+ with the `pymupdf` package (`pip install pymupdf`)
    * Tesseract OCR        (Mac: `brew install tesseract`,
                            Windows: https://github.com/UB-Mannheim/tesseract/wiki)
    * Ghostscript          (Mac: `brew install ghostscript`,
                            Windows: https://ghostscript.com/releases/gsdnld.html)

The default input/output folder is "<your Desktop>/automation". Drop the
Q<q><yy>_5_*.pdf files in there, run the script, and the matching
Q<q><yy>_6_*.pdf files appear next to them.

Classification logic (most -> least specific)
=============================================

1. Driver's log (Fahrtenbuch) is identified by perceptual-hash similarity
   to a known template OR by a text-based fingerprint (matches German
   "Fahrtenbuch" / English "driver's log" / column-header pair "Km Start"
   ... "Km Ende", etc.) -- either signal alone is sufficient, so a
   handwritten / rotated / poorly-scanned page that drifts from the
   template still gets caught. Each Fahrtenbuch page is attached to the
   NEAREST PRECEDING D12 ("Reimbursement private km") main voucher; this
   correctly distributes log pages across multiple D12 vouchers in the
   same bundle. Pages that came before any D12 are deferred and attached
   to the first D12 found, or dropped if none exists.

2. The page footer is OCR'd; a recognised D-code drives the category:

      D02   project funds report           -> SKIPPED entirely (per user)
      D07   universal standard voucher     -> MAIN
      D10   travel voucher (page 1)        -> MAIN
      D10   travel voucher (continuation)  -> MAIN-continuation (folded
                                              into the same _1 file)
      D12   reimbursement private km       -> MAIN
      D14   payment voucher private phone  -> SUPPORTING
      D15   collective voucher             -> MAIN
      D17   procurement voucher            -> SUPPORTING
      D21   reimbursement private (EUR)    -> SUPPORTING
      D22   per diem overview              -> SUPPORTING

3. Pages without a recognised D-code are checked against a supporting-
   document keyword fingerprint (participant lists, workshop agendas /
   itineraries, session plans, multi-lingual). A conservative match
   flips the page from the RECEIPT default to SUPPORTING. Contract /
   formal-agreement pages (e.g. a multi-page "Consultancy Contract"
   whose payment clauses mention invoices, totals and reimbursable
   expenses) are recognised by title / legal-boilerplate phrases and
   routed to SUPPORTING -- unless the page carries an actual receipt
   heading, in which case it is a receipt that merely references the
   contract (see looks_like_contract). Pages that
   look like photographs are also flipped to SUPPORTING -- BUT ONLY if
   the OCR text does not contain strong receipt keywords. Photograph
   detection RENDERS the page (deterministic across OS) and measures the
   largest contiguous region of continuous-tone pixels -- big for a
   photo or scanned ID card, near-zero for a text document -- so a small
   photo on a mostly-white page is still caught (see photo_region_score).
   The receipt keyword gate (positive matching against multi-lingual
   receipt vocabulary, including telecom prepaid cards like "telemor",
   "Konsulta saldu", "scratch card") protects colourful printed receipts
   from being mis-flipped.

4. Multi-page main forms (notably D08) print "1/N" / "2/N" markers in
   the footer. Any non-main page with marker (k>=2, N) is attached to
   the most recent main voucher -- but ONLY if that main has its own
   matching (1, N) marker AND no intervening page is advertising a
   competing (1, N) of its own.
   D08 specifically is ALWAYS a 2-page form by template. After the
   marker pass, every D08 main has its continuation force-attached:
   we prefer marker-based detection (handles scrambled bundles), and
   fall back to "the immediately next non-main page is the continuation"
   for canonical-order bundles where the marker OCR also failed.

5. Duplicate consecutive main pages (same D-code, near-identical
   perceptual hash) -- e.g. when a user accidentally scans a voucher
   form twice -- are merged automatically: the second copy is demoted
   to main_cont and folded into the same _1 file as the first.

Voucher numbering
=================
For each MAIN page we OCR the top-right "Voucher N°" cell with FOUR
Tesseract layout modes (PSM 6 prose, PSM 7 single line + digit
whitelist, PSM 8 single word + digit whitelist, PSM 11 sparse text +
digit whitelist) and collect every plausible candidate. Multi-PSM
agreement gives strong handwriting-recognition robustness.

Candidates are validated against the bundle's expected range (parsed
from the filename, e.g. 561-576). An out-of-range read like "54" when
the range is 561-576 is treated as an OCR error and discarded -- the
sequential fallback (anchored to range_start) takes over.

If the detected MAIN count differs from the expected count, the script
still produces output but emits a warning so the result can be
sanity-checked.
"""

from __future__ import annotations

import argparse
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional


def _missing_dep(name: str) -> None:
    sys.stderr.write(
        f"\nERROR: the '{name}' package is not installed for this Python interpreter.\n"
        f"  Interpreter in use: {sys.executable}\n\n"
        "Install the required Python packages into THIS interpreter with:\n"
        f'  "{sys.executable}" -m pip install pymupdf Pillow\n'
        "  (add --break-system-packages on macOS if pip refuses)\n\n"
        "Note: Tesseract OCR and Ghostscript are separate SYSTEM binaries,\n"
        "not pip packages -- install them per the README/Tutorial if missing.\n\n"
        "If you are running this from VS Code, the interpreter shown in the\n"
        "bottom-right status bar is the one being used -- either install into\n"
        "it (command above), or switch to one that already has the packages.\n"
    )
    sys.exit(1)


try:
    import fitz  # pymupdf
except ImportError:
    _missing_dep("pymupdf")

try:
    from PIL import Image  # type: ignore
except ImportError:
    _missing_dep("Pillow")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DEFAULT_FOLDER = Path.home() / "Desktop" / "automation"
# Match any quarter/year prefix: Q126, Q227, Q325, ...
INPUT_GLOB = "Q*_5_Vouchers_No_*.pdf"

OCR_DPI = 300                # full-page OCR
VOUCHER_NO_DPI = 400         # top-right crop OCR for voucher number

# ---------------------------------------------------------------------------
# Form-code patterns (matched against OCR'd footer / body text)
# ---------------------------------------------------------------------------

# Per form, two regexes are stored:
#   footer_pat -- the D-code printed in the page footer, e.g. "(D 09)".
#                 Highest-confidence signal; matched FIRST across all forms.
#   body_pat   -- a body-text fingerprint (form title, section heading, etc.).
#                 Matched only if no footer code was recognised on the page.
#
# Categories:
#   "main"          top-level voucher form, starts a new voucher group
#   "main_or_cont"  D10 only -- decided by D10_PAGE1_PAT below
#   "supporting"    supplementary form attached to most recent voucher (_3)
#   "fahrtenbuch"   driver's log -- attached to D12 voucher if present
#   "skip"          drop entirely
#
# Per-user rules baked in:
#   - Extension lists (D09) are ALWAYS supporting (TF2.2 / TF2.3).
#   - All TF5 forms (D17 procurement, D18 inventory, D19 segregation,
#     D20 handover) are ALWAYS supporting.
# OCR-confusable digit classes used inside footer patterns:
#   the digit "0" is regularly read as "O" or "Q" by Tesseract on small
#   footer text, "1" as "l" or "I", and "2" as "Z". The classes below are
#   case-insensitive (the patterns themselves carry re.I).
#
# BOTH digits of every two-digit code are REQUIRED (no optional tens digit).
# The printed footer always shows both (e.g. "(D 07)", "(D 12)"), so a
# pattern that allowed the tens digit to be dropped caused cross-code
# collisions on degraded OCR: "(D2)" matched D02 ("skip") before D12/D22
# and silently dropped the page; "(D7)"/"(D0)"/"(D1)" promoted supporting
# forms (D17/D20/D21) into D07/D10/D11 mains. With both digits required, a
# footer degraded enough to lose a digit simply falls through to the
# body-text fingerprint, and worst case defaults to "receipt" (kept,
# attached to the current voucher) rather than being dropped or misrouted.
_D0 = r"[0oq]"   # zero
_D1 = r"[1li]"   # one
_D2 = r"[2z]"    # two

FORM_PATTERNS: dict[str, tuple[re.Pattern, re.Pattern, str]] = {
    # --- Skipped --------------------------------------------------------
    "D02": (
        re.compile(rf"\(\s*d\s*{_D0}{_D2}\s*\)", re.I),
        re.compile(r"project\s+funds\s+report|forecast", re.I),
        "skip",
    ),
    # --- Driver's log (TF3) --------------------------------------------
    # Primary detection is by perceptual hash; this is a text-based fallback.
    "D04": (
        re.compile(rf"\(\s*d\s*{_D0}4\s*\)", re.I),
        re.compile(r"fahrtenbuch|driver.?s\s+log", re.I),
        "fahrtenbuch",
    ),
    # --- TF4 Exception forms -------------------------------------------
    "D06": (
        re.compile(rf"\(\s*d\s*{_D0}6\s*\)", re.I),
        re.compile(r"incident\s+report[\s\-]?voucher", re.I),
        "supporting",  # incident reports attach to a related main voucher
    ),
    # --- TF1 Universal Standard (main) ---------------------------------
    "D07": (
        re.compile(rf"\(\s*d\s*{_D0}7\s*\)|d\s*{_D0}7[\s\-]+standard", re.I),
        re.compile(r"universal\s+d\s*[o0q]?7\s+standard|"
                   r"standar[a-z\-\s]{0,4}voucher", re.I),
        "main",
    ),
    # --- TF2.2 Activity (main; ALWAYS a 2-page form) -------------------
    # Page 1 carries the activity header + tables A/B; page 2 carries
    # tables C/D/E. Both pages print the (D08) footer, so we resolve
    # page-1 vs continuation by CONTENT (see D08_PAGE1_PAT / D08_PAGE2_PAT
    # and _resolve_category) rather than treating every D08 page as a new
    # main. The body fingerprint excludes the extension-list title
    # ("Extension-List to Activity-voucher") which would otherwise collide.
    "D08": (
        re.compile(rf"\(\s*d\s*{_D0}8\s*\)|d\s*{_D0}8[\s\-]+activity|"
                   rf"activity\s*\(\s*d\s*{_D0}8\s*\)", re.I),
        re.compile(r"(?<!to )activity[\s\-]?voucher", re.I),
        "main_or_cont_d08",
    ),
    # --- Extension list (always supporting) ----------------------------
    # The D09 footer is frequently OCR'd with a doubled zero-like char,
    # e.g. "(DO09)" (letter-O + digit-0 before the 9), so we allow up to
    # two such chars. The body title "Extension-List to Activity-voucher"
    # is matched first in identify_form() as a high-priority override.
    "D09": (
        re.compile(rf"\(\s*d\s*{_D0}{{0,2}}9\s*\)", re.I),
        re.compile(r"extension[\s\-]?l?ist", re.I),
        "supporting",
    ),
    # --- TF2.3 Travel (main; multi-page via main_or_cont) --------------
    "D10": (
        re.compile(rf"\(\s*d\s*{_D1}{_D0}\s*\)|d\s*{_D1}{_D0}[\s\-]+travel", re.I),
        re.compile(r"travel[\s\-]?voucher|"
                   r"travel\s*\(\s*d\s*[1li]?[0oq]\s*\)", re.I),
        "main_or_cont",
    ),
    # --- TF4 Substitute (main: replaces a missing payment voucher) -----
    "D11": (
        re.compile(rf"\(\s*d\s*{_D1}{_D1}\s*\)", re.I),
        re.compile(r"substitute[\s\-]?voucher|"
                   r"internal[\s\-]voucher\s+replaces", re.I),
        "main",
    ),
    # --- TF3 Reimbursement private km (main) ---------------------------
    "D12": (
        re.compile(rf"\(\s*d\s*{_D1}{_D2}\s*\)", re.I),
        re.compile(r"private\s+mileage(?:\s+km)?|"
                   r"reimbursement\s+private\s+km|"
                   r"private\s+reimbursement\s+of\s+project[\s\-]?vehicle", re.I),
        "main",
    ),
    # --- TF4 Hospitality (main) ----------------------------------------
    "D13": (
        re.compile(rf"\(\s*d\s*{_D1}3\s*\)", re.I),
        re.compile(r"hospitality[\s\-]?voucher|hospitality\s+costs", re.I),
        "main",
    ),
    # --- TF2.1 Phone share (supporting) --------------------------------
    "D14": (
        re.compile(rf"\(\s*d\s*7?{_D1}\s*4\s*\)", re.I),  # OCR sometimes reads "(D714)"
        re.compile(r"payment[\s\-]?voucher\s+for\s+private|"
                   r"pri?v[\s\.]?teleph|telephone\s+share|"
                   r"voucher[\-\s]?no[\.,]?\s+of\s+telephone", re.I),
        "supporting",
    ),
    # --- TF2.1 Collective (main) ---------------------------------------
    "D15": (
        re.compile(rf"\(\s*d\s*{_D1}5\s*\)", re.I),
        re.compile(r"collectiv[a-z\-\s]{0,4}voucher", re.I),
        "main",
    ),
    # --- TF2.3 Foreign daily allowance reference (supporting) ----------
    "D16": (
        re.compile(rf"\(\s*d\s*{_D1}6\s*\)", re.I),
        re.compile(r"overseas\s+daily\s+allowance|auslandstagegeld|"
                   r"foreig[a-z\.]*\s*daily\s*allow", re.I),
        "supporting",
    ),
    # --- TF5 Procurement (always supporting per user rule) -------------
    "D17": (
        re.compile(rf"\(\s*d\s*{_D1}7\s*\)", re.I),
        re.compile(r"procurement\s+voucher|direct\s+award", re.I),
        "supporting",
    ),
    # --- TF5 Inventory list (always supporting per user rule) ----------
    "D18": (
        re.compile(rf"\(\s*d\s*{_D1}8\s*\)", re.I),
        re.compile(r"inventory\s+list|inventory\s+number", re.I),
        "supporting",
    ),
    # --- TF5 Segregation (always supporting per user rule) -------------
    "D19": (
        re.compile(rf"\(\s*d\s*{_D1}9\s*\)", re.I),
        re.compile(r"segregation\s+report|date\s+of\s+segregation", re.I),
        "supporting",
    ),
    # --- TF5 Handover (always supporting per user rule) ----------------
    "D20": (
        re.compile(rf"\(\s*d\s*{_D2}{_D0}\s*\)", re.I),
        re.compile(r"handover\s+report|"
                   r"hand[\s\-]?over\s+to\s+the\s+a\.?m\.?", re.I),
        "supporting",
    ),
    # --- TF3 Reimbursement private Euro expenses (supporting) ----------
    "D21": (
        re.compile(rf"\(\s*d\s*{_D2}{_D1}\s*\)", re.I),
        re.compile(r"reimbursement\s+private(?!\s+km)|"
                   r"private\s+euro\s+expenses", re.I),
        "supporting",
    ),
    # --- TF2.3 Per-diem overview (supporting) --------------------------
    "D22": (
        re.compile(rf"\(\s*d\s*{_D2}{_D2}\s*\)", re.I),
        re.compile(r"per\s+diem\s+overview", re.I),
        "supporting",
    ),
}

# D10 page 1 carries one of these phrases; continuation pages do not.
D10_PAGE1_PAT = re.compile(r"official\s+travel\s+by|purpose\s+of\s+travel", re.I)

# D08 Activity-voucher is a fixed 2-page form.
#   Page 1: activity header + budget tables A (Transport) / B (Accommodation).
#   Page 2: budget tables C (Meals) / D (Fees) / E (Miscellanous) + totals.
# A D08 page is a CONTINUATION (page 2) when it shows page-2 table headers
# AND lacks all page-1 markers; otherwise it is treated as a page-1 main.
D08_PAGE1_PAT = re.compile(
    r"theme\s*/?\s*titel|period\s+of\s+time|numb\w*\.?\s*o\.?\s*particip|"
    r"\b5\.\s*activit|a\.\s*transport|b\.\s*accommod|objectiv",
    re.I,
)
D08_PAGE2_PAT = re.compile(
    r"c\.\s*meals|d\.\s*fees|e\.\s*miscell|total\s+c\.\s*meals",
    re.I,
)

# D08 activity-voucher PAGE-1 HEADER markers, used by identify_form Pass 1b
# as a footer-less fallback (see there). These three fields appear ONLY on
# the activity-voucher header and NEVER on the D09 extension list, which
# shares the A./B./C. budget-table headers but carries no activity header.
# (Deliberately NOT "A. Transport"/"C. Meals" -- shared with D09 -- nor
# "Objectives" -- it leaks onto activity-report supporting docs.)
D08_ACTIVITY_HEADER_PAT = re.compile(
    r"theme\s*/?\s*titel|topic\s+of\s+the\s+activ|\b5\.\s*activit",
    re.I,
)

# The D09 Extension-List form's title reads "Extension-List to
# Activity-voucher". The "Activity-voucher" substring would otherwise be
# captured by the D08 body fingerprint, so identify_form() matches this
# title first and routes the page straight to D09 (supporting). Tolerates
# the common OCR dropout of the "l" ("extension-ist") and matches either
# the title context or the footer code in parentheses.
EXTENSION_LIST_PAT = re.compile(
    r"extension[\s\-]?l?ist\s+to\s+activ|"
    r"extension[\s\-]?l?ist\s*\(\s*d",
    re.I,
)

# Captures the printed value of the "Voucher N°" cell (top-right of forms).
# The digit group is bounded by non-digit look-arounds so a 5+ digit noisy
# run (two glued numbers, a stray stroke) is REJECTED rather than truncated
# to a 2-4 digit prefix that might fall in range and bypass the safety net.
VOUCHER_NO_PAT = re.compile(r"voucher\s*n[°o\*\.\s]*\W{0,3}(?<!\d)(\d{2,4})(?!\d)", re.I)

# Multi-page form footer markers, slash form only: "1/2", "2 / 2".
# Used to force the page(s) following a "1/N" main into "main_cont", because
# continuation pages of e.g. D08 sometimes lose their D-code title to layout
# overlap and would otherwise be mis-classified as receipt/supporting.
# The earlier "n of m" / "n von m" word forms were dropped: they match
# ordinary receipt prose ("3 of 5 copies") and produced spurious markers
# that could fold a receipt into the main sheet. A genuine word-form
# continuation is still recovered by the D08 always-2-page fallback.
PAGE_MARKER_PAT = re.compile(r"\b([1-9])\s*/\s*([1-9])\b")

# Receipt-like keyword fingerprints (third-party documents). Used by the
# fall-through heuristic to disambiguate "no D-code recognised" pages.
# We deliberately keep these multi-lingual because the project receives
# scans from Indonesian, German, Portuguese and Tetum-speaking offices.
RECEIPT_PAT = re.compile(
    r"\b(?:invoice|receipt|faktur|kwitansi|nota|rechnung|quittung|"
    r"recibo|factura|fatura|"
    r"sub[\s\-]?total|grand[\s\-]?total|total\s+amount|amount\s+due|"
    r"vat|ppn|mwst|ust|"
    r"iban|swift|bank\s+transfer|customer\s+id|"
    r"jumlah|tanggal|bayar|terbilang|harga|"
    r"qty|menge|preis|tax\s+invoice|"
    r"item\s+description|unit\s+price|"
    # Telecom prepaid scratch cards (Timor-Leste, Indonesia). These are
    # legitimately "receipts" because they're third-party purchase proof,
    # not internal supporting documents.
    r"prepaid|scratch[\s\-]?card|"
    r"telemor|telkomcel|timor[\s\-]?telecom|"
    r"sosa\s+pakote|fasil\s+liu|ransu|folin|"
    r"konsulta\s+saldu|cek\s+saldo|check\s+(?:remain\s+)?credit|"
    r"loron\s+\d+\s*(?:ba\s+oin|days?)|"
    r"expired\s+in\s+\d+\s*days?|"
    r"customer\s+supporte|apolu\s+ba\s+kliente)\b",
    re.I,
)

# STRONG receipt phrases: so unambiguously receipt-specific that a SINGLE
# hit is enough to keep a page in the receipt bucket. These rescue real
# receipts whose photo-region score is inflated by round stamps / paid
# seals / halftone logos (e.g. a "PAYMENT RECEIPT" with circular cantina
# stamps) but whose body OCR is too garbled to yield two ordinary receipt
# keywords. None of these phrases occur on the photos / ID cards /
# declaration letters we want routed to supporting.
STRONG_RECEIPT_PAT = re.compile(
    r"payment\s+receipt|payment\s+received|cash\s+receipt|official\s+receipt|"
    r"tax\s+invoice|\binvoice\b|"
    r"amount\s+in\s+(?:usd|idr|rp|\$)|"
    r"thank\s+you\s+for\s+(?:shopping|your)|"
    r"\bfaktur\b|kwitansi|"
    r"telemor|telkomcel|timor[\s\-]?telecom|sosa\s+pakote|fasil\s+liu",
    re.I,
)

# Supporting-document keyword fingerprints (internal: participant lists,
# agendas, workshop itineraries, session plans, etc.). Conservative -- we
# require multiple hits before overriding the receipt default.
SUPPORTING_PAT = re.compile(
    r"\b(?:participant|participants|attendee|attendees|signature|"
    r"lista\s+partisipante|lista\s+de\s+participantes|teilnehmer|"
    r"agenda|itinerary|programme|program|schedule|session|"
    r"workshop|seminar|training|facilitator|moderator|speaker|"
    r"day\s+[1-9]|hari\s+ke|tag\s+[1-9]|"
    r"opening|closing|coffee\s+break|lunch\s+break|tea\s+break|"
    r"introduction|presentation|discussion|q\s*&\s*a|wrap[\s\-]?up|"
    r"learning\s+objectives|expected\s+outcomes|methodology|"
    r"venue|location|date\s*:|time\s*:)\b",
    re.I,
)

# A clock-time slot like "08:30 - 09:00" is a strong agenda/itinerary signal.
TIME_SLOT_PAT = re.compile(r"\b\d{1,2}[:.]\d{2}\s*[-–]\s*\d{1,2}[:.]\d{2}\b")

# Official letters / declarations (NOT receipts). These are internal
# supporting documents -- e.g. a "Karta Deklarasaun" (declaration letter)
# certifying vehicle ownership, an authorization letter, an Indonesian
# "surat keterangan". The vocabulary below is specific to formal
# correspondence and does not occur on third-party purchase receipts, so
# a SINGLE hit is enough to route the page to supporting. (Generic words
# like "authorization" that could appear on card receipts are excluded;
# we keep the Tetum/Portuguese/Indonesian declaration-specific forms.)
OFFICIAL_LETTER_PAT = re.compile(
    r"deklarasaun|deklara\s+katak|karta\s+deklara|"
    r"letter\s+of\s+declaration|carta\s+de\s+declara|declaration\s+letter|"
    r"surat\s+(?:keterangan|pernyataan|deklara)|"
    r"autorizasaun|authorizasaun",
    re.I,
)

# Contracts / formal agreements (NOT receipts). A consultancy or service
# contract attached to a voucher (e.g. the FSP-DE "Consultancy Contract"
# package: cover page, payment terms, general conditions, acceptance page)
# is an internal supporting document (_3). Its payment clauses legitimately
# contain receipt vocabulary ("upon submission of a detailed invoice",
# totals, reimbursable-expense lists), which used to pull those pages into
# the receipt bucket, so the phrases below are TITLE / legal-boilerplate
# signals that occur on contract pages but never on third-party purchase
# receipts -- a SINGLE hit routes the page to supporting (mirrors
# OFFICIAL_LETTER_PAT). Bare "contract" / "contract no." are deliberately
# NOT matched: utility invoices print "Contract No." and the consultant
# fee invoices in these bundles say "contracted total 20 days".
CONTRACT_PAT = re.compile(
    r"consultancy\s+contract|consulting\s+service\s+contract|"
    r"contract\s+acceptance|hereinafter\s+referred|"
    r"terms\s+and\s+conditions\s+of\s+(?:this|the)\s+contract|"
    r"termination\s+of\s+(?:this|the)\s+contract|"
    r"terms\s+of\s+reference|"
    r"contrato\s+de\s+(?:presta[cç]|consultoria|servi[cç]|trabalho)|"
    r"kontratu\s+(?:serbisu|konsultoria)|"
    r"surat\s+perjanjian|perjanjian\s+kerja",
    re.I,
)

# Receipt TITLE phrases -- the gate that keeps a GENUINE receipt in _2 even
# when its line items reference the contract it bills against ("as per
# consultancy contract"). Deliberately NARROWER than STRONG_RECEIPT_PAT:
# bare "invoice" is excluded because contract payment clauses say "upon
# submission of a detailed invoice", and that mention must not keep the
# contract itself in the receipt bucket. Only phrases that function as a
# receipt's own heading/stamp qualify.
RECEIPT_TITLE_PAT = re.compile(
    r"payment\s+receipt|payment\s+received|cash\s+receipt|official\s+receipt|"
    r"tax\s+invoice|\bfaktur\b|kwitansi|nota\s+no",
    re.I,
)


def looks_like_contract(text: str) -> bool:
    """True for a page of a contract / formal agreement (supporting, _3).

    A single CONTRACT_PAT title/boilerplate hit decides -- contract pages
    routinely score ordinary receipt keywords in their payment clauses, so
    (unlike looks_supporting) this is NOT weighed against receipt hits.
    The only veto is RECEIPT_TITLE_PAT: a page that carries an actual
    receipt heading is a receipt that merely references the contract.
    """
    if not text:
        return False
    norm = text.lower()
    if not CONTRACT_PAT.search(norm):
        return False
    return not RECEIPT_TITLE_PAT.search(norm)

# Participant / attendance lists (signed lists attached to activity
# vouchers) are SUPPORTING documents, but they OCR to little more than a
# letterhead plus handwritten rows, so the multi-hit looks_supporting()
# heuristic often misses them. Detected two ways (see looks_like_participant_list):
#
#  (a) TITLE -- an unambiguous single-hit signal, multi-lingual.
#      NB: the bare phrase "participants list" is deliberately EXCLUDED -- the
#      D08 activity voucher itself prints "attach a signed participants list",
#      and matching that could mislabel a degraded D08. We match the list's
#      own titles ("List of Participants", "Lista Partisipante", Portuguese
#      "Lista Presença", Indonesian "Daftar Hadir", German attendance list).
PARTICIPANT_LIST_PAT = re.compile(
    r"list\s+of\s+participant|lista\s+partisipante|"
    r"lista\s+de\s+participantes|lista\s+(?:de\s+)?presen[cç]|"
    r"attendance\s+(?:list|sheet)|presence\s+list|"
    r"daftar\s+hadir|teilnehmerliste|anwesenheitsliste",
    re.I,
)

# Activity-report narrative pages. The report's TITLE page ("Activity
# Report <activity name>") classifies as supporting via the multi-hit
# looks_supporting heuristic, but its CONTINUATION pages are plain
# narrative prose ("Through participatory activities and guided
# self-reflection, teachers explored ...") with no form code and no
# keywords at all, so they used to default to receipt. Handled by a
# propagation pass analogous to the participant-list one (process_file
# pass 1d): a supporting page containing the report title opens a block;
# following receipt-default pages flip to supporting ONLY when they read
# like narrative prose (>= REPORT_PROSE_MIN_TOKENS legible words) AND
# carry zero receipt keywords. A garbled scratch-card / NOTA page (~20
# legible words in these bundles) or any real receipt (>= 1 receipt
# keyword) fails the gate, stays in _2 and ends the block.
REPORT_TITLE_PAT = re.compile(r"\bactivity\s+report\b", re.I)
REPORT_PROSE_MIN_TOKENS = 50

#  (b) STRUCTURE -- the typical column layout of such a list: a name column,
#      a position/structure column, a signature/contact column, etc., laid
#      out as a table. We recognise it by the COLUMN-HEADER vocabulary below;
#      two or more DISTINCT column words on one page indicate a list table.
#      These words are list-specific (Tetum form headers + Portuguese/English
#      name/signature headers) and rare on receipts, so requiring two keeps
#      false positives off invoices. A signed list usually also repeats the
#      institutional header on each page, but the column headers are the
#      structural fingerprint.
PLIST_COL_PAT = re.compile(
    r"\b(?:naran|pozisaun|instituisaun|organizasaun|kontaktu|asinatura|"
    r"assinatura|tanda\s+tangan|nome\s+completo|estrutura|"
    r"presen[cç]a|semester|curso|"
    r"signature|organisation|institution)\b",
    re.I,
)


def looks_like_participant_list(text: str) -> bool:
    """True for a participant / attendance / signature list page.

    Matches either the list's TITLE (one hit, unambiguous) or its
    structural COLUMN layout (>= 2 distinct list-column header words).
    Used by identify_form's fall-through and as the trigger for the
    continuation-propagation pass.

    The structural branch is GATED by `not looks_like_receipt`: a
    third-party receipt (e.g. a "FATURA / invoice and cash receipt") also
    prints a customer "Naran / Name" field and an "Asinatura / signature"
    line, so without this gate those columns would mis-flag the receipt as
    a list. A real receipt's invoice/total keywords win and keep it in _2.
    """
    if not text:
        return False
    norm = text.lower()
    if PARTICIPANT_LIST_PAT.search(norm):
        return True
    if looks_like_receipt(text):
        return False
    return len({m.group(0).strip() for m in PLIST_COL_PAT.finditer(norm)}) >= 2

# Source filename: Q<q><yy>_5_Vouchers_No_<start>-<end>.pdf
# e.g. Q126_5_Vouchers_No_238-241.pdf, Q227_5_Vouchers_No_12-19.pdf
FILENAME_RE = re.compile(
    r"^(?P<prefix>Q\d{3})_5_Vouchers_No_(?P<start>\d+)-(?P<end>\d+)\.pdf$",
    re.I,
)

# ---------------------------------------------------------------------------
# Driver's-log (Fahrtenbuch) detection by perceptual hash
# ---------------------------------------------------------------------------

FAHRTENBUCH_HASH = (
    "1111111001001111001100000000111111110000000011111111000000001111"
    "1111000000001111111100000000111111110000000011111111000000001111"
    "1110000000001111111000000000111111110000000011111111000000001111"
    "1111000000001111100100000000111110010000000011111011000000001111"
    "1111111111111111"
)
FAHRTENBUCH_THRESHOLD = 40   # 256-bit hash -- well below the ~120 baseline

# Text-based Fahrtenbuch backup: handwritten / rotated / poorly-scanned
# driver's-log pages can drift far enough from the template that the
# perceptual hash misses them. We use TWO tiers of evidence:
#
#   Strong single-signal: the German-only words "Fahrtenbuch" or
#   "lfd. Seite". Neither appears on any other form in the bundle, so
#   either match is decisive.
#
#   Multi-marker consensus: at least two of the distinctive header /
#   column markers below. Each individual marker can show up on other
#   forms by coincidence (D17 procurement carries "project country:" too;
#   D12 reimbursement-private-km mentions "logbook" and "licence plate"),
#   but the combination is unique to Fahrtenbuch.
_FAHRTENBUCH_STRONG = re.compile(r"fahrtenbuch|lfd\.?\s*seite", re.I)
_FAHRTENBUCH_WEAK_MARKERS = (
    r"licence\s+number(?!\s+plate)",   # NOT "licence plate" (D12)
    r"partner\s+organisation\s*:",
    r"route\s*/\s*grund",
    r"\bkm\s+start\b",
    r"name\s+of\s+cps[\s\-]?worker",
    r"fahrer\s*in",
    r"\bk[mn]\s+ende\b",               # "Km Ende" with OCR-tolerant n/m
)


def is_fahrtenbuch_text(text: str) -> bool:
    """True when the OCR text reads like a Fahrtenbuch page.

    Tier 1 (strong): German-only title/marker present -> always Fahrtenbuch.
    Tier 2 (consensus): >= 2 of the distinctive column/header markers.
                        Empirically zero false positives on D12 / D17 /
                        all other voucher forms in the test bundles.
    """
    if not text:
        return False
    norm = text.lower()
    if _FAHRTENBUCH_STRONG.search(norm):
        return True
    hits = sum(1 for pat in _FAHRTENBUCH_WEAK_MARKERS
               if re.search(pat, norm, re.I))
    return hits >= 2


# ---------------------------------------------------------------------------
# Locating external binaries (Tesseract + Ghostscript) cross-platform
# ---------------------------------------------------------------------------

def find_binary(names: list[str], extra_search: list[Path] | None = None) -> Optional[str]:
    """Return the first existing binary from *names*, optionally searching extra paths."""
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    for path in extra_search or []:
        if path.exists():
            return str(path)
    return None


def locate_tesseract() -> str:
    extras: list[Path] = []
    if platform.system() == "Windows":
        extras = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    found = find_binary(["tesseract"], extras)
    if not found:
        sys.exit(
            "tesseract not found.\n"
            "  macOS:   brew install tesseract\n"
            "  Windows: install from https://github.com/UB-Mannheim/tesseract/wiki\n"
            "           (and tick 'Add to PATH' during setup)"
        )
    return found


def locate_ghostscript() -> str:
    extras: list[Path] = []
    if platform.system() == "Windows":
        # Ghostscript installs to versioned subdirs; pick the newest available.
        # Sort by PARSED version, not lexicographically -- string sorting puts
        # "gs9.55" above "gs10.00" ('9' > '1'), picking the older release.
        def _gs_ver(p: Path) -> tuple[int, int]:
            m = re.search(r"(\d+)\.(\d+)", p.name)
            return (int(m.group(1)), int(m.group(2))) if m else (0, 0)
        for base in [Path(r"C:\Program Files\gs"), Path(r"C:\Program Files (x86)\gs")]:
            if base.exists():
                for sub in sorted(base.iterdir(), key=_gs_ver, reverse=True):
                    for exe in ("gswin64c.exe", "gswin32c.exe"):
                        cand = sub / "bin" / exe
                        if cand.exists():
                            extras.append(cand)
    found = find_binary(["gs", "gswin64c", "gswin32c"], extras)
    if not found:
        sys.exit(
            "Ghostscript not found.\n"
            "  macOS:   brew install ghostscript\n"
            "  Windows: install from https://ghostscript.com/releases/gsdnld.html"
        )
    return found


TESSERACT: str = ""
GHOSTSCRIPT: str = ""


# ---------------------------------------------------------------------------
# OCR + perceptual-hash helpers
# ---------------------------------------------------------------------------

def _run_tesseract(png_path: Path, psm: int = 6,
                   whitelist: Optional[str] = None) -> str:
    # Force UTF-8 decoding and replace undecodable bytes so we never bubble
    # up a None / UnicodeDecodeError on Windows (default cp1252) when the
    # OCR output contains non-ASCII glyphs from scanned receipts.
    args = [TESSERACT, str(png_path), "-", "-l", "eng", "--psm", str(psm)]
    if whitelist:
        args.extend(["-c", f"tessedit_char_whitelist={whitelist}"])
    out = subprocess.run(args, capture_output=True, timeout=180)
    # A non-zero exit (missing langpack, OOM-killed child, bad PSM) yields
    # empty stdout that is indistinguishable from "blank page" and would
    # silently disable every detector. Surface it as a one-line stderr
    # warning, but do NOT raise -- one bad page must not abort the file.
    if out.returncode != 0:
        err = (out.stderr or b"").decode("utf-8", errors="replace").strip()
        last = err.splitlines()[-1] if err else f"exit code {out.returncode}"
        print(f"  !! tesseract failed on {png_path.name}: {last}", file=sys.stderr)
    raw = out.stdout or b""
    return raw.decode("utf-8", errors="replace")


def ocr_page(page: fitz.Page, dpi: int = OCR_DPI, psm: int = 6,
             clip: Optional[fitz.Rect] = None,
             whitelist: Optional[str] = None) -> str:
    pix = page.get_pixmap(dpi=dpi, clip=clip) if clip else page.get_pixmap(dpi=dpi)
    # NB: on Windows, NamedTemporaryFile keeps an OS-level lock on the file
    # while its handle is open, which prevents PyMuPDF's pix.save() from
    # overwriting it. Use mkstemp + immediate close so the path is reserved
    # but no handle is held while we write/read it.
    fd, name = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    tmp = Path(name)
    try:
        pix.save(tmp)
        return _run_tesseract(tmp, psm=psm, whitelist=whitelist)
    finally:
        try:
            tmp.unlink()
        except OSError:
            pass


def page_phash(page: fitz.Page, size: int = 16) -> str:
    """256-bit average-hash of a downscaled grayscale render."""
    pix = page.get_pixmap(
        matrix=fitz.Matrix(size / page.rect.width, size / page.rect.height),
        colorspace=fitz.csGRAY,
    )
    data = pix.samples
    avg = sum(data) / len(data)
    return "".join("1" if b > avg else "0" for b in data)


def hamming(a: str, b: str) -> int:
    # NB: the comparison over the overlapping prefix (via zip) is
    # INTENTIONAL and load-bearing. FAHRTENBUCH_HASH is 272 bits (its
    # template was captured at a 16x17 grid) while page_phash() returns
    # 256 bits (16x16); the Fahrtenbuch check relies on comparing the
    # aligned top 16 rows. A landscape driver's log whose OCR and footer
    # are unreadable is detected by THIS hash alone, so do not "guard"
    # unequal lengths away -- doing so silently drops those pages.
    return sum(x != y for x, y in zip(a, b))


# ---------------------------------------------------------------------------
# Per-page classification + voucher-number extraction
# ---------------------------------------------------------------------------

def identify_form(text: str) -> tuple[Optional[str], str]:
    """Return (form_code, category) for an OCR'd page.

    Match order (most -> least confident):
      0. Extension-List override -- the D09 title "Extension-List to
         Activity-voucher" contains "Activity-voucher", which would
         otherwise be captured by the D08 body fingerprint. Resolve it
         first so an extension list is never mistaken for a D08 main.
      1. Footer D-code regex (highest confidence) -- the printed voucher
         code, e.g. "(D 08)".
      1b. D08 activity-voucher CONTENT fallback -- used ONLY when the
         footer yielded no D-code (see below).
      2. Body keyword fallback (used only if no footer matched).
    """
    norm = " ".join((text or "").lower().split())
    # Pass 0 -- extension-list title/footer override (D09 supporting).
    if EXTENSION_LIST_PAT.search(norm):
        return "D09", "supporting"
    # Pass 1 -- footer D-code (preferred signal, always wins when present).
    for code, (footer_re, _body_re, cat) in FORM_PATTERNS.items():
        if footer_re.search(norm):
            return code, _resolve_category(cat, norm)
    # Pass 1b -- D08 activity-voucher CONTENT fallback.
    #   The newer "(D08+D09 extension list)" template prints its FILE PATH
    #   in the footer; that path text both (a) fails the D08 footer code
    #   "(D08)" -- so Pass 1 finds nothing -- and (b) contains the literal
    #   words "extension list", which the D09 body fingerprint would grab
    #   in Pass 2 whenever the "Activity-voucher" title OCRs poorly,
    #   silently misfiling the activity voucher as a D09 supporting page.
    #   So, ONLY when the footer gave no D-code, fall back to the activity
    #   voucher's HEADER fields (Theme/Titel, Topic of the activity,
    #   5. Activities) -- markers that appear on the activity-voucher
    #   page 1 but NEVER on the D09 extension list (which shares the
    #   A./C. budget-table headers, so those can't be used here). The
    #   page-2 continuation is picked up afterwards by the D08 always-2-page
    #   post-pass. A clearly-read footer D-code above always takes priority;
    #   the proper user-side fix is to keep the printed voucher code in the
    #   footer clear of the file path.
    if D08_ACTIVITY_HEADER_PAT.search(norm):
        return "D08", "main"
    # Pass 2 -- body keyword fallback.
    for code, (_footer_re, body_re, cat) in FORM_PATTERNS.items():
        if body_re.search(norm):
            return code, _resolve_category(cat, norm)
    # No D-code recognised. Default is "receipt", but check first whether
    # the page reads like an internal supporting document (participant
    # list, contract, workshop agenda/itinerary, etc.) which would
    # otherwise be mis-categorised when the user stacks the bundle as
    # main / supporting / receipts instead of main / receipts / supporting.
    # Official letters / declarations, contracts / formal agreements,
    # participant/attendance lists (by title OR table structure) and the
    # multi-hit agenda/participant heuristic all route to supporting.
    if (OFFICIAL_LETTER_PAT.search(norm)
            or looks_like_contract(norm)
            or looks_like_participant_list(norm)
            or looks_supporting(norm)):
        return None, "supporting"
    return None, "receipt"


def _resolve_category(cat: str, norm: str) -> str:
    if cat == "main_or_cont":
        # D10 travel voucher: page 1 carries the travel-purpose phrases.
        return "main" if D10_PAGE1_PAT.search(norm) else "main_cont"
    if cat == "main_or_cont_d08":
        # D08 activity voucher: a page is the 2nd (continuation) page only
        # when it shows the page-2 budget tables (C/D/E) AND none of the
        # page-1 markers. Anything else is treated as a page-1 main, so a
        # standalone or page-1 D08 still starts a new voucher.
        if D08_PAGE2_PAT.search(norm) and not D08_PAGE1_PAT.search(norm):
            return "main_cont"
        return "main"
    return cat


def extract_voucher_number(page: fitz.Page,
                           expected_range: Optional[tuple[int, int]] = None
                           ) -> Optional[int]:
    """OCR the top-right of *page* and pull out the 'Voucher N°' value.

    Voucher numbers are often handwritten (especially in this project where
    field staff fill them in by hand). Standard prose OCR (PSM 6) reads
    handwriting unreliably -- "561" gets misread as "54", "S6/", "$6/" etc.
    To make this robust:

      1. We OCR a *narrow* crop (just the "Voucher N°" cell) at high DPI
         under THREE Tesseract layout modes:
           PSM 7  - treat region as a single text line
           PSM 8  - treat region as a single word
           PSM 11 - sparse text, no orientation assumptions
      2. PSM 8 + 7 are run with a digit-only whitelist
         (`tessedit_char_whitelist=0123456789`). This drops noise glyphs
         like "S", "$", "/", "O" that handwriting often gets confused for.
      3. The general PSM 6 pass (including the "Voucher N°" label match)
         is kept as a fourth signal so a printed-number form still works.
      4. We collect every plausible candidate (2-4 digits, in range
         [1, 9999]) and:
           - prefer candidates within the bundle's expected range
             ([range_start, range_end] from the filename); these are
             almost certainly correct.
           - among in-range candidates, take the most-frequent (multiple
             PSMs agreeing is strong evidence).
           - among out-of-range candidates, return None -- the sequential
             fallback in process_file() will use the position instead.
    """
    w, h = page.rect.width, page.rect.height
    clip = fitz.Rect(w * 0.55, h * 0.04, w, h * 0.20)

    candidates: list[int] = []

    # Pass 1: standard prose pass (catches printed "Voucher N° 562")
    text = " ".join(ocr_page(page, dpi=VOUCHER_NO_DPI, psm=6, clip=clip).split())
    m = VOUCHER_NO_PAT.search(text)
    if m:
        try:
            n = int(m.group(1))
            if 1 <= n <= 9999:
                candidates.append(n)
        except ValueError:
            pass

    # Pass 2-4: digit-only OCR under PSM 7, 8, 11 -- captures handwriting
    # the prose pass misreads. Each PSM uses a different layout assumption,
    # so disagreements expose unreliable reads.
    for psm in (7, 8, 11):
        digit_text = ocr_page(page, dpi=VOUCHER_NO_DPI, psm=psm, clip=clip,
                              whitelist="0123456789").strip()
        for token in re.findall(r"(?<!\d)\d{2,4}(?!\d)", digit_text):
            try:
                n = int(token)
                if 1 <= n <= 9999:
                    candidates.append(n)
            except ValueError:
                pass

    if not candidates:
        return None

    # Prefer candidates inside the expected range -- those are almost
    # certainly correct (the bundle's filename tells us what numbers to
    # expect, and any reading outside that range is far more likely to be
    # an OCR error than a real misnumbered voucher).
    if expected_range is not None:
        lo, hi = expected_range
        # Allow a small tolerance below/above the range to absorb any
        # legitimate gap between bundles, but reject wildly-wrong reads
        # like 54 when the range is 561-576.
        in_range = [c for c in candidates if lo <= c <= hi]
        if in_range:
            # Most-frequent in-range candidate (multi-PSM consensus). On a
            # genuine tie (no single value has the highest vote) the read is
            # ambiguous -- return None so process_file() uses the page's
            # positional number rather than arbitrarily biasing toward lo.
            counts: dict[int, int] = {}
            for c in in_range:
                counts[c] = counts.get(c, 0) + 1
            top = max(counts.values())
            winners = [k for k, v in counts.items() if v == top]
            return winners[0] if len(winners) == 1 else None
        # No in-range candidates -- the OCR clearly failed. Returning
        # None lets process_file() use position-based sequential numbering.
        return None

    # No range constraint; return the most-frequent candidate.
    counts = {}
    for c in candidates:
        counts[c] = counts.get(c, 0) + 1
    return max(counts, key=counts.get)


def extract_page_marker(page: fitz.Page) -> Optional[tuple[int, int]]:
    """OCR the bottom strip of *page*; return (k, n) if it shows 'k/n' / 'k of n'.

    Used to detect multi-page main forms (D08 typically prints "1/2", "2/2"
    in the footer). The continuation page often loses its D-code title to
    table-overlap during scanning, so the marker becomes the only reliable
    way to keep both pages glued to the same _1 file.
    """
    w, h = page.rect.width, page.rect.height
    clip = fitz.Rect(0, h * 0.88, w, h)
    text = " ".join(ocr_page(page, dpi=VOUCHER_NO_DPI, psm=6, clip=clip).split())
    m = PAGE_MARKER_PAT.search(text)
    if not m:
        return None
    try:
        k, n = int(m.group(1)), int(m.group(2))
    except ValueError:
        return None
    if 1 <= k <= n <= 9:
        return (k, n)
    return None


def is_low_signal(text: str) -> bool:
    """OCR returned mostly garbage -- too few legible words to classify on.

    Used (together with is_photo_page) to detect photos and badly-scanned
    pages. Threshold (20 alphabetic tokens of length >= 3) sits below
    what most forms and supporting docs produce, but receipts can also
    fall below it -- so callers must combine this with another signal
    before re-categorising the page.
    """
    tokens = re.findall(r"[A-Za-z]{3,}", text or "")
    return len(tokens) < 20


# ---------------------------------------------------------------------------
# Photograph detection (largest contiguous continuous-tone region)
# ---------------------------------------------------------------------------
#
# Why this approach (and why the old one failed):
#   These scans store each page as MANY embedded image XObjects (tiles,
#   masks, layers -- 6 to 49 per page in real bundles). Extracting "the"
#   embedded image and measuring it is unreliable: the largest-area
#   fragment is frequently NOT the visible scan, so the old stdev/midtone
#   numbers were measuring noise. We therefore RENDER the page (MuPDF is
#   bundled identically in every PyMuPDF wheel, so a fixed-DPI render is
#   deterministic across macOS/Linux/Windows) and analyse what is actually
#   printed.
#
# What distinguishes a photograph from a document/receipt:
#   A photo has a LARGE CONTIGUOUS region of continuous-tone (mid-gray)
#   pixels. A text document is bimodal (black ink on white) with almost
#   no mid-grays; a receipt's mid-grays (logos, halftone, coloured text)
#   are small and SCATTERED. So we measure the size of the single largest
#   connected blob of "continuous-tone" blocks -- big for photos, tiny
#   for documents. Halftone-heavy printed receipts (telecom scratch
#   cards) also score high here, but they are excluded by the
#   looks_like_receipt() text gate at the call site.
#
# Cross-OS robustness: block-averaging (each block ~= 144 rendered
# pixels) plus a wide acceptance band and a ~2x threshold margin absorb
# any sub-pixel rasteriser differences between platform builds.

_PHOTO_RENDER_DPI = 100      # fixed -> identical pixel grid on every OS
_PHOTO_BLOCK = 12            # render-pixels per analysis block
_PHOTO_TONE_LO = 55          # continuous-tone acceptance band (grayscale)
_PHOTO_TONE_HI = 200
_PHOTO_BLOCK_FILL = 0.5      # a block is "tone" if >50% of it is mid-gray
_PHOTO_REGION_MIN = 0.025    # largest tone-blob >2.5% of page -> photo-like


def _largest_blob(mask: list[int], w: int, h: int) -> int:
    """Largest 4-connected component of 1s in a row-major w*h grid."""
    seen = bytearray(len(mask))
    best = 0
    for start in range(len(mask)):
        if mask[start] and not seen[start]:
            seen[start] = 1
            stack = [start]
            size = 0
            while stack:
                p = stack.pop()
                size += 1
                x = p % w
                y = p // w
                if x > 0 and mask[p - 1] and not seen[p - 1]:
                    seen[p - 1] = 1; stack.append(p - 1)
                if x < w - 1 and mask[p + 1] and not seen[p + 1]:
                    seen[p + 1] = 1; stack.append(p + 1)
                if y > 0 and mask[p - w] and not seen[p - w]:
                    seen[p - w] = 1; stack.append(p - w)
                if y < h - 1 and mask[p + w] and not seen[p + w]:
                    seen[p + w] = 1; stack.append(p + w)
            if size > best:
                best = size
    return best


def photo_region_score(page: fitz.Page) -> float:
    """Fraction of the page covered by the single largest contiguous
    continuous-tone region. High for photographs, ~0 for text documents.

    Deterministic across operating systems: fixed-DPI MuPDF render ->
    Pillow LUT mask -> Pillow BOX downsample to blocks -> pure-Python
    connected-components. See the module comment above for rationale.
    """
    try:
        pix = page.get_pixmap(dpi=_PHOTO_RENDER_DPI, colorspace=fitz.csGRAY)
        im = Image.frombytes("L", (pix.width, pix.height), pix.samples)
    except Exception:
        return 0.0
    # 1 where the pixel is mid-gray (continuous tone), 0 elsewhere.
    lut = [255 if _PHOTO_TONE_LO <= v <= _PHOTO_TONE_HI else 0
           for v in range(256)]
    mask = im.point(lut)
    nbx = pix.width // _PHOTO_BLOCK
    nby = pix.height // _PHOTO_BLOCK
    if nbx < 2 or nby < 2:
        return 0.0
    # BOX downsample -> each output pixel = mean of its block (0..255),
    # i.e. 255 * (fraction of the block that is mid-gray).
    small = mask.resize((nbx, nby), Image.BOX)
    cutoff = _PHOTO_BLOCK_FILL * 255
    grid = [1 if b > cutoff else 0 for b in small.tobytes()]
    total = nbx * nby
    if not any(grid):
        return 0.0
    return _largest_blob(grid, nbx, nby) / total


def is_photo_page(page: fitz.Page) -> bool:
    """True if the page is dominated by a photograph rather than a
    document scan. Caller must still exclude printed receipts via the
    looks_like_receipt() text gate."""
    return photo_region_score(page) > _PHOTO_REGION_MIN


def looks_like_receipt(text: str) -> bool:
    """True if the page text contains strong, unambiguous receipt keywords.

    Used as a positive override against the photo flip: a colourful page
    full of telecom scratch-card art (high midtones, looks photo-like by
    image content) but whose OCR clearly reads "telemor", "Konsulta saldu",
    "Loron 30 days" should stay classified as a receipt -- not get flipped
    to supporting just because the page is colourful.

    Cross-OS-stable because it relies on POSITIVE matching of specific
    multi-lingual keywords (which exist on real receipts) rather than
    counting all OCR tokens (which varies wildly between Tesseract
    versions, especially on photos).

    Tiered: one STRONG, unambiguous receipt phrase (e.g. "payment
    receipt", a telecom brand) suffices; otherwise we require at least
    two ordinary receipt keywords. The strong tier rescues real receipts
    whose photo-region score is inflated by stamps / seals / logos but
    whose OCR yields only a single ordinary keyword.
    """
    if not text:
        return False
    norm = text.lower()
    if STRONG_RECEIPT_PAT.search(norm):
        return True
    return len(RECEIPT_PAT.findall(norm)) >= 2


def looks_supporting(text: str) -> bool:
    """Conservative heuristic for fall-through pages with no D-code footer.

    Returns True only when the page reads like an internal supporting
    document (participant list, workshop agenda/itinerary, session plan)
    rather than a third-party receipt. Requires at least two supporting
    signals AND for them to outweigh receipt signals -- so a real invoice
    that happens to mention "presentation" once still classifies as receipt.
    """
    if not text:
        return False
    norm = text.lower()
    receipt_hits = len(RECEIPT_PAT.findall(norm))
    support_hits = len(SUPPORTING_PAT.findall(norm))
    # An agenda-style time slot is worth two ordinary supporting hits.
    support_hits += 2 * len(TIME_SLOT_PAT.findall(norm))
    return support_hits >= 2 and support_hits > receipt_hits


# ---------------------------------------------------------------------------
# PDF assembly + ebook compression
# ---------------------------------------------------------------------------

def write_subset(src: fitz.Document, page_indices: list[int], out_path: Path) -> None:
    """Write *page_indices* to a new PDF, then compress it to ebook quality."""
    if not page_indices:
        return
    tmp_pdf = out_path.with_suffix(".uncompressed.pdf")
    sub = fitz.open()
    for idx in page_indices:
        sub.insert_pdf(src, from_page=idx, to_page=idx)
    sub.save(tmp_pdf, deflate=True, garbage=4)
    sub.close()

    gs_cmd = [
        GHOSTSCRIPT,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/ebook",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-sOutputFile={out_path}",
        str(tmp_pdf),
    ]
    # timeout: a malformed PDF can make Ghostscript loop forever, hanging an
    # unattended batch with no error. finally: always remove the uncompressed
    # temp, even if gs raises/times out, so it is never left in the folder.
    try:
        subprocess.run(gs_cmd, check=True, timeout=300)
    finally:
        tmp_pdf.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Main per-file pipeline
# ---------------------------------------------------------------------------

def parse_filename(path: Path) -> Optional[tuple[str, int, int]]:
    """Return (quarter_prefix, start, end) parsed from a Q<q><yy>_5_Vouchers_No_<a>-<b>.pdf name."""
    m = FILENAME_RE.match(path.name)
    if not m:
        return None
    return (m.group("prefix").upper(), int(m.group("start")), int(m.group("end")))


def process_file(src_path: Path, dry_run: bool = False) -> None:
    parsed = parse_filename(src_path)
    if not parsed:
        print(f"Cannot parse quarter/range from filename: {src_path.name} -- skipping.")
        return
    quarter_prefix, range_start, range_end = parsed
    output_prefix = f"{quarter_prefix}_6_Voucher"
    expected_count = range_end - range_start + 1

    print(f"\n=== {src_path.name}  ({quarter_prefix}, range {range_start}-{range_end}, "
          f"expecting {expected_count}) ===")
    doc = fitz.open(src_path)

    # 1. classify every page
    #
    # Per-page logic:
    #   a) Fahrtenbuch detection: perceptual-hash similarity OR
    #      text-based fingerprint (German "Fahrtenbuch", "driver's log",
    #      etc.). Either one is enough -- text rescues hand-written /
    #      rotated / poorly-scanned pages the hash misses.
    #   b) OCR full page; identify form by D-code / body keyword.
    #   c) OCR bottom strip for an "X/Y" page-number marker.
    #   d) Photo flip: a receipt-default page that looks like a
    #      photograph (image-content metrics) AND whose OCR text does
    #      NOT contain strong receipt keywords gets re-categorised as
    #      supporting. The receipt-keyword gate is positive matching
    #      against multi-lingual receipt vocabulary -- it's stable
    #      across OSes (unlike a token-count gate, which depends on
    #      Tesseract version), and it stops colourful printed
    #      documents (e.g. telecom prepaid scratch cards) from being
    #      mis-flipped.
    #   e) Backward-pull: a page with marker (k>=2, N) belongs to the
    #      most recent main voucher form -- but ONLY if that main has
    #      its OWN matching (1, N) marker AND no intervening page is
    #      already advertising itself as (1, N) of something else.
    expected_range = (range_start, range_end)
    page_info: list[dict] = []
    for i, page in enumerate(doc):
        ph = page_phash(page)
        text = ""
        # (a) Fahrtenbuch via hash first (cheap, no OCR needed)
        if hamming(ph, FAHRTENBUCH_HASH) <= FAHRTENBUCH_THRESHOLD:
            info = {"page": i, "form": "FAHRTENBUCH", "cat": "fahrtenbuch",
                    "vnum": None, "marker": None, "marker_tuple": None,
                    "photo_score": None, "phash": ph}
        else:
            text = ocr_page(page)
            # (a, fallback) Fahrtenbuch via text -- catches pages whose
            # template-deviation drove the hash off but whose form text
            # is unmistakable. See is_fahrtenbuch_text() for the
            # two-tier strong/consensus logic that avoids false
            # positives on D12 ("logbook" / "licence plate") and D17
            # ("project country" alone) pages.
            if is_fahrtenbuch_text(text):
                info = {"page": i, "form": "FAHRTENBUCH", "cat": "fahrtenbuch",
                        "vnum": None, "marker": None, "marker_tuple": None,
                        "photo_score": None, "phash": ph}
            else:
                form, cat = identify_form(text)
                marker_tuple = extract_page_marker(page)

                # (d) photo-like receipt-default page -> supporting,
                #     gated by absence of strong receipt keywords. The
                #     photo score is the largest contiguous continuous-
                #     tone region (see photo_region_score); the text gate
                #     keeps printed receipts (incl. telecom scratch cards,
                #     which also score high) in the receipt bucket.
                pscore = photo_region_score(page) if cat == "receipt" else None
                if pscore is not None:
                    if pscore > _PHOTO_REGION_MIN and not looks_like_receipt(text):
                        cat = "supporting"

                # (e) tightened backward-pull
                if (marker_tuple is not None and marker_tuple[0] >= 2
                        and cat in ("receipt", "supporting")):
                    n = marker_tuple[1]
                    coerce = False
                    for j in range(i - 1, -1, -1):
                        prev = page_info[j]
                        prev_m = prev.get("marker_tuple")
                        # Abort if we see a competing (1, N) on a non-main:
                        # that's a different multi-page document.
                        if (prev_m == (1, n)
                                and prev["cat"] not in ("main", "main_cont")):
                            break
                        if prev["cat"] in ("main", "main_cont"):
                            if prev_m == (1, n):
                                coerce = True
                            break  # nearest main decides
                    if coerce:
                        cat = "main_cont"
                        form = form or "FORCED"

                vnum = (extract_voucher_number(page, expected_range)
                        if cat == "main" else None)
                info = {"page": i, "form": form, "cat": cat, "vnum": vnum,
                        "marker": (f"{marker_tuple[0]}/{marker_tuple[1]}"
                                   if marker_tuple else None),
                        "marker_tuple": marker_tuple,
                        "photo_score": pscore,
                        "phash": ph,
                        # flags for the participant-list / activity-report
                        # continuation passes (1d)
                        "is_plist": looks_like_participant_list(text),
                        "strong_receipt": looks_like_receipt(text),
                        "is_report": bool(REPORT_TITLE_PAT.search(text)),
                        "report_prose": (
                            not RECEIPT_PAT.search(text)
                            and not looks_like_receipt(text)
                            and len(re.findall(r"[A-Za-z]{3,}", text))
                                >= REPORT_PROSE_MIN_TOKENS)}

        page_info.append(info)
        snippet = " ".join(text.split())[:80] if text else ""
        vn_disp = info["vnum"] if info["vnum"] is not None else "-"
        marker_disp = f" [{info['marker']}]" if info.get("marker") else ""
        ps_info = info.get("photo_score")
        photo_disp = f" img={ps_info:.3f}" if ps_info is not None else ""
        print(f"  p{i+1:02d}  {info['cat']:<11}  form={str(info['form'] or '-'):<5}  "
              f"vnum={vn_disp:<5}{marker_disp}{photo_disp}  {snippet}")

    # 1b. Post-pass: D08 always-2-page rule.
    #     D08 (Activity-voucher) is a TWO-page form by template: page 1
    #     carries the activity header and budget tables A/B (Transport,
    #     Accommodation), page 2 carries tables C/D/E (Meals, Fees,
    #     Miscellanous). The page-2 footer often gets garbled in scans
    #     (table-overlap, light "(D08)" footer text), so it lands as
    #     supporting/receipt by content alone.
    #     We force the continuation, preferring marker-based detection
    #     (any subsequent (2, N) marker before the next main) and falling
    #     back to "the immediately next non-main page" for canonical-order
    #     bundles where the marker OCR also failed.
    for i, info in enumerate(page_info):
        if info.get("form") == "D08" and info["cat"] == "main":
            # bound: stop at next main / fahrtenbuch / skip
            end = len(page_info)
            for j in range(i + 1, len(page_info)):
                if page_info[j]["cat"] in ("main", "fahrtenbuch", "skip"):
                    end = j
                    break
            cont_idx = None
            # marker-based first (handles scrambled bundles)
            for j in range(i + 1, end):
                mt = page_info[j].get("marker_tuple")
                if mt and mt[0] == 2:
                    cont_idx = j
                    break
            # canonical-order fallback: the next non-main page IS the
            # D08 continuation when the marker OCR failed on both pages.
            if cont_idx is None and i + 1 < end:
                nxt = page_info[i + 1]
                if nxt["cat"] in ("receipt", "supporting"):
                    cont_idx = i + 1
            if cont_idx is not None and page_info[cont_idx]["cat"] != "main_cont":
                page_info[cont_idx]["cat"] = "main_cont"
                if not page_info[cont_idx].get("form"):
                    page_info[cont_idx]["form"] = "FORCED"
                print(f"  -> p{cont_idx+1:02d} forced to main_cont "
                      f"(D08 continuation of p{i+1:02d})")

    # 1c. Post-pass: detect duplicate consecutive main pages.
    #     Field staff occasionally scan the same voucher form twice
    #     (carbon copy + signed original, or accidental rescan).
    #     Two consecutive "main" pages of the same form code with very
    #     similar perceptual hashes are almost certainly duplicates of
    #     the same voucher; we merge by demoting the second to main_cont.
    #     Threshold (hamming < 35 of 256 bits) is lenient enough to
    #     absorb signature/stamp differences but tight enough to avoid
    #     false merges of two legitimate sequential vouchers.
    for i in range(len(page_info) - 1):
        a, b = page_info[i], page_info[i + 1]
        if (a["cat"] == "main" and b["cat"] == "main"
                and a.get("form") and a.get("form") == b.get("form")):
            ph_a, ph_b = a.get("phash"), b.get("phash")
            if ph_a and ph_b and hamming(ph_a, ph_b) < 35:
                b["cat"] = "main_cont"
                if not b.get("form"):
                    b["form"] = "FORCED"
                print(f"  -> p{b['page']+1:02d} forced to main_cont "
                      f"(duplicate scan of p{a['page']+1:02d})")

    # 1d. Post-pass: participant-list / activity-report continuation
    #     propagation.
    #     A multi-page signed participant/attendance list is detectable on
    #     its first (titled / column-headed) page, but the continuation
    #     pages OCR to little more than handwriting and would default to
    #     "receipt". Likewise an activity report is detectable by its
    #     title page, while its continuation pages are plain narrative
    #     prose with no keywords at all. Per the bundle convention
    #     (main -> receipts -> supporting) and the user's rule that both
    #     document types are always _3, once a list/report page is seen we
    #     carry "supporting" forward through the immediately-following
    #     receipt-default pages until the block clearly ends.
    #     Strictly bounded so it never swallows a real receipt:
    #       - starts ONLY at a detected participant-list page (is_plist)
    #         or a supporting-classified activity-report page (is_report),
    #       - a LIST block propagates only to "receipt" pages WITHOUT
    #         strong receipt keywords (a real invoice/slip breaks the run
    #         and stays _2),
    #       - a REPORT block is stricter: only to pages that read like
    #         narrative prose AND carry ZERO receipt keywords
    #         (report_prose flag) -- a garbled low-text receipt does not
    #         qualify,
    #       - resets at any main/continuation/fahrtenbuch/skip boundary and
    #         at any other supporting page (photo, itinerary).
    block = None  # None | "plist" | "report"
    for info in page_info:
        cat = info["cat"]
        if cat in ("main", "main_cont", "fahrtenbuch", "skip"):
            block = None
        elif info.get("is_plist"):
            block = "plist"                  # already classified supporting
        elif cat == "supporting" and info.get("is_report"):
            block = "report"
        elif (cat == "receipt" and block == "plist"
              and not info.get("strong_receipt")):
            info["cat"] = "supporting"        # continuation of the list block
            print(f"  -> p{info['page']+1:02d} -> supporting "
                  f"(participant-list continuation)")
        elif (cat == "receipt" and block == "report"
              and info.get("report_prose")):
            info["cat"] = "supporting"        # narrative continuation
            print(f"  -> p{info['page']+1:02d} -> supporting "
                  f"(activity-report continuation)")
        else:
            block = None                      # strong receipt / other doc ends it

    # 2. group pages into voucher buckets (one per main)
    #
    # Voucher numbering: when extract_voucher_number() returned a value
    # within the expected range we trust it. When it returned None we
    # use sequential numbering anchored to range_start: this is robust
    # to the common case of handwritten numbers that defeat OCR -- the
    # filename's range is the ground truth, and walking it in page order
    # always produces the right numbers as long as no main was missed.
    groups: list[dict] = []
    deferred_fahrtenbuch: list[int] = []
    next_seq = range_start

    def attach_fahrtenbuch(p: int) -> None:
        """Attach a fahrtenbuch page to the most recent D12 voucher group."""
        for g in reversed(groups):
            if g["form"] == "D12":
                g["supporting"].append(p)
                return
        deferred_fahrtenbuch.append(p)  # save for later (no D12 yet)

    for info in page_info:
        cat = info["cat"]
        i = info["page"]
        if cat == "skip":
            continue
        if cat == "fahrtenbuch":
            attach_fahrtenbuch(i)
            continue
        if cat == "main":
            vnum = info["vnum"] if info["vnum"] is not None else next_seq
            # Advance the positional counter by exactly one per main, NEVER
            # re-anchored to the OCR read. A single wrong-but-in-range read
            # then mislabels only its own page instead of shifting (and
            # possibly colliding) every later voucher number.
            next_seq += 1
            groups.append({"vnum": vnum, "form": info["form"],
                           "main": [i], "receipt": [], "supporting": []})
        elif cat == "main_cont":
            if not groups:
                print(f"  !! page {i+1}: continuation with no preceding main -- skipping")
                continue
            groups[-1]["main"].append(i)
        elif cat == "supporting":
            if not groups:
                print(f"  !! page {i+1}: supporting form with no preceding main -- skipping")
                continue
            groups[-1]["supporting"].append(i)
        else:  # receipt
            if not groups:
                print(f"  !! page {i+1}: receipt with no preceding main -- skipping")
                continue
            groups[-1]["receipt"].append(i)

    # 3. handle any deferred fahrtenbuch pages (those that came BEFORE
    #    the file's first D12 voucher, if it has one). We attach them
    #    to the first D12 we found.
    if deferred_fahrtenbuch:
        d12 = next((g for g in groups if g["form"] == "D12"), None)
        if d12:
            d12["supporting"].extend(deferred_fahrtenbuch)
            print(f"  deferred driver's log p{[p+1 for p in deferred_fahrtenbuch]} "
                  f"-> voucher {d12['vnum']} supporting")
        else:
            print(f"  driver's log p{[p+1 for p in deferred_fahrtenbuch]} -> "
                  f"ignored (no D12 voucher in this file)")

    # 3b. Guard against duplicate voucher numbers. A wrong-but-in-range OCR
    #     read can land on another main's number; two groups sharing a vnum
    #     would write the same Q..._6_Voucher_<vnum>_*.pdf and the second
    #     would silently overwrite the first (Ghostscript -sOutputFile
    #     truncates). On any collision, fall back to position-based
    #     sequential numbering from range_start -- the filename range is the
    #     ground truth for a contiguous bundle -- so every output is unique.
    vnums = [g["vnum"] for g in groups]
    if len(set(vnums)) != len(vnums):
        print(f"  !! WARNING: duplicate voucher numbers {vnums} -- "
              f"renumbering sequentially from {range_start}.")
        for offset, g in enumerate(groups):
            g["vnum"] = range_start + offset

    # 4. sanity-check
    if len(groups) != expected_count:
        print(f"  !! WARNING: detected {len(groups)} main vouchers; "
              f"filename suggests {expected_count}.")
        print(f"     Numbers used: {[g['vnum'] for g in groups]}")

    # 5. emit per-voucher PDFs
    out_dir = src_path.parent
    for g in groups:
        vnum = g["vnum"]
        # Flag vouchers that end up with no _2 receipt file. A faint
        # receipt glued onto the voucher page (a scanning user-error) is
        # NOT reliably detectable on the page itself -- it OCRs to almost
        # nothing and carries no continuous-tone content -- so the only
        # robust signal is structural: the voucher has a main but no
        # separate receipt. The page is still (correctly) sorted to _1;
        # this flag prompts the operator to check whether a receipt was
        # glued onto the voucher sheet or is genuinely missing.
        no_receipt = "  [!] no _2 receipt -- receipt may be glued onto the voucher page or missing" \
            if not g["receipt"] else ""
        print(f"  voucher {vnum}: main={[p+1 for p in g['main']]}, "
              f"receipts={[p+1 for p in g['receipt']]}, "
              f"supporting={[p+1 for p in g['supporting']]}{no_receipt}")
        if dry_run:
            continue
        for suffix, kind in [("1", "main"), ("2", "receipt"), ("3", "supporting")]:
            pages = g[kind]
            if not pages:
                continue
            out_path = out_dir / f"{output_prefix}_{vnum}_{suffix}.pdf"
            write_subset(doc, pages, out_path)
            print(f"    -> {out_path.name}  ({out_path.stat().st_size:,} bytes)")

    doc.close()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Split scanned voucher PDFs by category and compress to ebook quality.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("pdfs", nargs="*", type=Path,
                    help="Specific PDFs to process. "
                         f"Default: every {INPUT_GLOB} in --folder.")
    ap.add_argument("--folder", type=Path, default=DEFAULT_FOLDER,
                    help=f"Folder to scan when no PDFs are given (default: {DEFAULT_FOLDER}).")
    ap.add_argument("--dry-run", action="store_true",
                    help="Print classification + planned grouping without writing files.")
    args = ap.parse_args()

    global TESSERACT, GHOSTSCRIPT
    TESSERACT = locate_tesseract()
    GHOSTSCRIPT = locate_ghostscript()
    print(f"using tesseract:   {TESSERACT}")
    print(f"using ghostscript: {GHOSTSCRIPT}")

    if args.pdfs:
        pdfs = args.pdfs
    else:
        if not args.folder.exists():
            sys.exit(f"Folder not found: {args.folder}")
        pdfs = sorted(args.folder.glob(INPUT_GLOB))
    if not pdfs:
        sys.exit(f"No input PDFs matched in {args.folder} (pattern '{INPUT_GLOB}').")

    failures = 0
    for pdf in pdfs:
        # Isolate each file: a corrupt/encrypted PDF, a Tesseract timeout, a
        # Ghostscript error or a MemoryError must not abort the remaining
        # files. Report it and keep going; exit non-zero so callers notice.
        try:
            process_file(pdf, dry_run=args.dry_run)
        except Exception as e:
            failures += 1
            print(f"!! FAILED {pdf.name}: {type(e).__name__}: {e}", file=sys.stderr)
    if failures:
        sys.exit(f"\n{failures} file(s) failed -- see messages above.")


if __name__ == "__main__":
    main()
