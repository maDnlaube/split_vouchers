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

1. Driver's log (Fahrtenbuch): perceptual-hash match against a template
   OR text fingerprint ("Fahrtenbuch", "driver's log", column headers) --
   either alone suffices, so degraded/rotated scans still get caught.
   Attached to the NEAREST PRECEDING D12 main; pages before any D12 are
   deferred to the first D12 found, or dropped if none exists.

2. The OCR'd footer D-code drives the category: D02 skip (per user);
   D07/D11/D12/D13/D15 main; D10 main with continuations folded into the
   same _1; D06/D09/D14/D16-D22 supporting.

3. No D-code: default is receipt, flipped to supporting when the page
   reads like an internal document -- participant list (title or column
   structure), agenda/itinerary keyword consensus, official letter or
   declaration, or contract package (title/legal-boilerplate phrases,
   vetoed by an actual receipt heading so an invoice that merely
   references the contract stays _2). Photograph-like pages (largest
   contiguous continuous-tone region of a deterministic fixed-DPI
   render) also flip to supporting unless strong multi-lingual receipt
   keywords (incl. telecom prepaid cards) say otherwise.

4. Multi-page mains print "k/N" footer markers: a non-main page with
   (k>=2, N) re-attaches to the most recent main -- only if that main
   advertised its own (1, N) and no intervening page claims a competing
   (1, N). D08 is ALWAYS a 2-page form: its continuation is force-
   attached (marker first; next-non-main fallback for canonical order).

5. Duplicate consecutive mains (same D-code, near-identical hash, e.g.
   an accidental rescan) merge: the second page demotes to main_cont.

Voucher numbering
=================
The top-right "Voucher N°" cell is OCR'd under four Tesseract layout
modes (PSM 6/7/8/11, digit whitelists); multi-PSM agreement handles
handwriting. Candidates outside the filename's range are discarded and
sequential numbering (anchored to range_start, advanced once per main)
takes over. A main-count/filename mismatch emits a warning.
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

# Per form: footer_pat (the printed D-code, e.g. "(D 09)" -- highest
# confidence, matched first) and body_pat (title/heading fallback, used only
# when no footer code matched). Categories: "main" starts a voucher group,
# "main_or_cont*" resolved by content, "supporting" -> _3, "fahrtenbuch"
# -> attached to a D12, "skip" -> dropped. User rules baked in: D09
# extension lists and all TF5 forms (D17-D20) are ALWAYS supporting.
#
# The digit classes absorb OCR confusions (0->O/Q, 1->l/I, 2->Z). BOTH
# digits of every code are REQUIRED: an optional tens digit caused
# cross-code collisions on degraded OCR ("(D2)" hit D02 "skip" and dropped
# D12/D22 pages; "(D7)"/"(D0)"/"(D1)" promoted supporting forms to mains).
# A footer missing a digit now falls through to the body fingerprint,
# worst case defaulting to "receipt" (kept) rather than dropped/misrouted.
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
    # Both pages print (D08); page-1 vs continuation is resolved by CONTENT
    # (D08_PAGE1_PAT / D08_PAGE2_PAT in _resolve_category). The body pat
    # excludes the D09 extension-list title, which would otherwise collide.
    "D08": (
        re.compile(rf"\(\s*d\s*{_D0}8\s*\)|d\s*{_D0}8[\s\-]+activity|"
                   rf"activity\s*\(\s*d\s*{_D0}8\s*\)", re.I),
        re.compile(r"(?<!to )activity[\s\-]?voucher", re.I),
        "main_or_cont_d08",
    ),
    # --- Extension list (always supporting) ----------------------------
    # Footer often OCRs with doubled zero-like chars ("(DO09)") -- allow up
    # to two. The body title is matched first in identify_form() (Pass 0).
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

# D08 page 1 = activity header + budget tables A/B; page 2 = tables C/D/E.
# A D08 page is a continuation only when it shows page-2 headers AND lacks
# all page-1 markers; otherwise it is treated as a page-1 main.
D08_PAGE1_PAT = re.compile(
    r"theme\s*/?\s*titel|period\s+of\s+time|numb\w*\.?\s*o\.?\s*particip|"
    r"\b5\.\s*activit|a\.\s*transport|b\.\s*accommod|objectiv",
    re.I,
)
D08_PAGE2_PAT = re.compile(
    r"c\.\s*meals|d\.\s*fees|e\.\s*miscell|total\s+c\.\s*meals",
    re.I,
)

# D08 PAGE-1 HEADER markers for identify_form Pass 1b (footer-less
# fallback). These appear ONLY on the activity-voucher header, never on the
# D09 list (deliberately not "A. Transport"/"C. Meals" -- shared with D09 --
# nor "Objectives", which leaks onto activity reports).
D08_ACTIVITY_HEADER_PAT = re.compile(
    r"theme\s*/?\s*titel|topic\s+of\s+the\s+activ|\b5\.\s*activit",
    re.I,
)

# The D09 title "Extension-List to Activity-voucher" contains the D08 body
# fingerprint, so identify_form() matches it first (Pass 0) and routes the
# page to D09 supporting. Tolerates the "l" dropout ("extension-ist").
EXTENSION_LIST_PAT = re.compile(
    r"extension[\s\-]?l?ist\s+to\s+activ|"
    r"extension[\s\-]?l?ist\s*\(\s*d",
    re.I,
)

# "Voucher N°" cell value (top-right of forms). Non-digit look-arounds
# REJECT 5+ digit noisy runs rather than truncating to an in-range prefix.
VOUCHER_NO_PAT = re.compile(r"voucher\s*n[°o\*\.\s]*\W{0,3}(?<!\d)(\d{2,4})(?!\d)", re.I)

# Footer page markers, slash form ONLY ("1/2", "2 / 2"). The "n of m" /
# "n von m" word forms were dropped: they match receipt prose ("3 of 5
# copies") and folded receipts into the main sheet; genuine word-form
# continuations are still recovered by the D08 always-2-page fallback.
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
    # Telecom prepaid scratch cards -- third-party purchase proof, so _2.
    r"prepaid|scratch[\s\-]?card|"
    r"telemor|telkomcel|timor[\s\-]?telecom|"
    r"sosa\s+pakote|fasil\s+liu|ransu|folin|"
    r"konsulta\s+saldu|cek\s+saldo|check\s+(?:remain\s+)?credit|"
    r"loron\s+\d+\s*(?:ba\s+oin|days?)|"
    r"expired\s+in\s+\d+\s*days?|"
    r"customer\s+supporte|apolu\s+ba\s+kliente)\b",
    re.I,
)

# STRONG receipt phrases: a SINGLE hit keeps the page in _2. Rescues real
# receipts whose photo score is inflated by stamps/seals/logos but whose
# garbled OCR yields fewer than two ordinary keywords. None occur on the
# photos / ID cards / letters we route to supporting.
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

# Official letters / declarations ("Karta Deklarasaun", authorization
# letters, "surat keterangan") are internal supporting docs. The vocabulary
# is correspondence-specific and absent from purchase receipts, so a SINGLE
# hit routes to supporting (generic words like "authorization" excluded).
OFFICIAL_LETTER_PAT = re.compile(
    r"deklarasaun|deklara\s+katak|karta\s+deklara|"
    r"letter\s+of\s+declaration|carta\s+de\s+declara|declaration\s+letter|"
    r"surat\s+(?:keterangan|pernyataan|deklara)|"
    r"autorizasaun|authorizasaun",
    re.I,
)

# Contracts / formal agreements are supporting (_3). Their payment clauses
# legitimately contain receipt vocabulary ("upon submission of a detailed
# invoice", totals), which used to pull them into _2 -- so these are TITLE /
# legal-boilerplate phrases that never occur on purchase receipts; a SINGLE
# hit routes to supporting. Bare "contract" / "contract no." deliberately
# NOT matched (utility invoices print "Contract No.", consultant invoices
# say "contracted total 20 days").
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

# Receipt TITLE phrases -- keep a GENUINE receipt in _2 even when it
# references the contract it bills against. Narrower than STRONG_RECEIPT_PAT:
# bare "invoice" is excluded because contract payment clauses mention
# submitting one; only a receipt's own heading/stamp qualifies.
RECEIPT_TITLE_PAT = re.compile(
    r"payment\s+receipt|payment\s+received|cash\s+receipt|official\s+receipt|"
    r"tax\s+invoice|\bfaktur\b|kwitansi|nota\s+no",
    re.I,
)


def looks_like_contract(text: str) -> bool:
    """Contract / formal-agreement page (-> supporting, _3).

    One CONTRACT_PAT hit decides -- NOT weighed against ordinary receipt
    keywords (payment clauses contain them legitimately). Only veto:
    RECEIPT_TITLE_PAT, i.e. an actual receipt referencing the contract.
    """
    if not text:
        return False
    norm = text.lower()
    if not CONTRACT_PAT.search(norm):
        return False
    return not RECEIPT_TITLE_PAT.search(norm)

# Participant / attendance lists are SUPPORTING but OCR to little more than
# a letterhead + handwritten rows, so looks_supporting() often misses them.
# Detected two ways (see looks_like_participant_list):
#  (a) TITLE -- single-hit, multi-lingual. The bare phrase "participants
#      list" is EXCLUDED: the D08 form itself prints "attach a signed
#      participants list" and could be mislabelled when degraded.
PARTICIPANT_LIST_PAT = re.compile(
    r"list\s+of\s+participant|lista\s+partisipante|"
    r"lista\s+de\s+participantes|lista\s+(?:de\s+)?presen[cç]|"
    r"attendance\s+(?:list|sheet)|presence\s+list|"
    r"daftar\s+hadir|teilnehmerliste|anwesenheitsliste",
    re.I,
)

# Activity reports: the TITLE page classifies as supporting via the
# multi-hit heuristic, but its continuations are plain narrative prose with
# no keywords and used to default to receipt. Pass 1d propagates supporting
# through them, gated on >= REPORT_PROSE_MIN_TOKENS legible words AND zero
# receipt keywords -- a garbled receipt (~20 legible words) or a real one
# (>= 1 keyword) fails the gate, stays in _2 and ends the block.
REPORT_TITLE_PAT = re.compile(r"\bactivity\s+report\b", re.I)
REPORT_PROSE_MIN_TOKENS = 50

#  (b) STRUCTURE -- >= 2 DISTINCT column-header words below (Tetum /
#      Portuguese / English name, position, signature headers). They are
#      list-specific and rare on receipts, so requiring two keeps false
#      positives off invoices.
PLIST_COL_PAT = re.compile(
    r"\b(?:naran|pozisaun|instituisaun|organizasaun|kontaktu|asinatura|"
    r"assinatura|tanda\s+tangan|nome\s+completo|estrutura|"
    r"presen[cç]a|semester|curso|"
    r"signature|organisation|institution)\b",
    re.I,
)


def looks_like_participant_list(text: str) -> bool:
    """Participant / attendance / signature list page.

    TITLE match (single hit) OR >= 2 distinct column-header words. The
    structural branch is GATED by `not looks_like_receipt`: receipts also
    print "Naran / Name" and "Asinatura / signature" fields, and their
    invoice/total keywords must win and keep them in _2.
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

# Text-based Fahrtenbuch backup for pages the perceptual hash misses.
# Tier 1 (strong): German-only "Fahrtenbuch" / "lfd. Seite" -- decisive.
# Tier 2 (consensus): >= 2 of the column/header markers below; singly they
# occur on other forms (D17 "project country:", D12 "licence plate"), the
# combination only on a driver's log.
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
    """OCR text reads like a Fahrtenbuch page (tiers above). Empirically
    zero false positives on all other forms in the test bundles."""
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
    if platform.system() == "Windows":
        extras = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    else:
        # GUI-launched shells on macOS often miss /opt/homebrew/bin in PATH.
        extras = [Path(p) / "tesseract"
                  for p in ("/opt/homebrew/bin", "/usr/local/bin", "/usr/bin")]
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
    if platform.system() != "Windows":
        extras = [Path(p) / "gs"
                  for p in ("/opt/homebrew/bin", "/usr/local/bin", "/usr/bin")]
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
    args = [TESSERACT, str(png_path), "-", "-l", "eng", "--psm", str(psm)]
    if whitelist:
        args.extend(["-c", f"tessedit_char_whitelist={whitelist}"])
    out = subprocess.run(args, capture_output=True, timeout=180)
    # Non-zero exit yields empty stdout, indistinguishable from a blank
    # page and silently disabling every detector -- surface a one-line
    # warning but do NOT raise (one bad page must not abort the file).
    if out.returncode != 0:
        err = (out.stderr or b"").decode("utf-8", errors="replace").strip()
        last = err.splitlines()[-1] if err else f"exit code {out.returncode}"
        print(f"  !! tesseract failed on {png_path.name}: {last}", file=sys.stderr)
    # UTF-8 with replacement: Windows (cp1252) must never raise on
    # non-ASCII glyphs from scanned receipts.
    raw = out.stdout or b""
    return raw.decode("utf-8", errors="replace")


def ocr_page(page: fitz.Page, dpi: int = OCR_DPI, psm: int = 6,
             clip: Optional[fitz.Rect] = None,
             whitelist: Optional[str] = None) -> str:
    pix = page.get_pixmap(dpi=dpi, clip=clip) if clip else page.get_pixmap(dpi=dpi)
    # mkstemp + immediate close: NamedTemporaryFile would hold a Windows
    # lock that blocks pix.save() from writing the path.
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
    # INTENTIONAL prefix comparison (zip stops at the shorter string):
    # FAHRTENBUCH_HASH is 272 bits (16x17 template) vs page_phash's 256
    # (16x16); the check relies on the aligned top 16 rows. A landscape
    # driver's log with unreadable text is caught by THIS hash alone --
    # do NOT "guard" unequal lengths away or those pages are dropped.
    return sum(x != y for x, y in zip(a, b))


# ---------------------------------------------------------------------------
# Per-page classification + voucher-number extraction
# ---------------------------------------------------------------------------

def identify_form(text: str) -> tuple[Optional[str], str]:
    """Return (form_code, category) for an OCR'd page.

    Match order (most -> least confident):
      0. Extension-List title override (contains "Activity-voucher" and
         would otherwise be captured by the D08 body fingerprint).
      1. Footer D-code regex (highest confidence; always wins).
      1b. D08 header-content fallback, only when no footer code matched.
      2. Body keyword fallback.
    """
    norm = " ".join((text or "").lower().split())
    # Pass 0 -- extension-list title/footer override (D09 supporting).
    if EXTENSION_LIST_PAT.search(norm):
        return "D09", "supporting"
    # Pass 1 -- footer D-code (preferred signal, always wins when present).
    for code, (footer_re, _body_re, cat) in FORM_PATTERNS.items():
        if footer_re.search(norm):
            return code, _resolve_category(cat, norm)
    # Pass 1b -- D08 CONTENT fallback, only when the footer gave no D-code.
    #   The newer D08 template prints its FILE PATH in the footer: Pass 1
    #   then finds no "(D08)", and the path's literal "extension list"
    #   would let the D09 body pat grab the page in Pass 2 whenever the
    #   title OCRs poorly. So match the activity-voucher HEADER fields
    #   instead -- present on D08 page 1, never on the D09 list. Page 2
    #   is recovered by the always-2-page post-pass.
    if D08_ACTIVITY_HEADER_PAT.search(norm):
        return "D08", "main"
    # Pass 2 -- body keyword fallback.
    for code, (_footer_re, body_re, cat) in FORM_PATTERNS.items():
        if body_re.search(norm):
            return code, _resolve_category(cat, norm)
    # No D-code: default receipt, but route internal supporting documents
    # (official letters, contracts, participant lists by title or column
    # structure, agenda/itinerary keyword consensus) to supporting first.
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
        # Continuation only with page-2 tables AND no page-1 markers;
        # anything else starts a new voucher as a page-1 main.
        if D08_PAGE2_PAT.search(norm) and not D08_PAGE1_PAT.search(norm):
            return "main_cont"
        return "main"
    return cat


def extract_voucher_number(page: fitz.Page,
                           expected_range: Optional[tuple[int, int]] = None
                           ) -> Optional[int]:
    """OCR the top-right of *page*; return the 'Voucher N°' value or None.

    Numbers are often handwritten, so a narrow high-DPI crop is OCR'd
    under four layout modes: PSM 6 prose (catches printed forms) plus
    PSM 7/8/11 with a digit-only whitelist (drops the S/$/O glyphs
    handwriting is misread as). All plausible 2-4 digit candidates are
    collected; in-range candidates (per the filename's range) win by
    multi-PSM majority. A tie or an out-of-range-only read returns None
    so the caller's sequential fallback takes over.
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
    """OCR the bottom strip of *page*; return (k, n) for a 'k/n' marker.

    Continuation pages often lose their D-code title to table overlap;
    the marker is then the only signal gluing them to the same _1 file.
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
    """Too few legible words to classify on (< 20 alphabetic tokens).
    Receipts can also fall below this -- callers must combine it with
    another signal before re-categorising the page.
    """
    tokens = re.findall(r"[A-Za-z]{3,}", text or "")
    return len(tokens) < 20


# ---------------------------------------------------------------------------
# Photograph detection (largest contiguous continuous-tone region)
# ---------------------------------------------------------------------------
#
# These scans embed MANY image XObjects per page (6-49 in real bundles),
# so measuring "the" extracted image is unreliable. Instead RENDER the page
# (MuPDF is bundled identically in every PyMuPDF wheel -> a fixed-DPI
# render is deterministic across OSes) and measure the single largest
# connected blob of continuous-tone (mid-gray) blocks: large for photos
# and ID cards, ~zero for bimodal text documents, small and SCATTERED for
# receipt logos/halftone. Halftone-heavy printed receipts also score high;
# the looks_like_receipt() text gate at the call site keeps them in _2.
# Block-averaging, the wide tone band and a ~2x threshold margin absorb
# sub-pixel rasteriser differences between platform builds.

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
    """Fraction of the page covered by the largest contiguous
    continuous-tone region -- high for photographs, ~0 for text.
    Deterministic across OSes (fixed-DPI render, LUT mask, BOX
    downsample, pure-Python connected components).
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
    """Text contains unambiguous receipt keywords (gate for the photo flip).

    POSITIVE multi-lingual matching -- stable across Tesseract versions,
    unlike token counting. Tiered: one STRONG phrase (receipt heading,
    telecom brand) suffices; otherwise two ordinary keywords. Keeps
    colourful printed receipts (scratch cards, stamped receipts) in _2.
    """
    if not text:
        return False
    norm = text.lower()
    if STRONG_RECEIPT_PAT.search(norm):
        return True
    return len(RECEIPT_PAT.findall(norm)) >= 2


def looks_supporting(text: str) -> bool:
    """Multi-hit heuristic for no-D-code pages: >= 2 supporting signals
    (a time slot counts double) AND more of them than receipt hits -- an
    invoice mentioning "presentation" once still classifies as receipt.
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
    #   a) Fahrtenbuch: hash first (cheap, no OCR), then text fallback.
    #   b) full-page OCR -> identify_form (D-code / body keywords).
    #   c) bottom-strip OCR -> "k/N" page marker.
    #   d) photo flip: a receipt-default page that renders photo-like
    #      flips to supporting unless strong receipt keywords veto it.
    #   e) backward-pull: a (k>=2, N) page re-attaches to the most recent
    #      main -- only if that main advertised its own (1, N) and no
    #      intervening page claims a competing (1, N).
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
            # (a, fallback) Fahrtenbuch via text -- rescues pages whose
            # template drift defeats the hash; see is_fahrtenbuch_text().
            if is_fahrtenbuch_text(text):
                info = {"page": i, "form": "FAHRTENBUCH", "cat": "fahrtenbuch",
                        "vnum": None, "marker": None, "marker_tuple": None,
                        "photo_score": None, "phash": ph}
            else:
                form, cat = identify_form(text)
                marker_tuple = extract_page_marker(page)

                # (d) photo flip, gated by the receipt-keyword veto (see
                #     photo_region_score / looks_like_receipt).
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

    # 1b. Post-pass: D08 is a 2-page form by template, but the page-2
    #     footer often garbles and lands as supporting/receipt. Force the
    #     continuation: marker-based first (handles scrambled bundles),
    #     else the immediately-next non-main page (canonical order).
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

    # 1c. Post-pass: duplicate consecutive mains (accidental rescans).
    #     Same form code + hamming < 35/256 (lenient enough for signature/
    #     stamp differences, tight enough not to merge two real vouchers)
    #     -> demote the second to main_cont.
    for i in range(len(page_info) - 1):
        a, b = page_info[i], page_info[i + 1]
        if (a["cat"] == "main" and b["cat"] == "main"
                and a.get("form") and a.get("form") == b.get("form")):
            ph_a, ph_b = a.get("phash"), b.get("phash")
            if ph_a and ph_b and hamming(ph_a, ph_b) < 35:
                b["cat"] = "main_cont"
                print(f"  -> p{b['page']+1:02d} forced to main_cont "
                      f"(duplicate scan of p{a['page']+1:02d})")

    # 1d. Post-pass: participant-list / activity-report continuation
    #     propagation. Both document types are detectable on their first
    #     page only; continuations OCR to handwriting / plain prose and
    #     would default to receipt. Once a list/report page is seen, carry
    #     "supporting" forward through following receipt-default pages.
    #     Bounded so it never swallows a real receipt: a LIST block skips
    #     pages with strong receipt keywords; a REPORT block additionally
    #     requires narrative prose with ZERO receipt keywords
    #     (report_prose). Any main/fahrtenbuch/skip boundary or other
    #     supporting page (photo, itinerary) resets the block.
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

    # 2. group pages into voucher buckets (one per main). In-range OCR
    #    numbers are trusted; otherwise sequential numbering anchored to
    #    range_start (the filename is ground truth for a contiguous bundle).
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
            # Advance exactly once per main, NEVER re-anchored to the OCR
            # read -- a wrong read then mislabels only its own page.
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

    # 3b. Duplicate voucher numbers (a wrong-but-in-range OCR read) would
    #     silently overwrite output files (gs -sOutputFile truncates).
    #     On any collision renumber every group sequentially from
    #     range_start so each output is unique.
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
        # A receipt glued onto the voucher sheet (scan user-error) is not
        # detectable on the page itself; the structural signal -- a main
        # with no separate _2 receipt -- prompts the operator to check.
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
        pdfs = []
        for p in args.pdfs:
            if p.is_file():
                pdfs.append(p)
            else:
                print(f"!! skipping missing file: {p}", file=sys.stderr)
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
