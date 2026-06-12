# Voucher Splitter

Splits scanned **AGIAMONDO quarterly voucher bundles** into one tidy PDF per
voucher, sorted into three categories:

| Output suffix | Contents |
|---------------|----------|
| `_1` | the main voucher form (Standard / Collective / Travel / Activity / Reimbursement-private-km) |
| `_2` | the original third-party receipts, invoices and slips |
| `_3` | supporting documents (driver's logs, extension lists, participant lists, contracts and offer letters, ID cards, etc.) |

You drop a bundle named like `Q126_5_Vouchers_No_238-241.pdf` into the working
folder, run one command, and the matching `Q126_6_Voucher_238_1.pdf`,
`…_2.pdf`, `…_3.pdf` files appear next to it. The `Q<q><yy>` quarter/year
prefix is read from the input filename and reused on every output, so it works
for any quarter and year.

> **A full, illustrated walkthrough is in [`Tutorial.pdf`](Tutorial.pdf).**
> Non-technical users should follow that guide — this README is the short
> version for quick reference.

---

## What it does (under the hood)

- **OCR** (Tesseract) reads each page's printed footer D-code (D07, D08, …) and
  body text to classify it.
- **Multi-page main forms** (e.g. the 2-page D08 Activity voucher) are kept
  together; continuation pages are folded into the same `_1` file.
- **Photographs and ID cards** are detected by rendering the page and measuring
  the largest contiguous continuous-tone region, then routed to `_3` —
  independent of OCR, so it behaves identically on macOS, Linux and Windows.
- **Receipts** (including telecom prepaid scratch cards) are kept in `_2` via a
  multi-lingual keyword gate.
- **Driver's logs** (Fahrtenbuch) are attached to the matching D12
  reimbursement voucher.
- Each output is compressed with Ghostscript so the files email easily.

## Requirements (one-time install)

- **Python 3.9+**
- **Tesseract OCR** and **Ghostscript** (system binaries, not pip packages)
- Python packages: `pymupdf`, `Pillow`

**macOS**

```bash
brew install python tesseract ghostscript
python3 -m pip install --user --break-system-packages pymupdf Pillow
```

**Windows** — install Python (tick *Add python.exe to PATH*), Tesseract OCR
(UB-Mannheim build) and Ghostscript (AGPL release) from their official
installers, then:

```bat
py -m pip install pymupdf Pillow
```

## Usage

Put `split_vouchers.py` (and, on Windows, `split_vouchers.bat`) in a folder on
your Desktop named exactly `automation`, drop the scan bundles in the same
folder, then:

```bash
# macOS / Linux
cd ~/Desktop/automation
python3 split_vouchers.py

# Windows: double-click split_vouchers.bat
```

Useful flags:

```bash
python3 split_vouchers.py --dry-run        # classify & preview, write nothing
python3 split_vouchers.py FILE.pdf ...      # process specific files
python3 split_vouchers.py --folder PATH     # use a different working folder
```

## Privacy

**Never commit scanned vouchers, receipts or their outputs to this
repository.** They contain personal and financial data. The included
[`.gitignore`](.gitignore) excludes every `*.pdf` (except this guide's
`Tutorial.pdf`) and the `test/` folder by default — keep it that way.
