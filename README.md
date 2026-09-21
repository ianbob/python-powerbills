# KPLC Power Bill Renamer

Finds the first table in each KPLC bill that contains an account number, reads
the account number with OCR, looks it up in the `august` sheet of the Excel
workbook, and renames the PDF to `<station-name>.pdf`. If the account number is
not in the workbook, it uses `<account-number>.pdf`.

The supplied PDFs are image-only scans, so the project uses Tesseract OCR.

## Setup

Install the system OCR and Python dependencies:

```bash
sudo apt install tesseract-ocr
python -m pip install -e .
```

The first command installs the Tesseract OCR engine used to read scanned
account numbers. The second installs this project and its Python dependencies
in editable mode, so changes to the source are reflected immediately.

## Usage

Preview planned renames without changing files:

```bash
python -m powerbills
```

Apply the renames:

```bash
python -m powerbills --apply
```

The command always uses `raw_files` when no input is supplied. The `raw-2`
directory is a backup/test directory and is never scanned automatically. To
run a manual test against it, pass it explicitly: `python -m powerbills raw-2`.

By default the command reads `excel/Power Bills Template.xlsx` and its
`August` sheet. The requested sheet name can be overridden with
`--lookup-sheet` and `--output-sheet`; if the workbook has only one sheet, that
sheet is used automatically when the requested name is absent.

You can also pass individual PDFs or a glob. Existing destination files are
never overwritten; use `--on-conflict skip` (the default) or `--on-conflict
error` to stop at the first conflict.