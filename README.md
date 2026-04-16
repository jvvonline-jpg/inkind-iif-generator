# Cornerstones In-Kind JE → IIF Generator

A small app that reads the quarterly in-kind gift Excel workbooks and produces
a QuickBooks-ready `.iif` journal entry.

## Files

| File | Purpose |
| --- | --- |
| `in_kind_iif_generator.py` | The Python program. Runs as a GUI or CLI. |
| `Generate In-Kind IIF.command` | macOS double-click launcher — opens the GUI. |
| `Sample 1 Folder/` | Q2 FY26 example inputs (All Gifts, ERCS split, per-program sheets) plus the original `In-Kind Q2 IIF.iif` and the generated version `In-Kind Q2 IIF (Generated).iif` for comparison. |

## How to use

1. Put every quarter's gift workbooks in one folder. At minimum the folder must contain:
   - `All Gifts Q# FY## REV##.xlsx` — master list of every gift (by Fund Description)
   - `ERCS Gifts Q# FY## REV##.xlsx` — with two sheets: one starting `351…` (Shelter), one starting `381…` (HH)
2. Double-click **Generate In-Kind IIF.command**. A small window opens.
3. Point it at the folder, confirm posting date and quarter label, click **Generate IIF**.
4. Import the resulting `.iif` into QuickBooks: *File → Utilities → Import → IIF Files*.

## The mapping the app applies

All program totals come from the **All Gifts** workbook, grouped by `Fund Description`.

| Fund Description | Class | JE line label |
| --- | --- | --- |
| Assistance Services & Pantry Program | 41000 | ASAPP |
| Community Connected Sites | 43000 | Community Connected Sites (omitted when $0) |
| Emergency Support Housing Program | 51000 | ESHP (HOUSE) |
| Food Hub | 81000 | Food Hub |
| Hypothermia | 32200 | Hypo - TOS |
| General Cornerstones Fund | 21000 | General CS |
| Herndon NRC … Adult Services | 42000 | HNRC -OAS |
| Laurel Learning Center | 61000 | LLC |
| **Embry Rucker Community Shelter** | — | split 3 ways, see below |

**ERCS split** — the ERCS workbook must have two sheets whose names start with:

- `351…` → ERCS Shelter, class `32100` (full total)
- `381…` → HH total, split **50/50** into `32302` (HH w/ Children) and `32303` (HH w/o Children)

Every JE line produces two IIF rows: a debit to
`DIRECT PROGRAM EXPENSES:IN-KIND Program Expense:In Kind - Programs` (9026.1)
and an equal credit to
`In-Kind Contributions/Donations:In Kind Goods & Prof. Services` (8111),
both tagged with the full QuickBooks class path and a memo of the form
`"Q# FY## in-kind program exp/rev per Q# summary - {program}"`.

## Running from the command line

```bash
python3 in_kind_iif_generator.py \
    --folder "Sample 1 Folder" \
    --posting-date 12/31/2025 \
    --quarter "Q2 FY26" \
    --out "Sample 1 Folder/In-Kind Q2 IIF (Generated).iif"
```

## A note on the Q2 FY26 sample

The app reproduces the template/PDF totals (851,467.06). The original
`In-Kind Q2 IIF.iif` shipped with a smaller ERCS figure (79,882.67 shelter,
33,133.255 each HH) that appears to have been a one-off manual adjustment
(those reduced values are hardcoded in the template's column K with no
formula tying them to anything). The app outputs the gross amounts that
match the PDF general-journal report.
