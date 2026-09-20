---
name: invoice-organizer-skill
description: Non-destructively organize Chinese reimbursement invoices from an immutable input directory into renamed output copies with structured extraction, DiDi pairing, ZIP PDF handling, and one Excel summary. Use for Chinese invoice sorting and reimbursement-file organization; require explicit confirmation before creating output.
metadata:
  short-description: Safely organize Chinese invoices into output and Excel
---

# Invoice Organizer Skill

Organize invoices through a strict `input -> .runtime -> output` workflow. Input files are immutable and must never be modified, moved, renamed, overwritten, or deleted.

## Workflow

1. Ask the user to place original files in `input/`, or accept explicit `--input` and `--output` paths.
2. Run `python scripts/scan_invoices.py`. Scanning is read-only and writes only the internal plan and preview under `.runtime/`.
3. Review `.runtime/发票整理预览.md`. Surface every unresolved extraction or pairing issue.
4. Wait for the user's explicit confirmation.
5. Run `python scripts/organize_invoices.py --plan .runtime/.invoice-organization-plan.json --confirm <TOKEN>`.
6. Copy ordinary PDFs and extract selected ZIP PDF members into `output/`. Keep every input file, including ZIP archives, unchanged.
7. Create only `output/发票整理清单.xlsx` as the user-facing summary. The last row totals only explicitly extracted ticket totals.
8. Run `python scripts/validate_organization.py --plan .runtime/.invoice-organization-plan.json`.

The approval token authorizes creation of output only. It never authorizes input mutation.

## Extraction Safety

When extraction is uncertain, leave the field blank.

Never fabricate, infer, calculate, or guess financial or tax values merely to complete the spreadsheet. Do not derive one financial field from another, use filename amounts as ticket totals, or substitute a business date for an invoice date. The only financial arithmetic allowed is the final Decimal sum of nonblank `total_amount` values.

Read [references/extraction-rules.md](references/extraction-rules.md) before changing field extraction or Excel mappings.

## Naming and Ordering

Keep the established category order, business-date sorting, filename format, sequence rules, and shared DiDi invoice/report sequence. Read [references/naming-and-ordering.md](references/naming-and-ordering.md) before changing these rules.

## Internal Plan

The JSON plan, approval token, preview, state, and temporary files belong only in `.runtime/`. Read [references/manifest-schema.md](references/manifest-schema.md) before changing the internal schema.

## Output Contract

- `output/` contains final ticket files plus `发票整理清单.xlsx` only. The repository's `.gitkeep` is an inert directory marker.
- Ordinary PDFs are copied with metadata; ZIP PDFs are read from the archive and written as new output files.
- OFD and other non-PDF ZIP members never enter output.
- Existing identical ticket outputs may be skipped. Any same-name different-content file is a collision and stops execution before writes.
- An existing workbook is accepted only when the current plan records its matching SHA256.
- `validate_organization.py` is read-only.

## Boundaries

Do not edit PDF contents, make tax or reimbursement-policy judgments, recurse into unrelated directories, create category subfolders, or expose internal paths, hashes, statuses, or categories in the Excel workbook.
