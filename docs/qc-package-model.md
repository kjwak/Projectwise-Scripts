# QC Package Model

> **Note (Phase 2):** Production package grouping uses SQL **`sheet_packages`** and **`sheet_package_qc_pdfs`** via `Core.Database.psm1`. The in-memory `QC.Package*` modules described below are **not wired into the production pipeline** (test/documentation only). See Branch 3 decision doc when available: `docs/phase-2-package-model-decision.md`.

The QC automation framework treats related ProjectWise artifacts as a single QC Package instead of making workflow and attribute decisions on a single triggering file.

A QC Package contains:

- `DgnDocument` for the source `*.dgn` design artifact.
- `PdfDocument` for the production `*.pdf`, which is the preferred source of truth for user-controlled QC metadata.
- `QcPdfDocument` for `*_QC.pdf`, `-QC.pdf`, or other configured QC history/review PDFs.
- `PackageId`, `PackageKey`, `DocumentKey`, `BaseName`, `SheetId`, and `PackageRootFolder`.
- Package-level workflow state, review type, role metadata, automation status fields, warnings, and conflicts.

## Resolution strategy

`Resolve-QCPackage` accepts an audited ProjectWise document and candidate sibling documents. The resolver classifies the triggering document, derives a stable package key, and groups siblings using extension points for explicit package attributes, GUID cache, SQL package cache, folder proximity, and configurable naming rules. Filename matching is only one strategy and is intentionally isolated behind configuration.

## Metadata authority

`Get-QCPackageCanonicalDocument` selects the canonical metadata source by configuration. The default preference is:

1. Production PDF
2. QC PDF
3. DGN

If the production PDF is missing, the fallback is deterministic and emits a warning. SQL cache rows are not treated as authoritative ProjectWise metadata; they are only relationship/job-history telemetry.

## Attribute ownership

`QC.AttributePolicy.psm1` separates user-owned fields from automation-owned fields. User fields are read by automation and are not blindly copied across all documents. Automation status writes are allowlisted and default to production PDF + QC PDF. DGN receives only lightweight package/status fields when explicitly configured.

Validation checks include missing environment columns, unavailable ProjectWise discovery cmdlets, bad emails, and invalid/missing configured attributes. Attribute writes use `Update-PWDocumentAttributes` and support dry-run planning.

## State policy

`QC.StatePolicy.psm1` resolves package-level state using configurable precedence and reports conflicting sibling states. Package state changes are idempotent and use ProjectWise state APIs (`Set-PWDocumentState`) when mutation is not dry-run.

## SQL/cache contract

`QC.Package.Database.psm1` defines additive cache fields:

- `PackageId`
- `DgnGuid`
- `PdfGuid`
- `QcPdfGuid`
- `PackageKey`
- `LastResolvedUtc`
- `LastCanonicalGuid`
- `LastKnownState`
- conflict flags

Manual ProjectWise/admin setup still required:

1. Ensure target ProjectWise environments contain configured user-owned and automation-owned attributes.
2. Ensure configured workflow state names exactly match ProjectWise workflow state labels.
3. Grant the automation account permission to update allowlisted automation fields and perform allowed workflow state changes.
4. Create the additive package-cache table or migration equivalent if SQL relationship caching is enabled.
5. Review naming suffixes and any future explicit GUID/package-link relationship attributes per datasource.
