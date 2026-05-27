# `docs/`

Project documentation and reference material for this repository.

If you're looking for "how the pipeline works", start at:

- repo root `README.md`
- `modules/README.md`
- `scripts/README.md`

## Architecture

- `hybrid-polling.md` - hybrid audit-trail + reconciliation polling architecture (replaces pure folder scanning).
- `audit-trail-architecture.md` - original analysis and design for transitioning from directory polling to event-driven processing.

## Database & Telemetry

- `database-telemetry.md` - SQL Server telemetry layer: schema, tables, views, fire-and-forget pattern, and `sheet_index`.
- `pw-environment-email-attributes.md` - reading designer/reviewer email attributes from ProjectWise (multi-environment).

## QC Workflow & Reporting

- `qc-workflow-framework.md` - configurable ProjectWise QC workflow/state/attribute writeback framework and rollout plan.
- `qc-notifications.md` - QC workflow email notification system (Mock + future Microsoft Graph).
- `qc-reporting.md` - attribute-first QC reporting snapshots and QC_REPORTING_SCAN architecture.

## Engineering

- `module-contracts.md` - module conventions, result object schema, error codes, export policy.
- `scripts-organization-review.md` - script refactoring plan and progress.
- `dashboard-upgrade-options.md` - terminal dashboard rendering improvements (Phase 1 complete).
- `status-set-av-refactor.md` - status set AV file-churn mitigation notes.

## ProjectWise Reference

- `project_id_from_workarea.md` - extracting Project ID from work area type properties.
