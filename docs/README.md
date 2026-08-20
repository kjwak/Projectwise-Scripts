# Documentation index

Project documentation for the Prepend PDF QC pipeline. For a quick repo overview, see the root [`README.md`](../README.md). For agent/AI conventions, see [`AGENTS.md`](../AGENTS.md).

**Module and script inventories:** [`modules/README.md`](../modules/README.md), [`scripts/README.md`](../scripts/README.md).

---

## Reference (operators & config)

| Document | Description |
|----------|-------------|
| [`reference/appsettings-reference.md`](reference/appsettings-reference.md) | Full `appsettings.json` section guide |
| [`reference/testing-config.md`](reference/testing-config.md) | Test profile, merge chain, `appsettings.test.json` |
| [`reference/dry-run-policy.md`](reference/dry-run-policy.md) | Layered dry-run flags and effective side-effect policy |
| [`reference/module-contracts.md`](reference/module-contracts.md) | Result envelope, error codes, module conventions |
| [`reference/project_id_from_workarea.md`](reference/project_id_from_workarea.md) | Project ID from PW work area properties |
| [`reference/pw-extractable-from-projectwise.json`](reference/pw-extractable-from-projectwise.json) | PW extractable field reference (JSON) |

---

## Architecture

| Document | Description |
|----------|-------------|
| [`architecture/hybrid-polling.md`](architecture/hybrid-polling.md) | **Current:** audit-driven watcher, reconciliation, watermarks |
| [`architecture/qc-package-model.md`](architecture/qc-package-model.md) | SQL sheet packages and lane QC PDF registry |
| [`architecture/audit-trail-architecture.md`](architecture/audit-trail-architecture.md) | Original audit-trail design (historical + as-built notes) |
| [`architecture/watcher-architecture-refactor-summary.md`](architecture/watcher-architecture-refactor-summary.md) | Watcher refactor summary (historical) |

---

## Workflow & notifications

| Document | Description |
|----------|-------------|
| [`workflow/qc-workflow-framework.md`](workflow/qc-workflow-framework.md) | PW workflow/state/attribute writeback |
| [`workflow/qc-comment-status-sync.md`](workflow/qc-comment-status-sync.md) | Comment-driven `QC_COMMENT_STATUS_SYNC` |
| [`workflow/qc-notifications.md`](workflow/qc-notifications.md) | Email notifications (Mock + Microsoft Graph) |
| [`workflow/qc-reporting.md`](workflow/qc-reporting.md) | `QC_REPORTING_SCAN` and reporting views |
| [`workflow/pw-environment-email-attributes.md`](workflow/pw-environment-email-attributes.md) | Designer/reviewer/checker email attributes |

---

## Data & telemetry

| Document | Description |
|----------|-------------|
| [`data/database-telemetry.md`](data/database-telemetry.md) | SQL schema, fire-and-forget writes, `sheet_index`, execution invariants |

---

## Engineering & audits

| Document | Description |
|----------|-------------|
| [`engineering/dead-code-and-deprecation-audit.md`](engineering/dead-code-and-deprecation-audit.md) | Dead code, legacy bridges, cleanup priorities |
| [`engineering/phase-2-native-prepend-parity-plan.md`](engineering/phase-2-native-prepend-parity-plan.md) | Legacy vs native prepend migration plan |
| [`engineering/dead-code-phase-1-patch-summary.md`](engineering/dead-code-phase-1-patch-summary.md) | Phase 1 dead-code patch log |
| [`engineering/scripts-organization-review.md`](engineering/scripts-organization-review.md) | Script organization review (historical) |
| [`engineering/dashboard-upgrade-options.md`](engineering/dashboard-upgrade-options.md) | Terminal dashboard improvements |
| [`engineering/status-set-av-refactor.md`](engineering/status-set-av-refactor.md) | Status set AV churn notes |
| [`engineering/remote-worker-architecture-intent.md`](engineering/remote-worker-architecture-intent.md) | **Intent:** remote QC workers on modelling PC while server coordinates queue |
| [`engineering/remote-worker-host-guardrails.md`](engineering/remote-worker-host-guardrails.md) | **In force:** modelling-PC clone is a processor host only; no watcher/dashboard; no UNC claims until host-aware locks are on the server |

---

## Archive — phase migration summaries

> **Historical only.** Files under `docs/archive/` may describe older paths, states, or architecture. Do **not** treat them as current implementation guidance unless another **current** document (outside `archive/`) explicitly links to them as authoritative. For day-to-day operations, use **Reference**, **Architecture**, **Workflow**, **Data**, and **Engineering** sections above.

Historical records of Phase 2–4 refactors:

| Document | Description |
|----------|-------------|
| [`archive/phase/phase-2-docs-refresh-summary.md`](archive/phase/phase-2-docs-refresh-summary.md) | TYPSA three-lane doc refresh |
| [`archive/phase/phase-2-config-loader-consolidation-summary.md`](archive/phase/phase-2-config-loader-consolidation-summary.md) | Config loader consolidation |
| [`archive/phase/phase-2-package-model-decision.md`](archive/phase/phase-2-package-model-decision.md) | In-memory vs SQL package model decision |
| [`archive/phase/phase-3-package-model-archive-summary.md`](archive/phase/phase-3-package-model-archive-summary.md) | Package model v1 archive |
| [`archive/phase/phase-3-jobfactory-package-dedupe-decision.md`](archive/phase/phase-3-jobfactory-package-dedupe-decision.md) | JobFactory package dedupe removal |
| [`archive/phase/phase-4-module-script-organization-plan.md`](archive/phase/phase-4-module-script-organization-plan.md) | Phase 4 module/script organization master plan |
| [`archive/phase/phase-4-server-path-audit.md`](archive/phase/phase-4-server-path-audit.md) | Server path audit & worker checklist |
| Other `archive/phase/phase-4-*.md` | Per-phase move/bootstrap summaries |

---

## Folder layout

```
docs/
  README.md           ← this index
  reference/          ← config, testing, conventions
  architecture/       ← polling, packages, design history
  workflow/           ← QC workflow, notifications, comments
  data/               ← SQL telemetry
  engineering/        ← audits, active plans, tooling notes
  archive/phase/      ← historical migration summaries (not current guidance)
```
