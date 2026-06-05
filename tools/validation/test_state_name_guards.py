#!/usr/bin/env python3
"""Static guard check for variable -StateName call sites in PowerShell files.

The audit watcher must never pass an empty string into state-mutating calls. This
check scans PowerShell source for variable-based -StateName arguments and requires
nearby guard evidence for state-write calls. Predicate/helper calls that perform
their own normalization are allow-listed explicitly.
"""
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[2]
PATHS = [ROOT / 'modules', ROOT / 'scripts']
CALL_RE = re.compile(r'(?P<cmd>[A-Za-z0-9_\-.]+)\s+.*-StateName\s+(?P<arg>\$[A-Za-z_][\w.]*)')
GUARD_TOKENS = (
    'IsNullOrWhiteSpace',
    '_PWD-TestStateNameNotEmpty',
    'WATCH_AUDIT_EMPTY_STATE_GUARDED',
    'QC_WORKFLOW_STATE_MISSING',
    'QC_WORKFLOW_STATE_EMPTY_SKIPPED',
    'PACKAGE_STATE_EMPTY_TARGET',
)
ALLOWLIST_COMMANDS = {
    # Predicate/decision helpers normalize empty state internally or are not state writes.
    'Test-QCWorkflowStateIsQcInitiated',
    'Test-QCWorkflowStateIsQcFinalizing',
    '_QCN-TestWorkflowStateNameIsQcInitiated',
    'Resolve-QCWorkflowAssignee',
}
RISKY_COMMANDS = {
    'Set-PWDocumentState',
    'Set-PWQCWorkflowState',
    '_PWD-InvokeSetPwDocumentState',
    '_RSO-SetPwDocumentState',
}

failures = []
for base in PATHS:
    for path in sorted(base.rglob('*')):
        if path.suffix.lower() not in ('.ps1', '.psm1'):
            continue
        lines = path.read_text(encoding='utf-8', errors='ignore').splitlines()
        for idx, line in enumerate(lines):
            if '-StateName' not in line:
                continue
            m = CALL_RE.search(line)
            if not m:
                continue
            cmd = m.group('cmd')
            if cmd in ALLOWLIST_COMMANDS:
                continue
            if cmd not in RISKY_COMMANDS and not cmd.lower().endswith('setpwdocumentstate'):
                continue
            start = max(0, idx - 28)
            end = min(len(lines), idx + 4)
            window = '\n'.join(lines[start:end])
            if any(tok in window for tok in GUARD_TOKENS):
                continue
            failures.append(f'{path.relative_to(ROOT)}:{idx+1}: {line.strip()}')

if failures:
    print('Missing empty StateName guard near these call sites:', file=sys.stderr)
    for failure in failures:
        print(f'  {failure}', file=sys.stderr)
    sys.exit(1)
print('StateName guard validation passed')
