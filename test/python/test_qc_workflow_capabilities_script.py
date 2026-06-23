from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SCRIPT = REPO_ROOT / "tools" / "discovery" / "Test-QCWorkflowCapabilities.ps1"


def test_capability_script_exists_and_has_expected_output_shape():
    text = SCRIPT.read_text(encoding="utf-8")
    expected_keys = [
        "pwpsLoaded",
        "pwpsDabLoaded",
        "availableCmdlets",
        "missingCmdlets",
        "candidateWriteCmdlets",
        "readCapabilities",
        "writeCapabilities",
        "warnings",
    ]
    for key in expected_keys:
        assert key in text


def test_capability_script_lists_required_candidate_cmdlets():
    text = SCRIPT.read_text(encoding="utf-8")
    candidates = [
        "Get-PWWorkflows",
        "Get-PWWorkflowStateLinks",
        "Get-PWDocumentEAttributes",
        "Get-PWEnvironmentColumns",
        "Get-PWDocumentsBySearch",
        "Get-PWDocumentsBySearchExtended",
        "Get-PWDocumentsBySearchWithReturnColumns",
        "Update-PWDocumentAttributes",
        "Set-PWDocumentState",
        "Set-PWFolderWorkflow",
        "Set-PWWorkflowByFolderPath",
        "Set-PWDocumentWorkflow",
    ]
    for candidate in candidates:
        assert candidate in text


def test_capability_script_is_read_only_for_write_candidates():
    text = SCRIPT.read_text(encoding="utf-8")
    forbidden_invocations = [
        "Set-PWDocumentState ",
        "Set-PWFolderWorkflow ",
        "Set-PWWorkflowByFolderPath ",
        "Set-PWDocumentWorkflow ",
        "Set-PWDocumentWorkflowState ",
        "Update-PWDocumentAttributes ",
        "Update-PWDocumentProperties ",
    ]
    executable_lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip() and not line.strip().startswith("#") and not line.strip().startswith("'")
    ]
    for line in executable_lines:
        for forbidden in forbidden_invocations:
            assert not line.startswith(forbidden), f"Unsafe direct write invocation found: {line}"

WRITEBACK_SCRIPT = REPO_ROOT / "tools" / "discovery" / "Test-QCWorkflowWriteback.ps1"


def test_controlled_writeback_script_requires_explicit_safety_flags():
    text = WRITEBACK_SCRIPT.read_text(encoding="utf-8")
    assert "[switch]$ConfirmWrites" in text
    assert "[string]$TestDocumentPath" in text
    assert "Refusing to run" in text
    assert "Planned operations (dry-run)" in text
    assert "-Rollback" in text or "$Rollback" in text
    assert "StateOnly" in text and "AttributeOnly" in text and "Combined" in text
