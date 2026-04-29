# Module Contracts and Conventions

## Standard Result Object Schema
All public module functions should return a result object with this shape:

- `IsSuccess` `[bool]`: Indicates success/failure of the function call.
- `Code` `[string]`: Machine-readable outcome code (success, validation, or error category).
- `Message` `[string]`: Human-readable summary message.
- `Data` `[object]`: Optional payload with function-specific details.

Helper constructors are defined in `modules/Core.Results.psm1`:
- `New-QCResult`
- `New-QCSuccessResult`
- `New-QCFailureResult`

## Error Code Naming Convention
Recommended pattern:

`<AREA>_<CATEGORY>_<DETAIL>`

Examples:
- `CONFIG_VALIDATION_MISSING_KEY`
- `QUEUE_DUPLICATE_DETECTED`
- `PW_SESSION_UNAVAILABLE`
- `TRIGGER_NO_MATCH`

Guidelines:
- Use uppercase snake case.
- Keep stable codes for dashboarding and alert rules.
- Prefer specific, deterministic codes over generic ones.

## Public/Private Function Naming
- Public function naming: approved Verb-Noun cmdlet style.
- Internal/private helper naming recommendation:
  - Prefix private helpers with underscore (for example: `_Resolve-InternalRule`).
  - Avoid exporting private helpers.

## Module Export Policy Recommendation
Current scaffolding exports all functions using:

`Export-ModuleMember -Function *`

Recommendation for implementation phase:
- Keep broad export during rapid prototyping.
- Move to curated exports once public/private split is clear.
- Add module-level tests to enforce export boundaries.

## ProjectWise Safety Rule
- No ProjectWise write operations are allowed without explicit approval.
- Read-only discovery/session checks are permitted.
- Any future write-capable function must be reviewed separately and clearly documented.
