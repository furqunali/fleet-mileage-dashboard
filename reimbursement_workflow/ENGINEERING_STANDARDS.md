# Engineering Standards

## Deterministic calculations
- Validate numeric inputs before calculations.
- Reject NaN, infinity, and negative distances/rates.
- Keep business rules in reusable Python modules rather than duplicating logic in scripts.

## Data integrity
- Preserve source records and make transformations auditable.
- Normalize user-provided text before matching.
- Prefer explicit validation failures over silently accepting malformed values.

## Testing
- Add regression tests for business rules and boundary conditions.
- Keep pure calculation helpers deterministic and easy to test.

## Security
- Never commit credentials, tokens, or private customer exports.
- Keep generated/local artifacts outside version control.
