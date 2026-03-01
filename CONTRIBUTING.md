# Project Guidelines & Documentation

All contributors (including AI Agents) must follow these protocols. All logs are maintained within this file under the relevant sections.

---

## Section 1: Source Code Changes (Changelog)
Every non-trivial modification must be appended to the bottom of this section using this template:

### [vX.X.X] | YYYY-MM-DD
- **Type:** (feat | fix | refactor | docs)
- **Description:** One-sentence summary.
- **Motivation:** Why was this change necessary?
- **Files Modified:** `path/to/file`
- **Test Strategy:** How was this verified?

### v1.1.0 | 2026-03-01
- **Type:** fix
- **Description:** JD CSV expense sign parsing for `收/支` header variant.
- **Motivation:** JD export header uses `收/支` while parser only read `收支`, causing expense amounts to be inserted as positive instead of negative.
- **Files Modified:** `insert_transactions_by_date.py` (direction header fallback + `--repair-jd-legacy-sign` flag), `tests/test_helpers.py` (regression tests for both header variants).
- **Test Strategy:** Added `JdCsvSignNormalizationTests` to verify `收/支` header with `支出` yields negative amount; added `RepairLegacyJdSignTests` to verify opt-in repair updates legacy-positive JD expenses to negative in place.

---

## Section 2: Error & Solution Log
Every unique error encountered during development or deployment must be recorded here:

### [ERROR-ID] | Short Descriptive Title
- **Context:** Where/when did it happen?
- **Symptoms:** Error message or stack trace.
- **Root Cause:** Detailed explanation of the failure.
- **Resolution:** Step-by-step fix implemented.
- **Prevention:** How to avoid this in the future.

### JD-CSV-SIGN-001 | JD Expense Sign Misparsing
- **Context:** Parsing JD CSV exports during transaction import on 2026-03-01.
- **Symptoms:** Expenses with header `收/支` and value `支出` were inserted as positive amounts instead of negative.
- **Root Cause:** Direction header fallback missing - parser only recognized `收支`, not `收/支` variant used by JD exports.
- **Resolution:** Added fallback in `parse_jd_csv_file()`: `direction = cell('收/支') or cell('收支')`; implemented `--repair-jd-legacy-sign` flag to repair legacy-positive JD expenses in place.
- **Prevention:** Keep regression tests for both header variants (`收/支` and `收支`); recommend using `--repair-jd-legacy-sign` to avoid duplicates for legacy-inserted rows.

---

## Section 3: Workflow for Agents
1. **Analyze** the task or error.
2. **Execute** code changes.
3. **Verify** with tests.
4. **Append** documentation to Section 1 (for changes) or Section 2 (for errors) of **this file** before finalizing.