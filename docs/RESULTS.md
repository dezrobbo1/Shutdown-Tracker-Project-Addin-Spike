# Office.js Project add-in spike results

Status: **Not yet executed in Microsoft Project**

This file records observed host behaviour only. Do not convert untested assumptions into product authority.

## Milestone 1 — read-only host access

### Environment

- Microsoft Project version: Pending
- Windows version: Pending
- Office bitness: Pending
- Add-in commit SHA: Pending
- Synthetic Project file: Pending
- Local time zone: Pending

### Results

Complete `docs/TEST_MATRIX.md` during the manual host trial, then summarize the evidence here.

### Raw value observations

| Field | Example raw value | JavaScript type | Notes |
| --- | --- | --- | --- |
| Project GUID | Pending | Pending | |
| Project ReadOnly | Pending | Pending | |
| Task GUID | Pending | Pending | |
| ID | Pending | Pending | |
| Name | Pending | Pending | |
| WBS | Pending | Pending | |
| Summary | Pending | Pending | |
| Percent Complete | Pending | Pending | |
| Actual Start | Pending | Pending | Do not normalize yet |
| Actual Finish | Pending | Pending | Do not normalize yet |
| Start | Pending | Pending | Do not normalize yet |
| Finish | Pending | Pending | Do not normalize yet |
| Duration | Pending | Pending | |
| Remaining Duration | Pending | Pending | |

### Identity comparison

Office.js selected task GUID versus MSPDI/XML `<Task><GUID>`:

Pending.

### Decision

Pending.

Milestone 2 must not begin until the read-only identity and field contract is understood well enough to define a controlled write experiment.
