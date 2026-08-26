# Read-only Office.js host test matrix

## Milestone 1 objective

Prove that a Microsoft Project task-pane add-in can identify the active project and selected task, and can read the exact Project fields needed for later controlled progress experiments.

No task field writes are permitted in this milestone.

## Test environment record

Record before testing:

- Microsoft Project edition and version
- Windows version
- Office bitness
- Project file name
- Project file origin: synthetic only
- Add-in commit SHA
- Local time zone

## Tests

| ID | Test | Expected result | Result | Evidence / notes |
| --- | --- | --- | --- | --- |
| R01 | Add-in task pane loads in Microsoft Project | Task pane renders without host error | Pending | |
| R02 | Read active project GUID | Non-empty Project GUID returned | Pending | |
| R03 | Read active project read-only state | Boolean-like host value returned and raw JS type recorded | Pending | |
| R04 | Select a leaf task and read host task GUID | Non-empty task GUID returned | Pending | |
| R05 | Read leaf task identity | ID, Name and WBS returned | Pending | |
| R06 | Read Summary | Leaf task resolves as non-summary | Pending | |
| R07 | Read Percent Complete | Current Project value returned | Pending | |
| R08 | Read Actual Start and Actual Finish | Raw host values and JS types recorded | Pending | |
| R09 | Read Start and Finish | Raw host values and JS types recorded | Pending | |
| R10 | Read Duration and Remaining Duration | Values returned and raw JS types recorded | Pending | |
| R11 | Select a summary task | Summary resolves as true and remaining fields are observable | Pending | |
| R12 | Change selected task and re-read | Add-in reports the newly selected task, not cached values | Pending | |
| R13 | Compare Office.js Task GUID with MSPDI Task GUID | Same synthetic task GUID matches exported MSPDI/XML | Pending | |
| R14 | Save, close and reopen synthetic project, then re-read | Task GUID remains the expected Project task GUID | Pending | |

## Mandatory capture

For every field read, retain:

- displayed field label
- raw value returned by Office.js
- JavaScript type
- selected task GUID supplied by `getSelectedTaskAsync`
- timestamp of the probe

Do not normalize date values in Milestone 1. The purpose is to observe the host contract before defining a date codec.

## Stop conditions

Stop the spike and record the behaviour if:

- Project GUID cannot be read reliably;
- selected task GUID cannot be read reliably;
- required task fields fail inconsistently;
- Office.js task GUID does not correspond to the same task GUID in exported MSPDI/XML;
- simply reading fields noticeably destabilizes Project.

## Exit gate

Milestone 1 passes only when R01-R14 are completed with no unexplained identity mismatch. Controlled writes belong to a separate follow-up PR.
