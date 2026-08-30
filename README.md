# Shutdown Tracker Project Add-in Spike

Disposable proof-of-concept repository for testing Microsoft Project Office.js task-pane capabilities before any integration is added to the main Shutdown Tracker repository.

## Milestone 1

The first milestone is deliberately read-only. It proves that a task-pane add-in can:

- load inside Microsoft Project;
- read the active Project GUID;
- read the active Project `ReadOnly` state;
- obtain the selected task GUID;
- read Task GUID, ID, Name, WBS and Summary;
- read Percent Complete, Actual Start, Actual Finish, Start, Finish, Duration and Remaining Duration;
- display the raw Office.js value and JavaScript type for every field;
- compare the Office.js task GUID with the same synthetic task in MSPDI/XML.

The Project Common API exposes the fields used here through `getProjectFieldAsync`, `getSelectedTaskAsync`, and `getTaskFieldAsync`. This spike intentionally makes task reads sequential so the first host-behaviour evidence is easy to interpret.

## Non-goals

Milestone 1 has no:

- Shutdown Tracker backend;
- database;
- authentication;
- XML generation;
- production UI;
- `setTaskFieldAsync` calls;
- batching;
- schedule calculation;
- automatic schedule changes;
- production integration.

## Local development

Prerequisites:

- Windows with Microsoft Project desktop installed;
- Node.js and npm;
- permission to install/trust a localhost development certificate;
- a synthetic `.mpp` created specifically for this spike.

In Windows PowerShell, use the `.cmd` entry points so a restrictive PowerShell script execution policy does not block `npm.ps1` or `npx.ps1`:

```powershell
npm.cmd install
npx.cmd office-addin-dev-certs install
npm.cmd run dev
```

If the development certificate has already been installed and trusted, the middle command does not need to be repeated.

The Vite development server uses `office-addin-dev-certs` and serves the task pane at:

```text
https://localhost:3000/index.html
```

Sideload `manifest.xml` into Microsoft Project using the supported Office add-in sideloading process for the installed Project/Office version. Do not use a production Project schedule for the spike.

## Test evidence

Use:

- `docs/TEST_MATRIX.md` for the manual host test;
- `docs/RESULTS.md` for observed raw values, identity comparison and the milestone decision.

Do not normalize date values yet. Milestone 1 exists partly to determine exactly what Microsoft Project returns for its date fields before a date codec is designed.

## Exit gate

Do not add write functionality until the read-only test matrix is complete and the Office.js task GUID has been reconciled against the matching MSPDI/XML `<Task><GUID>` in the synthetic schedule.
