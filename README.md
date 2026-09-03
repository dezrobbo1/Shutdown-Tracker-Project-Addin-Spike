# Shutdown Tracker Project Add-in Spike

Disposable proof-of-concept repository for testing Microsoft Project Office.js task-pane capabilities before any integration is added to the main Shutdown Tracker repository.

## Current milestone — guarded Percent Complete write

Milestone 1 proved that the add-in can load in Microsoft Project, read the active Project GUID, read the selected task GUID, and reconcile those identities with a Project-generated MSPDI/XML export.

Observed synthetic identities:

- Project GUID: `CADADEB7-9DA7-F111-9812-4C56DF4490A6`
- `Remove cover` task GUID: `07709767-9EA7-F111-9812-4C56DF4490A6`
- task ID: `2`
- WBS: `1.1`

Milestone 2 asks one narrow question:

> Does `Office.context.document.setTaskFieldAsync` setting Percent Complete from 0% to 25% cause Microsoft Project to produce the same meaningful tracking state as entering 25% natively in Project?

The add-in therefore contains exactly one write path. It is hard-coded to the synthetic fixture above and refuses to write unless every guard passes:

- active Project GUID matches the synthetic fixture;
- Project is not read-only;
- selected task GUID and Task GUID field both match `Remove cover`;
- ID is `2`;
- Name is `Remove cover`;
- WBS is `1.1`;
- Summary is `No`;
- Percent Complete is exactly `0%` immediately before the write.

If all guards pass, the add-in calls `setTaskFieldAsync` for `Office.ProjectTaskFields.PercentComplete` with numeric value `25`, then captures immediate and 500 ms read-back snapshots.

## Native Project control already observed

The synthetic native-Project 0% → 25% control produced these task-level effects in the Project-generated XML:

- Percent Complete: `0` → `25`;
- Actual Start: absent → `2026-09-03T08:00:00`;
- Actual Duration: `PT0H0M0S` → `PT2H0M0S`;
- Remaining Duration: `PT8H0M0S` → `PT6H0M0S`;
- Stop / Resume: absent → `2026-09-03T10:00:00`;
- planned Start, Finish and Duration remained unchanged;
- Project GUID and task GUID remained unchanged.

The assignment attached to UID 2 also gained 25% work progress, Actual Start, 2h Actual Work, 6h Remaining Work and corresponding timephased data. The Office.js task pane does not attempt to reproduce these fields itself. Microsoft Project remains responsible for all host-side tracking calculations.

## Non-goals

Milestone 2 has no:

- Shutdown Tracker backend;
- database;
- authentication;
- batch writes;
- Actual Start or Actual Finish writes;
- date conversion logic;
- schedule calculation;
- automatic task selection;
- production UI;
- production integration.

## Windows local development

Use the `.cmd` entry points in PowerShell so a restrictive PowerShell execution policy does not block `npm.ps1` or `npx.ps1`.

```powershell
cd C:\Users\dez16\Shutdown-Tracker-Project-Addin-Spike
git switch spike/officejs-write-percent-complete
git pull --ff-only
npm.cmd install
npx.cmd office-addin-dev-certs install
npm.cmd run dev
```

If the development certificate is already installed and trusted, do not reinstall it.

The Vite server runs at:

```text
https://localhost:3000/
```

## Project shared-folder sideload

Project on Windows should load this add-in from the trusted shared-folder catalog already configured for the spike.

Refresh the catalog copy after switching branches:

```powershell
$catalog = "\\$env:COMPUTERNAME\OfficeAddinManifests"
Copy-Item .\manifest.xml "$catalog\manifest.xml" -Force
```

The write milestone requires `ReadWriteDocument` permission, so Project must load this milestone's updated manifest rather than the earlier read-only manifest. If necessary, fully close Project before reopening **Project → My Add-ins → Shared Folder**.

## Mandatory experiment procedure

1. Start from a fresh 0% copy of the synthetic `OfficeJsSpikeTest` schedule. Do not use the native 25% control file.
2. Open the add-in and click **Read project**.
3. Select `Remove cover` and click **Read selected task**.
4. Verify the pane shows the expected Project GUID, task GUID and `0%` value.
5. Click **Write 25% to guarded task**.
6. Accept the confirmation prompt.
7. Preserve the Before, After — immediate, After — 500 ms settled and probe-log evidence.
8. Save the Project file.
9. Export another copy as Microsoft Project XML.
10. Compare that XML with both the 0% baseline and the native-Project 25% control.

Use `docs/WRITE_PERCENT_COMPLETE_TEST.md` as the evidence checklist.

## Exit gate

Do not add any second writable field until this single-field experiment is understood.

The milestone passes only if:

- the guarded Office.js write succeeds on the intended synthetic task;
- Project reads back `25%`;
- task and Project GUIDs remain unchanged;
- no planned schedule-authority field is unexpectedly changed;
- exported XML shows host-side tracking changes that are explainable and acceptably equivalent to the native Project 25% control.
