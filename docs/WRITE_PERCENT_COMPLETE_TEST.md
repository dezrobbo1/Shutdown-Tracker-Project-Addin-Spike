# Milestone 2 — controlled Percent Complete write test

## Objective

Compare one guarded Office.js write against the already-observed native Microsoft Project control.

The only permitted Office.js write in this milestone is:

```text
Remove cover
Project GUID: CADADEB7-9DA7-F111-9812-4C56DF4490A6
Task GUID:    07709767-9EA7-F111-9812-4C56DF4490A6
Expected:     0%
Requested:    25%
```

No other task and no other field may be written by the spike.

## Native Project control

The native Project 0% → 25% control has already been exported to MSPDI/XML. The meaningful observed task-level effects were:

| Field | 0% baseline | Native Project 25% control |
| --- | --- | --- |
| Percent Complete | 0 | 25 |
| Actual Start | absent | 2026-09-03T08:00:00 |
| Actual Duration | PT0H0M0S | PT2H0M0S |
| Remaining Duration | PT8H0M0S | PT6H0M0S |
| Stop | absent | 2026-09-03T10:00:00 |
| Resume | absent | 2026-09-03T10:00:00 |
| Start | 2026-09-03T08:00:00 | unchanged |
| Finish | 2026-09-03T17:00:00 | unchanged |
| Duration | PT8H0M0S | unchanged |
| Project GUID | CADADEB7-9DA7-F111-9812-4C56DF4490A6 | unchanged |
| Task GUID | 07709767-9EA7-F111-9812-4C56DF4490A6 | unchanged |

The corresponding UID 2 assignment also changed from 0% to 25% work complete, gained Actual Start at 08:00, Actual Work of 2h, Remaining Work of 6h, Stop/Resume at 10:00, and split actual/remaining timephased data.

These are comparison evidence, not values the add-in should set directly.

## Preconditions

Before clicking the write button:

- use a disposable 0% synthetic baseline only;
- Project GUID must match the fixture above;
- selected task must be `Remove cover`;
- task GUID must match the fixture above;
- Project must not be read-only;
- Percent Complete must read as `0%`;
- preserve a separate native 25% control file/XML;
- do not use real shutdown data.

## Test sequence

| ID | Test | Expected result | Result | Evidence / notes |
| --- | --- | --- | --- | --- |
| W01 | Load write-spike manifest | Add-in opens with controlled-write warning | Pending | |
| W02 | Read active project | Expected Project GUID and read-only=false | Pending | |
| W03 | Read selected Remove cover task | Expected GUID, ID 2, Name, WBS 1.1, Summary No, 0% | Pending | |
| W04 | Select wrong task and attempt write | Guard rejects; no write occurs | Pending | |
| W05 | Re-select Remove cover and run write | Confirmation appears | Pending | |
| W06 | Confirm guarded write | `setTaskFieldAsync` succeeds | Pending | |
| W07 | Immediate read-back | Percent Complete reads 25%; snapshot captured | Pending | |
| W08 | 500 ms read-back | Percent Complete remains 25%; recalculated fields captured | Pending | |
| W09 | Repeat write without resetting to 0% | Guard rejects because expected current value is no longer 0% | Pending | |
| W10 | Save Project file | File saves normally | Pending | |
| W11 | Export Project-generated XML | XML export succeeds | Pending | |
| W12 | Compare Project/task identities | Project GUID and Task GUID unchanged | Pending | |
| W13 | Compare planned Start/Finish/Duration | No unexpected schedule-authority changes | Pending | |
| W14 | Compare tracking fields with native Project control | Effects are equivalent or differences are explicitly understood | Pending | |
| W15 | Compare assignment/timephased effects | Effects are equivalent or differences are explicitly understood | Pending | |

## Mandatory pane capture

Retain screenshots or copied values for:

- Before;
- After — immediate;
- After — 500 ms settled;
- probe log;
- Project GUID;
- task GUID;
- Percent Complete;
- Actual Start;
- Actual Finish;
- Actual Duration;
- Start;
- Finish;
- Duration;
- Remaining Duration;
- Percent Work Complete;
- Actual Work;
- Remaining Work.

## Stop conditions

Stop and preserve evidence if:

- any guard unexpectedly passes on the wrong task or wrong project;
- write succeeds when current Percent Complete is not 0%;
- Project/task GUID changes;
- planned Start, Finish or Duration changes unexpectedly;
- Project becomes unstable or unresponsive;
- write callback succeeds but read-back does not reach 25%;
- immediate and settled snapshots differ in unexplained ways.

Do not attempt another field write as a workaround.

## Final XML comparison

The decision is based on three states:

```text
0% baseline
    vs
native Project 25% control
    vs
Office.js 25% result
```

The Office.js path does not need byte-identical XML. Microsoft Project may rewrite metadata and derived records. The required result is equivalent task identity and acceptable Project-owned tracking semantics without unexpected schedule-authority changes.
