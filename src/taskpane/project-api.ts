export type RawFieldValue = {
  label: string;
  raw: unknown;
  jsType: string;
};

export type ProjectSnapshot = {
  guid: RawFieldValue;
  readOnly: RawFieldValue;
};

export type TaskSnapshot = {
  taskGuid: string;
  fields: RawFieldValue[];
};

export type ControlledPercentWriteResult = {
  project: ProjectSnapshot;
  requestedPercent: number;
  before: TaskSnapshot;
  afterImmediate: TaskSnapshot;
  afterSettled: TaskSnapshot;
};

const EXPECTED_TASK_ID = 2;
const EXPECTED_TASK_NAME = "Remove cover";
const EXPECTED_TASK_WBS = "1.1";
const EXPECTED_PERCENT_COMPLETE = "0%";
const TARGET_PERCENT_COMPLETE = 25;

function jsTypeOf(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  return typeof value;
}

function normalizeGuid(value: unknown): string {
  return typeof value === "string" ? value.trim().toUpperCase() : "";
}

function getProjectField(fieldId: Office.ProjectProjectFields): Promise<unknown> {
  return new Promise((resolve, reject) => {
    Office.context.document.getProjectFieldAsync(fieldId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value?.fieldValue);
        return;
      }
      reject(new Error(`${result.error?.code ?? "OFFICE_ERROR"}: ${result.error?.message ?? "Unable to read project field"}`));
    });
  });
}

function getSelectedTaskGuid(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedTaskAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
        return;
      }
      reject(new Error(`${result.error?.code ?? "NO_TASK"}: ${result.error?.message ?? "No selected task"}`));
    });
  });
}

function getTaskField(taskGuid: string, fieldId: Office.ProjectTaskFields): Promise<unknown> {
  return new Promise((resolve, reject) => {
    Office.context.document.getTaskFieldAsync(taskGuid, fieldId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value?.fieldValue);
        return;
      }
      reject(new Error(`${result.error?.code ?? "OFFICE_ERROR"}: ${result.error?.message ?? "Unable to read task field"}`));
    });
  });
}

function setTaskField(taskGuid: string, fieldId: Office.ProjectTaskFields, value: string | number | boolean | object): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setTaskFieldAsync(taskGuid, fieldId, value, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }
      reject(new Error(`${result.error?.code ?? "OFFICE_ERROR"}: ${result.error?.message ?? "Unable to write task field"}`));
    });
  });
}

function rawField(label: string, raw: unknown): RawFieldValue {
  return { label, raw, jsType: jsTypeOf(raw) };
}

function fieldValue(snapshot: TaskSnapshot, label: string): unknown {
  return snapshot.fields.find((field) => field.label === label)?.raw;
}

function expectEqual(label: string, actual: unknown, expected: unknown): void {
  if (actual !== expected) {
    throw new Error(`Guard failed for ${label}. Expected ${String(expected)}, observed ${String(actual)}.`);
  }
}

function requireGuid(label: string, value: unknown): string {
  const normalized = normalizeGuid(value);
  if (!normalized) {
    throw new Error(`Guard failed for ${label}. A non-empty GUID was required, observed ${String(value)}.`);
  }
  return normalized;
}

function expectGuid(label: string, actual: unknown, expected: string): void {
  if (normalizeGuid(actual) !== expected) {
    throw new Error(`Guard failed for ${label}. Expected ${expected}, observed ${String(actual)}.`);
  }
}

function wait(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

export async function readProjectSnapshot(): Promise<ProjectSnapshot> {
  const [guid, readOnly] = await Promise.all([
    getProjectField(Office.ProjectProjectFields.GUID),
    getProjectField(Office.ProjectProjectFields.ReadOnly),
  ]);

  return {
    guid: rawField("Project GUID", guid),
    readOnly: rawField("Read only", readOnly),
  };
}

const TASK_FIELDS: Array<[string, Office.ProjectTaskFields]> = [
  ["Task GUID", Office.ProjectTaskFields.TaskGUID],
  ["ID", Office.ProjectTaskFields.ID],
  ["Name", Office.ProjectTaskFields.Name],
  ["WBS", Office.ProjectTaskFields.WBS],
  ["Summary", Office.ProjectTaskFields.Summary],
  ["Percent Complete", Office.ProjectTaskFields.PercentComplete],
  ["Percent Work Complete", Office.ProjectTaskFields.PercentWorkComplete],
  ["Actual Start", Office.ProjectTaskFields.ActualStart],
  ["Actual Finish", Office.ProjectTaskFields.ActualFinish],
  ["Actual Duration", Office.ProjectTaskFields.ActualDuration],
  ["Start", Office.ProjectTaskFields.Start],
  ["Finish", Office.ProjectTaskFields.Finish],
  ["Duration", Office.ProjectTaskFields.Duration],
  ["Remaining Duration", Office.ProjectTaskFields.RemainingDuration],
  ["Actual Work", Office.ProjectTaskFields.ActualWork],
  ["Remaining Work", Office.ProjectTaskFields.RemainingWork],
];

export async function readSelectedTaskSnapshot(): Promise<TaskSnapshot> {
  const taskGuid = await getSelectedTaskGuid();
  const fields: RawFieldValue[] = [];

  // Intentionally sequential so host-side recalculation behaviour is easy to inspect.
  for (const [label, fieldId] of TASK_FIELDS) {
    const raw = await getTaskField(taskGuid, fieldId);
    fields.push(rawField(label, raw));
  }

  return { taskGuid, fields };
}

function validateSyntheticTask(snapshot: TaskSnapshot, expectedTaskGuid: string, expectedPercent: string): void {
  expectGuid("selected task GUID", snapshot.taskGuid, expectedTaskGuid);
  expectGuid("Task GUID field", fieldValue(snapshot, "Task GUID"), expectedTaskGuid);
  expectEqual("Task ID", fieldValue(snapshot, "ID"), EXPECTED_TASK_ID);
  expectEqual("Task Name", fieldValue(snapshot, "Name"), EXPECTED_TASK_NAME);
  expectEqual("Task WBS", fieldValue(snapshot, "WBS"), EXPECTED_TASK_WBS);
  expectEqual("Summary", fieldValue(snapshot, "Summary"), "No");
  expectEqual("Percent Complete", fieldValue(snapshot, "Percent Complete"), expectedPercent);
}

export async function runControlledPercentCompleteWrite(): Promise<ControlledPercentWriteResult> {
  const project = await readProjectSnapshot();
  const projectGuid = requireGuid("Project GUID", project.guid.raw);
  expectEqual("Project read-only state", project.readOnly.raw, false);

  // Copies/re-opened synthetic files can legitimately receive a different Project GUID.
  // Anchor this run to the GUIDs observed now, then require them to remain stable across
  // the complete read → compare → write → read-back sequence.
  const before = await readSelectedTaskSnapshot();
  const taskGuid = requireGuid("selected task GUID", before.taskGuid);
  validateSyntheticTask(before, taskGuid, EXPECTED_PERCENT_COMPLETE);

  // Re-read the whole selected task immediately before writing. This rejects both a
  // selection change and any progress/identity drift after the first snapshot.
  const finalBefore = await readSelectedTaskSnapshot();
  validateSyntheticTask(finalBefore, taskGuid, EXPECTED_PERCENT_COMPLETE);

  const finalProject = await readProjectSnapshot();
  expectGuid("final pre-write Project GUID", finalProject.guid.raw, projectGuid);
  expectEqual("final pre-write read-only state", finalProject.readOnly.raw, false);

  await setTaskField(taskGuid, Office.ProjectTaskFields.PercentComplete, TARGET_PERCENT_COMPLETE);

  const afterImmediate = await readSelectedTaskSnapshot();
  await wait(500);
  const afterSettled = await readSelectedTaskSnapshot();
  const projectAfter = await readProjectSnapshot();

  expectGuid("post-write Project GUID", projectAfter.guid.raw, projectGuid);
  expectGuid("post-write selected task GUID", afterSettled.taskGuid, taskGuid);
  expectGuid("post-write Task GUID field", fieldValue(afterSettled, "Task GUID"), taskGuid);
  expectEqual("post-write Task ID", fieldValue(afterSettled, "ID"), EXPECTED_TASK_ID);
  expectEqual("post-write Task Name", fieldValue(afterSettled, "Name"), EXPECTED_TASK_NAME);
  expectEqual("post-write Task WBS", fieldValue(afterSettled, "WBS"), EXPECTED_TASK_WBS);
  expectEqual("post-write Summary", fieldValue(afterSettled, "Summary"), "No");
  expectEqual("post-write Percent Complete", fieldValue(afterSettled, "Percent Complete"), "25%");

  return {
    project,
    requestedPercent: TARGET_PERCENT_COMPLETE,
    before,
    afterImmediate,
    afterSettled,
  };
}
