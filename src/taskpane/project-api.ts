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

export function getSelectedTaskGuid(): Promise<string> {
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

function expectGuid(label: string, actual: unknown, expected: string): void {
  if (normalizeGuid(actual) !== normalizeGuid(expected)) {
    throw new Error(`Guard failed for ${label}. Expected ${expected}, observed ${String(actual)}.`);
  }
}

function expectNonEmptyGuid(label: string, actual: unknown): string {
  const normalized = normalizeGuid(actual);
  if (!normalized) {
    throw new Error(`Guard failed for ${label}. A non-empty GUID is required.`);
  }
  return normalized;
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

export async function readTaskSnapshot(taskGuid: string): Promise<TaskSnapshot> {
  const normalizedTaskGuid = expectNonEmptyGuid("cached selected task GUID", taskGuid);
  const fields: RawFieldValue[] = [];

  for (const [label, fieldId] of TASK_FIELDS) {
    const raw = await getTaskField(normalizedTaskGuid, fieldId);
    fields.push(rawField(label, raw));
  }

  return { taskGuid: normalizedTaskGuid, fields };
}

export async function readSelectedTaskSnapshot(): Promise<TaskSnapshot> {
  const taskGuid = await getSelectedTaskGuid();
  return readTaskSnapshot(taskGuid);
}

function validateSyntheticTask(snapshot: TaskSnapshot, expectedTaskGuid: string, expectedPercent: string): void {
  expectGuid("cached task GUID", snapshot.taskGuid, expectedTaskGuid);
  expectGuid("Task GUID field", fieldValue(snapshot, "Task GUID"), expectedTaskGuid);
  expectEqual("Task ID", fieldValue(snapshot, "ID"), EXPECTED_TASK_ID);
  expectEqual("Task Name", fieldValue(snapshot, "Name"), EXPECTED_TASK_NAME);
  expectEqual("Task WBS", fieldValue(snapshot, "WBS"), EXPECTED_TASK_WBS);
  expectEqual("Summary", fieldValue(snapshot, "Summary"), "No");
  expectEqual("Percent Complete", fieldValue(snapshot, "Percent Complete"), expectedPercent);
}

export async function runControlledPercentCompleteWrite(cachedTaskGuid: string): Promise<ControlledPercentWriteResult> {
  const expectedTaskGuid = expectNonEmptyGuid("cached selected task GUID", cachedTaskGuid);

  const project = await readProjectSnapshot();
  const runtimeProjectGuid = expectNonEmptyGuid("Project GUID", project.guid.raw);
  expectEqual("Project read-only state", project.readOnly.raw, false);

  const before = await readTaskSnapshot(expectedTaskGuid);
  validateSyntheticTask(before, expectedTaskGuid, EXPECTED_PERCENT_COMPLETE);

  // Re-read the exact cached task by GUID immediately before writing. This does not
  // depend on task-pane focus or on getSelectedTaskAsync after the button is clicked.
  const finalBefore = await readTaskSnapshot(expectedTaskGuid);
  validateSyntheticTask(finalBefore, expectedTaskGuid, EXPECTED_PERCENT_COMPLETE);

  const finalProject = await readProjectSnapshot();
  expectGuid("final pre-write Project GUID", finalProject.guid.raw, runtimeProjectGuid);
  expectEqual("final pre-write Project read-only state", finalProject.readOnly.raw, false);

  await setTaskField(expectedTaskGuid, Office.ProjectTaskFields.PercentComplete, TARGET_PERCENT_COMPLETE);

  const afterImmediate = await readTaskSnapshot(expectedTaskGuid);
  await wait(500);
  const afterSettled = await readTaskSnapshot(expectedTaskGuid);

  const projectAfter = await readProjectSnapshot();
  expectGuid("post-write Project GUID", projectAfter.guid.raw, runtimeProjectGuid);
  validateSyntheticTask(afterSettled, expectedTaskGuid, "25%");

  return {
    project,
    requestedPercent: TARGET_PERCENT_COMPLETE,
    before,
    afterImmediate,
    afterSettled,
  };
}
