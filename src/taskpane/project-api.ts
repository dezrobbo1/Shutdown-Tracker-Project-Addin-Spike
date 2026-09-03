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

const EXPECTED_PROJECT_GUID = "CADADEB7-9DA7-F111-9812-4C56DF4490A6";
const EXPECTED_TASK_GUID = "07709767-9EA7-F111-9812-4C56DF4490A6";
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

export async function runControlledPercentCompleteWrite(): Promise<ControlledPercentWriteResult> {
  const project = await readProjectSnapshot();
  expectGuid("Project GUID", project.guid.raw, EXPECTED_PROJECT_GUID);
  expectEqual("Project read-only state", project.readOnly.raw, false);

  const before = await readSelectedTaskSnapshot();
  expectGuid("selected task GUID", before.taskGuid, EXPECTED_TASK_GUID);
  expectGuid("Task GUID field", fieldValue(before, "Task GUID"), EXPECTED_TASK_GUID);
  expectEqual("Task ID", fieldValue(before, "ID"), EXPECTED_TASK_ID);
  expectEqual("Task Name", fieldValue(before, "Name"), EXPECTED_TASK_NAME);
  expectEqual("Task WBS", fieldValue(before, "WBS"), EXPECTED_TASK_WBS);
  expectEqual("Summary", fieldValue(before, "Summary"), "No");
  expectEqual("Percent Complete", fieldValue(before, "Percent Complete"), EXPECTED_PERCENT_COMPLETE);

  // Final compare immediately before the write. If the user or Project changed the
  // value after the snapshot was read, refuse to write rather than overwrite it.
  const finalPercent = await getTaskField(before.taskGuid, Office.ProjectTaskFields.PercentComplete);
  expectEqual("final pre-write Percent Complete", finalPercent, EXPECTED_PERCENT_COMPLETE);

  await setTaskField(before.taskGuid, Office.ProjectTaskFields.PercentComplete, TARGET_PERCENT_COMPLETE);

  const afterImmediate = await readSelectedTaskSnapshot();
  await wait(500);
  const afterSettled = await readSelectedTaskSnapshot();

  expectGuid("post-write selected task GUID", afterSettled.taskGuid, EXPECTED_TASK_GUID);
  expectEqual("post-write Percent Complete", fieldValue(afterSettled, "Percent Complete"), "25%");

  return {
    project,
    requestedPercent: TARGET_PERCENT_COMPLETE,
    before,
    afterImmediate,
    afterSettled,
  };
}
