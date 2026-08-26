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

function jsTypeOf(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "array";
  return typeof value;
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

function rawField(label: string, raw: unknown): RawFieldValue {
  return { label, raw, jsType: jsTypeOf(raw) };
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
  ["Actual Start", Office.ProjectTaskFields.ActualStart],
  ["Actual Finish", Office.ProjectTaskFields.ActualFinish],
  ["Start", Office.ProjectTaskFields.Start],
  ["Finish", Office.ProjectTaskFields.Finish],
  ["Duration", Office.ProjectTaskFields.Duration],
  ["Remaining Duration", Office.ProjectTaskFields.RemainingDuration],
];

export async function readSelectedTaskSnapshot(): Promise<TaskSnapshot> {
  const taskGuid = await getSelectedTaskGuid();
  const fields: RawFieldValue[] = [];

  // Intentionally sequential for the first host-behaviour spike.
  for (const [label, fieldId] of TASK_FIELDS) {
    const raw = await getTaskField(taskGuid, fieldId);
    fields.push(rawField(label, raw));
  }

  return { taskGuid, fields };
}
