import "./styles.css";
import {
  getSelectedTaskGuid,
  readProjectSnapshot,
  readTaskSnapshot,
  runControlledPercentCompleteWrite,
  type RawFieldValue,
  type TaskSnapshot,
} from "./project-api";

const projectFields = document.querySelector<HTMLDListElement>("#project-fields");
const taskFields = document.querySelector<HTMLDivElement>("#task-fields");
const taskEmpty = document.querySelector<HTMLParagraphElement>("#task-empty");
const log = document.querySelector<HTMLPreElement>("#log");
const refreshProject = document.querySelector<HTMLButtonElement>("#refresh-project");
const refreshTask = document.querySelector<HTMLButtonElement>("#refresh-task");
const writePercent = document.querySelector<HTMLButtonElement>("#write-percent");
const writeResult = document.querySelector<HTMLDivElement>("#write-result");

let writeArmed = false;
let cachedTaskGuid: string | null = null;

function appendLog(message: string): void {
  if (!log) return;
  const timestamp = new Date().toISOString();
  log.textContent = `${log.textContent ?? ""}[${timestamp}] ${message}\n`;
}

function printable(value: unknown): string {
  if (value === undefined) return "undefined";
  if (value === null) return "null";
  if (typeof value === "string") return value || "(empty string)";
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function renderFieldList(container: HTMLElement, fields: RawFieldValue[]): void {
  container.replaceChildren();

  for (const field of fields) {
    const row = document.createElement("div");
    row.className = "field-row";

    const label = document.createElement("div");
    label.className = "field-label";
    label.textContent = field.label;

    const value = document.createElement("code");
    value.className = "field-value";
    value.textContent = printable(field.raw);

    const type = document.createElement("span");
    type.className = "field-type";
    type.textContent = `JS type: ${field.jsType}`;

    row.append(label, value, type);
    container.append(row);
  }
}

function renderSnapshotBlock(parent: HTMLElement, titleText: string, snapshot: TaskSnapshot): void {
  const wrapper = document.createElement("div");
  wrapper.className = "write-result-block";

  const title = document.createElement("h3");
  title.textContent = titleText;

  const fields = document.createElement("div");
  renderFieldList(fields, snapshot.fields);

  wrapper.append(title, fields);
  parent.append(wrapper);
}

async function loadProject(): Promise<void> {
  if (!projectFields || !refreshProject) return;
  refreshProject.disabled = true;
  try {
    appendLog("Reading active Project fields...");
    const project = await readProjectSnapshot();
    renderFieldList(projectFields, [project.guid, project.readOnly]);
    appendLog("Active Project fields read successfully.");
  } catch (error) {
    appendLog(`Project read failed: ${error instanceof Error ? error.message : String(error)}`);
  } finally {
    refreshProject.disabled = false;
  }
}

async function captureTaskSelection(reason: string): Promise<void> {
  try {
    const guid = await getSelectedTaskGuid();
    cachedTaskGuid = guid;
    appendLog(`Captured Project task selection (${reason}): ${guid}`);

    if (taskEmpty) {
      taskEmpty.hidden = true;
      taskEmpty.textContent = "Select a task row in Microsoft Project, then read it.";
    }

    if (taskFields) {
      const task = await readTaskSnapshot(guid);
      renderFieldList(taskFields, task.fields);
    }
  } catch (error) {
    appendLog(`Task selection capture skipped (${reason}): ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function loadTask(): Promise<void> {
  if (!taskFields || !refreshTask || !taskEmpty) return;
  refreshTask.disabled = true;

  try {
    if (!cachedTaskGuid) {
      throw new Error("No cached Project task selection. Click the Remove cover row in Project first; the add-in will capture it on TaskSelectionChanged.");
    }

    appendLog(`Reading cached task GUID ${cachedTaskGuid} sequentially...`);
    const task = await readTaskSnapshot(cachedTaskGuid);
    taskEmpty.hidden = true;
    renderFieldList(taskFields, task.fields);
    appendLog(`Cached task read successfully. Task GUID: ${task.taskGuid}`);
  } catch (error) {
    taskEmpty.hidden = false;
    taskEmpty.textContent = error instanceof Error ? error.message : String(error);
    taskFields.replaceChildren();
    appendLog(`Task read failed: ${error instanceof Error ? error.message : String(error)}`);
  } finally {
    refreshTask.disabled = false;
  }
}

function armWrite(): void {
  if (!writePercent || !writeResult) return;
  writeArmed = true;
  writeResult.replaceChildren();

  const warning = document.createElement("p");
  warning.className = "write-warning";
  warning.textContent = "Write armed. Click the button again to perform the guarded 0% → 25% write. All identity/state guards will still be rechecked immediately before writing.";
  writeResult.append(warning);

  writePercent.textContent = "Confirm guarded write 25%";
  appendLog("Controlled write armed. Waiting for second click confirmation.");
}

function resetWriteArm(): void {
  writeArmed = false;
  if (writePercent) writePercent.textContent = "Write 25% to guarded task";
}

async function runWrite(): Promise<void> {
  if (!writePercent || !writeResult) return;

  if (!writeArmed) {
    armWrite();
    return;
  }

  resetWriteArm();
  writePercent.disabled = true;
  refreshProject && (refreshProject.disabled = true);
  refreshTask && (refreshTask.disabled = true);
  writeResult.replaceChildren();

  try {
    if (!cachedTaskGuid) {
      throw new Error("No cached Project task selection. Click the Remove cover row in Project first so TaskSelectionChanged can capture its GUID.");
    }

    appendLog(`Starting guarded Percent Complete write using cached task GUID ${cachedTaskGuid}...`);
    const result = await runControlledPercentCompleteWrite(cachedTaskGuid);
    appendLog(`Guard checks passed. Requested Percent Complete = ${result.requestedPercent}%.`);
    appendLog("Office.js write succeeded; immediate and 500 ms read-back snapshots captured.");

    const requested = document.createElement("p");
    requested.className = "write-requested";
    requested.textContent = `Requested write: Percent Complete = ${result.requestedPercent}%`;
    writeResult.append(requested);

    renderSnapshotBlock(writeResult, "Before", result.before);
    renderSnapshotBlock(writeResult, "After — immediate", result.afterImmediate);
    renderSnapshotBlock(writeResult, "After — 500 ms settled", result.afterSettled);

    await loadTask();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    appendLog(`Controlled write stopped: ${message}`);

    const errorText = document.createElement("p");
    errorText.className = "write-error";
    errorText.textContent = `No successful controlled write result: ${message}`;
    writeResult.append(errorText);
  } finally {
    writePercent.disabled = false;
    refreshProject && (refreshProject.disabled = false);
    refreshTask && (refreshTask.disabled = false);
  }
}

Office.onReady((info) => {
  appendLog(`Office ready. Host: ${info.host ?? "unknown"}; platform: ${info.platform ?? "unknown"}.`);

  if (info.host !== Office.HostType.Project) {
    appendLog("This spike must be opened inside Microsoft Project.");
    return;
  }

  refreshProject?.addEventListener("click", () => void loadProject());
  refreshTask?.addEventListener("click", () => void loadTask());
  writePercent?.addEventListener("click", () => void runWrite());

  Office.context.document.addHandlerAsync(
    Office.EventType.TaskSelectionChanged,
    () => void captureTaskSelection("TaskSelectionChanged"),
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        appendLog("TaskSelectionChanged handler registered. Click a task row in Project to cache its GUID before using task-pane buttons.");
      } else {
        appendLog(`Unable to register TaskSelectionChanged handler: ${result.error?.code ?? "OFFICE_ERROR"}: ${result.error?.message ?? "Unknown error"}`);
      }
    },
  );

  void loadProject();
  void captureTaskSelection("initial selection");
});
