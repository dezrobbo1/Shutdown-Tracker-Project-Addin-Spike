import "./styles.css";
import {
  readProjectSnapshot,
  readSelectedTaskSnapshot,
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

async function loadTask(): Promise<void> {
  if (!taskFields || !refreshTask || !taskEmpty) return;
  refreshTask.disabled = true;
  try {
    appendLog("Reading selected task sequentially...");
    const task = await readSelectedTaskSnapshot();
    taskEmpty.hidden = true;
    renderFieldList(taskFields, task.fields);
    appendLog(`Selected task read successfully. Host task GUID: ${task.taskGuid}`);
  } catch (error) {
    taskEmpty.hidden = false;
    taskFields.replaceChildren();
    appendLog(`Task read failed: ${error instanceof Error ? error.message : String(error)}`);
  } finally {
    refreshTask.disabled = false;
  }
}

async function runWrite(): Promise<void> {
  if (!writePercent || !writeResult) return;

  const confirmed = window.confirm(
    "Synthetic spike only. This will write Percent Complete = 25% to the guarded Remove cover task if every identity and state check passes. Continue?",
  );
  if (!confirmed) {
    appendLog("Controlled write cancelled by user.");
    return;
  }

  writePercent.disabled = true;
  refreshProject && (refreshProject.disabled = true);
  refreshTask && (refreshTask.disabled = true);
  writeResult.replaceChildren();

  try {
    appendLog("Starting guarded Percent Complete write. Reading and validating project/task state...");
    const result = await runControlledPercentCompleteWrite();
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
  void loadProject();
});
