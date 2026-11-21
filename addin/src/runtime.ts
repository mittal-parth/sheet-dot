import { describeError, recordDialogResult, recordRuntimeStart, runWorkbookChecks } from "./excel/compatibility";

const status = document.querySelector<HTMLElement>("#status")!;
const runButton = document.querySelector<HTMLButtonElement>("#run-checks")!;
const dialogButton = document.querySelector<HTMLButtonElement>("#open-dialog")!;
let runningChecks: Promise<void> | undefined;

function runChecks(): Promise<void> {
  runningChecks ??= runWorkbookChecks()
    .then(() => { status.textContent = "Check the results on the Sheetdot Compatibility worksheet."; })
    .catch(async (error) => {
      console.error("Sheetdot compatibility checks failed", error);
      status.textContent = describeError(error);
      await Office.addin.showAsTaskpane();
    })
    .finally(() => {
      runButton.disabled = false;
      runningChecks = undefined;
    });
  runButton.disabled = true;
  return runningChecks;
}

async function runRibbonChecks(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await runChecks();
  } finally {
    event.completed();
  }
}

function showDialogResult(message: string): void {
  status.textContent = message;
  void recordDialogResult(message).catch((error) => console.error("Could not report the dialog result", error));
}

function openTestDialog(): void {
  dialogButton.disabled = true;
  const url = new URL("dialog.html", window.location.href);
  Office.context.ui.displayDialogAsync(url.href, { height: 42, width: 40 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      dialogButton.disabled = false;
      showDialogResult(`Dialog failed · ${result.error.message}`);
      return;
    }
    result.value.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
      if ("message" in event && event.message === "sheetdot-dialog-ready") {
        showDialogResult("Office pop-up opened · follow the Microsoft Learn link to verify external navigation");
      }
    });
    result.value.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
      dialogButton.disabled = false;
      // 12006 is the normal close notification.
      if ("error" in event && event.error !== 12006) showDialogResult(`Dialog event error · ${event.error}`);
    });
  });
}

Office.actions.associate("runSheetdotChecks", runRibbonChecks);

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) {
    status.textContent = "Open this add-in from Excel to run the compatibility checks.";
    return;
  }
  status.textContent = "Run the checks to create a test worksheet.";
  runButton.disabled = false;
  dialogButton.disabled = false;
  runButton.addEventListener("click", () => void runChecks());
  dialogButton.addEventListener("click", openTestDialog);
  void recordRuntimeStart().catch((error) => console.error("Sheetdot runtime startup check failed", error));
});
