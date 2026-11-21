const report = message => document.querySelector("#status").append(message + "\n");
Office.onReady(async () => {
  const run = document.querySelector("#run");
  const dialog = document.querySelector("#dialog");
  const stopWatch = document.querySelector("#stop-watch");
  run.disabled = dialog.disabled = stopWatch.disabled = false;
  stopWatch.onclick = async () => {
    stopWatch.disabled = true;
    stopWatch.setAttribute("aria-busy", "true");
    try {
      const address = await Excel.run(async context => {
        const cell = context.workbook.getActiveCell();
        cell.load("address");
        await context.sync();
        return cell.address;
      });
      const { stopped } = await requestSheetdot("/watch", { workbook: await getWorkbookId(), address }, "DELETE");
      watchSheetStatus(stopped ? `Stopped WATCH at ${address}.` : `No active WATCH at ${address}.`);
      globalThis.refreshWatchSheet?.();
    } catch (error) {
      watchSheetStatus(error.message, "error");
    } finally {
      stopWatch.removeAttribute("aria-busy");
      stopWatch.disabled = false;
    }
  };
  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    report("Shared runtime loaded at " + new Date().toLocaleTimeString());
  } catch (error) { report("Load on open: " + error.message); }
  run.onclick = () => Excel.run(async context => {
    const sheet = context.workbook.worksheets.add();
    sheet.activate();
    const cell = sheet.getRange("A1");
    cell.values = [["Sheetdot compatibility"]];
    cell.format.fill.color = "yellow";
    const flag = sheet.getRange("B1");
    flag.values = [[false]];
    cell.conditionalFormats.add("ContainsText").textComparison.rule =
      { operator: "Contains", text: "Sheetdot" };
    await context.sync();
    report("Created sheet, fill and conditional format.");
    try {
      sheet.notes.add(cell, "Sheetdot source test");
      await context.sync();
      report("Note added.");
    } catch (error) { report("Note: " + error.message); }
    try {
      flag.control = { type: Excel.CellControlType.checkbox };
      await context.sync();
      report("Toggle B1 to test change detection.");
    } catch (error) {
      report("Checkbox: " + error.message);
      flag.dataValidation.rule = { list: { inCellDropDown: true, source: "TRUE,FALSE" } };
    }
    sheet.onChanged.add(event => { report("Changed " + event.address); });
    await context.sync();
  }).catch(error => report(error.message));
  dialog.onclick = () => Office.context.ui.displayDialogAsync(location.origin + "/dialog.html", result => {
    report(result.status === Office.AsyncResultStatus.Succeeded ? "Dialog opened." : result.error.message);
  });
});
