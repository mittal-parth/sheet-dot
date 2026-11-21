const sheetName = "Sheetdot Compatibility";
const title = "Sheetdot compatibility checks";
const approvalCell = "B8";
const approvalStatus = "B9";

async function getSheet(context: Excel.RequestContext): Promise<Excel.Worksheet | null> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
  await context.sync();
  if (sheet.isNullObject) return null;

  const header = sheet.getRange("A1").load("values");
  await context.sync();
  if (header.values[0][0] !== title) {
    throw new Error(`Rename the existing “${sheetName}” worksheet before running the checks.`);
  }
  return sheet;
}

async function writeStatus(address: string, status: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getSheet(context);
    if (!sheet) return;
    sheet.getRange(address).values = [[status]];
    await context.sync();
  });
}

export function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function addNote(): Promise<void> {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
    await writeStatus("B6", "Notes API unavailable · requires ExcelApi 1.18");
    return;
  }
  try {
    await Excel.run(async (context) => {
      const sheet = await getSheet(context);
      if (!sheet) return;
      const address = "B6";
      const existing = sheet.notes.getItemOrNullObject(address);
      await context.sync();
      const content = "Source: Sheetdot compatibility test";
      if (existing.isNullObject) sheet.notes.add(address, content);
      else existing.content = content;
      sheet.getRange(address).values = [["Note added · open it to verify"]];
      await context.sync();
    });
  } catch (error) {
    await writeStatus("B6", `Note failed · ${describeError(error)}`);
  }
}

async function addApprovalControl(): Promise<void> {
  let fallback = "Checkbox unavailable · choose TRUE or FALSE";
  if (Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
    try {
      await Excel.run(async (context) => {
        const sheet = await getSheet(context);
        if (!sheet) return;
        const cell = sheet.getRange(approvalCell);
        cell.dataValidation.clear();
        cell.control = { type: Excel.CellControlType.checkbox };
        sheet.getRange(approvalStatus).values = [["Tick the checkbox to verify the change event"]];
        await context.sync();
      });
      return;
    } catch (error) {
      fallback = `Checkbox failed · choose TRUE or FALSE (${describeError(error)})`;
    }
  }

  await Excel.run(async (context) => {
    const sheet = await getSheet(context);
    if (!sheet) return;
    const cell = sheet.getRange(approvalCell);
    cell.dataValidation.rule = { list: { inCellDropDown: true, source: "TRUE,FALSE" } };
    sheet.getRange(approvalStatus).values = [[fallback]];
    await context.sync();
  });
}

let watchedSheetId: string | undefined;

export async function recordRuntimeStart(): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await getSheet(context);
    if (!sheet) return;
    sheet.load("id");
    await context.sync();
    if (watchedSheetId === sheet.id) return;

    sheet.onChanged.add(async (event) => {
      const address = event.address.split("!").at(-1)?.replaceAll("$", "");
      if (address !== approvalCell) return;
      try {
        await Excel.run(async (eventContext) => {
          const eventSheet = eventContext.workbook.worksheets.getItem(event.worksheetId);
          const cell = eventSheet.getRange(approvalCell).load("values");
          await eventContext.sync();
          const value = cell.values[0][0];
          eventSheet.getRange(approvalStatus).values = [[`Change event received · ${String(value).toUpperCase()}`]];
          await eventContext.sync();
        });
      } catch (error) {
        console.error("Sheetdot could not read the approval change", error);
      }
    });
    sheet.getRange("B11").values = [[`Started ${new Date().toLocaleString()}`]];
    watchedSheetId = sheet.id;
    try {
      await context.sync();
    } catch (error) {
      watchedSheetId = undefined;
      throw error;
    }
  });
}

export async function runWorkbookChecks(): Promise<void> {
  await Excel.run(async (context) => {
    const existing = await getSheet(context);
    const sheet = existing ?? context.workbook.worksheets.add(sheetName);
    sheet.getRange("A1:B7").values = [
      [title, "Result"],
      ["Scalar formula", ""],
      ["Streaming formula", ""],
      ["Spilled array", ""],
      ["", ""],
      ["Cell note", "Checking…"],
      ["Cell fill and conditional format", 1],
    ];
    sheet.getRange("A8:A12").values = [
      ["Approval control"], ["Approval change event"], ["Load when workbook opens"],
      ["Last runtime start"], ["Office dialog"],
    ];
    if (!existing) sheet.getRange(approvalCell).values = [[false]];
    sheet.getRange("B9:B10").values = [["Checking…"], ["Checking…"]];
    sheet.getRange("B12").values = [["Try Home → Sheetdot → Test dialog"]];
    sheet.getRange("C4:C5").clear(Excel.ClearApplyTo.contents);
    sheet.getRange("A1:C1").format.font.bold = true;
    sheet.getRange("A:A").format.columnWidth = 195;
    sheet.getRange("B:B").format.columnWidth = 320;
    sheet.getRange("B2:B4").formulas = [
      ['=AI.ECHO("Value returned")'],
      ['=AI.STREAM("Stream finished")'],
      ["=AI.SPILL()"],
    ];
    const fill = sheet.getRange("B7");
    fill.format.fill.color = "#FFF2CC";
    fill.conditionalFormats.clearAll();
    const condition = fill.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    condition.custom.rule.formula = "=B7=2";
    condition.custom.format.fill.color = "#C6EFCE";
    sheet.activate();
    await context.sync();
  });

  await addNote();
  await addApprovalControl();
  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    await writeStatus("B10", "Enabled · save, close, and reopen this workbook");
  } catch (error) {
    await writeStatus("B10", `Startup registration failed · ${describeError(error)}`);
  }
  await recordRuntimeStart();
}

export async function recordDialogResult(status: string): Promise<void> {
  await writeStatus("B12", status);
}
