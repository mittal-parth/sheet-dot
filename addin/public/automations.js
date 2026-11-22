const automationsSheet = "Automations";
const automationsTable = "SheetdotAutomations";
const automationHeaders = ["WATCH cell", "Question", "Schedule", "Status", "Last run"];
let previousAutomations;
let automationsQueue = Promise.resolve();
let subscribedSheet;
let retryAutomations;
let billWatchIds = new Map();

function watchSheetStatus(message, tone = "success") {
  const status = document.querySelector("#automations-status");
  status.dataset.tone = tone;
  status.textContent = message;
}

function automationRows(values) {
  return values.slice(1).filter(row => typeof row[0] === "string" && row[0].includes("!"));
}

async function readAutomationRows(create, activate) {
  return Excel.run(async context => {
    const workbook = context.workbook;
    const existing = workbook.tables.getItemOrNullObject(automationsTable);
    await context.sync();
    if (!existing.isNullObject) {
      const range = existing.getRange();
      const sheet = existing.worksheet;
      sheet.load("id");
      range.load("values");
      if (activate) sheet.activate();
      await context.sync();
      if (subscribedSheet !== sheet.id) {
        sheet.onChanged.add(() => refreshWatchSheet());
        await context.sync();
        subscribedSheet = sheet.id;
      }
      return automationRows(range.values);
    }
    if (!create) return;

    let sheet = workbook.worksheets.getItemOrNullObject(automationsSheet);
    await context.sync();
    if (sheet.isNullObject) {
      sheet = workbook.worksheets.add(automationsSheet);
    } else {
      const used = sheet.getUsedRangeOrNullObject(true);
      await context.sync();
      if (!used.isNullObject) throw new Error("The existing Automations sheet already contains data.");
    }
    const header = sheet.getRange("A1:E1");
    header.values = [automationHeaders];
    const table = sheet.tables.add(header, true);
    table.name = automationsTable;
    sheet.getRange("A:A").format.columnWidth = 110;
    sheet.getRange("B:B").format.columnWidth = 280;
    sheet.getRange("C:E").format.columnWidth = 135;
    sheet.onChanged.add(() => refreshWatchSheet());
    sheet.load("id");
    if (activate) sheet.activate();
    await context.sync();
    subscribedSheet = sheet.id;
    return [];
  });
}

async function writeAutomationRows(before, after) {
  if (JSON.stringify(before) === JSON.stringify(after)) return true;
  return Excel.run(async context => {
    const table = context.workbook.tables.getItem(automationsTable);
    const current = table.getRange();
    current.load("values");
    await context.sync();
    if (JSON.stringify(automationRows(current.values)) !== JSON.stringify(before)) return false;

    const sheet = table.worksheet;
    const rowCount = Math.max(1, after.length);
    const width = automationHeaders.length;
    const oldCount = current.values.length - 1;
    if (rowCount !== oldCount) table.resize(sheet.getRangeByIndexes(0, 0, rowCount + 1, width));
    const body = sheet.getRangeByIndexes(1, 0, rowCount, width);
    body.numberFormat = Array.from({ length: rowCount }, () => automationHeaders.map(() => "@"));
    body.values = after.length ? after.map(row => row.map(value =>
      typeof value === "string" && /^[=+\-@]/.test(value) ? "'" + value : value
    )) : [automationHeaders.map(() => "")];
    if (oldCount > rowCount) {
      sheet.getRangeByIndexes(rowCount + 1, 0, oldCount - rowCount, width).clear("Contents");
    }
    await context.sync();
    return true;
  });
}

async function reconcileAutomationEdits(workbook, rows) {
  if (!previousAutomations) return;
  const current = new Map(rows.map(row => [row[0], row[2]]));
  const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  for (const [address, oldSchedule] of previousAutomations) {
    const billId = billWatchIds.get(address);
    try {
      if (!current.has(address)) {
        if (billId) await requestBills("stop", { id: billId });
        else await requestSheetdot("/automations", { workbook, address }, "DELETE");
        watchSheetStatus(`Stopped and removed ${address}.`);
      } else if (current.get(address) !== oldSchedule) {
        const schedule = current.get(address);
        if (billId) await requestBills("schedule", { id: billId, schedule });
        else await requestSheetdot("/automations", { workbook, address, schedule, timezone }, "PATCH");
        watchSheetStatus(`Updated the schedule for ${address}.`);
      }
    } catch (error) {
      if (error.retryable || error instanceof TypeError) throw error;
      watchSheetStatus(error.message, "error");
    }
  }
}

async function synchronizeAutomations(activate) {
  const workbook = await getWorkbookId();
  let rows = await readAutomationRows(activate, activate);
  if (rows) await reconcileAutomationEdits(workbook, rows);
  const [formulas, bills] = await Promise.all([
    requestSheetdot("/automations", { workbook }), requestBills("status"),
  ]);
  const billSchedules = bills.schedules.map(item => ({ ...item, address: `Bill table!${item.name}` }));
  const schedules = [...formulas.schedules, ...billSchedules];
  if (!rows && schedules.length === 0) return;
  rows ??= await readAutomationRows(true, false);

  const desired = schedules.map(item => [
    item.address, item.question, item.schedule, item.status,
    item.lastRun ? new Date(item.lastRun * 1000).toLocaleString() : "",
  ]);
  if (await writeAutomationRows(rows, desired)) {
    previousAutomations = new Map(desired.map(row => [row[0], row[2]]));
    billWatchIds = new Map(billSchedules.map(item => [item.address, item.id]));
  } else {
    refreshWatchSheet();
  }
}

function refreshWatchSheet(activate = false) {
  automationsQueue = automationsQueue.then(() => synchronizeAutomations(activate)).catch(error => {
    watchSheetStatus(error.message, "error");
    if ((error.retryable || error instanceof TypeError) && !retryAutomations) {
      retryAutomations = setTimeout(() => {
        retryAutomations = undefined;
        refreshWatchSheet();
      }, 5000);
    }
  });
  return automationsQueue;
}

globalThis.refreshWatchSheet = refreshWatchSheet;
Office.onReady(() => {
  const open = document.querySelector("#open-automations");
  open.disabled = false;
  open.onclick = async () => {
    open.disabled = true;
    open.setAttribute("aria-busy", "true");
    try {
      await refreshWatchSheet(true);
    } finally {
      open.removeAttribute("aria-busy");
      open.disabled = false;
    }
  };
  refreshWatchSheet();
});
