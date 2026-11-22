const billColumns = {
  merchant: ["merchant", "expense", "biller"],
  amount: ["amount due", "amount", "amount ₹"],
  due: ["due date", "due"],
  month: ["billing month", "billing period"],
  status: ["payment status"],
  id: ["sheetdot bill id"],
};
const extraBillColumns = { month: "Billing month", status: "Payment status", id: "Sheetdot bill ID" };
let gmailDialog;
let billKey;
let billPolling;
let billScheduleSignature;
let gmailConnecting = false;
const appliedBillRuns = new Map();
const syncedBillRows = new Map();

function billStatus(message, tone = "success") {
  const status = document.getElementById("bill-watch-status");
  status.textContent = message;
  status.dataset.tone = tone;
}

async function billAccessKey(create = false) {
  const keyName = "sheetdot.billWatchKey";
  const settings = Office.context.document.settings;
  if (!billKey && settings.get(keyName)) billKey = Promise.resolve(settings.get(keyName));
  if (!billKey && create) {
    billKey = (async () => {
      const key = Array.from(crypto.getRandomValues(new Uint8Array(32)), byte => byte.toString(16).padStart(2, "0")).join("");
      settings.set(keyName, key);
      await new Promise((resolve, reject) => settings.saveAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error("Could not save the Gmail connection key in this workbook."));
      }));
      return key;
    })().catch(error => {
      settings.remove(keyName);
      billKey = undefined;
      throw error;
    });
  }
  return billKey;
}

async function requestBills(action, data = {}) {
  const key = await billAccessKey(action === "connect");
  if (!key && action === "status") return { connected: false, schedules: [] };
  if (!key) throw new Error("Connect Gmail first.");
  const response = await fetch("/gmail-bills", {
    method: "POST",
    headers: { Authorization: `Bearer ${key}`, "Content-Type": "application/json" },
    body: JSON.stringify({ action, workbook: await getWorkbookId(), ...data }),
  });
  const result = await response.json();
  if (!response.ok) throw Object.assign(new Error(result.error), { retryable: response.status >= 500 });
  return result;
}

async function connectBillGmail() {
  const { url } = await requestBills("connect");
  gmailConnecting = true;
  const callback = new URL("/gmail-connected.html", location.origin);
  callback.hash = encodeURIComponent(url);
  await new Promise((resolve, reject) => Office.context.ui.displayDialogAsync(callback.href,
    { height: 60, width: 45, displayInIframe: false }, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error.message));
        return;
      }
      gmailDialog = result.value;
      gmailDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        gmailDialog = undefined;
        refreshBillWatches();
      });
      resolve();
    }));
  billStatus("Finish connecting Gmail in the dialog.");
  setTimeout(refreshBillWatches, 2_000);
}

function billIndexes(headers) {
  const names = headers.map(value => String(value).trim().toLowerCase());
  return Object.fromEntries(Object.entries(billColumns).map(([key, aliases]) =>
    [key, names.findIndex(name => aliases.includes(name))]));
}

function billDate(value) {
  if (typeof value === "number" && value >= 36_526 && value < 73_050) {
    // The Excel web workbooks we target use the standard 1900 date system.
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86_400_000).toISOString().slice(0, 10);
  }
  if (typeof value === "string" && /^20\d\d-\d\d-\d\d$/.test(value)) {
    const date = new Date(value);
    if (Number.isFinite(date.getTime()) && date.toISOString().startsWith(value)) return value;
  }
  return "";
}

function billMonth(value, due) {
  if (value === "") return due.slice(0, 7);
  if (typeof value === "string" && /^20\d\d-(0[1-9]|1[0-2])$/.test(value)) return value;
  return billDate(value).slice(0, 7);
}

function billSerial(value) {
  return (Date.parse(value) - Date.UTC(1899, 11, 30)) / 86_400_000;
}

const isBillFormula = (formula, value) => typeof formula === "string" && formula.startsWith("=") && formula !== value;
const billUuid = /^[a-f0-9-]{36}$/;

async function emptyBillExpansion(context, range) {
  const tables = range.getTables();
  tables.load("items");
  range.load("values,formulas");
  await context.sync();
  return !tables.items.length && range.values.every((row, r) =>
    row.every((value, c) => value === "" && !isBillFormula(range.formulas[r][c], value)));
}

async function prepareSelectedBillTable() {
  return Excel.run(async context => {
    const tables = context.workbook.getActiveCell().getTables();
    tables.load("items/id,items/name,items/showHeaders,items/showTotals");
    await context.sync();
    const table = tables.items[0];
    if (!table || !table.showHeaders || table.showTotals || ["SheetdotAutomations", "SheetdotAboutMe"].includes(table.name)) {
      throw new Error("Select a cell in an Excel bill table with headers and its totals row turned off.");
    }
    const header = table.getHeaderRowRange();
    const range = table.getRange();
    header.load("values");
    range.load("rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();
    const indexes = billIndexes(header.values[0]);
    if (["merchant", "amount", "due"].some(key => indexes[key] < 0)) {
      throw new Error("Add Merchant, Amount due and Due date headers to this table.");
    }
    const body = table.getDataBodyRange();
    body.load("values");
    await context.sync();
    const merchants = [...new Set(body.values.map(row => String(row[indexes.merchant]).trim()).filter(Boolean))];
    if (!merchants.length || merchants.length > 12 || merchants.some(value => value.length > 100 || !/[\p{L}\p{N}]/u.test(value))) {
      throw new Error("Enter between 1 and 12 merchants in the table.");
    }
    if (body.values.length > 400) throw new Error("This bill watch supports up to 400 rows.");
    if (body.values.some(row => {
      const merchant = String(row[indexes.merchant]).trim();
      if (!merchant) return false;
      const due = billDate(row[indexes.due]);
      const month = billMonth(indexes.month < 0 ? "" : row[indexes.month], due);
      const amount = row[indexes.amount];
      return !due || !month || typeof amount !== "number" || amount <= 0 || amount > 1e9 ||
        Math.abs(amount * 100 - Math.round(amount * 100)) > 0.00001;
    })) throw new Error("Each bill needs a positive numeric amount, a valid due date, and a valid billing month if supplied.");
    const extras = Object.entries(extraBillColumns).filter(([key]) => indexes[key] < 0).map(([, name]) => name);
    if (extras.length) {
      if (range.columnIndex + range.columnCount + extras.length > 16_384) throw new Error("Make room for the bill columns.");
      const expansion = table.worksheet.getRangeByIndexes(range.rowIndex, range.columnIndex + range.columnCount,
        range.rowCount, extras.length);
      if (!await emptyBillExpansion(context, expansion)) throw new Error("Make room to the right of the bill table for its new columns.");
      table.resize(range.getResizedRange(0, extras.length));
      expansion.getRow(0).values = [extras];
      await context.sync();
    }
    return { id: table.id, name: table.name, merchants };
  });
}

async function billTableContext(context, id) {
  const table = context.workbook.tables.getItemOrNullObject(id);
  table.load("showHeaders,showTotals");
  await context.sync();
  if (table.isNullObject) throw new Error("The tracked Excel table is missing. Remove its schedule in Automations.");
  if (!table.showHeaders || table.showTotals) throw new Error("Show the tracked table’s headers and turn off its totals row.");
  const header = table.getHeaderRowRange();
  const body = table.getDataBodyRange();
  const range = table.getRange();
  header.load("values");
  body.load("values,formulas,address");
  range.load("rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
  const indexes = billIndexes(header.values[0]);
  if (Object.values(indexes).some(index => index < 0)) throw new Error("A tracked bill column is missing.");
  return { table, body, range, indexes };
}

async function snapshotBillTable(id) {
  return Excel.run(async context => {
    const { body, indexes } = await billTableContext(context, id);
    const rows = [];
    const ids = new Set();
    for (const [index, values] of body.values.entries()) {
      const merchant = String(values[indexes.merchant]).trim();
      const due = billDate(values[indexes.due]);
      const month = billMonth(values[indexes.month], due);
      const amount = values[indexes.amount];
      if (!merchant || !due || !month || typeof amount !== "number" || amount <= 0 || amount > 1e9 ||
        Math.abs(amount * 100 - Math.round(amount * 100)) > 0.00001) continue;
      let id = values[indexes.id];
      if (!billUuid.test(String(id)) || ids.has(id)) {
        if (isBillFormula(body.formulas[index][indexes.id], id)) throw new Error("Keep the Sheetdot bill ID column free of formulas.");
        id = crypto.randomUUID();
        body.getCell(index, indexes.id).values = [[id]];
      }
      ids.add(id);
      rows.push({ id, merchant, month, due, amount, paid: /^paid\b/i.test(String(values[indexes.status])) });
    }
    if (rows.length > 400) throw new Error("This bill watch supports up to 400 rows.");
    await context.sync();
    return rows;
  });
}

function billSource(evidence, title, checkedAt) {
  return { title: [title, evidence.subject || "Gmail message", evidence.date.slice(0, 10),
    checkedAt && `Checked ${checkedAt}`].filter(Boolean).join(" · "), url: evidence.url };
}

function billCommentAddress(bodyAddress, rowIndex, columnIndex) {
  const separator = bodyAddress.lastIndexOf("!");
  const start = bodyAddress.slice(separator + 1).split(":")[0];
  const firstRow = Number(start.replace(/[^0-9]/g, ""));
  const firstColumn = [...start.replace(/[^A-Z]/gi, "").toUpperCase()]
    .reduce((column, letter) => column * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  return bodyAddress.slice(0, separator + 1) + columnLabel(firstColumn + columnIndex) + (firstRow + rowIndex);
}

async function applyBillRows(id, bills) {
  const comments = await Excel.run(async context => {
    const { table, body, range, indexes } = await billTableContext(context, id);
    const byId = new Map(body.values.flatMap((row, index) => row[indexes.id] ? [[row[indexes.id], index]] : []));
    const available = body.values.flatMap((values, index) => values.every(value => value === "") &&
      body.formulas[index].every(value => value === "") ? [index] : []);
    const additions = bills.filter(bill => !byId.has(bill.id) && bill.notice);
    const grow = Math.max(0, additions.length - available.length);
    if (grow) {
      if (range.rowIndex + range.rowCount + grow > 1048576) throw new Error("Make room below the bill table.");
      const expansion = table.worksheet.getRangeByIndexes(range.rowIndex + range.rowCount,
        range.columnIndex, grow, range.columnCount);
      if (!await emptyBillExpansion(context, expansion)) throw new Error("Make room below the bill table for new bills.");
      table.resize(range.getResizedRange(grow, 0));
    }

    const notes = [];
    const addNote = (row, column, evidence, title, checkedAt) => {
      if (!evidence) return;
      const address = billCommentAddress(body.address, row, column);
      notes.push({ address, answer: { answeredAt: checkedAt || evidence.date, sources: [billSource(evidence, title, checkedAt)] } });
    };
    const updates = table.getDataBodyRange();
    const set = (row, column, value, fresh = false, highlight = true) => {
      const before = body.values[row]?.[column];
      if (!fresh && (before === value || isBillFormula(body.formulas[row]?.[column], before))) return;
      const cell = updates.getCell(row, column);
      cell.values = [[literalWatchValue(value)]];
      if (highlight) markWatchCell(cell);
    };
    const paymentStatus = bill => bill.review ? "Needs review" : bill.paid ?
      (bill.paidOn ? `Paid · ${bill.paidOn}` : "Paid") : "No payment evidence found";
    const annotate = (row, bill) => {
      addNote(row, indexes.amount, bill.notice, `Bill · ₹${bill.amount} due ${bill.due}`, bill.noticeCheckedAt);
      addNote(row, indexes.status, bill.review ?? bill.payment, bill.review ? "Conflicting bill or payment evidence" :
        `Payment matched · ${bill.merchant} · ₹${bill.amount} · ${bill.month}`, bill.checkedAt);
    };
    for (const bill of bills) {
      const row = byId.get(bill.id);
      if (row === undefined) continue;
      const values = body.values[row];
      if (String(values[indexes.merchant]).trim() !== bill.merchant ||
        billMonth(values[indexes.month], billDate(values[indexes.due])) !== bill.month) continue;
      const sheetAmount = values[indexes.amount];
      const sheetDue = billDate(values[indexes.due]);
      const currentTerms = sheetAmount === bill.amount && sheetDue === bill.due;
      const previousTerms = bill.termsHistory?.some(terms => terms.amount === sheetAmount && terms.due === sheetDue);
      if (!currentTerms && !previousTerms) continue;
      if (previousTerms) {
        if (sheetAmount !== bill.amount) set(row, indexes.amount, bill.amount);
        if (sheetDue !== bill.due) {
          if (!isBillFormula(body.formulas[row][indexes.due], values[indexes.due])) {
            updates.getCell(row, indexes.due).numberFormat = [["dd mmm yyyy"]];
            set(row, indexes.due, billSerial(bill.due));
          }
        }
      }
      if (!values[indexes.month]) set(row, indexes.month, bill.month, false, false);
      if (bill.review || bill.payment || !values[indexes.status]) set(row, indexes.status, paymentStatus(bill));
      annotate(row, bill);
    }
    for (const [index, bill] of additions.entries()) {
      const row = available[index] ?? body.values.length + index - available.length;
      set(row, indexes.id, bill.id, true, false);
      set(row, indexes.merchant, bill.merchant, true);
      set(row, indexes.month, bill.month, true);
      set(row, indexes.amount, bill.amount, true);
      updates.getCell(row, indexes.due).numberFormat = [["dd mmm yyyy"]];
      set(row, indexes.due, billSerial(bill.due), true);
      set(row, indexes.status, paymentStatus(bill), true);
      annotate(row, bill);
    }
    await context.sync();
    return notes;
  });
  for (const { address, answer } of comments) await writeAnswerComment(address, answer, () => false);
}

async function synchronizeBillWatches() {
  const { connected, connecting, schedules } = await requestBills("status");
  const connect = document.getElementById("connect-gmail");
  connect.textContent = connected ? "Reconnect Gmail" : "Connect Gmail";
  if (connected && !connecting && gmailConnecting) {
    gmailDialog?.close();
    gmailDialog = undefined;
    gmailConnecting = false;
    billStatus("Gmail connected.");
  }
  const signature = JSON.stringify(schedules);
  if (signature !== billScheduleSignature && (billScheduleSignature !== undefined || schedules.length)) {
    refreshWatchSheet();
  }
  billScheduleSignature = signature;
  for (const schedule of schedules) {
    try {
      const { id, lastRun } = schedule;
      const rows = await snapshotBillTable(id);
      const snapshot = JSON.stringify(rows);
      if (syncedBillRows.get(id) !== snapshot) {
        await requestBills("sync", { id, rows });
        syncedBillRows.set(id, snapshot);
      }
      if (!lastRun || appliedBillRuns.get(id) === lastRun) continue;
      const { rows: bills } = await requestBills("read", { id });
      if (!bills) continue;
      const job = watchEditQueue.then(() => applyBillRows(id, bills));
      watchEditQueue = job.catch(logWatchEditError);
      await job;
      appliedBillRuns.set(id, lastRun);
    } catch (error) {
      billStatus(error.message, "error");
    }
  }
}

function refreshBillWatches() {
  if (!billPolling) {
    billPolling = synchronizeBillWatches().catch(error => billStatus(error.message, "error"))
      .finally(() => { billPolling = undefined; });
  }
  return billPolling;
}

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel || !Office.context.requirements.isSetSupported("ExcelApi", "1.13")) return;
  const connect = document.getElementById("connect-gmail");
  const start = document.getElementById("start-bill-watch");
  connect.disabled = start.disabled = false;
  connect.onclick = async () => {
    connect.disabled = true;
    try { await connectBillGmail(); }
    catch (error) { billStatus(error.message, "error"); }
    finally { connect.disabled = false; }
  };
  document.getElementById("bill-watch-form").onsubmit = async event => {
    event.preventDefault();
    start.disabled = true;
    try {
      const { connected } = await requestBills("status");
      if (!connected) throw new Error("Connect Gmail before tracking bills.");
      const table = await withSuggestionPreviewCleared(prepareSelectedBillTable);
      const rows = await snapshotBillTable(table.id);
      const schedule = document.getElementById("bill-watch-schedule").value;
      const label = document.getElementById("bill-watch-label").value.trim();
      const importMonth = document.getElementById("bill-watch-import").value.trim();
      const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
      await requestBills("start", { ...table, rows, schedule, label, importMonth, timezone });
      billStatus("Bill tracking is active. Edit its schedule in Automations.");
      refreshWatchSheet();
      refreshBillWatches();
    } catch (error) {
      billStatus(error.message, "error");
    } finally {
      start.disabled = false;
    }
  };
  refreshBillWatches();
  setInterval(refreshBillWatches, 15_000);
});
