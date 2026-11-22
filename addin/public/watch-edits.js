const watchFormula = /(?:AI\.WATCH|_xldudf_AI_WATCH)\s*\(/i;
const watchHighlight = '=N("Sheetdot WATCH change")=0';
const watchEditReceipts = new Map();
let watchEditQueue = Promise.resolve();

function watchEditsSupported() {
  return Office.context.requirements.isSetSupported("ExcelApi", "1.9");
}

async function findWatchTable(context, address) {
  const separator = address?.lastIndexOf("!") ?? -1;
  if (separator < 1) return;
  const name = address.slice(0, separator).replace(/^'(.*)'$/, "$1").replaceAll("''", "'");
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  if (sheet.isNullObject) return;

  const anchor = sheet.getRange(address.slice(separator + 1));
  const tables = anchor.getTables();
  tables.load("items/id,items/name");
  anchor.load("rowIndex,columnIndex,formulas");
  await context.sync();
  const [table] = tables.items;
  if (!table || table.name.toLowerCase() === "sheetdotautomations" || name.toLowerCase() === "about me" ||
      !watchFormula.test(anchor.formulas[0][0])) return;

  const header = table.getHeaderRowRange();
  const body = table.getDataBodyRange();
  const rows = table.rows;
  header.load("values");
  header.load("rowIndex,columnIndex");
  rows.load("count");
  await context.sync();
  const rowIndex = anchor.rowIndex - header.rowIndex - 1;
  const columnIndex = anchor.columnIndex - header.columnIndex;
  if (rowIndex < 0 || rowIndex >= rows.count) return;

  const columns = header.values[0];
  if (columns.length > 256) return;
  return {
    sheet, table, body, rowIndex, rowCount: rows.count, columnIndex,
    firstRow: header.rowIndex + 1, firstColumn: header.columnIndex,
    scope: { id: table.id, columns, watchColumn: columns[columnIndex] },
  };
}

async function readWatchTable(address) {
  if (!address || !watchEditsSupported()) return;
  return Excel.run(async context => (await findWatchTable(context, address))?.scope ?? null);
}

function literalWatchValue(value) {
  // A leading apostrophe keeps model-supplied strings literal, including dates and formulas.
  return typeof value === "string" && value !== "" ? "'" + value : value;
}

function markWatchCell(cell) {
  const format = cell.conditionalFormats.add("Custom");
  format.custom.rule.formula = watchHighlight;
  format.custom.format.fill.color = "#FFF2CC";
}

function watchFieldIndexes(fields, columns) {
  return Object.entries(fields).map(([column, value]) => [columns.indexOf(column), value]);
}

async function applyWatchPatch(address, patch, isCanceled) {
  return Excel.run(async context => {
    const target = await findWatchTable(context, address);
    if (isCanceled()) return;
    if (!target || JSON.stringify(target.scope) !== JSON.stringify(patch.table)) {
      return "WATCH could not apply edits because its table changed.";
    }
    const { sheet, table, body, rowIndex, rowCount, columnIndex, firstRow, firstColumn, scope } = target;
    const row = body.getRow(rowIndex);
    const cells = watchFieldIndexes(patch.cells, scope.columns);
    const additions = patch.rows.map(fields => watchFieldIndexes(fields, scope.columns));
    const protectedColumn = ([index]) => index < 0 || index === columnIndex;
    if (cells.some(protectedColumn) || additions.some(fields => fields.some(protectedColumn))) {
      return "WATCH edits would overwrite a formula.";
    }

    const isFormula = (formula, value) => typeof formula === "string" && formula.startsWith("=") && formula !== value;
    row.load("values,formulas");
    if (additions.length) body.load("values,formulas");
    await context.sync();
    if (isCanceled()) return;
    if (cells.some(([index]) => isFormula(row.formulas[0][index], row.values[0][index]))) {
      return "WATCH edits would overwrite a formula.";
    }

    const rowsToAdd = [];
    const watchColumns = new Set([columnIndex]);
    if (additions.length) {
      const formulaColumns = new Set();
      for (const [row, formulas] of body.formulas.entries()) {
        formulas.forEach((formula, index) => {
          if (!isFormula(formula, body.values[row][index])) return;
          formulaColumns.add(index);
          if (watchFormula.test(formula)) watchColumns.add(index);
        });
      }
      if (additions.some(fields => fields.some(([index]) => formulaColumns.has(index)))) {
        return "WATCH edits would overwrite a formula column.";
      }
      const existing = [...body.values];
      for (const fields of additions) {
        if (existing.some(values => fields.every(([index, value]) => values[index] === value))) continue;
        const values = scope.columns.map(() => "");
        for (const [index, value] of fields) values[index] = value;
        existing.push(values);
        rowsToAdd.push({ fields, values });
      }
    }

    const nativeAppend = Office.context.requirements.isSetSupported("ExcelApi", "1.15");
    if (rowsToAdd.length) {
      if (!nativeAppend && !Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
        return "This version of Excel cannot add WATCH rows.";
      }
      const fullRange = table.getRange();
      fullRange.load("rowIndex,rowCount");
      await context.sync();
      const afterTable = fullRange.rowIndex + fullRange.rowCount;
      if (afterTable + rowsToAdd.length > 1048576) return "WATCH needs more empty space below its table.";
      const expansion = sheet.getRangeByIndexes(afterTable, firstColumn, rowsToAdd.length, scope.columns.length);
      const overlaps = expansion.getTables();
      overlaps.load("items");
      expansion.load("values,formulas");
      await context.sync();
      if (isCanceled()) return;
      if (overlaps.items.length || expansion.values.some((values, r) => values.some(
        (value, c) => value !== "" || isFormula(expansion.formulas[r][c], value)
      ))) return "WATCH needs empty space below its table to add rows.";
    }

    context.workbook.application.suspendApiCalculationUntilNextSync();
    for (const [index, value] of cells) {
      if (row.values[0][index] === value) continue;
      const cell = row.getCell(0, index);
      cell.values = [[literalWatchValue(value)]];
      markWatchCell(cell);
    }
    if (rowsToAdd.length) {
      const values = rowsToAdd.map(row => row.values.map(literalWatchValue));
      if (nativeAppend) {
        table.rows.add(-1, values, false);
      } else {
        table.resize(table.getRange().getResizedRange(rowsToAdd.length, 0));
        sheet.getRangeByIndexes(firstRow + rowCount, firstColumn, rowsToAdd.length, scope.columns.length).values = values;
      }
      for (const [offset, { fields }] of rowsToAdd.entries()) {
        const newRow = sheet.getRangeByIndexes(firstRow + rowCount + offset, firstColumn, 1, scope.columns.length);
        // Excel can extend a calculated column to new rows, including WATCH itself.
        for (const index of watchColumns) newRow.getCell(0, index).clear("Contents");
        for (const [index, value] of fields) {
          if (value !== "") markWatchCell(newRow.getCell(0, index));
        }
      }
    }
    await context.sync();
    return "";
  });
}

function queueWatchEdits(address, patches, isCanceled, onIssue) {
  // The task pane applies patches after the streaming formula has returned its result.
  setTimeout(() => {
    watchEditQueue = watchEditQueue.then(() => applyWatchEdits(address, patches, isCanceled))
      .then(onIssue).catch(logWatchEditError);
  }, 0);
}

async function applyWatchEdits(address, patches, isCanceled) {
  if (isCanceled()) return;
  const receipts = watchEditReceipts.get(address) ?? new Map();
  watchEditReceipts.set(address, receipts);
  let issue;
  for (const patch of patches) {
    if (isCanceled()) break;
    if (!receipts.has(patch.id)) {
      const result = await applyWatchPatch(address, patch, isCanceled);
      if (result === undefined) break;
      receipts.set(patch.id, result);
    }
    issue ||= receipts.get(patch.id);
  }
  return issue;
}

function logWatchEditError(error) {
  console.warn("Could not apply Sheetdot WATCH changes", error.code ?? error.name);
}

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel || !watchEditsSupported()) return;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    const job = watchEditQueue.then(() => Excel.run(async context => {
      const formats = context.workbook.getSelectedRanges().conditionalFormats;
      formats.load("items/type");
      await context.sync();
      const custom = formats.items.filter(format => format.type === "Custom");
      if (!custom.length) return;
      for (const format of custom) format.custom.rule.load("formula");
      await context.sync();
      for (const format of custom) {
        if (format.custom.rule.formula === watchHighlight) format.delete();
      }
      await context.sync();
    }));
    watchEditQueue = job.catch(logWatchEditError);
  });
});
