const suggestionsSetting = "sheetdot.columnSuggestions";
const previewSetting = "sheetdot.columnSuggestionPreview";
const suggestionCache = new Map();
const suggestionGray = "#696969";
let suggestionsEnabled = false;
let suggestionPreview;
let suggestionQueue = Promise.resolve();
let suggestionTimer;
let selectionRevision = 0;

function suggestionStatus(message, tone = "success") {
  const status = document.querySelector("#suggestion-status");
  status.textContent = message;
  status.dataset.tone = tone;
}

function saveSuggestionSetting(name, value) {
  const settings = Office.context.document.settings;
  if (value === undefined) settings.remove(name);
  else settings.set(name, value);
  return new Promise((resolve, reject) => settings.saveAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
    else reject(new Error("Could not save column suggestions in this workbook."));
  }));
}

function queueSuggestionWork(work) {
  suggestionQueue = suggestionQueue.then(work).catch(error => suggestionStatus(error.message, "error"));
  return suggestionQueue;
}

function columnLabel(index) {
  let label = "";
  for (let number = index + 1; number > 0; number = Math.floor((number - 1) / 26)) {
    label = String.fromCharCode(65 + (number - 1) % 26) + label;
  }
  return label;
}

function literalCell(text) {
  return /^[=+\-@']/.test(text) ? "'" + text : text;
}

function inSuggestionArea(area, selection, extraColumns = 0) {
  return selection.sheet === area.sheet && selection.row >= area.row &&
    selection.row < area.row + Math.max(area.height, extraColumns ? 3 : 0) &&
    selection.column >= area.column && selection.column < area.column + area.width + extraColumns;
}

function activeSuggestionCell(context) {
  const cell = context.workbook.getActiveCell();
  cell.load(["rowIndex", "columnIndex", "worksheet/id", "worksheet/name"]);
  return cell;
}

function cellPosition(cell) {
  return { sheet: cell.worksheet.id, row: cell.rowIndex, column: cell.columnIndex };
}

function cutSuggestionText(value) {
  return typeof value === "string" ? value.slice(0, 120) : value;
}

async function readSuggestionArea() {
  return Excel.run(async context => {
    const cell = activeSuggestionCell(context);
    const tables = cell.getTables();
    tables.load("items/id,items/showHeaders");
    await context.sync();
    if (["About me", "Automations"].includes(cell.worksheet.name)) return;

    const table = tables.items[0];
    if (table && !table.showHeaders) return;
    const range = table ? table.getRange() : cell.getSurroundingRegion();
    const rowCount = table?.rows.getCount();
    range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();
    const dataRows = rowCount?.value ?? range.rowCount - 1;
    if (!dataRows || range.columnCount > 30 || range.columnIndex + range.columnCount >= 16_384) return;

    const sample = cell.worksheet.getRangeByIndexes(range.rowIndex, range.columnIndex,
      Math.min(dataRows, 5) + 1, range.columnCount);
    sample.load("values");
    await context.sync();
    const [headers, ...rows] = sample.values;
    if (!headers.every(header => typeof header === "string" && header.trim()) ||
      !rows.some(row => row.some(value => value !== ""))) return;
    return {
      area: { sheet: cell.worksheet.id, table: table?.id, row: range.rowIndex, column: range.columnIndex,
        width: range.columnCount, height: range.rowCount, dataRows, headers },
      rows: rows.map(row => row.map(cutSuggestionText)),
    };
  });
}

function getPreviewColumn(sheet, area, index, height = 3) {
  return sheet.getRangeByIndexes(area.row, area.column + area.width + index, height, 1);
}

function readPreviewCells(range, checkbox) {
  return [0, 1, 2].map((row) => {
    const cell = range.getCell(row, 0);
    cell.load(["values", "formulas", "format/font/color", checkbox ? "control" : "dataValidation/type"]);
    return cell;
  });
}

async function previewTableRange(context, sheet, area) {
  if (!area.table) return;
  const table = sheet.tables.getItemOrNullObject(area.table);
  await context.sync();
  if (table.isNullObject) return;
  const range = table.getRange();
  range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
  return { table, range };
}

function restorePreviewTable(sheet, preview, current, previewHeaders) {
  if (!current) return;
  const { area, entries } = preview;
  const range = current.range;
  const expanded = range.columnCount - area.width;
  if (range.rowIndex !== area.row || range.columnIndex !== area.column || range.rowCount !== area.height ||
    expanded < 1 || expanded > entries.length || !previewHeaders.slice(0, expanded).every(Boolean)) return;
  current.table.resize(sheet.getRangeByIndexes(area.row, area.column, area.height, area.width));
}

async function clearSuggestionPreview() {
  const preview = suggestionPreview;
  if (!preview) return;
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(preview.area.sheet);
    await context.sync();
    if (sheet.isNullObject) return;
    const table = await previewTableRange(context, sheet, preview.area);
    const entries = preview.entries.map((entry, index) =>
      readPreviewCells(getPreviewColumn(sheet, preview.area, index), preview.checkbox));
    for (const [index, cells] of entries.entries()) {
      if (preview.entries[index].width) cells[0].format.load("columnWidth");
    }
    await context.sync();
    restorePreviewTable(sheet, preview, table, entries.map((cells, index) =>
      cells[0].values[0][0] === preview.entries[index].suggestion.title));
    for (const [index, cells] of entries.entries()) {
      const entry = preview.entries[index];
      const expected = [entry.suggestion.title, entry.suggestion.description];
      const owned = cells.slice(0, 2).map((cell, row) =>
        cell.values[0][0] === expected[row] && cell.formulas[0][0] === expected[row]);
      if (entry.width && owned.some(Boolean) && Math.abs(cells[0].format.columnWidth - entry.width.preview) < 1) {
        cells[0].format.columnWidth = entry.width.original;
      }
      for (const [row, cell] of cells.entries()) {
        const value = cell.values[0][0];
        if (row < 2 && !owned[row]) continue;
        if (row === 2 && !owned.some(Boolean)) continue;
        if (row === 2 && preview.checkbox && cell.control.type !== Excel.CellControlType.checkbox) continue;
        if (row === 2 && !preview.checkbox && cell.dataValidation.type !== "List") continue;
        if (row === 2) {
          if (preview.checkbox) cell.control = { type: Excel.CellControlType.empty };
          else cell.dataValidation.clear();
          if (value !== false && value !== true && value !== "Add column") continue;
        }
        cell.clear("Contents");
        if (row < 2 && cell.format.font.color?.toUpperCase() === suggestionGray.toUpperCase()) {
          cell.format.font.color = entry.colors[row];
        }
      }
    }
    await context.sync();
  });
  await saveSuggestionSetting(previewSetting);
  suggestionPreview = undefined;
}

async function showSuggestionPreview(area, suggestions, revision) {
  if (!suggestions.length) return;
  await Excel.run(async context => {
    const active = activeSuggestionCell(context);
    const sheet = context.workbook.worksheets.getItem(area.sheet);
    const checkbox = Office.context.requirements.isSetSupported("ExcelApi", "1.18") && !!Excel.CellControlType?.checkbox;
    const candidates = [];
    const count = Math.min(suggestions.length, 16_384 - area.column - area.width);
    for (let index = 0; index < count; index++) {
      const column = getPreviewColumn(sheet, area, index, Math.max(area.height, 3));
      const cells = readPreviewCells(column, checkbox);
      cells[0].format.load("columnWidth");
      candidates.push({
        used: column.getUsedRangeOrNullObject(true),
        usedColumn: column.getEntireColumn().getUsedRangeOrNullObject(true),
        formulas: column.getSpecialCellsOrNullObject("Formulas"),
        cells,
      });
    }
    await context.sync();
    if (revision !== selectionRevision || !inSuggestionArea(area, cellPosition(active))) return;

    const entries = [];
    for (const [index, candidate] of candidates.entries()) {
      const { used, usedColumn, formulas, cells } = candidate;
      if (!used.isNullObject || !formulas.isNullObject ||
        cells.some(cell => checkbox ? cell.control.type !== Excel.CellControlType.empty : cell.dataValidation.type !== "None")) break;
      const suggestion = suggestions[index];
      const originalWidth = cells[0].format.columnWidth;
      const desiredWidth = Math.min(205, Math.max(117, Math.max(suggestion.title.length, suggestion.description.length) * 5 + 12));
      const width = usedColumn.isNullObject && originalWidth > 0 && originalWidth < desiredWidth
        ? { original: originalWidth, preview: desiredWidth }
        : undefined;
      entries.push({ suggestion, width, colors: cells.slice(0, 2).map(cell => cell.format.font.color || "#000000") });
    }
    if (!entries.length) return;

    suggestionPreview = { area, checkbox, entries };
    await saveSuggestionSetting(previewSetting, suggestionPreview);
    if (revision !== selectionRevision) {
      await saveSuggestionSetting(previewSetting);
      suggestionPreview = undefined;
      return;
    }
    for (const [index, entry] of entries.entries()) {
      const [header, detail, add] = candidates[index].cells;
      header.values = [[literalCell(entry.suggestion.title)]];
      detail.values = [[literalCell(entry.suggestion.description)]];
      header.format.font.color = detail.format.font.color = suggestionGray;
      if (entry.width) header.format.columnWidth = entry.width.preview;
      if (checkbox) {
        add.values = [[false]];
        add.control = { type: Excel.CellControlType.checkbox };
      } else {
        add.dataValidation.rule = { list: { inCellDropDown: true, source: "Add column" } };
      }
    }
    await context.sync();
    // Excel can automatically include adjacent writes in a table. Keep previews outside its filters and data range.
    if (area.table) {
      const table = await previewTableRange(context, sheet, area);
      for (const { cells } of candidates.slice(0, entries.length)) cells[0].load("values");
      await context.sync();
      restorePreviewTable(sheet, suggestionPreview, table, entries.map((entry, index) =>
        candidates[index].cells[0].values[0][0] === entry.suggestion.title));
      await context.sync();
    }
    suggestionStatus(checkbox ? "Tick a checkbox beside the table to add its column."
      : "Choose Add column from the dropdown below a suggestion.");
  });
}

function suggestionFormula(suggestion, area, row) {
  const refs = suggestion.columns.map(index => columnLabel(area.column + index) + row);
  switch (suggestion.kind) {
    case "sum": return `=SUM(${refs.join(",")})`;
    case "product": return `=PRODUCT(${refs.join(",")})`;
    case "difference": return `=${refs[0]}-${refs[1]}`;
    case "ratio": return `=IFERROR(${refs[0]}/${refs[1]},"")`;
    default: {
      const input = refs.length === 1 ? refs[0] : `HSTACK(${refs.join(",")})`;
      const question = suggestion.question.replaceAll('"', '""');
      return `=AI.${suggestion.kind === "check" ? "CHECK" : "ASK"}("${question}",${input})`;
    }
  }
}

async function appendSuggestedColumn(area, suggestion) {
  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(area.sheet);
    const header = sheet.getRangeByIndexes(area.row, area.column, 1, area.width);
    header.load("values");
    const target = getPreviewColumn(sheet, area, 0, area.height);
    const used = target.getUsedRangeOrNullObject(true);
    const formulas = target.getSpecialCellsOrNullObject("Formulas");
    const table = area.table ? sheet.tables.getItem(area.table) : undefined;
    const range = table ? table.getRange() : header.getCell(0, 0).getSurroundingRegion();
    range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    const rowCount = table?.rows.getCount();
    await context.sync();
    if (!used.isNullObject || !formulas.isNullObject) throw new Error("The next column already contains data.");
    if (JSON.stringify(header.values[0]) !== JSON.stringify(area.headers) ||
      range.rowIndex !== area.row || range.columnIndex !== area.column ||
      range.rowCount !== area.height || range.columnCount !== area.width) {
      throw new Error("The table changed. Select it again to get suggestions for the updated columns.");
    }

    if (table) table.resize(range.getResizedRange(0, 1));
    target.getCell(0, 0).values = [[literalCell(suggestion.title)]];
    const rows = rowCount?.value ?? area.dataRows;
    target.getCell(1, 0).getResizedRange(rows - 1, 0).formulas = Array.from({ length: rows },
      (_, index) => [suggestionFormula(suggestion, area, area.row + index + 2)]);
    header.getCell(0, 0).select();
    await context.sync();
    return { ...area, width: area.width + 1, headers: [...area.headers, suggestion.title] };
  });
}

function suggestionAreaKey(area) {
  return JSON.stringify([area.sheet, area.table, area.row, area.column, area.headers]);
}

async function refreshColumnSuggestions(revision) {
  if (revision !== selectionRevision) return;
  if (suggestionPreview) {
    const selection = await Excel.run(async context => {
      const cell = activeSuggestionCell(context);
      await context.sync();
      return cellPosition(cell);
    });
    if (suggestionsEnabled && inSuggestionArea(suggestionPreview.area, selection, suggestionPreview.entries.length)) return;
    await clearSuggestionPreview();
    suggestionStatus("");
  }
  if (!suggestionsEnabled || revision !== selectionRevision) return;
  const data = await readSuggestionArea();
  if (!data || revision !== selectionRevision) return;
  const { area, rows } = data;
  const key = suggestionAreaKey(area);
  let suggestions = suggestionCache.get(key);
  if (!suggestions) {
    suggestionStatus("Looking for useful columns…");
    const preferences = (await readWorkbookPreferences()).slice(0, 20).map(pair => pair.map(cutSuggestionText));
    if (!suggestionsEnabled || revision !== selectionRevision) return;
    const { suggestions: choices } = await requestSheetdot("/suggest-columns", {
      workbook: await getWorkbookId(), headers: area.headers.map(cutSuggestionText), rows, preferences,
    });
    suggestions = choices;
    suggestionCache.set(key, suggestions);
  }
  if (suggestionsEnabled && revision === selectionRevision) {
    suggestionStatus("");
    await showSuggestionPreview(area, suggestions, revision);
  }
}

function scheduleColumnSuggestions() {
  const revision = ++selectionRevision;
  clearTimeout(suggestionTimer);
  suggestionTimer = setTimeout(() => queueSuggestionWork(() => refreshColumnSuggestions(revision)), 350);
}

function withSuggestionPreviewCleared(work) {
  ++selectionRevision;
  clearTimeout(suggestionTimer);
  const operation = suggestionQueue.then(async () => {
    await clearSuggestionPreview();
    suggestionStatus("");
    return work();
  });
  suggestionQueue = operation.catch(() => {});
  return operation.finally(() => {
    if (suggestionsEnabled) scheduleColumnSuggestions();
  });
}

function acceptChangedSuggestion({ worksheetId, address, source }) {
  if (!suggestionsEnabled || source === "Remote" || !suggestionPreview || worksheetId !== suggestionPreview.area.sheet) return;
  const preview = suggestionPreview;
  const { area, entries } = preview;
  const cellAddress = address.split("!").at(-1).replaceAll("$", "").toUpperCase();
  const index = entries.findIndex((_, offset) => columnLabel(area.column + area.width + offset) + (area.row + 3) === cellAddress);
  if (index < 0) return;
  queueSuggestionWork(async () => {
    if (!suggestionsEnabled || suggestionPreview !== preview) return;
    const selected = await Excel.run(async context => {
      const column = getPreviewColumn(context.workbook.worksheets.getItem(area.sheet), area, index);
      column.load("values");
      await context.sync();
      const [[title], [description], [accept]] = column.values;
      return title === entries[index].suggestion.title && description === entries[index].suggestion.description &&
        (preview.checkbox ? accept === true : accept === "Add column");
    });
    if (!selected) return;
    const accepted = entries[index].suggestion;
    const remaining = (suggestionCache.get(suggestionAreaKey(area)) ?? entries.map(entry => entry.suggestion))
      .filter(suggestion => suggestion !== accepted && suggestion.title !== accepted.title);
    await clearSuggestionPreview();
    const updated = await appendSuggestedColumn(area, accepted);
    suggestionCache.set(suggestionAreaKey(updated), remaining);
    suggestionStatus(`Added ${accepted.title}.`);
    scheduleColumnSuggestions();
  });
}

Office.onReady(() => {
  const toggle = document.querySelector("#enable-suggestions");
  const settings = Office.context.document.settings;
  suggestionsEnabled = settings.get(suggestionsSetting) === true;
  suggestionPreview = settings.get(previewSetting);
  const updateToggle = () => {
    toggle.checked = suggestionsEnabled;
    toggle.disabled = false;
  };
  queueSuggestionWork(async () => {
    await clearSuggestionPreview();
    await Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.onSelectionChanged.add(scheduleColumnSuggestions);
      sheets.onActivated.add(scheduleColumnSuggestions);
      sheets.onChanged.add(acceptChangedSuggestion);
      await context.sync();
    });
    updateToggle();
    if (suggestionsEnabled) scheduleColumnSuggestions();
  });

  toggle.onchange = () => {
    const enabled = toggle.checked;
    toggle.disabled = true;
    // Invalidate an in-flight request before waiting for the workbook write queue.
    suggestionsEnabled = enabled;
    ++selectionRevision;
    queueSuggestionWork(async () => {
      try {
        await saveSuggestionSetting(suggestionsSetting, enabled);
        suggestionStatus("");
        if (!enabled) await clearSuggestionPreview();
        else scheduleColumnSuggestions();
      } catch (error) {
        suggestionsEnabled = settings.get(suggestionsSetting) === true;
        throw error;
      } finally {
        updateToggle();
      }
    });
  };
});
