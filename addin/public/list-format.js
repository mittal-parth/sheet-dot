const pendingListFormats = new Map();

function queueListFormatting(address, values) {
  if (!address || !Office.context.requirements.isSetSupported("ExcelApi", "1.12")) return;
  const separator = address.lastIndexOf("!");
  if (separator < 1) return;
  const sheet = address.slice(0, separator).replace(/^'(.*)'$/, "$1").replaceAll("''", "'");
  const cell = address.slice(separator + 1);
  const job = { sheet, cell, values, expires: Date.now() + 5_000 };
  pendingListFormats.set(address, job);
  void retryListFormatting(address, job);
}

async function retryListFormatting(address, job) {
  // A fast cached result can spill after the last calculation event.
  for (const delay of [200, 450, 1_200]) {
    await new Promise(resolve => setTimeout(resolve, delay));
    if (pendingListFormats.get(address) !== job) return;
    await formatCalculatedLists(job.sheet).catch(logListFormatError);
  }
}

function logListFormatError(error) {
  console.warn("Could not format Sheetdot LIST results", error.code ?? error.name);
}

async function formatCalculatedLists(worksheet) {
  for (const [address, job] of pendingListFormats) {
    if (Date.now() > job.expires) pendingListFormats.delete(address);
  }
  if (!pendingListFormats.size) return;

  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(worksheet);
    sheet.load("name");
    await context.sync();
    if (sheet.isNullObject) return;

    const lists = [];
    for (const [address, job] of pendingListFormats) {
      if (job.sheet !== sheet.name) continue;
      const anchor = sheet.getRange(job.cell);
      anchor.load("formulas");
      const singleCell = job.values.length === 1 && job.values[0].length === 1;
      const range = singleCell ? anchor : anchor.getSpillingToRangeOrNullObject();
      range.load(["values", "text", "numberFormat", "columnIndex"]);
      lists.push({ address, job, anchor, range });
    }
    if (!lists.length) return;
    await context.sync();

    const widths = new Map();
    const wrappedColumns = [];
    for (const { address, job, anchor, range } of lists) {
      if (pendingListFormats.get(address) !== job || range.isNullObject) continue;
      const formula = anchor.formulas[0][0];
      if (typeof formula !== "string" || !formula.startsWith("=") || !/(?:AI\.LIST|_xldudf_AI_LIST)\s*\(/i.test(formula)) continue;
      const values = range.values;
      if (values.length !== job.values.length || values.some((row, index) =>
        row.length !== job.values[index].length || row.some((value, column) => value !== job.values[index][column]))) continue;

      pendingListFormats.delete(address);
      const header = range.getRow(0);
      header.format.fill.color = "#107c41";
      header.format.font.color = "#ffffff";
      header.format.font.bold = true;

      for (const [index, label] of values[0].entries()) {
        const column = range.getColumn(index);
        const rupees = /\bINR\b|₹/i.test(label);
        let longest = label.length + 2;
        for (let row = 1; row < values.length; row++) {
          const value = values[row][index];
          const isAmount = rupees && typeof value === "number";
          if (isAmount && range.numberFormat[row][index] === "General") {
            const amountFormat = Number.isInteger(value) ? '"₹"#,##0' : '"₹"#,##0.00';
            range.getCell(row, index).numberFormat = [[amountFormat]];
          }
          const text = isAmount ? `₹${value.toLocaleString("en-IN")}` : range.text[row][index];
          longest = Math.max(longest, ...text.split("\n").map(line => line.length));
        }
        const width = Math.min(240, Math.max(56, longest * 7 + 12));
        if (width === 240) wrappedColumns.push(column);
        const position = range.columnIndex + index;
        const current = widths.get(position);
        if (current) {
          current.width = Math.max(current.width, width);
        } else {
          const format = column.format;
          format.load("columnWidth");
          widths.set(position, { format, width });
        }
      }
    }
    if (!widths.size) return;
    await context.sync();
    for (const { format, width } of widths.values()) {
      if (format.columnWidth > 0 && format.columnWidth < width) format.columnWidth = width;
    }
    await context.sync();
    if (wrappedColumns.length) {
      for (const column of wrappedColumns) column.format.wrapText = true;
      await context.sync();
    }
  });
}

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel || !Office.context.requirements.isSetSupported("ExcelApi", "1.12")) return;
  Excel.run(async context => {
    context.workbook.worksheets.onCalculated.add(({ worksheetId }) =>
      formatCalculatedLists(worksheetId).catch(logListFormatError));
    await context.sync();
  }).catch(error => console.warn("Could not register Sheetdot LIST formatting", error.code ?? error.name));
});
