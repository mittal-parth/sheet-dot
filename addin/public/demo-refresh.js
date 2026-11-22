const demoFeedMarker = "Sheetdot public feed";
const liveDemoName = "Live Demo India";
const liveDemoMarker = "Sheetdot live demo India";
const previousDemoSheets = new Map([
  ["Live Demo", ["AA1", [["Sheetdot live demo"]]]],
  ["Start Here", ["B2:B3", [["Sheetdot demos"], ["Seven Excel examples using real public data"]]]],
  ...[
    ["Flat Hunt", "housing"], ["Bengaluru Weather", "weather"], ["Bengaluru Air Quality", "air"],
    ["London Tube", "tube"], ["Earthquake Monitor", "earthquakes"], ["ECB Exchange Rates", "fx"],
    ["Node Release Watch", "releases"], ["Weekend Plans", "weather"], ["Find a Flat", "housing"],
  ].map(([name, id]) => [name, ["B10:C10", [[demoFeedMarker, id]]]]),
  ["Demo Guide", ["B10:C10", [["Sheetdot everyday demo", "guide"]]]],
  ["Family Bills", ["B10:C10", [["Sheetdot everyday demo", "bills"]]]],
]);

async function addLiveDemo() {
  const response = await fetch("/assets/live-demos.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("The demo workbook could not be loaded.");
  const bytes = new Uint8Array(await response.arrayBuffer());
  const binary = Array.from(bytes, byte => String.fromCharCode(byte)).join("");
  return Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const names = new Set(sheets.items.map(sheet => sheet.name.toLowerCase()));
    const exists = names.has(liveDemoName.toLowerCase());
    const existingMarker = exists ? sheets.getItem(liveDemoName).getRange("AA1") : undefined;
    existingMarker?.load("values");
    const previous = [];
    for (const [name, [address, marker]] of previousDemoSheets) {
      if (!names.has(name.toLowerCase())) continue;
      const sheet = sheets.getItem(name);
      const range = sheet.getRange(address);
      range.load("values");
      previous.push({ sheet, range, marker });
    }
    await context.sync();
    if (existingMarker && !sameDemoMarker(existingMarker, [[liveDemoMarker]])) {
      throw new Error(`A worksheet named ${liveDemoName} already exists. Rename it before adding the demo.`);
    }
    if (!exists) {
      context.workbook.insertWorksheetsFromBase64(btoa(binary), { sheetNamesToInsert: [liveDemoName] });
      await context.sync();
    }
    const demo = sheets.getItem(liveDemoName);
    demo.visibility = Excel.SheetVisibility.visible;
    demo.activate();
    for (const { sheet, range, marker } of previous) {
      if (sameDemoMarker(range, marker)) sheet.visibility = Excel.SheetVisibility.hidden;
    }
    await context.sync();
    return exists ? `Opened ${liveDemoName}.` : `Added ${liveDemoName}.`;
  });
}

function sameDemoMarker(range, marker) {
  return marker.every((row, i) => row.every((value, j) => range.values[i][j] === value));
}

async function refreshDemoFeed() {
  const { sheetId, id } = await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("id");
    const marker = sheet.getRange("B10:C10");
    marker.load("values");
    await context.sync();
    const [[label, id]] = marker.values;
    if (label !== demoFeedMarker || typeof id !== "string" || !id) {
      throw new Error("Open a public-data sheet in the Sheetdot demo workbook.");
    }
    return { sheetId: sheet.id, id };
  });

  const response = await fetch("/demo-feed?" + new URLSearchParams({ id }), { cache: "no-store" });
  const feed = await response.json();
  if (!response.ok) throw new Error(feed.error);

  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetId);
    const marker = sheet.getRange("B10:C10");
    const header = sheet.getRangeByIndexes(11, 1, 1, feed.headers.length);
    marker.load("values");
    header.load("values");
    await context.sync();
    if (marker.values[0][0] !== demoFeedMarker || marker.values[0][1] !== id ||
      !feed.headers.every((label, column) => header.values[0][column] === label)) {
      throw new Error("The demo sheet layout changed. Open an original demo sheet and try again.");
    }
    const capacity = id === "weather" ? 7 : id === "housing" ? 12 : 20;
    if (feed.rows.length > capacity) throw new Error("The source returned more rows than fit in this demo.");
    const data = sheet.getRangeByIndexes(12, 1, capacity, feed.headers.length);
    data.values = Array.from({ length: capacity }, (_, row) => feed.headers.map((_, column) => {
      const value = feed.rows[row]?.[column] ?? "";
      return typeof value === "string" ? literalCell(value) : value;
    }));
    if (id === "tube") data.format.autofitRows();
    sheet.getRange("B4").values = [[`Fetched ${formatDemoFeedTime(feed.retrieved)} · ${feed.rows.length} rows`]];
    await context.sync();
  });
  return feed.rows.length;
}

function formatDemoFeedTime(timestamp) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Kolkata", dateStyle: "medium", timeStyle: "short",
  }).format(new Date(timestamp));
  return parts + " IST";
}

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) return;
  const status = document.getElementById("demo-feed-status");
  const actions = [{
    button: document.getElementById("refresh-demo-feed"),
    pending: "Fetching public data…",
    run: async () => `Refreshed ${await refreshDemoFeed()} rows.`,
  }];
  if (Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
    actions.push({ button: document.getElementById("add-demo-sheets"), pending: "Opening the live demo…", run: addLiveDemo });
  }
  for (const { button, pending, run } of actions) {
    button.disabled = false;
    button.onclick = async () => {
      for (const action of actions) action.button.disabled = true;
      button.setAttribute("aria-busy", "true");
      status.dataset.tone = "";
      status.textContent = pending;
      try {
        status.textContent = await run();
      } catch (error) {
        status.dataset.tone = "error";
        status.textContent = error.message;
      } finally {
        button.removeAttribute("aria-busy");
        for (const action of actions) action.button.disabled = false;
      }
    };
  };
});
