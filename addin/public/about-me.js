const aboutSheetName = "About me";
const aboutPreferencesRange = "A2:B101";
let preferenceRead;
let syncedPreferences;
let aboutQueue = Promise.resolve();
let aboutSheetId;
let aboutEventsRegistered = false;
let aboutTimer;

function readWorkbookPreferences() {
  preferenceRead ??= Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(aboutSheetName);
    await context.sync();
    if (sheet.isNullObject) return [];
    const range = sheet.getRange(aboutPreferencesRange);
    range.load("text");
    await context.sync();
    return range.text
      .map(([name, value]) => [name.trim(), value.trim()])
      .filter(([name, value]) => name && value);
  }).finally(() => { preferenceRead = undefined; });
  return preferenceRead;
}

function syncWorkbookPreferences() {
  const sync = aboutQueue.catch(() => undefined).then(async () => {
    const workbook = await getWorkbookId();
    const preferences = await readWorkbookPreferences();
    const signature = JSON.stringify(preferences);
    if (signature === syncedPreferences) return false;
    await requestSheetdot("/about-me", { workbook, preferences });
    syncedPreferences = signature;
    return true;
  });
  aboutQueue = sync;
  return sync;
}

async function prepareAboutMeSheet(open = false) {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItemOrNullObject(aboutSheetName);
    await context.sync();
    if (sheet.isNullObject && open) {
      sheet = sheets.add(aboutSheetName);
      const header = sheet.getRange("A1:B1");
      header.values = [["Preference", "Value"]];
      header.format.fill.color = "#107c41";
      header.format.font.color = "#ffffff";
      header.format.font.bold = true;
      sheet.getRange("A:A").format.columnWidth = 160;
      sheet.getRange("B:B").format.columnWidth = 300;
      await context.sync();
    }
    if (!aboutEventsRegistered) {
      sheets.onAdded.add(() => queueAboutMeSync(true));
      sheets.onDeleted.add(() => queueAboutMeSync(true));
      sheets.onNameChanged.add(() => queueAboutMeSync(true));
      await context.sync();
      aboutEventsRegistered = true;
    }
    if (sheet.isNullObject) {
      aboutSheetId = undefined;
      return;
    }
    if (open) sheet.activate();
    sheet.load("id");
    await context.sync();
    if (aboutSheetId !== sheet.id) {
      sheet.onChanged.add(() => queueAboutMeSync());
      await context.sync();
      aboutSheetId = sheet.id;
    }
  });
}

function queueAboutMeSync(checkSheet = false, delay = 250) {
  clearTimeout(aboutTimer);
  aboutTimer = setTimeout(async () => {
    try {
      if (checkSheet) await prepareAboutMeSheet();
      if (await syncWorkbookPreferences()) {
        await Excel.run(async context => {
          context.application.calculate(Excel.CalculationType.full);
          await context.sync();
        });
      }
      aboutMeStatus("");
    } catch (error) {
      aboutMeStatus(error.message);
      if (error.retryable || error instanceof TypeError) queueAboutMeSync(checkSheet, 5000);
    }
  }, delay);
}

function aboutMeStatus(message) {
  document.querySelector("#about-me-status").textContent = message;
}

Office.onReady(async () => {
  const open = document.querySelector("#open-about-me");
  open.disabled = false;
  open.onclick = async () => {
    open.disabled = true;
    try {
      await prepareAboutMeSheet(true);
      await syncWorkbookPreferences();
      aboutMeStatus("");
    } catch (error) {
      aboutMeStatus(error.message);
      if (error.retryable || error instanceof TypeError) queueAboutMeSync(true, 5000);
    } finally {
      open.disabled = false;
    }
  };
  try {
    await prepareAboutMeSheet();
    await syncWorkbookPreferences();
  } catch (error) {
    aboutMeStatus(error.message);
    if (error.retryable || error instanceof TypeError) queueAboutMeSync(true, 5000);
  }
});
