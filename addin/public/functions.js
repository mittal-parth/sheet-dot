const answerNotePrefix = "Sheetdot answer\n";
let workbookId;

async function identifyWorkbook() {
  const document = Office.context.document;
  let identity = document.url;
  if (!identity) {
    const settings = document.settings;
    const setting = "sheetdot.workbookId";
    identity = settings.get(setting);
    if (!identity) {
      identity = crypto.randomUUID();
      settings.set(setting, identity);
      await new Promise((resolve, reject) => settings.saveAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error("Could not save the Sheetdot workbook ID."));
      }));
    }
  }
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(identity));
  return Array.from(new Uint8Array(hash), byte => byte.toString(16).padStart(2, "0")).join("");
}

async function getWorkbookId() {
  await Office.onReady();
  workbookId ??= identifyWorkbook().catch(error => {
    workbookId = undefined;
    throw error;
  });
  return workbookId;
}

async function requestSheetdot(path, input, method = "POST", signal) {
  const response = await fetch(path, {
    method,
    signal,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });
  const result = await response.json();
  if (!response.ok) throw Object.assign(new Error(result.error), { retryable: response.status >= 500 });
  return result;
}

async function writeAnswerNote(address, answer, isCanceled) {
  if (!address || !Office.context.requirements.isSetSupported("ExcelApi", "1.18")) return;

  const sources = answer.sources.map(({ title, url }) => `${title || url}\n${url}`);
  const content = answerNotePrefix + [answer.answeredAt, ...sources].join("\n\n");
  await Excel.run(async context => {
    const notes = context.workbook.notes;
    const note = notes.getItemOrNullObject(address);
    note.load("content");
    await context.sync();
    if (isCanceled()) return;

    if (note.isNullObject) notes.add(address, content);
    else if (note.content.startsWith(answerNotePrefix) && note.content !== content) note.content = content;
    else return;
    await context.sync();
  });
}

CustomFunctions.associate("ASK", async (question, row, invocation) => {
  let canceled = false;
  invocation.onCanceled = () => { canceled = true; };
  const working = setTimeout(() => {
    if (!canceled) invocation.setResult("Working…");
  }, 300);

  try {
    const workbook = await getWorkbookId();
    if (canceled) return;
    const result = await requestSheetdot("/ask", { workbook, question, row: row ?? [] });
    clearTimeout(working);
    if (canceled) return;
    invocation.setResult(result.answer.value);
    await writeAnswerNote(invocation.address, result.answer, () => canceled).catch(error => {
      console.warn("Could not add a Sheetdot source note", error.code ?? error.name);
    });
  } catch (error) {
    if (!canceled) invocation.setResult(new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable, error.message
    ));
  } finally {
    clearTimeout(working);
  }
});

CustomFunctions.associate("WATCH", (question, row, schedule, invocation) => {
  let canceled = false;
  let timer;
  let displayed;
  let recurring = false;
  const controller = new AbortController();
  const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  invocation.onCanceled = () => {
    canceled = true;
    clearTimeout(timer);
    controller.abort();
  };

  async function poll() {
    try {
      const workbook = await getWorkbookId();
      if (canceled) return;
      const watch = await requestSheetdot("/watch", {
        workbook, question, row, schedule, timezone, address: invocation.address,
      }, "POST", controller.signal);
      if (canceled) return;
      if (watch.status === "failed") throw new Error(watch.error);
      const value = watch.value ?? (watch.status === "stopped" ? "Stopped" : "Scheduled");
      if (value !== displayed) {
        invocation.setResult(value);
        displayed = value;
      }
      recurring = watch.recurring && watch.status !== "stopped";
      if (watch.status !== "scheduled" && !recurring) return;
      timer = setTimeout(poll, 10_000);
    } catch (error) {
      if (canceled) return;
      if (recurring && (error.retryable || error instanceof TypeError)) {
        timer = setTimeout(poll, 10_000);
        return;
      }
      invocation.setResult(new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable, error.message
      ));
    }
  }

  void poll();
});

CustomFunctions.associate("PROBE", () => "Sheetdot ready");
CustomFunctions.associate("PROBEARRAY", () => [[1, 2], [3, 4]]);
CustomFunctions.associate("PROBESTREAM", invocation => {
  invocation.setResult("Working");
  const timer = setTimeout(() => invocation.setResult("Ready"), 1500);
  invocation.onCanceled = () => clearTimeout(timer);
});
