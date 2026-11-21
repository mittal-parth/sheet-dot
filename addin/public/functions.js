const answerAnnotationPrefix = "Sheetdot answer\n";
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

function isLegacyAnswerNote(text) {
  const normalized = text.replace(/\r\n?/g, "\n").trimEnd();
  if (!normalized.startsWith(answerAnnotationPrefix)) return false;
  const [timestamp, ...sources] = normalized.slice(answerAnnotationPrefix.length).split("\n\n");
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(timestamp) &&
    sources.every(source => /\nhttps?:\/\/\S+$/.test(source));
}

async function writeAnswerComment(address, answer, isCanceled) {
  if (!address || !Office.context.requirements.isSetSupported("ExcelApi", "1.10")) return;

  const sources = answer.sources.map(({ title, url }) => `${title || url}\n${url}`);
  const content = answerAnnotationPrefix + [answer.answeredAt, ...sources].join("\n\n");
  const previous = await Excel.run(async context => {
    const comment = context.workbook.comments.getItemByCell(address);
    comment.load(["id", "content"]);
    await context.sync();
    return { id: comment.id, content: comment.content };
  }).catch(error => {
    // Excel for the web reports a missing comment as InvalidOperation.
    if (error.code !== "ItemNotFound" && error.code !== "InvalidOperation") throw error;
  });

  if (isCanceled() || (previous && !previous.content.startsWith(answerAnnotationPrefix))) return;
  if (previous?.content === content) return;

  let migratedNote;
  try {
    await Excel.run(async context => {
      const { comments, notes } = context.workbook;
      if (previous) {
        const comment = comments.getItem(previous.id);
        comment.load("content");
        await context.sync();
        if (isCanceled() || comment.content !== previous.content) return;
        comment.content = content;
      } else {
        if (Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
          const legacy = notes.getItemOrNullObject(address);
          legacy.load("content");
          await context.sync();
          if (isCanceled()) return;
          if (!legacy.isNullObject) {
            const original = legacy.content;
            if (!isLegacyAnswerNote(original)) return;
            // Excel allows only one annotation per cell.
            legacy.delete();
            await context.sync();
            migratedNote = original;
          }
        }
        comments.add(address, migratedNote ?? content);
      }
      await context.sync();
    });
  } catch (error) {
    if (migratedNote) {
      await Excel.run(async context => {
        context.workbook.notes.add(address, migratedNote);
        await context.sync();
      });
    }
    throw error;
  }
}

CustomFunctions.associate("ASK", async (question, row, invocation) => {
  let canceled = false;
  invocation.onCanceled = () => { canceled = true; };
  const working = setTimeout(() => {
    if (!canceled) invocation.setResult("Working…");
  }, 300);

  try {
    const workbook = await getWorkbookId();
    const preferences = await readWorkbookPreferences();
    if (canceled) return;
    const result = await requestSheetdot("/ask", { workbook, question, row: row ?? [], preferences });
    clearTimeout(working);
    if (canceled) return;
    invocation.setResult(result.answer.value);
    await writeAnswerComment(invocation.address, result.answer, () => canceled).catch(error => {
      console.warn("Could not add a Sheetdot source comment", error.code ?? error.name, error.debugInfo?.errorLocation);
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
  let lastWatch;
  let preferencesSynced = false;
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
      if (!preferencesSynced) {
        await syncWorkbookPreferences();
        preferencesSynced = true;
      }
      if (canceled) return;
      const watch = await requestSheetdot("/watch", {
        workbook, question, row, schedule, timezone, address: invocation.address,
      }, "POST", controller.signal);
      if (canceled) return;
      const revision = JSON.stringify([watch.id, watch.status, watch.lastRunAt]);
      if (revision !== lastWatch) {
        lastWatch = revision;
        globalThis.refreshWatchSheet?.();
      }
      if (watch.status === "failed") throw new Error(watch.error);
      const value = watch.value ?? (watch.status === "stopped" ? "Stopped" : "Scheduled");
      if (value !== displayed) {
        invocation.setResult(value);
        displayed = value;
      }
      timer = setTimeout(poll, 10_000);
    } catch (error) {
      if (canceled) return;
      if (lastWatch && (error.retryable || error instanceof TypeError)) {
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
